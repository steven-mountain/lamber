//! Human-in-the-loop approval: suspend a dsh tool call until the user decides.
//!
//! The dsh answerer plugin posts a question to `POST /lamber-bridge/approval`
//! and holds that HTTP request open. This module parks the bridge worker thread
//! on a condition variable, emits the question to the frontend, and wakes the
//! thread when `ai_resolve_approval` arrives with the user's answer.
//!
//! Every ambiguous outcome is a denial. A timeout resolves `rejected` with an
//! explicit reason rather than leaving the request hanging or falling through to
//! a grant — an agent that silently gets its way when nobody answered would be
//! worse than one that is told "no".
//!
//! Every settled question is handed to a recorder so it can be persisted; see
//! `approval_log.rs`. The gate itself stays free of database coupling, which
//! keeps it testable and keeps a storage failure from blocking a decision.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant};

/// Route serving approval questions.
pub const APPROVAL_ROUTE: &str = "/lamber-bridge/approval";

/// Frontend event carrying one pending approval question.
pub const APPROVAL_EVENT: &str = "ai://approval-request";

/// Default wait for a human before failing closed.
///
/// Shorter than the answerer's own bound (180s) so the normal timeout path is
/// this explicit `rejected` reply rather than the plugin's abort.
pub const DEFAULT_APPROVAL_TIMEOUT: Duration = Duration::from_secs(90);

/// Overrides the wait, in seconds. Tests use it to exercise the timeout path
/// without stalling for a minute and a half.
pub const APPROVAL_TIMEOUT_ENV: &str = "LAMBER_APPROVAL_TIMEOUT_SECS";

/// The effective wait: the override when it parses to a positive value, else the default.
pub fn approval_timeout() -> Duration {
    std::env::var(APPROVAL_TIMEOUT_ENV)
        .ok()
        .and_then(|raw| raw.trim().parse::<u64>().ok())
        .filter(|secs| *secs > 0)
        .map(Duration::from_secs)
        .unwrap_or(DEFAULT_APPROVAL_TIMEOUT)
}

/// Question posted by the dsh answerer.
#[derive(Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct ApprovalRequest {
    pub tool_name: String,
    #[serde(default)]
    pub call_id: Option<String>,
    #[serde(default)]
    pub reason: Option<String>,
    /// Parsed tool arguments, forwarded so the dialog can show what would run.
    #[serde(default)]
    pub args: serde_json::Value,
}

/// Decision returned to the answerer. `approved` is true only for an explicit grant.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ApprovalDecision {
    pub approved: bool,
    pub reason: String,
}

impl ApprovalDecision {
    fn denied(reason: impl Into<String>) -> Self {
        Self {
            approved: false,
            reason: reason.into(),
        }
    }
}

/// What the frontend is shown for one pending question.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ApprovalPrompt {
    /// lamber-issued id the frontend echoes back through `ai_resolve_approval`.
    pub request_id: String,
    pub tool_name: String,
    pub call_id: Option<String>,
    pub reason: Option<String>,
    pub args: serde_json::Value,
    /// Seconds before lamber auto-rejects, so the dialog can show a countdown.
    pub timeout_seconds: u64,
}

/// Who or what settled a question.
#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum DecidedBy {
    /// A human answered through the approval dialog.
    User,
    /// Nobody answered within the gate's wait.
    Timeout,
    /// The runtime was torn down while the question was still open.
    Shutdown,
    /// The question could not be tracked (poisoned lock, lost slot).
    Internal,
}

impl DecidedBy {
    /// Stable string written to the audit log.
    pub fn as_str(self) -> &'static str {
        match self {
            DecidedBy::User => "user",
            DecidedBy::Timeout => "timeout",
            DecidedBy::Shutdown => "shutdown",
            DecidedBy::Internal => "internal",
        }
    }
}

/// One settled question, handed to the recorder for persistence.
///
/// `Deserialize` exists so a decision buffered while no workspace was open can
/// be read back and backfilled; see `approval_log`.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ApprovalRecord {
    pub request_id: String,
    pub tool_name: String,
    pub call_id: Option<String>,
    pub reason: Option<String>,
    /// Tool arguments as JSON text, exactly as shown to the user.
    pub args_json: String,
    pub approved: bool,
    pub decided_by: DecidedBy,
    pub decision_reason: String,
    pub requested_at: String,
    pub decided_at: String,
}

/// Persists settled questions. Called on the bridge worker thread.
pub type ApprovalRecorder = Arc<dyn Fn(&ApprovalRecord) + Send + Sync>;

/// What the gate remembers about a question while it is open.
struct Slot {
    decision: Option<ApprovalDecision>,
    decided_by: DecidedBy,
    prompt: ApprovalPrompt,
    requested_at: String,
}

/// Registry of questions currently awaiting a human.
pub struct ApprovalGate {
    slots: Mutex<GateState>,
    signal: Condvar,
    /// How long each question waits. Held per gate rather than read from the
    /// environment at wait time, so tests can shorten it without mutating
    /// process-global state that parallel tests would race on.
    timeout: Duration,
    recorder: Mutex<Option<ApprovalRecorder>>,
}

/// Mutex-protected gate state.
#[derive(Default)]
struct GateState {
    slots: HashMap<String, Slot>,
    /// Set by `shutdown`; every open question is denied and no new one is accepted.
    closed: bool,
}

impl Default for ApprovalGate {
    fn default() -> Self {
        Self::new(approval_timeout())
    }
}

impl ApprovalGate {
    /// Build a gate with an explicit wait.
    pub fn new(timeout: Duration) -> Self {
        Self {
            slots: Mutex::new(GateState::default()),
            signal: Condvar::new(),
            timeout,
            recorder: Mutex::new(None),
        }
    }

    /// The wait applied to each question on this gate.
    pub fn timeout(&self) -> Duration {
        self.timeout
    }

    /// Install the sink that persists settled questions.
    ///
    /// Set once at launch. A recorder that panics or fails must not affect the
    /// decision, so its errors are the recorder's own responsibility.
    pub fn set_recorder(&self, recorder: ApprovalRecorder) {
        if let Ok(mut slot) = self.recorder.lock() {
            *slot = Some(recorder);
        }
    }

    /// Park the calling thread until the user answers, or the timeout elapses.
    ///
    /// @param prompt - the question, retained so a settled record can describe it.
    /// @returns the decision; a timeout, a shutdown, or a lost slot yields a denial.
    pub fn wait(&self, prompt: &ApprovalPrompt) -> ApprovalDecision {
        let request_id = prompt.request_id.clone();
        let requested_at = chrono::Utc::now().to_rfc3339();
        let deadline = Instant::now() + self.timeout;
        let Ok(mut state) = self.slots.lock() else {
            return self.settle(
                prompt,
                &requested_at,
                ApprovalDecision::denied("审批状态锁已中毒，按拒绝处理"),
                DecidedBy::Internal,
            );
        };

        if state.closed {
            drop(state);
            return self.settle(
                prompt,
                &requested_at,
                ApprovalDecision::denied("AI 运行时正在关闭，按拒绝处理"),
                DecidedBy::Shutdown,
            );
        }

        state.slots.insert(
            request_id.clone(),
            Slot {
                decision: None,
                decided_by: DecidedBy::Internal,
                prompt: prompt.clone(),
                requested_at: requested_at.clone(),
            },
        );

        loop {
            match state.slots.get_mut(&request_id) {
                Some(slot) => {
                    if let Some(decision) = slot.decision.take() {
                        let decided_by = slot.decided_by;
                        state.slots.remove(&request_id);
                        drop(state);
                        return self.settle(prompt, &requested_at, decision, decided_by);
                    }
                }
                // Only `resolve`, `shutdown` and this function touch the map, and
                // neither of the others removes a slot it did not fill.
                None => {
                    drop(state);
                    return self.settle(
                        prompt,
                        &requested_at,
                        ApprovalDecision::denied("审批请求已失效，按拒绝处理"),
                        DecidedBy::Internal,
                    );
                }
            }

            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                state.slots.remove(&request_id);
                drop(state);
                return self.settle(
                    prompt,
                    &requested_at,
                    ApprovalDecision::denied(format!(
                        "等待用户确认超时（{} 秒），按拒绝处理",
                        self.timeout.as_secs()
                    )),
                    DecidedBy::Timeout,
                );
            }
            match self.signal.wait_timeout(state, remaining) {
                Ok((next, _)) => state = next,
                Err(_) => {
                    return self.settle(
                        prompt,
                        &requested_at,
                        ApprovalDecision::denied("审批状态锁已中毒，按拒绝处理"),
                        DecidedBy::Internal,
                    )
                }
            }
        }
    }

    /// Hand one settled question to the recorder and return the decision unchanged.
    fn settle(
        &self,
        prompt: &ApprovalPrompt,
        requested_at: &str,
        decision: ApprovalDecision,
        decided_by: DecidedBy,
    ) -> ApprovalDecision {
        let recorder = self.recorder.lock().ok().and_then(|r| r.clone());
        if let Some(recorder) = recorder {
            recorder(&ApprovalRecord {
                request_id: prompt.request_id.clone(),
                tool_name: prompt.tool_name.clone(),
                call_id: prompt.call_id.clone(),
                reason: prompt.reason.clone(),
                args_json: prompt.args.to_string(),
                approved: decision.approved,
                decided_by,
                decision_reason: decision.reason.clone(),
                requested_at: requested_at.to_string(),
                decided_at: chrono::Utc::now().to_rfc3339(),
            });
        }
        decision
    }

    /// Deliver the user's answer and wake the parked bridge thread.
    ///
    /// @param request_id - the id from the emitted prompt.
    /// @param approved - whether the user granted this one call.
    /// @returns an error when the question already timed out or was answered.
    pub fn resolve(&self, request_id: &str, approved: bool) -> Result<(), String> {
        let mut state = self
            .slots
            .lock()
            .map_err(|_| "审批状态锁已中毒".to_string())?;
        let Some(slot) = state.slots.get_mut(request_id) else {
            return Err("该审批请求不存在或已超时".to_string());
        };
        if slot.decision.is_some() {
            return Err("该审批请求已被响应".to_string());
        }
        slot.decision = Some(ApprovalDecision {
            approved,
            reason: if approved {
                "用户已确认".to_string()
            } else {
                "用户已拒绝".to_string()
            },
        });
        slot.decided_by = DecidedBy::User;
        self.signal.notify_all();
        Ok(())
    }

    /// Deny every open question and refuse new ones.
    ///
    /// Called when the AI runtime is torn down, so a parked bridge thread is
    /// released instead of holding its worker (and the answerer's HTTP request)
    /// until the full timeout. Idempotent.
    ///
    /// @returns how many open questions were denied.
    pub fn shutdown(&self) -> usize {
        let Ok(mut state) = self.slots.lock() else {
            return 0;
        };
        state.closed = true;
        let mut denied = 0;
        for slot in state.slots.values_mut() {
            if slot.decision.is_none() {
                slot.decision = Some(ApprovalDecision::denied(
                    "AI 运行时已关闭，未完成的审批按拒绝处理",
                ));
                slot.decided_by = DecidedBy::Shutdown;
                denied += 1;
            }
        }
        self.signal.notify_all();
        denied
    }

    /// Reopen a gate that was shut down, so the next launch can use it again.
    pub fn reopen(&self) {
        if let Ok(mut state) = self.slots.lock() {
            state.closed = false;
        }
    }

    /// Number of questions currently parked. Used by tests and diagnostics.
    pub fn pending_count(&self) -> usize {
        self.slots.lock().map(|s| s.slots.len()).unwrap_or(0)
    }
}

/// Handle one approval question: announce it, then wait for the answer.
///
/// `announce` is separate from the Tauri layer so tests can drive the gate with
/// a plain closure instead of a running app.
///
/// @param gate - the shared registry of pending questions.
/// @param request - the question posted by the answerer.
/// @param announce - called with the prompt before parking; must not block.
/// @returns the decision to serialize back to the answerer.
pub fn handle_request(
    gate: &Arc<ApprovalGate>,
    request: ApprovalRequest,
    announce: impl FnOnce(&ApprovalPrompt),
) -> ApprovalDecision {
    let request_id = uuid::Uuid::new_v4().to_string();
    let prompt = ApprovalPrompt {
        request_id: request_id.clone(),
        tool_name: request.tool_name,
        call_id: request.call_id,
        reason: request.reason,
        args: request.args,
        timeout_seconds: gate.timeout().as_secs(),
    };
    announce(&prompt);
    gate.wait(&prompt)
}
