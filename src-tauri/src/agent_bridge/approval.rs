//! Human-in-the-loop approval: suspend a dsh tool call until the user decides.
//!
//! The question arrives as an ACP `session/requestPermission` request on the
//! dsh connection (see `dsh_session.rs`). This module parks a blocking task on
//! a condition variable, emits the question to the frontend, and wakes the task
//! when `ai_resolve_approval` arrives with the user's answer.
//!
//! Until the ACP rewrite the question arrived instead over a loopback HTTP
//! route the plugin posted to. That route is gone: under `--profile acp`,
//! `dsh-acp` answers the `approval/request` Cordis event itself and forwards it
//! to the client, so nothing could ever have reached lamber's own answerer. The
//! gate below is unchanged by that move — only its trigger is new.
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

/// Frontend event carrying one pending approval question.
pub const APPROVAL_EVENT: &str = "ai://approval-request";

/// Default wait for a human before failing closed.
///
/// dsh imposes no deadline of its own on a `session/requestPermission`: it
/// waits for whatever the client answers. So this bound is the only thing that
/// keeps an unattended turn from parking forever, and it must always fire.
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

/// Tools that require a human decision, with the text shown in the dialog.
///
/// This is a *mirror* of the plugin's own `GATED_TOOLS`
/// (`agent-bridge/dsh-tool-lamber/src/approval.ts`), and the duplication is
/// forced by the protocol rather than chosen. The plugin's `tools/pre-execute`
/// guard still decides *which* calls need a human — that half is untouched by
/// the ACP move. But `dsh-acp` builds its `session/requestPermission` params
/// from the tool call id alone and drops the guard's `reason`
/// (`dsh-acp/lib/index.js:1123-1134`), so the explanation the user reads has
/// nowhere to travel and must exist on this side.
///
/// The two tables are kept honest by `gated_tool_names_match_the_plugin`, which
/// reads the plugin source and fails if either side gains or loses a tool.
const GATED_TOOLS: &[(&str, &str)] = &[(
    "write_test_marker",
    "该工具会写入文件（测试标记文件，位于系统临时目录），需要你确认后才执行。",
)];

/// Why a tool needs confirming, or `None` when lamber does not gate it.
///
/// @param tool_name - the tool name announced on the ACP tool call.
/// @returns the dialog text for a gated tool.
pub fn gated_tool_reason(tool_name: &str) -> Option<&'static str> {
    GATED_TOOLS
        .iter()
        .find(|(name, _)| *name == tool_name)
        .map(|(_, reason)| *reason)
}

/// Every tool name lamber expects to be gated. Used by the contract test.
pub fn gated_tool_names() -> Vec<&'static str> {
    GATED_TOOLS.iter().map(|(name, _)| *name).collect()
}

/// One question, as the ACP permission handler assembled it.
///
/// Built from two ACP messages rather than one payload: the id and the session
/// come from `session/requestPermission`, while the tool name and arguments come
/// from the `session/update` that announced the call (see `tool_calls.rs`).
#[derive(Debug, Clone)]
pub struct ApprovalQuestion {
    pub tool_name: String,
    pub call_id: Option<String>,
    pub reason: Option<String>,
    /// Arguments the tool would run with, so the dialog can show what happens.
    pub args: serde_json::Value,
}

/// Decision handed back to the ACP handler. `approved` is true only for an explicit grant.
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

/// Persists settled questions. Called on the blocking approval task.
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
    /// `announce` runs once the question is registered and while its lock is
    /// still held. That ordering is the point: `resolve` needs the same lock, so
    /// an answer cannot arrive before the slot it would fill exists. Announcing
    /// first instead would leave a window — however small — in which a fast
    /// answer is told the request does not exist.
    ///
    /// The cost of that ordering is that `announce` must not call back into this
    /// gate — `resolve` and `shutdown` want the same lock, and the mutex is not
    /// reentrant. Emitting an event is fine; waiting on the answer is not.
    ///
    /// @param prompt - the question, retained so a settled record can describe it.
    /// @param announce - surfaces the question; runs under the lock, must not block or re-enter.
    /// @returns the decision; a timeout, a shutdown, or a lost slot yields a denial.
    pub fn wait(
        &self,
        prompt: &ApprovalPrompt,
        announce: impl FnOnce(&ApprovalPrompt),
    ) -> ApprovalDecision {
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
        announce(prompt);

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
/// Blocking by design — the ACP permission handler runs it on a blocking task
/// so the connection's dispatch loop stays free to deliver `session/update`
/// notifications while the dialog is open.
///
/// `announce` is separate from the Tauri layer so tests can drive the gate with
/// a plain closure instead of a running app.
///
/// @param gate - the shared registry of pending questions.
/// @param question - the question the permission handler assembled.
/// @param announce - called with the prompt once it is registered; must not block.
/// @returns the decision to answer `session/requestPermission` with.
pub fn handle_request(
    gate: &Arc<ApprovalGate>,
    question: ApprovalQuestion,
    announce: impl FnOnce(&ApprovalPrompt),
) -> ApprovalDecision {
    let request_id = uuid::Uuid::new_v4().to_string();
    let prompt = ApprovalPrompt {
        request_id: request_id.clone(),
        tool_name: question.tool_name,
        call_id: question.call_id,
        reason: question.reason,
        args: question.args,
        timeout_seconds: gate.timeout().as_secs(),
    };
    gate.wait(&prompt, announce)
}
