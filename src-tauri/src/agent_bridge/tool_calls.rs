//! Correlates an ACP permission request back to the tool call it is about.
//!
//! ACP's `session/requestPermission` is deliberately thin: its `toolCall` field
//! is a `ToolCallUpdate`, and dsh fills in nothing but `toolCallId`
//! (`dsh-acp/lib/index.js:1123-1126`). The tool's name and arguments arrive
//! earlier, on the `session/update` notification that announced the call —
//! dsh awaits `drainUpdates()` before asking, so that notification is always
//! already delivered by the time the question shows up.
//!
//! lamber's approval dialog has to show the user what would actually run, so
//! this index remembers each announced call and hands it back by id. It is the
//! Rust-side successor to the plugin's `pendingCalls.ts`: the same correlation,
//! moved to the side of the seam that now asks the question.
//!
//! Bounded and self-cleaning, like its predecessor: the permission handler
//! consumes the entry it reads, and calls that never raise a question (ungated
//! tools, an abandoned turn) are evicted once the map exceeds its cap.

use std::collections::VecDeque;
use std::sync::Mutex;

/// Announced calls kept for correlation; the oldest is dropped first.
const MAX_TRACKED: usize = 64;

/// One tool call the agent announced, as the dialog will describe it.
#[derive(Debug, Clone, PartialEq)]
pub struct TrackedCall {
    /// The tool's name. dsh puts it in the update's `title`.
    pub tool_name: String,
    /// Arguments the tool would run with, from the update's `rawInput`.
    pub args: serde_json::Value,
}

/// Recent tool calls, keyed by ACP `toolCallId`.
#[derive(Default)]
pub struct ToolCallIndex {
    state: Mutex<VecDeque<(String, TrackedCall)>>,
}

impl ToolCallIndex {
    /// Record a call the agent just announced.
    ///
    /// Re-announcing the same id replaces the entry rather than duplicating it,
    /// so a `tool_call_update` that restates the call cannot leave two rows
    /// behind for one question to pick between.
    ///
    /// @param tool_call_id - the ACP id the permission request will carry.
    /// @param call - the name and arguments to show the user.
    pub fn record(&self, tool_call_id: &str, call: TrackedCall) {
        let Ok(mut state) = self.state.lock() else {
            return;
        };
        state.retain(|(id, _)| id != tool_call_id);
        state.push_back((tool_call_id.to_string(), call));
        while state.len() > MAX_TRACKED {
            state.pop_front();
        }
    }

    /// Consume the call a permission request is about.
    ///
    /// @param tool_call_id - the id from `RequestPermissionRequest::tool_call`.
    /// @returns the announced call, or `None` when no update announced it.
    pub fn take(&self, tool_call_id: &str) -> Option<TrackedCall> {
        let mut state = self.state.lock().ok()?;
        let at = state.iter().position(|(id, _)| id == tool_call_id)?;
        state.remove(at).map(|(_, call)| call)
    }

    /// Number of calls currently tracked. Used by tests and diagnostics.
    pub fn len(&self) -> usize {
        self.state.lock().map(|s| s.len()).unwrap_or(0)
    }
}
