use std::collections::HashMap;
use std::path::Path;
use std::sync::Arc;

use a3s_office::NativeOfficeEditor;
use a3s_use_core::{UseError, UseResult, UseSessionId};
use tokio::sync::{Mutex, RwLock};

const MAX_OPEN_SESSIONS: usize = 64;

#[derive(Debug)]
pub(super) struct NativeOfficeSession {
    pub(super) editor: NativeOfficeEditor,
    pub(super) read_only: bool,
    closed: bool,
}

impl NativeOfficeSession {
    pub(super) fn ensure_open(&self, session: &UseSessionId) -> UseResult<()> {
        if self.closed {
            return Err(UseError::new(
                "use.office.session_closed",
                format!("Native Office session '{}' is closed.", session.as_str()),
            ));
        }
        Ok(())
    }

    pub(super) fn ensure_mutable(&self, session: &UseSessionId) -> UseResult<()> {
        self.ensure_open(session)?;
        if self.read_only {
            return Err(UseError::new(
                "use.office.read_only",
                format!("Native Office session '{}' is read-only.", session.as_str()),
            ));
        }
        Ok(())
    }
}

type SharedSession = Arc<Mutex<NativeOfficeSession>>;

#[derive(Debug, Clone, Default)]
pub(super) struct NativeOfficeSessions {
    entries: Arc<RwLock<HashMap<UseSessionId, SharedSession>>>,
    open_gate: Arc<Mutex<()>>,
}

impl NativeOfficeSessions {
    pub(super) async fn create(
        &self,
        session: String,
        path: impl AsRef<Path>,
    ) -> UseResult<(UseSessionId, SharedSession)> {
        self.open(session, path, false, true).await
    }

    pub(super) async fn open_existing(
        &self,
        session: String,
        path: impl AsRef<Path>,
        read_only: bool,
    ) -> UseResult<(UseSessionId, SharedSession)> {
        self.open(session, path, read_only, false).await
    }

    async fn open(
        &self,
        session: String,
        path: impl AsRef<Path>,
        read_only: bool,
        create: bool,
    ) -> UseResult<(UseSessionId, SharedSession)> {
        let session = UseSessionId::parse(session)?;
        let _gate = self.open_gate.lock().await;
        {
            let entries = self.entries.read().await;
            if entries.contains_key(&session) {
                return Err(UseError::new(
                    "use.office.session_exists",
                    format!(
                        "Native Office session '{}' is already open.",
                        session.as_str()
                    ),
                ));
            }
            if entries.len() >= MAX_OPEN_SESSIONS {
                return Err(UseError::new(
                    "use.office.session_limit",
                    format!(
                        "Native Office MCP supports at most {MAX_OPEN_SESSIONS} open sessions."
                    ),
                )
                .with_suggestion(
                    "Save and close an existing Office session before opening another.",
                ));
            }
        }

        let editor = if create {
            NativeOfficeEditor::create(path).await?
        } else {
            NativeOfficeEditor::open(path).await?
        };
        let entry = Arc::new(Mutex::new(NativeOfficeSession {
            editor,
            read_only,
            closed: false,
        }));
        self.entries
            .write()
            .await
            .insert(session.clone(), Arc::clone(&entry));
        Ok((session, entry))
    }

    pub(super) async fn get(&self, value: &str) -> UseResult<(UseSessionId, SharedSession)> {
        let session = UseSessionId::parse(value.to_string())?;
        let entry = self
            .entries
            .read()
            .await
            .get(&session)
            .cloned()
            .ok_or_else(|| {
                UseError::new(
                    "use.office.session_missing",
                    format!("Native Office session '{}' is not open.", session.as_str()),
                )
            })?;
        Ok((session, entry))
    }

    pub(super) async fn list(&self) -> Vec<(UseSessionId, SharedSession)> {
        self.entries
            .read()
            .await
            .iter()
            .map(|(session, entry)| (session.clone(), Arc::clone(entry)))
            .collect()
    }

    pub(super) async fn close(
        &self,
        value: &str,
        discard: bool,
    ) -> UseResult<(UseSessionId, SharedSession)> {
        let (session, entry) = self.get(value).await?;
        {
            let mut state = entry.lock().await;
            state.ensure_open(&session)?;
            if state.editor.is_dirty() && !discard {
                return Err(UseError::new(
                    "use.office.unsaved_changes",
                    format!(
                        "Native Office session '{}' has unsaved changes.",
                        session.as_str()
                    ),
                )
                .with_suggestion(
                    "Call office_save first, or call office_close with discard=true.",
                ));
            }
            state.closed = true;
        }
        self.entries.write().await.remove(&session);
        Ok((session, entry))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[tokio::test]
    async fn close_requires_explicit_discard_for_dirty_sessions() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("report.docx");
        let sessions = NativeOfficeSessions::default();
        let (_, entry) = sessions.create("report".to_string(), &path).await.unwrap();
        entry
            .lock()
            .await
            .editor
            .add_paragraph("/body", "unsaved")
            .unwrap();

        let error = sessions.close("report", false).await.unwrap_err();
        assert_eq!(error.code, "use.office.unsaved_changes");
        sessions.close("report", true).await.unwrap();
    }
}
