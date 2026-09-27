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

fn same_office_path(open_path: &Path, requested: &Path) -> bool {
    if open_path == requested {
        return true;
    }
    match (
        std::fs::canonicalize(open_path),
        std::fs::canonicalize(requested),
    ) {
        (Ok(left), Ok(right)) => left == right,
        _ => false,
    }
}

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
        let session_id = UseSessionId::parse(session.clone())?;
        let requested = path.as_ref();
        let existing = self.entries.read().await.get(&session_id).map(Arc::clone);
        if let Some(entry) = existing {
            let state = entry.lock().await;
            state.ensure_open(&session_id)?;
            let same_file = same_office_path(state.editor.package().path(), requested);
            if same_file && state.read_only == read_only {
                drop(state);
                return Ok((session_id, entry));
            }
            return Err(UseError::new(
                "use.office.session_exists",
                format!(
                    "Native Office session '{}' is already open.",
                    session_id.as_str()
                ),
            )
            .with_detail("path", state.editor.package().path().display().to_string()));
        }
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

        a3s_office::live_replica_blocks_office_session(path.as_ref())?;
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

    #[tokio::test]
    async fn open_existing_reuses_the_same_file() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("letter.docx");
        let other = temp.path().join("other.docx");
        let sessions = NativeOfficeSessions::default();
        sessions.create("letter".to_string(), &path).await.unwrap();

        let (id, _) = sessions
            .open_existing("letter".to_string(), &path, false)
            .await
            .unwrap();
        assert_eq!(id.as_str(), "letter");

        let read_only = sessions
            .open_existing("letter".to_string(), &path, true)
            .await
            .unwrap_err();
        assert_eq!(read_only.code, "use.office.session_exists");

        sessions.create("other".to_string(), &other).await.unwrap();
        let different_file = sessions
            .open_existing("letter".to_string(), &other, false)
            .await
            .unwrap_err();
        assert_eq!(different_file.code, "use.office.session_exists");

        let duplicate_create = sessions
            .create("letter".to_string(), &path)
            .await
            .unwrap_err();
        assert_eq!(duplicate_create.code, "use.office.session_exists");
    }

    #[tokio::test]
    async fn office_open_refuses_a_package_whose_replica_is_live() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("letter.docx");
        std::fs::write(&path, b"snapshot").unwrap();
        let document = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body><w:p w14:paraId=\"00000001\" w14:textId=\"00000002\"><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>";
        let mut parts = std::collections::BTreeMap::new();
        parts.insert("word/document.xml".to_owned(), document.as_bytes().to_vec());
        a3s_office::import_document_snapshot(
            &path,
            &parts,
            &temp.path().join("replica"),
            "letter-artifact",
        )
        .unwrap();

        let sessions = NativeOfficeSessions::default();
        let opened = sessions
            .open_existing("letter".to_string(), &path, false)
            .await
            .unwrap_err();
        assert_eq!(opened.code, "office.collaboration.replica_live");
        let created = sessions
            .create("fresh".to_string(), &path)
            .await
            .unwrap_err();
        assert_eq!(created.code, "office.collaboration.replica_live");
    }
}
