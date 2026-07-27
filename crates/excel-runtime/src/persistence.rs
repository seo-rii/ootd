use std::fs;
#[cfg(unix)]
use std::fs::File;
use std::io::{self, Write};
#[cfg(unix)]
use std::os::unix::fs::PermissionsExt;
use std::path::Path;
use tempfile::Builder;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum AtomicWriteMode {
    Replace,
    CreateNew,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum PersistenceFailurePoint {
    CreateTemporary,
    WriteTemporary,
    FlushTemporary,
    SyncTemporary,
    ReplaceTarget,
    SyncParentDirectory,
}

pub(crate) fn durable_atomic_write(
    target: &Path,
    bytes: &[u8],
    mode: AtomicWriteMode,
    failure_point: Option<PersistenceFailurePoint>,
) -> io::Result<()> {
    let parent = target
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));

    fail_if_requested(failure_point, PersistenceFailurePoint::CreateTemporary)?;
    let target_permissions = if mode == AtomicWriteMode::Replace {
        match fs::metadata(target) {
            Ok(metadata) if metadata.is_file() => Some(metadata.permissions()),
            Ok(_) => None,
            Err(error) if error.kind() == io::ErrorKind::NotFound => None,
            Err(error) => return Err(with_stage(error, "inspect target permissions")),
        }
    } else {
        None
    };
    let mut builder = Builder::new();
    if let Some(permissions) = target_permissions {
        builder.permissions(permissions);
    } else {
        #[cfg(unix)]
        builder.permissions(fs::Permissions::from_mode(0o666));
    }
    let mut temporary = builder
        .tempfile_in(parent)
        .map_err(|error| with_stage(error, "create same-directory temporary file"))?;

    fail_if_requested(failure_point, PersistenceFailurePoint::WriteTemporary)?;
    temporary
        .write_all(bytes)
        .map_err(|error| with_stage(error, "write temporary file"))?;

    fail_if_requested(failure_point, PersistenceFailurePoint::FlushTemporary)?;
    temporary
        .flush()
        .map_err(|error| with_stage(error, "flush temporary file"))?;

    fail_if_requested(failure_point, PersistenceFailurePoint::SyncTemporary)?;
    temporary
        .as_file()
        .sync_all()
        .map_err(|error| with_stage(error, "sync temporary file"))?;

    fail_if_requested(failure_point, PersistenceFailurePoint::ReplaceTarget)?;
    let persisted_file = match mode {
        AtomicWriteMode::Replace => temporary.persist(target),
        AtomicWriteMode::CreateNew => temporary.persist_noclobber(target),
    }
    .map_err(|error| with_stage(error.error, "atomically persist temporary file"))?;

    fail_if_requested(failure_point, PersistenceFailurePoint::SyncParentDirectory)?;
    sync_parent_directory(parent)?;
    persisted_file
        .sync_all()
        .map_err(|error| with_stage(error, "sync persisted file"))?;
    Ok(())
}

fn fail_if_requested(
    configured: Option<PersistenceFailurePoint>,
    current: PersistenceFailurePoint,
) -> io::Result<()> {
    if configured == Some(current) {
        Err(io::Error::other(format!(
            "injected persistence failure at {current:?}"
        )))
    } else {
        Ok(())
    }
}

fn with_stage(error: io::Error, stage: &str) -> io::Error {
    io::Error::new(error.kind(), format!("{stage}: {error}"))
}

#[cfg(unix)]
fn sync_parent_directory(parent: &Path) -> io::Result<()> {
    File::open(parent)
        .and_then(|directory| directory.sync_all())
        .map_err(|error| with_stage(error, "sync parent directory"))
}

#[cfg(not(unix))]
fn sync_parent_directory(_parent: &Path) -> io::Result<()> {
    Ok(())
}
