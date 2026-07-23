use std::path::{Path, PathBuf};

use a3s_use_core::{UseError, UseResult};
use tokio::io::AsyncReadExt;

use crate::CommandOutput;

const SKILLS_ENV: &str = "A3S_USE_OFFICE_SKILLS_DIR";
const PRIMARY_SKILL: &str = "a3s-office";
const MAX_SKILLS: usize = 64;
const MAX_SKILL_BYTES: u64 = 256 * 1024;
const MAX_FULL_SKILL_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone)]
struct OfficeSkill {
    name: String,
    description: String,
    directory: PathBuf,
    skill_path: PathBuf,
}

pub(crate) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let args = args
        .iter()
        .filter(|argument| argument.as_str() != "--json")
        .map(String::as_str)
        .collect::<Vec<_>>();
    let Some((command, command_args)) = args.split_first() else {
        return list(&[]).await;
    };
    match *command {
        "list" => list(command_args).await,
        "get" => get(command_args).await,
        "path" => path(command_args).await,
        "help" | "--help" | "-h" => Ok(help()),
        command => Err(skills_usage_error(format!(
            "unknown Office skills command '{command}'"
        ))),
    }
}

pub(crate) async fn primary_skill_surface() -> Option<(PathBuf, PathBuf)> {
    let (package_root, skills_root) = locate_skills_root().await?;
    let skill = discover_skills_at(&skills_root)
        .await
        .ok()?
        .into_iter()
        .find(|skill| skill.name == PRIMARY_SKILL && skill.skill_path.starts_with(&package_root))?;
    Some((package_root, skill.skill_path))
}

async fn list(args: &[&str]) -> UseResult<CommandOutput> {
    if !args.is_empty() {
        return Err(skills_usage_error(
            "office skills list does not accept positional arguments",
        ));
    }
    let skills = discover_skills().await?;
    let human = if skills.is_empty() {
        "No packaged Office Skills found.".to_string()
    } else {
        skills
            .iter()
            .map(|skill| format!("{}  {}", skill.name, skill.description))
            .collect::<Vec<_>>()
            .join("\n")
    };
    let values = skills
        .into_iter()
        .map(|skill| {
            serde_json::json!({
                "name": skill.name,
                "description": skill.description,
                "path": skill.skill_path
            })
        })
        .collect::<Vec<_>>();
    Ok(CommandOutput::success(
        human,
        serde_json::Value::Array(values),
    ))
}

async fn get(args: &[&str]) -> UseResult<CommandOutput> {
    let full = args.contains(&"--full");
    let names = args
        .iter()
        .copied()
        .filter(|argument| *argument != "--full")
        .collect::<Vec<_>>();
    if names.len() != 1 {
        return Err(skills_usage_error(
            "office skills get requires exactly one <name> and optional --full",
        ));
    }
    let skill = find_skill(names[0]).await?;
    let mut content = read_bounded(&skill.skill_path, MAX_SKILL_BYTES).await?;
    if full {
        append_references(&skill, &mut content).await?;
    }
    Ok(CommandOutput::success(
        content.clone(),
        serde_json::json!({
            "name": skill.name,
            "description": skill.description,
            "path": skill.skill_path,
            "full": full,
            "content": content
        }),
    ))
}

async fn path(args: &[&str]) -> UseResult<CommandOutput> {
    if args.len() > 1 {
        return Err(skills_usage_error(
            "office skills path accepts at most one [name]",
        ));
    }
    let path = match args.first() {
        Some(name) => find_skill(name).await?.directory,
        None => locate_skills_root()
            .await
            .map(|(_, skills_root)| skills_root)
            .ok_or_else(skills_missing)?,
    };
    Ok(CommandOutput::success(
        path.display().to_string(),
        serde_json::json!({ "path": path }),
    ))
}

fn help() -> CommandOutput {
    CommandOutput::success(
        concat!(
            "a3s-office skills — packaged Office agent guidance\n\n",
            "usage:\n",
            "  a3s-office skills list [--json]\n",
            "  a3s-office skills get <name> [--full] [--json]\n",
            "  a3s-office skills path [name] [--json]"
        ),
        serde_json::json!({
            "commands": ["list", "get", "path"],
            "primary": PRIMARY_SKILL
        }),
    )
}

async fn discover_skills() -> UseResult<Vec<OfficeSkill>> {
    let (_, root) = locate_skills_root().await.ok_or_else(skills_missing)?;
    discover_skills_at(&root).await
}

async fn discover_skills_at(root: &Path) -> UseResult<Vec<OfficeSkill>> {
    let mut directory = tokio::fs::read_dir(root)
        .await
        .map_err(|error| skills_io_error("list", root, error))?;
    let mut entries = Vec::new();
    while let Some(entry) = directory
        .next_entry()
        .await
        .map_err(|error| skills_io_error("list", root, error))?
    {
        entries.push(entry);
        if entries.len() > MAX_SKILLS {
            return Err(UseError::new(
                "use.office.skills_too_many",
                format!("Office Skill collection contains more than {MAX_SKILLS} entries."),
            ));
        }
    }
    entries.sort_by_key(|entry| entry.file_name());
    let mut skills = Vec::new();
    for entry in entries {
        let file_type = entry
            .file_type()
            .await
            .map_err(|error| skills_io_error("inspect", &entry.path(), error))?;
        if !file_type.is_dir() || file_type.is_symlink() {
            continue;
        }
        let directory = entry.path();
        let Some(directory_name) = directory.file_name().and_then(|name| name.to_str()) else {
            continue;
        };
        if !valid_skill_name(directory_name) {
            continue;
        }
        let skill_path = directory.join("SKILL.md");
        let metadata = match tokio::fs::symlink_metadata(&skill_path).await {
            Ok(metadata) if metadata.is_file() && !metadata.file_type().is_symlink() => metadata,
            Ok(_) => continue,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue,
            Err(error) => return Err(skills_io_error("inspect", &skill_path, error)),
        };
        if metadata.len() > MAX_SKILL_BYTES {
            return Err(UseError::new(
                "use.office.skill_too_large",
                format!(
                    "Office Skill '{}' exceeds the {MAX_SKILL_BYTES}-byte limit.",
                    skill_path.display()
                ),
            ));
        }
        let content = read_bounded(&skill_path, MAX_SKILL_BYTES).await?;
        let (name, description) = parse_frontmatter(&content).ok_or_else(|| {
            UseError::new(
                "use.office.skill_invalid",
                format!(
                    "Office Skill '{}' has invalid name or description frontmatter.",
                    skill_path.display()
                ),
            )
        })?;
        if name != directory_name {
            return Err(UseError::new(
                "use.office.skill_invalid",
                format!(
                    "Office Skill frontmatter name '{name}' does not match directory '{directory_name}'."
                ),
            ));
        }
        skills.push(OfficeSkill {
            name,
            description,
            directory,
            skill_path,
        });
    }
    Ok(skills)
}

async fn find_skill(name: &str) -> UseResult<OfficeSkill> {
    if !valid_skill_name(name) {
        return Err(UseError::new(
            "use.office.skill_name_invalid",
            format!("Office Skill name '{name}' is invalid."),
        ));
    }
    discover_skills()
        .await?
        .into_iter()
        .find(|skill| skill.name == name)
        .ok_or_else(|| {
            UseError::new(
                "use.office.skill_not_found",
                format!("Packaged Office Skill '{name}' was not found."),
            )
            .with_suggestion("Run 'a3s use office skills list'.")
        })
}

async fn locate_skills_root() -> Option<(PathBuf, PathBuf)> {
    if let Some(skills_root) = std::env::var_os(SKILLS_ENV).map(PathBuf::from) {
        if !skills_root.is_absolute() {
            return None;
        }
        let package_root = skills_root.parent()?.to_path_buf();
        return canonical_skills_root(package_root, skills_root).await;
    }
    if let Ok(executable) = std::env::current_exe() {
        if let Some(package_root) = executable.parent() {
            if let Some(root) = canonical_skills_root(
                package_root.to_path_buf(),
                package_root.join("office-skills"),
            )
            .await
            {
                return Some(root);
            }
        }
    }
    let package_root = Path::new(env!("CARGO_MANIFEST_DIR")).to_path_buf();
    canonical_skills_root(package_root.clone(), package_root.join("skills")).await
}

async fn canonical_skills_root(
    package_root: PathBuf,
    skills_root: PathBuf,
) -> Option<(PathBuf, PathBuf)> {
    let package_root = tokio::fs::canonicalize(package_root).await.ok()?;
    let skills_root = tokio::fs::canonicalize(skills_root).await.ok()?;
    let metadata = tokio::fs::metadata(&skills_root).await.ok()?;
    if metadata.is_dir() && skills_root.starts_with(&package_root) {
        Some((package_root, skills_root))
    } else {
        None
    }
}

async fn append_references(skill: &OfficeSkill, content: &mut String) -> UseResult<()> {
    let references = skill.directory.join("references");
    match tokio::fs::symlink_metadata(&references).await {
        Ok(metadata) if metadata.is_dir() && !metadata.file_type().is_symlink() => {}
        Ok(_) => return Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(error) => return Err(skills_io_error("inspect", &references, error)),
    }
    let mut entries = tokio::fs::read_dir(&references)
        .await
        .map_err(|error| skills_io_error("list", &references, error))?;
    let mut paths = Vec::new();
    while let Some(entry) = entries
        .next_entry()
        .await
        .map_err(|error| skills_io_error("list", &references, error))?
    {
        let path = entry.path();
        if path.extension().and_then(|extension| extension.to_str()) == Some("md") {
            paths.push(path);
        }
    }
    paths.sort();
    for path in paths {
        let metadata = tokio::fs::symlink_metadata(&path)
            .await
            .map_err(|error| skills_io_error("inspect", &path, error))?;
        if !metadata.is_file() || metadata.file_type().is_symlink() {
            continue;
        }
        let reference = read_bounded(&path, MAX_SKILL_BYTES).await?;
        let name = path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("reference.md");
        let additional = reference
            .len()
            .saturating_add(name.len())
            .saturating_add(40);
        if content.len().saturating_add(additional) > MAX_FULL_SKILL_BYTES {
            return Err(UseError::new(
                "use.office.skill_too_large",
                format!(
                    "Office Skill '{}' with references exceeds the {MAX_FULL_SKILL_BYTES}-byte limit.",
                    skill.name
                ),
            ));
        }
        content.push_str("\n\n---\n\n## Bundled reference: references/");
        content.push_str(name);
        content.push_str("\n\n");
        content.push_str(&reference);
    }
    Ok(())
}

async fn read_bounded(path: &Path, limit: u64) -> UseResult<String> {
    let metadata = tokio::fs::symlink_metadata(path)
        .await
        .map_err(|error| skills_io_error("inspect", path, error))?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err(UseError::new(
            "use.office.skill_invalid",
            format!(
                "Office Skill file '{}' is not a regular file.",
                path.display()
            ),
        ));
    }
    if metadata.len() > limit {
        return Err(UseError::new(
            "use.office.skill_too_large",
            format!(
                "Office Skill file '{}' exceeds {limit} bytes.",
                path.display()
            ),
        ));
    }
    let file = tokio::fs::File::open(path)
        .await
        .map_err(|error| skills_io_error("open", path, error))?;
    let mut reader = file.take(limit.saturating_add(1));
    let mut bytes = Vec::new();
    reader
        .read_to_end(&mut bytes)
        .await
        .map_err(|error| skills_io_error("read", path, error))?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > limit {
        return Err(UseError::new(
            "use.office.skill_too_large",
            format!(
                "Office Skill file '{}' exceeds {limit} bytes.",
                path.display()
            ),
        ));
    }
    String::from_utf8(bytes).map_err(|error| {
        UseError::new(
            "use.office.skill_invalid",
            format!(
                "Office Skill file '{}' is not valid UTF-8: {error}",
                path.display()
            ),
        )
    })
}

fn parse_frontmatter(content: &str) -> Option<(String, String)> {
    let mut lines = content.lines();
    if lines.next()? != "---" {
        return None;
    }
    let mut name = None;
    let mut description = None;
    let mut closed = false;
    for line in lines {
        if line == "---" {
            closed = true;
            break;
        }
        if let Some(value) = line.strip_prefix("name:") {
            name = Some(value.trim().to_string());
        } else if let Some(value) = line.strip_prefix("description:") {
            description = Some(value.trim().to_string());
        }
    }
    if !closed {
        return None;
    }
    let name = name.filter(|value| valid_skill_name(value))?;
    let description = description.filter(|value| !value.is_empty())?;
    Some((name, description))
}

fn valid_skill_name(value: &str) -> bool {
    !value.is_empty()
        && value.len() <= 64
        && !value.starts_with('-')
        && !value.ends_with('-')
        && !value.contains("--")
        && value
            .bytes()
            .all(|byte| byte.is_ascii_lowercase() || byte.is_ascii_digit() || byte == b'-')
}

fn skills_missing() -> UseError {
    UseError::new(
        "use.office.skills_missing",
        "Packaged A3S Office Skills could not be located.",
    )
    .with_suggestion(format!(
        "Repair A3S Use or set {SKILLS_ENV} to the absolute packaged Office Skills directory."
    ))
}

fn skills_io_error(action: &str, path: &Path, error: std::io::Error) -> UseError {
    UseError::new(
        "use.office.skill_unreadable",
        format!(
            "Failed to {action} Office Skill path '{}': {error}",
            path.display()
        ),
    )
}

fn skills_usage_error(message: impl Into<String>) -> UseError {
    UseError::new("use.cli.invalid_usage", message)
        .with_suggestion("Run 'a3s use office skills help' or 'a3s use office skills list'.")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[tokio::test]
    async fn source_skill_is_discoverable_and_frontmatter_matches_its_directory() {
        let skills = discover_skills().await.unwrap();
        let primary = skills
            .iter()
            .find(|skill| skill.name == PRIMARY_SKILL)
            .unwrap();
        assert!(primary.skill_path.is_absolute());
        assert!(primary.skill_path.ends_with("a3s-office/SKILL.md"));
        assert!(primary.description.contains(".docx"));
    }

    #[tokio::test]
    async fn primary_skill_surface_stays_inside_its_package_root() {
        let (root, path) = primary_skill_surface().await.unwrap();
        assert!(path.starts_with(root));
        assert!(path.ends_with("a3s-office/SKILL.md"));
    }

    #[test]
    fn skill_names_and_frontmatter_fail_closed() {
        assert!(valid_skill_name("a3s-office"));
        assert!(!valid_skill_name("../office"));
        assert!(!valid_skill_name("Office"));
        assert!(parse_frontmatter("---\nname: mismatch\ndescription: Useful\n---\n").is_some());
        assert!(parse_frontmatter("---\nname: bad/name\ndescription: Useful\n---\n").is_none());
        assert!(parse_frontmatter("---\nname: incomplete\ndescription: Useful\n").is_none());
    }

    #[tokio::test]
    async fn discovery_rejects_collections_above_the_entry_bound() {
        let temp = tempfile::tempdir().unwrap();
        for index in 0..=MAX_SKILLS {
            tokio::fs::create_dir(temp.path().join(format!("skill-{index}")))
                .await
                .unwrap();
        }

        let error = discover_skills_at(temp.path()).await.unwrap_err();
        assert_eq!(error.code, "use.office.skills_too_many");
    }

    #[tokio::test]
    async fn bounded_reads_reject_oversized_skill_content() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("SKILL.md");
        tokio::fs::write(&path, b"12345").await.unwrap();

        let error = read_bounded(&path, 4).await.unwrap_err();
        assert_eq!(error.code, "use.office.skill_too_large");
    }

    #[tokio::test]
    async fn no_subcommand_lists_skills_and_full_get_includes_references() {
        let listed = run(&[]).await.unwrap();
        let skills = listed.json["data"].as_array().unwrap();
        assert!(skills.iter().any(|skill| skill["name"] == PRIMARY_SKILL));

        let output = get(&[PRIMARY_SKILL, "--full"]).await.unwrap();
        let content = output.json["data"]["content"].as_str().unwrap();
        assert!(content.contains("## Bundled reference: references/word.md"));
        assert!(content.contains("## Bundled reference: references/spreadsheet.md"));
        assert!(content.contains("## Bundled reference: references/presentation.md"));
        assert!(content.contains("## Bundled reference: references/mcp.md"));
        assert_eq!(output.json["data"]["full"], true);
        assert!(content.len() <= MAX_FULL_SKILL_BYTES);
    }
}
