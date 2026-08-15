use std::io::Write as _;

use super::read_library;

#[tokio::test]
async fn source_library_is_mutation_locked_until_engine_initialization() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("pdfium.dll");
    std::fs::write(&path, b"exact-pdfium-library-bytes").unwrap();

    let input = read_library(&path).await.unwrap();
    assert_eq!(input.bytes, b"exact-pdfium-library-bytes");
    assert_eq!(input.canonical_path, std::fs::canonicalize(&path).unwrap());
    assert!(std::fs::OpenOptions::new().write(true).open(&path).is_err());

    drop(input);
    std::fs::OpenOptions::new()
        .write(true)
        .truncate(true)
        .open(&path)
        .unwrap()
        .write_all(b"replacement")
        .unwrap();
    assert_eq!(std::fs::read(path).unwrap(), b"replacement");
}
