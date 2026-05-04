/// build.rs — Platform-specific C++ runtime linking for the native engine FFI,
/// and auto-copy of engine resources from `resources/` into the target dir.
///
/// The engine library (`libNativeClientEngine`) is loaded at runtime via
/// `libloading`, so we do NOT link it at build time.  However, the C++
/// standard library used by the engine must match at the ABI level:
///
///  - **macOS**:   libc++ (ships with the system, linked by default by Clang)
///  - **Linux**:   libstdc++ (GCC's C++ runtime)
///  - **Windows**: MSVC CRT or MinGW libstdc++ (handled by the toolchain)

use std::path::{Path, PathBuf};

fn main() {
    // ── macOS: link libc++ ──────────────────────────────────────────
    #[cfg(target_os = "macos")]
    {
        println!("cargo:rustc-link-lib=dylib=c++");
    }

    // ── Linux: link libstdc++ ───────────────────────────────────────
    #[cfg(target_os = "linux")]
    {
        println!("cargo:rustc-link-lib=dylib=stdc++");
    }

    // ── Copy engine resources into the target output dir ────────────
    //
    // Files like libNativeClientEngine.dll, icudt76l.dat, EngineFonts/, etc.
    // live in `resources/` at the project root so they survive `cargo clean`.
    // This step copies them next to the built executable automatically.
    let manifest_dir = PathBuf::from(std::env::var("CARGO_MANIFEST_DIR").unwrap());
    let resources_dir = manifest_dir.join("resources");

    // Re-run this build script whenever the resources directory changes.
    println!("cargo:rerun-if-changed=resources");
    println!("cargo:rerun-if-changed=build.rs");

    if resources_dir.is_dir() {
        let out_dir = PathBuf::from(std::env::var("OUT_DIR").unwrap());
        // OUT_DIR is something like target/debug/build/<crate>/out
        // Walk up to find target/debug (or target/release).
        let target_dir = out_dir
            .ancestors()
            .find(|p| p.file_name().map_or(false, |n| n == "debug" || n == "release"))
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| out_dir.clone());

        copy_dir_recursive(&resources_dir, &target_dir);
    }
}

/// Recursively copies all files/directories from `src` into `dst`.
/// Existing files are overwritten only if the source is newer.
fn copy_dir_recursive(src: &Path, dst: &Path) {
    let entries = match std::fs::read_dir(src) {
        Ok(e) => e,
        Err(e) => {
            eprintln!("cargo:warning=Cannot read resources dir {}: {}", src.display(), e);
            return;
        }
    };

    for entry in entries.flatten() {
        let src_path = entry.path();
        let dst_path = dst.join(entry.file_name());

        if src_path.is_dir() {
            let _ = std::fs::create_dir_all(&dst_path);
            copy_dir_recursive(&src_path, &dst_path);
        } else {
            // Skip copy if destination is already up-to-date.
            let needs_copy = match (src_path.metadata(), dst_path.metadata()) {
                (Ok(src_meta), Ok(dst_meta)) => {
                    src_meta.len() != dst_meta.len()
                        || src_meta.modified().ok() > dst_meta.modified().ok()
                }
                _ => true,
            };
            if needs_copy {
                if let Err(e) = std::fs::copy(&src_path, &dst_path) {
                    eprintln!(
                        "cargo:warning=Failed to copy {} -> {}: {}",
                        src_path.display(),
                        dst_path.display(),
                        e
                    );
                }
            }
        }
    }
}
