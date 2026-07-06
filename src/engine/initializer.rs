/// Initializes the native spreadsheet engine — sets app data path, ICU data,
/// and font resource path.  Must be called once before any engine operation.
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};

use crate::engine::ffi::EngineHandle;

static INITIALIZED: AtomicBool = AtomicBool::new(false);

/// Holds the memory-mapped ICU data so it stays alive for the engine's lifetime.
static ICU_MMAP: std::sync::OnceLock<memmap2::Mmap> = std::sync::OnceLock::new();

/// Performs one-time engine initialization.
/// Sets the application data path, loads ICU data, and sets the font resource path.
pub fn initialize(engine: &EngineHandle) -> bool {
    if INITIALIZED.load(Ordering::Relaxed) {
        return true;
    }

    // Engine needs a writable folder for temporary files and recovery data.
    let app_data_path = get_engine_resources_dir();
    if let Err(e) = std::fs::create_dir_all(&app_data_path) {
        eprintln!("Failed to create engine resources directory: {}", e);
        return false;
    }

    if let Err(e) = engine.set_app_data_path(&app_data_path) {
        eprintln!("Failed to set engine app data path: {}", e);
        return false;
    }

    // ── ICU data loading ────────────────────────────────────────────
    //
    // The engine's dylib links ICU with *stub data* (~32 bytes).  We must
    // provide the real ICU 76 data blob via SetDataDirectory(const uint8_t*)
    // before any locale / number-format operation occurs.
    if !load_icu_data(engine) {
        eprintln!("Failed to load ICU data — locale operations will crash.");
        return false;
    }

    // Engine needs access to font files for text measurement during formula
    // evaluation and export.  Probe several directories for "EngineFonts":
    //   1. Next to the executable  (distribution layout)
    //   2. Current working directory (dev / script layout)
    //   3. Next to the engine DLL   (shared layout)
    let candidate_dirs: Vec<std::path::PathBuf> = [
        std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|d| d.to_path_buf())),
        #[cfg(all(target_os = "linux", target_arch = "aarch64"))]
        std::env::current_dir().ok().map(|d| d.join("resources").join("nativeLib").join("linux-arm64")),
        std::env::current_dir().ok().map(|d| d.join("resources").join("nativeLib")),
        std::env::current_dir().ok(),
    ]
    .into_iter()
    .flatten()
    .collect();

     // Suppress verbose engine logs during normal CLI usage.
    let _ = engine.set_file_logging();


    let mut font_path_set = false;
    for dir in &candidate_dirs {
        let fonts_path = dir.join("EngineFonts");
        if fonts_path.is_dir() {
            // The engine concatenates the font filename directly to this path,
            // so we must include a trailing separator.
            let mut fonts_str = fonts_path.to_string_lossy().to_string();
            if !fonts_str.ends_with(std::path::MAIN_SEPARATOR) {
                fonts_str.push(std::path::MAIN_SEPARATOR);
            }
            if let Err(e) = engine.set_font_resource_path(&fonts_str) {
                eprintln!("Warning: failed to set font resource path: {}", e);
            } else {
                font_path_set = true;
            }
            break;
        }
    }
    if !font_path_set {
        eprintln!("Warning: EngineFonts directory not found. Font rendering may fail.");
        eprintln!("  Searched: {:?}", candidate_dirs.iter().map(|d| d.join("EngineFonts")).collect::<Vec<_>>());
    }

   
    INITIALIZED.store(true, Ordering::Relaxed);
    true
}

/// Returns the platform-appropriate engine resources directory path.
pub fn get_engine_resources_dir() -> String {
    let base = if cfg!(target_os = "windows") {
        dirs::data_local_dir()
    } else if cfg!(target_os = "macos") {
        dirs::data_dir() // ~/Library/Application Support
    } else {
        dirs::data_dir() // ~/.local/share
    };

    let dir = base
        .unwrap_or_else(|| std::path::PathBuf::from("."))
        .join("ZsCLI")
        .join("EngineResources");

    dir.to_string_lossy().to_string()
}

/// Returns the full path to the engine resources working directory.
pub fn engine_working_dir() -> std::path::PathBuf {
    Path::new(&get_engine_resources_dir()).to_path_buf()
}

/// Finds the system ICU 76 data file (`icudt76l.dat`), memory-maps it,
/// and passes the raw pointer to the engine via `SetDataDirectory`.
///
/// The mapping is stored in a global static so it remains valid for the
/// entire process lifetime.
fn load_icu_data(engine: &EngineHandle) -> bool {
    const ICU_DAT: &str = "icudt76l.dat";

    // Well-known system paths for ICU data (platform-dependent)
    let mut dynamic_candidates: Vec<String> = Vec::new();

    if cfg!(target_os = "macos") {
        dynamic_candidates.extend([
            "/usr/share/icu/icudt76l.dat".to_string(),
            "/usr/local/share/icu/76.1/icudt76l.dat".to_string(),
            "/opt/homebrew/share/icu/76.1/icudt76l.dat".to_string(),
        ]);
    } else if cfg!(target_os = "linux") {
        dynamic_candidates.extend([
            "/usr/share/icu/76.1/icudt76l.dat".to_string(),
            "/usr/share/icu/icudt76l.dat".to_string(),
            "/usr/local/share/icu/76.1/icudt76l.dat".to_string(),
        ]);
    } else if cfg!(target_os = "windows") {
        // The engine DLL is MinGW-compiled; if the user has MSYS2 installed,
        // ICU data lives under the MinGW prefix.
        if let Ok(prefix) = std::env::var("MINGW_PREFIX") {
            let base = std::path::PathBuf::from(&prefix);
            dynamic_candidates.push(base.join("share").join("icu").join("76.1").join(ICU_DAT).to_string_lossy().to_string());
            dynamic_candidates.push(base.join("bin").join(ICU_DAT).to_string_lossy().to_string());
        }
        // Common MSYS2 installation paths
        for prefix in &[
            r"C:\msys64\mingw64",
            r"C:\msys64\ucrt64",
            r"C:\msys64\clang64",
        ] {
            let base = std::path::PathBuf::from(prefix);
            dynamic_candidates.push(base.join("share").join("icu").join("76.1").join(ICU_DAT).to_string_lossy().to_string());
            dynamic_candidates.push(base.join("bin").join(ICU_DAT).to_string_lossy().to_string());
        }
    }

    // On all platforms (especially Windows where there are no well-known
    // system paths), look next to the executable and in the cwd.
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            dynamic_candidates.push(
                dir.join(ICU_DAT).to_string_lossy().to_string(),
            );
        }
    }
    if let Ok(cwd) = std::env::current_dir() {
        dynamic_candidates.push(
            cwd.join("resources").join(ICU_DAT).to_string_lossy().to_string(),
        );
        dynamic_candidates.push(
            cwd.join(ICU_DAT).to_string_lossy().to_string(),
        );
    }

    // Also check the ICU_DATA environment variable (highest priority)
    let env_path = std::env::var("ICU_DATA").ok().map(|dir| {
        let p = std::path::PathBuf::from(dir).join(ICU_DAT);
        p.to_string_lossy().to_string()
    });

    let all_paths: Vec<String> = env_path
        .into_iter()
        .chain(dynamic_candidates)
        .collect();

    for path in &all_paths {
        let p = Path::new(path.as_str());
        if !p.is_file() {
            continue;
        }

        match std::fs::File::open(p) {
            Ok(file) => {
                // Safety: the file is read-only mapped; the engine only reads it.
                match unsafe { memmap2::Mmap::map(&file) } {
                    Ok(mmap) => {
                        let ptr = mmap.as_ptr();
                        let _len = mmap.len();
                        // Store the mapping so it won't be dropped.
                        let _ = ICU_MMAP.set(mmap);

                        if let Err(e) = engine.set_data_directory(ptr) {
                            eprintln!("Warning: SetDataDirectory failed: {}", e);
                            continue;
                        }

                        // Also set the directory path so the engine's
                        // ClientHelper knows where ICU data lives.
                        let parent = p.parent().map(|d| d.to_string_lossy().to_string())
                            .unwrap_or_default();
                        if !parent.is_empty() {
                            let _ = engine.set_data_directory_path(&parent);
                        }

                        //eprintln!("Loaded ICU data from: {} ({} bytes)", path, len);
                        return true;
                    }
                    Err(e) => {
                        eprintln!("Warning: failed to mmap {}: {}", path, e);
                    }
                }
            }
            Err(e) => {
                eprintln!("Warning: failed to open {}: {}", path, e);
            }
        }
    }

    eprintln!("Error: ICU data file (icudt76l.dat) not found.");
    eprintln!("  Searched: {:?}", all_paths);
    eprintln!("  Set ICU_DATA env variable to the directory containing icudt76l.dat");
    false
}
