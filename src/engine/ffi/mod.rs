/// Engine FFI — dynamically loads `libNativeClientEngine` (.dll/.so/.dylib)
/// and calls C++ mangled symbols from `ZSEngine::RequestManager` (static
/// methods) and `ZSResponse` (instance method) directly.
///
/// # Platform-specific ABI handling
///
/// Platform-specific details (NativeString layout, symbol mangling, library
/// names) are isolated into sub-modules selected at compile time:
///
///  - **macOS**   → `macos.rs`   — Clang/libc++ alternate string layout (24 B)
///  - **Linux**   → `linux.rs`   — GCC __cxx11 string layout (32 B)
///  - **Windows** → `windows.rs` — MinGW/GCC __cxx11 string layout (32 B)
///
/// This module re-exports the correct platform implementation and provides
/// the shared `EngineHandle` public API that the rest of the crate uses.

use std::path::{Path, PathBuf};

use libloading::Library;

// ═══════════════════════════════════════════════════════════════════════════════
// Platform-specific module selection
// ═══════════════════════════════════════════════════════════════════════════════

#[cfg(target_os = "macos")]
#[path = "macos.rs"]
mod platform;

#[cfg(target_os = "linux")]
#[path = "linux.rs"]
mod platform;

#[cfg(target_os = "windows")]
#[path = "windows.rs"]
mod platform;

use platform::{
    NativeString, RefNativeString, TransferNativeString,
    sym, LIB_NAMES,
};

// ─── Error type ──────────────────────────────────────────────────────────────

#[derive(Debug)]
pub enum EngineError {
    LibraryNotFound(String),
    SymbolNotFound(String),
    NullResponse,
    CallFailed(String),
}

impl std::fmt::Display for EngineError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LibraryNotFound(m) => write!(f, "Engine library not found: {}", m),
            Self::SymbolNotFound(m)  => write!(f, "Engine symbol not found: {}", m),
            Self::NullResponse       => write!(f, "Engine returned null response"),
            Self::CallFailed(m)      => write!(f, "Engine call failed: {}", m),
        }
    }
}

impl std::error::Error for EngineError {}

// ─── Opaque return buffer for ZSResponse ─────────────────────────────────────
//
// We don't know the exact layout of `ZSResponse`, but it is a fixed-size C++
// object.  512 bytes is a generous over-allocation (typical size ~80-120 B).

#[repr(C, align(16))]
struct ZSResponseBuf {
    _data: [u8; 512],
}

impl ZSResponseBuf {
    fn zeroed() -> Self {
        Self { _data: [0u8; 512] }
    }
}

// ─── Function-pointer types ──────────────────────────────────────────────────
//
// On all platforms, non-trivially-destructible C++ objects (std::string,
// ZSResponse) are passed/returned indirectly.  By declaring large return
// types directly in the function signature, Rust's `extern "C"` ABI
// automatically passes the hidden struct-return pointer in the correct
// platform register (rdi on x86-64 SysV, rcx on Win64, x8 on ARM64).
//
// Static method  SetAppDataPath(std::string):
//   Caller constructs the string and passes a pointer (indirect by-value).
type SetPathFn = unsafe extern "C" fn(*mut NativeString);

// Static method  SetDataDirectory(unsigned char const*):
//   Simple raw pointer to ICU data blob.
type SetDataDirectoryFn = unsafe extern "C" fn(*const u8);

// Static method  ProcessRequest(const std::string&) -> ZSResponse:
//   Reference param = pointer; large return via sret (handled by Rust).
type RequestFn = unsafe extern "C" fn(*const NativeString) -> ZSResponseBuf;

// Static method  SetFileLogging() / EnableLogging() / DisableLogging():
//   No parameters, no return value.
type VoidFn = unsafe extern "C" fn();

// Instance method  ZSResponse::GetResponseString() -> std::string:
//   this pointer as first param; large return via sret (handled by Rust).
type GetResponseStringFn = unsafe extern "C" fn(*mut ZSResponseBuf) -> NativeString;

// ─── EngineHandle ────────────────────────────────────────────────────────────

pub struct EngineHandle {
    _lib: Library,
    fn_set_app_data_path:       SetPathFn,
    fn_set_font_resource_path:  SetPathFn,
    fn_set_data_directory_path: SetPathFn,
    fn_set_data_directory:      SetDataDirectoryFn,
    fn_process_request:         RequestFn,
    fn_process_request_with_flat_buffers: RequestFn,
    fn_fetch:                   RequestFn,
    fn_doc_fetch:               RequestFn,
    fn_get_response_string:     GetResponseStringFn,
    fn_set_file_logging:        VoidFn,
    fn_enable_logging:          VoidFn,
    fn_disable_logging:         VoidFn,
}

// The engine is internally thread-safe.
unsafe impl Send for EngineHandle {}
unsafe impl Sync for EngineHandle {}

impl EngineHandle {
    /// Load `libNativeClientEngine` from the given directory, the exe
    /// directory, or the current working directory.
    pub fn load(search_dir: Option<&Path>) -> Result<Self, EngineError> {
        let dirs = build_search_dirs(search_dir);
        let mut last_err = String::new();

        for lib_name in LIB_NAMES {
            for dir in &dirs {
                let full_path = dir.join(lib_name);
                if !full_path.exists() { continue; }
                match unsafe { Library::new(full_path.as_os_str()) } {
                    Ok(lib) => return Self::resolve(lib),
                    Err(e) => last_err = format!("{}: {}", full_path.display(), e),
                }
            }
            // Also try bare name (system search path / LD_LIBRARY_PATH)
            match unsafe { Library::new(lib_name) } {
                Ok(lib) => return Self::resolve(lib),
                Err(e) => last_err = format!("{}: {}", lib_name, e),
            }
        }

        Err(EngineError::LibraryNotFound(format!(
            "Could not load {:?}. Last error: {}", LIB_NAMES, last_err
        )))
    }

    /// Resolve all required mangled symbols from the loaded library.
    fn resolve(lib: Library) -> Result<Self, EngineError> {
        unsafe {
            let fn1  = load_sym::<SetPathFn>(&lib, sym::SET_APP_DATA_PATH, "SetAppDataPath")?;
            let fn2  = load_sym::<SetPathFn>(&lib, sym::SET_FONT_RESOURCE_PATH, "SetFontResourcePath")?;
            let fn2b = load_sym::<SetPathFn>(&lib, sym::SET_DATA_DIRECTORY_PATH, "SetDataDirectoryPath")?;
            let fn2c = load_sym::<SetDataDirectoryFn>(&lib, sym::SET_DATA_DIRECTORY, "SetDataDirectory")?;
            let fn3  = load_sym::<RequestFn>(&lib, sym::PROCESS_REQUEST, "ProcessRequest")?;
            let fn3b = load_sym::<RequestFn>(&lib, sym::PROCESS_REQUEST_WITH_FLAT_BUFFERS, "ProcessRequestWithFlatBuffers")?;
            let fn4  = load_sym::<RequestFn>(&lib, sym::FETCH, "Fetch")?;
            let fn5  = load_sym::<RequestFn>(&lib, sym::DOC_FETCH, "DocFetch")?;
            let fn6  = load_sym::<GetResponseStringFn>(&lib, sym::RESPONSE_GET_STRING, "GetResponseString")?;
            let fn7  = load_sym::<VoidFn>(&lib, sym::SET_FILE_LOGGING, "SetFileLogging")?;
            let fn8  = load_sym::<VoidFn>(&lib, sym::ENABLE_LOGGING, "EnableLogging")?;
            let fn9  = load_sym::<VoidFn>(&lib, sym::DISABLE_LOGGING, "DisableLogging")?;
            Ok(Self {
                _lib: lib,
                fn_set_app_data_path:       fn1,
                fn_set_font_resource_path:  fn2,
                fn_set_data_directory_path: fn2b,
                fn_set_data_directory:      fn2c,
                fn_process_request:         fn3,
                fn_process_request_with_flat_buffers: fn3b,
                fn_fetch:                   fn4,
                fn_doc_fetch:               fn5,
                fn_get_response_string:     fn6,
                fn_set_file_logging:        fn7,
                fn_enable_logging:          fn8,
                fn_disable_logging:         fn9,
            })
        }
    }

    // ── Public API ───────────────────────────────────────────────────

    pub fn set_app_data_path(&self, path: &str) -> Result<(), EngineError> {
        let mut gs = TransferNativeString::new(path);
        unsafe { (self.fn_set_app_data_path)(gs.as_mut_ptr()); }
        Ok(())
    }

    pub fn set_font_resource_path(&self, path: &str) -> Result<(), EngineError> {
        let mut gs = TransferNativeString::new(path);
        unsafe { (self.fn_set_font_resource_path)(gs.as_mut_ptr()); }
        Ok(())
    }

    /// Sets the data directory path (std::string overload).
    pub fn set_data_directory_path(&self, path: &str) -> Result<(), EngineError> {
        let mut gs = TransferNativeString::new(path);
        unsafe { (self.fn_set_data_directory_path)(gs.as_mut_ptr()); }
        Ok(())
    }

    /// Directs engine logging output to a file.
    pub fn set_file_logging(&self) -> Result<(), EngineError> {
        unsafe { (self.fn_set_file_logging)(); }
        Ok(())
    }

    /// Enables engine logging.
    pub fn enable_logging(&self) -> Result<(), EngineError> {
        unsafe { (self.fn_enable_logging)(); }
        Ok(())
    }

    /// Disables engine logging.
    pub fn disable_logging(&self) -> Result<(), EngineError> {
        unsafe { (self.fn_disable_logging)(); }
        Ok(())
    }

    /// Passes raw ICU data blob pointer to the engine.
    /// The data must remain valid for the lifetime of the engine.
    pub fn set_data_directory(&self, data: *const u8) -> Result<(), EngineError> {
        unsafe { (self.fn_set_data_directory)(data); }
        Ok(())
    }

    /// Sends a mutation request (write cell, insert row, undo, ...).
    pub fn process_request_json(&self, json: &str) -> Result<String, EngineError> {
        self.call_returning_string(self.fn_process_request, json)
    }

    /// Sends a request via ProcessRequestWithFlatBuffers (open workbook, etc.).
    pub fn process_request_with_flat_buffers(&self, json: &str) -> Result<String, EngineError> {
        self.call_returning_string(self.fn_process_request_with_flat_buffers, json)
    }

    /// Sends a read request (cell values via fetch schema).
    pub fn fetch_json(&self, json: &str) -> Result<String, EngineError> {
        self.call_returning_string(self.fn_fetch, json)
    }

    /// Sends a document-level metadata request (sheet lists, etc.).
    pub fn doc_fetch_json(&self, json: &str) -> Result<String, EngineError> {
        self.call_returning_string(self.fn_doc_fetch, json)
    }

    // ── Internals ────────────────────────────────────────────────────

    /// Call a `RequestManager` static method that takes `const std::string&`
    /// and returns `ZSResponse`, then extract the response string.
    fn call_returning_string(&self, func: RequestFn, json: &str) -> Result<String, EngineError> {
        let gs_input = RefNativeString::new(json);
        let ns_ptr = gs_input.as_ptr();

        unsafe {
            let mut response: ZSResponseBuf = (func)(ns_ptr);
            let result_str: NativeString = (self.fn_get_response_string)(&mut response);
            let rust_string = result_str.to_rust_string();
            result_str.dispose();
            Ok(rust_string)
        }
    }
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn build_search_dirs(extra: Option<&Path>) -> Vec<PathBuf> {
    let mut dirs = Vec::new();
    if let Some(d) = extra {
        dirs.push(d.to_path_buf());
    }
    if let Ok(exe) = std::env::current_exe() {
        if let Some(d) = exe.parent() {
            dirs.push(d.to_path_buf());
        }
    }
    if let Ok(cwd) = std::env::current_dir() {
        // On Linux arm64, the native library lives in a separate subdirectory.
        #[cfg(all(target_os = "linux", target_arch = "aarch64"))]
        dirs.push(cwd.join("resources").join("nativeLib").join("linux-arm64"));

        dirs.push(cwd.join("resources").join("nativeLib"));
        dirs.push(cwd);
    }
    dirs
}

/// Load a symbol from the library using its mangled name.
/// `friendly` is only used in error messages.
unsafe fn load_sym<T: Copy>(
    lib: &Library,
    mangled: &[u8],
    friendly: &str,
) -> Result<T, EngineError> {
    let sym: libloading::Symbol<T> = lib
        .get(mangled)
        .map_err(|e| EngineError::SymbolNotFound(format!("{}: {}", friendly, e)))?;
    Ok(*sym)
}
