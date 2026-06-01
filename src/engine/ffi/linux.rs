/// Linux FFI internals — GCC `__cxx11::basic_string` layout (32 bytes)
///
///  struct basic_string {
///      char*  _M_dataplus;          // 8 B — pointer to char data
///      size_t _M_string_length;     // 8 B
///      union {
///          char   _M_local_buf[16]; // SSO (for len <= 15)
///          size_t _M_allocated_capacity;
///      };
///  };                               // total: 32 bytes
///
/// SSO:  data_ptr -> _M_local_buf  (self-referential, within the struct)
/// Heap: data_ptr -> malloc'd buffer; capacity in _M_allocated_capacity

use std::ffi::c_void;

// ─── NativeString: GCC __cxx11 layout (32 bytes) ────────────────────────────

pub(crate) const NATIVE_SSO_CAP: usize = 15;

#[repr(C)]
pub(crate) struct NativeString {
    pub data_ptr:   *mut u8,
    pub length:     usize,
    pub buf_or_cap: [u8; 16],
}

impl NativeString {
    pub fn zeroed() -> Self {
        Self {
            data_ptr:   std::ptr::null_mut(),
            length:     0,
            buf_or_cap: [0u8; 16],
        }
    }

    pub fn to_rust_string(&self) -> String {
        if self.length == 0 || self.data_ptr.is_null() {
            return String::new();
        }
        let bytes = unsafe { std::slice::from_raw_parts(self.data_ptr, self.length) };
        String::from_utf8_lossy(bytes).into_owned()
    }

    /// Free heap-allocated string data (no-op for SSO strings).
    pub unsafe fn dispose(&self) {
        if self.length > NATIVE_SSO_CAP && !self.data_ptr.is_null() {
            extern "C" { fn free(ptr: *mut c_void); }
            free(self.data_ptr as *mut c_void);
        }
    }
}

// ─── RefNativeString: heap-pinned, for passing `const std::string&` ──────────

pub(crate) struct RefNativeString {
    inner: Box<NativeString>,
    /// Keeps heap data alive for non-SSO strings.
    _heap: Option<Vec<u8>>,
}

impl RefNativeString {
    pub fn new(s: &str) -> Self {
        let mut gs = Box::new(NativeString::zeroed());
        gs.length = s.len();

        let heap = if s.len() <= NATIVE_SSO_CAP {
            // SSO: copy into local buffer.  The Box is heap-pinned so the
            // self-referential data_ptr -> buf_or_cap is stable.
            gs.buf_or_cap[..s.len()].copy_from_slice(s.as_bytes());
            if s.len() < 16 { gs.buf_or_cap[s.len()] = 0; }
            let buf_ptr = gs.buf_or_cap.as_mut_ptr();
            gs.data_ptr = buf_ptr;
            None
        } else {
            let mut data = Vec::with_capacity(s.len() + 1);
            data.extend_from_slice(s.as_bytes());
            data.push(0);
            gs.data_ptr = data.as_mut_ptr();
            gs.buf_or_cap[..8].copy_from_slice(&s.len().to_ne_bytes());
            Some(data)
        };
        Self { inner: gs, _heap: heap }
    }

    pub fn as_ptr(&self) -> *const NativeString {
        &*self.inner as *const NativeString
    }
}

// ─── TransferNativeString: by-value transfer, C++ callee destroys ────────────

pub(crate) struct TransferNativeString {
    inner: Box<NativeString>,
}

impl TransferNativeString {
    pub fn new(s: &str) -> Self {
        let alloc_len = s.len() + 1;
        let layout = std::alloc::Layout::from_size_align(alloc_len, 1).unwrap();
        let heap_ptr = unsafe { std::alloc::alloc(layout) };
        if heap_ptr.is_null() {
            std::alloc::handle_alloc_error(layout);
        }
        unsafe {
            std::ptr::copy_nonoverlapping(s.as_bytes().as_ptr(), heap_ptr, s.len());
            *heap_ptr.add(s.len()) = 0;
        }
        let mut gs = Box::new(NativeString::zeroed());
        gs.data_ptr = heap_ptr;
        gs.length = s.len();
        // ABI: store capacity (usable chars, not including NUL).
        gs.buf_or_cap[..8].copy_from_slice(&s.len().to_ne_bytes());
        Self { inner: gs }
    }

    pub fn as_mut_ptr(&mut self) -> *mut NativeString {
        &mut *self.inner as *mut NativeString
    }
}

// ─── Mangled symbol names (GCC / libstdc++ Itanium ABI) ─────────────────────

pub(crate) mod sym {
    /// `ZSEngine::RequestManager::SetAppDataPath(std::string)`   (GCC __cxx11)
    pub const SET_APP_DATA_PATH: &[u8] =
        b"_ZN8ZSEngine14RequestManager14SetAppDataPathENSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::RequestManager::SetFontResourcePath(std::string)`  (GCC __cxx11)
    pub const SET_FONT_RESOURCE_PATH: &[u8] =
        b"_ZN8ZSEngine14RequestManager19SetFontResourcePathENSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::RequestManager::SetDataDirectoryPath(std::string)`  (GCC __cxx11)
    pub const SET_DATA_DIRECTORY_PATH: &[u8] =
        b"_ZN8ZSEngine14RequestManager20SetDataDirectoryPathENSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::RequestManager::SetDataDirectory(unsigned char const*)`
    pub const SET_DATA_DIRECTORY: &[u8] =
        b"_ZN8ZSEngine14RequestManager16SetDataDirectoryEPKh\0";

    /// `ZSEngine::RequestManager::ProcessRequest(const std::string&)` -> ZSResponse
    pub const PROCESS_REQUEST: &[u8] =
        b"_ZN8ZSEngine14RequestManager14ProcessRequestERKNSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::RequestManager::ProcessRequestWithFlatBuffers(const std::string&)` -> ZSResponse
    pub const PROCESS_REQUEST_WITH_FLAT_BUFFERS: &[u8] =
        b"_ZN8ZSEngine14RequestManager29ProcessRequestWithFlatBuffersERKNSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::RequestManager::Fetch(const std::string&)` -> ZSResponse
    pub const FETCH: &[u8] =
        b"_ZN8ZSEngine14RequestManager5FetchERKNSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::RequestManager::DocFetch(const std::string&)` -> ZSResponse
    pub const DOC_FETCH: &[u8] =
        b"_ZN8ZSEngine14RequestManager8DocFetchERKNSt7__cxx1112basic_stringIcSt11char_traitsIcESaIcEEE\0";

    /// `ZSEngine::Tools::SetFileLogging()`
    pub const SET_FILE_LOGGING: &[u8] =
        b"_ZN8ZSEngine5Tools14SetFileLoggingEv\0";

    /// `ZSEngine::Tools::EnableLogging()`
    pub const ENABLE_LOGGING: &[u8] =
        b"_ZN8ZSEngine5Tools13EnableLoggingEv\0";

    /// `ZSEngine::Tools::DisableLogging()`
    pub const DISABLE_LOGGING: &[u8] =
        b"_ZN8ZSEngine5Tools14DisableLoggingEv\0";

    /// `ZSResponse::GetResponseString()` -> std::string
    pub const RESPONSE_GET_STRING: &[u8] =
        b"_ZN10ZSResponse17GetResponseStringB5cxx11Ev\0";
}

/// Library probe names for Linux.
pub(crate) const LIB_NAMES: &[&str] = &["libNativeClientEngine.so"];
