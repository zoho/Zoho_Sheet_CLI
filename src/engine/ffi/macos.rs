/// macOS FFI internals — libc++ alternate string layout (24 bytes)
///
/// Apple's libc++ uses `_LIBCPP_ABI_ALTERNATE_STRING_LAYOUT`.  On little-endian:
///
///  __long  (is_long = 1):
///      bytes 0-7:   char* __data_       — pointer to heap buffer
///      bytes 8-15:  size_t __size_      — string length
///      bytes 16-23: __cap_ | (1 << 63)  — capacity with is_long flag in bit 63
///
///  __short (is_long = 0):
///      bytes 0-22:  char __data_[23]    — inline buffer (SSO, null-terminated)
///      byte  23:    __size_ (bits 0-6), __is_long_ = 0 (bit 7)
///
///  Discriminant: byte 23, bit 7.  Same position in both layouts on LE.

use std::ffi::c_void;

// ─── NativeString: libc++ alternate layout (24 bytes) ────────────────────────

pub(crate) const NATIVE_SSO_CAP: usize = 22;

#[repr(C, align(8))]
pub(crate) struct NativeString {
    pub _data: [u8; 24],
}

impl NativeString {
    pub fn zeroed() -> Self {
        Self { _data: [0u8; 24] }
    }

    fn is_long(&self) -> bool {
        (self._data[23] & 0x80) != 0
    }

    fn data_ptr(&self) -> *const u8 {
        if self.is_long() {
            usize::from_ne_bytes(self._data[0..8].try_into().unwrap()) as *const u8
        } else {
            self._data.as_ptr()
        }
    }

    fn length(&self) -> usize {
        if self.is_long() {
            usize::from_ne_bytes(self._data[8..16].try_into().unwrap())
        } else {
            (self._data[23] & 0x7F) as usize
        }
    }

    pub fn to_rust_string(&self) -> String {
        let len = self.length();
        if len == 0 {
            return String::new();
        }
        let ptr = self.data_ptr();
        if ptr.is_null() {
            return String::new();
        }
        let bytes = unsafe { std::slice::from_raw_parts(ptr, len) };
        String::from_utf8_lossy(bytes).into_owned()
    }

    /// Free heap-allocated string data (no-op for SSO strings).
    pub unsafe fn dispose(&self) {
        if self.is_long() {
            let ptr = self.data_ptr();
            if !ptr.is_null() {
                extern "C" { fn free(ptr: *mut c_void); }
                free(ptr as *mut c_void);
            }
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
        let mut ns = Box::new(NativeString::zeroed());

        if s.len() <= NATIVE_SSO_CAP {
            // SSO: copy into inline buffer (bytes 0..22), set size in byte 23
            ns._data[..s.len()].copy_from_slice(s.as_bytes());
            if s.len() < 23 { ns._data[s.len()] = 0; }
            ns._data[23] = s.len() as u8; // is_long=0 (bit 7 clear), size in bits 0-6
            Self { inner: ns, _heap: None }
        } else {
            let mut data = Vec::with_capacity(s.len() + 1);
            data.extend_from_slice(s.as_bytes());
            data.push(0); // null-terminate

            // __long layout: data ptr, size, cap|flag
            let ptr = data.as_ptr() as usize;
            ns._data[0..8].copy_from_slice(&ptr.to_ne_bytes());
            ns._data[8..16].copy_from_slice(&s.len().to_ne_bytes());
            let cap_with_flag = s.len() | (1usize << 63);
            ns._data[16..24].copy_from_slice(&cap_with_flag.to_ne_bytes());

            Self { inner: ns, _heap: Some(data) }
        }
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

        let mut ns = Box::new(NativeString::zeroed());
        // __long layout: data ptr, size, cap|flag
        ns._data[0..8].copy_from_slice(&(heap_ptr as usize).to_ne_bytes());
        ns._data[8..16].copy_from_slice(&s.len().to_ne_bytes());
        let cap_with_flag = s.len() | (1usize << 63);
        ns._data[16..24].copy_from_slice(&cap_with_flag.to_ne_bytes());

        Self { inner: ns }
    }

    pub fn as_mut_ptr(&mut self) -> *mut NativeString {
        &mut *self.inner as *mut NativeString
    }
}

// ─── Mangled symbol names (Clang / libc++ Itanium ABI) ───────────────────────

pub(crate) mod sym {
    /// `ZSEngine::RequestManager::SetAppDataPath(std::string)`
    pub const SET_APP_DATA_PATH: &[u8] =
        b"_ZN8ZSEngine14RequestManager14SetAppDataPathENSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

    /// `ZSEngine::RequestManager::SetFontResourcePath(std::string)`
    pub const SET_FONT_RESOURCE_PATH: &[u8] =
        b"_ZN8ZSEngine14RequestManager19SetFontResourcePathENSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

    /// `ZSEngine::RequestManager::SetDataDirectoryPath(std::string)`
    pub const SET_DATA_DIRECTORY_PATH: &[u8] =
        b"_ZN8ZSEngine14RequestManager20SetDataDirectoryPathENSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

    /// `ZSEngine::RequestManager::SetDataDirectory(unsigned char const*)`
    pub const SET_DATA_DIRECTORY: &[u8] =
        b"_ZN8ZSEngine14RequestManager16SetDataDirectoryEPKh\0";

    /// `ZSEngine::RequestManager::ProcessRequest(const std::string&)` -> ZSResponse
    pub const PROCESS_REQUEST: &[u8] =
        b"_ZN8ZSEngine14RequestManager14ProcessRequestERKNSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

    /// `ZSEngine::RequestManager::ProcessRequestWithFlatBuffers(const std::string&)` -> ZSResponse
    pub const PROCESS_REQUEST_WITH_FLAT_BUFFERS: &[u8] =
        b"_ZN8ZSEngine14RequestManager29ProcessRequestWithFlatBuffersERKNSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

    /// `ZSEngine::RequestManager::Fetch(const std::string&)` -> ZSResponse
    pub const FETCH: &[u8] =
        b"_ZN8ZSEngine14RequestManager5FetchERKNSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

    /// `ZSEngine::RequestManager::DocFetch(const std::string&)` -> ZSResponse
    pub const DOC_FETCH: &[u8] =
        b"_ZN8ZSEngine14RequestManager8DocFetchERKNSt3__112basic_stringIcNS1_11char_traitsIcEENS1_9allocatorIcEEEE\0";

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
        b"_ZN10ZSResponse17GetResponseStringEv\0";
}

/// Library probe names for macOS.
pub(crate) const LIB_NAMES: &[&str] = &["libNativeClientEngine.dylib"];
