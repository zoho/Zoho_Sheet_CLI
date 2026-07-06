/// In-memory state of the currently active CLI workbook session.
#[derive(Debug, Default)]
pub struct CliSession {
    /// Resource identifier returned by the engine after opening/creating a workbook.
    pub rid: Option<String>,
    /// Display name of the currently open workbook.
    pub workbook_name: Option<String>,
    /// Zero-based index of the currently active sheet.
    pub active_sheet_index: usize,
    /// Name of the currently active sheet.
    pub active_sheet_name: Option<String>,
    /// File path of the currently open workbook on disk.
    pub file_path: Option<String>,
    /// Total number of sheets in the workbook.
    pub sheet_count: usize,
    /// List of sheet names in the workbook.
    pub sheet_names: Vec<String>,
    /// List of engine sheet identifiers aligned with `sheet_names`.
    pub sheet_ids: Vec<String>,
    /// Whether the workbook has unsaved changes.
    pub is_dirty: bool,
    /// Cache of chart renames: maps user-assigned name → chart_id.
    /// Cleared on sheet switch or workbook close.
    pub chart_name_cache: std::collections::HashMap<String, String>,
}

impl CliSession {
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether a workbook is currently open.
    pub fn is_active(&self) -> bool {
        self.rid.as_ref().map_or(false, |r| !r.is_empty())
    }

    /// Gets the active sheet identifier for engine requests.
    pub fn get_active_sheet_id_or_default(&self) -> String {
        if self.active_sheet_index < self.sheet_ids.len() {
            let id = &self.sheet_ids[self.active_sheet_index];
            if !id.trim().is_empty() && id != &self.active_sheet_index.to_string() {
                return id.clone();
            }
        }

        if let Some(ref name) = self.active_sheet_name {
            if !name.trim().is_empty() {
                return name.clone();
            }
        }

        self.active_sheet_index.to_string()
    }

    /// Resets the session to an inactive state, clearing all workbook information.
    pub fn clear(&mut self) {
        self.rid = None;
        self.workbook_name = None;
        self.active_sheet_index = 0;
        self.active_sheet_name = None;
        self.file_path = None;
        self.sheet_count = 0;
        self.is_dirty = false;
        self.sheet_names.clear();
        self.sheet_ids.clear();
        self.chart_name_cache.clear();
    }
}
