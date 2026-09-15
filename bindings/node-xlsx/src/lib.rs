#![deny(clippy::all)]

use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{
    CalculationOptions, CellEdit, CellRange, CellRef, CellValue as CoreCellValue,
    HorizontalAlignment, MAX_COLLABORATION_CLIENT_ID, MutationResult as CoreMutationResult,
    NumberFormatMutation, Proposal as CoreProposal, ProposalEditInput as CoreProposalEditInput,
    ProposalRequest, RenderOptions, SheetId, StylePatch, TextWrapping, VerticalAlignment, Workbook,
};
use napi::bindgen_prelude::{AsyncTask, Buffer, Task};
use napi::{Env, Error, Result};
use napi_derive::napi;

fn error(reason: impl ToString) -> Error {
    Error::from_reason(reason.to_string())
}

fn lock<T>(value: &Mutex<T>) -> Result<std::sync::MutexGuard<'_, T>> {
    value.lock().map_err(|_| error("workbook lock is poisoned"))
}

fn cell_ref(address: &str) -> Result<CellRef> {
    CellRef::parse_a1(address).map_err(error)
}

fn cell_range(range: &str) -> Result<CellRange> {
    CellRange::parse_a1(range).map_err(error)
}

fn client_id(value: f64) -> Result<u64> {
    if !value.is_finite()
        || value.fract() != 0.0
        || value < 1.0
        || value > MAX_COLLABORATION_CLIENT_ID as f64
    {
        return Err(error("clientId must be a positive safe integer"));
    }
    Ok(value as u64)
}

#[napi(object)]
pub struct OpenWorkbookOptions {
    pub client_id: Option<f64>,
    pub read_only: Option<bool>,
    pub recalculate: Option<bool>,
    pub now_serial: Option<f64>,
}

#[napi(object)]
pub struct RenderSheetOptions {
    pub sheet: Option<u32>,
    pub range: Option<String>,
    pub scale: Option<f64>,
    pub max_width: Option<u32>,
    pub max_height: Option<u32>,
}

#[napi(object)]
pub struct RenderedSheet {
    pub data: Buffer,
    pub width: u32,
    pub height: u32,
}

#[napi(object)]
pub struct CellValue {
    pub address: String,
    pub input: String,
    pub formula: bool,
}

#[napi(object)]
pub struct CellInput {
    pub address: String,
    pub input: String,
}

#[napi(object)]
pub struct ProposalInput {
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposalCellInput>,
}

#[napi(object)]
pub struct ProposalCellInput {
    pub sheet: u32,
    pub address: String,
    pub input: String,
}

#[napi(object)]
pub struct Proposal {
    pub id: String,
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposedCell>,
}

#[napi(object)]
pub struct ProposedCell {
    pub sheet: u32,
    pub address: String,
    pub input: String,
    pub before: String,
    pub after: String,
}

#[napi(object)]
pub struct CellComputedValue {
    pub kind: String,
    pub number: Option<f64>,
    pub text: Option<String>,
    pub boolean: Option<bool>,
    pub error: Option<String>,
}

#[napi(object)]
pub struct StyleInput {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strikethrough: Option<bool>,
    pub font_family: Option<String>,
    pub font_size: Option<f64>,
    pub text_color: Option<String>,
    pub fill_color: Option<String>,
    pub horizontal_alignment: Option<String>,
    pub vertical_alignment: Option<String>,
    pub text_wrapping: Option<String>,
}

#[napi(object)]
pub struct MutationResult {
    pub applied: bool,
    pub changed: Vec<String>,
    pub cycle_cells: Vec<String>,
    pub limited_cells: Vec<String>,
}

#[napi(object)]
pub struct CalculationResult {
    pub changed: Vec<String>,
    pub cycle_cells: Vec<String>,
    pub limited_cells: Vec<String>,
}

#[napi(object)]
pub struct HistoryState {
    pub can_undo: bool,
    pub can_redo: bool,
    pub undo_depth: u32,
    pub redo_depth: u32,
}

#[napi(object)]
pub struct SheetInfo {
    pub ids: Vec<String>,
    pub names: Vec<String>,
    pub active_sheet: u32,
    pub content_width: f64,
    pub content_height: f64,
    pub frozen_rows: u32,
    pub frozen_columns: u32,
    pub initial_scroll_x: f64,
    pub initial_scroll_y: f64,
}

fn map_cell(address: String, value: CellEdit) -> CellValue {
    CellValue {
        address,
        input: value.input,
        formula: value.is_formula,
    }
}

fn addresses(workbook: &Workbook, values: Vec<betteroffice_xlsx::CellAddress>) -> Vec<String> {
    values
        .into_iter()
        .map(|address| workbook.format_address(address))
        .collect()
}

fn map_mutation(workbook: &Workbook, value: CoreMutationResult) -> MutationResult {
    MutationResult {
        applied: value.applied,
        changed: addresses(workbook, value.changed),
        cycle_cells: addresses(workbook, value.cycle_cells),
        limited_cells: addresses(workbook, value.limited_cells),
    }
}

fn map_computed_value(value: &CoreCellValue) -> CellComputedValue {
    match value {
        CoreCellValue::Empty => CellComputedValue {
            kind: "empty".to_owned(),
            number: None,
            text: None,
            boolean: None,
            error: None,
        },
        CoreCellValue::Number { value } => CellComputedValue {
            kind: "number".to_owned(),
            number: Some(*value),
            text: None,
            boolean: None,
            error: None,
        },
        CoreCellValue::Text { value } => CellComputedValue {
            kind: "text".to_owned(),
            number: None,
            text: Some(value.clone()),
            boolean: None,
            error: None,
        },
        CoreCellValue::Bool { value } => CellComputedValue {
            kind: "boolean".to_owned(),
            number: None,
            text: None,
            boolean: Some(*value),
            error: None,
        },
        CoreCellValue::Error { value } => CellComputedValue {
            kind: "error".to_owned(),
            number: None,
            text: None,
            boolean: None,
            error: Some(value.as_str().to_owned()),
        },
    }
}

fn map_proposal(value: &CoreProposal) -> Proposal {
    Proposal {
        id: value.id.clone(),
        agent_id: value.agent_id.clone(),
        note: value.note.clone(),
        edits: value
            .edits
            .iter()
            .map(|edit| ProposedCell {
                sheet: edit.sheet,
                address: CellRef::new(edit.row, edit.col).to_a1(),
                input: edit.input.clone(),
                before: edit.old_text.clone(),
                after: edit.new_text.clone(),
            })
            .collect(),
    }
}

fn horizontal_alignment(value: &str) -> Result<HorizontalAlignment> {
    match value {
        "left" => Ok(HorizontalAlignment::Left),
        "center" => Ok(HorizontalAlignment::Center),
        "right" => Ok(HorizontalAlignment::Right),
        _ => Err(error("horizontalAlignment must be left, center, or right")),
    }
}

fn vertical_alignment(value: &str) -> Result<VerticalAlignment> {
    match value {
        "top" => Ok(VerticalAlignment::Top),
        "middle" => Ok(VerticalAlignment::Middle),
        "bottom" => Ok(VerticalAlignment::Bottom),
        _ => Err(error("verticalAlignment must be top, middle, or bottom")),
    }
}

fn text_wrapping(value: &str) -> Result<TextWrapping> {
    match value {
        "overflow" => Ok(TextWrapping::Overflow),
        "wrap" => Ok(TextWrapping::Wrap),
        "clip" => Ok(TextWrapping::Clip),
        _ => Err(error("textWrapping must be overflow, wrap, or clip")),
    }
}

pub struct OpenTask {
    bytes: Vec<u8>,
    options: OpenWorkbookOptions,
}

impl Task for OpenTask {
    type Output = Workbook;
    type JsValue = XlsxWorkbook;

    fn compute(&mut self) -> Result<Self::Output> {
        let calculation = CalculationOptions {
            now_serial: self.options.now_serial,
        };
        let client_id = self.options.client_id.map(client_id).transpose()?;
        match (
            client_id,
            self.options.read_only.unwrap_or(false),
            self.options.recalculate.unwrap_or(false),
        ) {
            (Some(client_id), _, true) => {
                Workbook::open_collaborative_recalculated(&self.bytes, client_id, calculation)
            }
            (Some(client_id), _, false) => Workbook::open_collaborative(&self.bytes, client_id),
            (None, true, _) => Workbook::open_for_read(&self.bytes),
            (None, false, true) => Workbook::open_recalculated(&self.bytes, calculation),
            (None, false, false) => Workbook::open(&self.bytes),
        }
        .map_err(error)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(XlsxWorkbook {
            inner: Arc::new(Mutex::new(output)),
        })
    }
}

pub struct RenderTask {
    workbook: Arc<Mutex<Workbook>>,
    sheet: SheetId,
    options: RenderOptions,
}

impl Task for RenderTask {
    type Output = betteroffice_xlsx::RenderedPng;
    type JsValue = RenderedSheet;

    fn compute(&mut self) -> Result<Self::Output> {
        lock(&self.workbook)?
            .render_sheet(self.sheet, &self.options)
            .map_err(error)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(RenderedSheet {
            data: output.bytes.into(),
            width: output.width,
            height: output.height,
        })
    }
}

pub struct SaveTask {
    workbook: Arc<Mutex<Workbook>>,
}

impl Task for SaveTask {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        lock(&self.workbook)?.save().map_err(error)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output.into())
    }
}

#[napi(js_name = "Workbook")]
pub struct XlsxWorkbook {
    inner: Arc<Mutex<Workbook>>,
}

#[napi]
impl XlsxWorkbook {
    #[napi(getter)]
    pub fn client_id(&self) -> Result<f64> {
        Ok(lock(&self.inner)?.client_id() as f64)
    }

    #[napi(getter)]
    pub fn collaborative(&self) -> Result<bool> {
        Ok(lock(&self.inner)?.is_collaborative())
    }

    #[napi(getter)]
    pub fn sheet_count(&self) -> Result<u32> {
        Ok(lock(&self.inner)?.sheet_count() as u32)
    }

    #[napi(getter)]
    pub fn active_sheet(&self) -> Result<u32> {
        Ok(lock(&self.inner)?.active_sheet().0)
    }

    #[napi(setter)]
    pub fn set_active_sheet(&self, sheet: u32) -> Result<()> {
        lock(&self.inner)?
            .set_active_sheet(SheetId(sheet))
            .map_err(error)
    }

    #[napi]
    pub fn sheet_info(&self) -> Result<SheetInfo> {
        let value = lock(&self.inner)?.sheet_info().map_err(error)?;
        Ok(SheetInfo {
            ids: value.sheet_ids,
            names: value.sheet_names,
            active_sheet: value.active_sheet.0,
            content_width: f64::from(value.content_width),
            content_height: f64::from(value.content_height),
            frozen_rows: value.frozen_rows,
            frozen_columns: value.frozen_cols,
            initial_scroll_x: f64::from(value.initial_scroll_x),
            initial_scroll_y: f64::from(value.initial_scroll_y),
        })
    }

    #[napi]
    pub fn cell(&self, sheet: u32, address: String) -> Result<CellValue> {
        let value = lock(&self.inner)?
            .cell(SheetId(sheet), cell_ref(&address)?)
            .map_err(error)?;
        Ok(map_cell(address, value))
    }

    #[napi]
    pub fn value(&self, sheet: u32, address: String) -> Result<CellComputedValue> {
        let workbook = lock(&self.inner)?;
        let sheet = workbook.sheet(SheetId(sheet)).map_err(error)?;
        Ok(sheet
            .cell(cell_ref(&address)?)
            .map(|cell| map_computed_value(&cell.value))
            .unwrap_or_else(|| map_computed_value(&CoreCellValue::Empty)))
    }

    #[napi]
    pub fn formula(&self, sheet: u32, address: String) -> Result<Option<String>> {
        let workbook = lock(&self.inner)?;
        let sheet = workbook.sheet(SheetId(sheet)).map_err(error)?;
        Ok(sheet
            .cell(cell_ref(&address)?)
            .and_then(|cell| cell.formula.clone()))
    }

    #[napi]
    pub fn range(&self, sheet: u32, range: String) -> Result<Vec<Vec<CellValue>>> {
        let parsed = cell_range(&range)?;
        let workbook = lock(&self.inner)?;
        let cells = workbook
            .range_cells(SheetId(sheet), parsed)
            .map_err(error)?;
        Ok(cells
            .into_iter()
            .enumerate()
            .map(|(row, values)| {
                values
                    .into_iter()
                    .enumerate()
                    .map(|(column, value)| {
                        let address = CellRef::new(
                            parsed.start.row + row as u32,
                            parsed.start.col + column as u32,
                        )
                        .to_a1();
                        map_cell(address, value)
                    })
                    .collect()
            })
            .collect())
    }

    #[napi]
    pub fn set(
        &self,
        sheet: u32,
        address: String,
        input: String,
        now_serial: Option<f64>,
    ) -> Result<MutationResult> {
        let mut workbook = lock(&self.inner)?;
        let result = workbook
            .edit_cell(
                SheetId(sheet),
                cell_ref(&address)?,
                &input,
                CalculationOptions { now_serial },
            )
            .map_err(error)?;
        Ok(map_mutation(&workbook, result))
    }

    #[napi]
    pub fn set_many(
        &self,
        sheet: u32,
        edits: Vec<CellInput>,
        now_serial: Option<f64>,
    ) -> Result<MutationResult> {
        let edits = edits
            .into_iter()
            .map(|edit| {
                Ok(betteroffice_xlsx::CellInput {
                    cell: cell_ref(&edit.address)?,
                    input: edit.input,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let mut workbook = lock(&self.inner)?;
        let result = workbook
            .edit_cells(SheetId(sheet), &edits, CalculationOptions { now_serial })
            .map_err(error)?;
        Ok(map_mutation(&workbook, result))
    }

    #[napi]
    pub fn recalculate(&self, now_serial: Option<f64>) -> Result<CalculationResult> {
        let mut workbook = lock(&self.inner)?;
        let value = workbook.recalculate_all(CalculationOptions { now_serial });
        Ok(CalculationResult {
            changed: addresses(&workbook, value.changed),
            cycle_cells: addresses(&workbook, value.cycle_cells),
            limited_cells: addresses(&workbook, value.limited_cells),
        })
    }

    #[napi(getter)]
    pub fn history(&self) -> Result<HistoryState> {
        let value = lock(&self.inner)?.history_state();
        Ok(HistoryState {
            can_undo: value.can_undo,
            can_redo: value.can_redo,
            undo_depth: value.undo_depth as u32,
            redo_depth: value.redo_depth as u32,
        })
    }

    #[napi(getter)]
    pub fn can_undo(&self) -> Result<bool> {
        Ok(lock(&self.inner)?.can_undo())
    }

    #[napi(getter)]
    pub fn can_redo(&self) -> Result<bool> {
        Ok(lock(&self.inner)?.can_redo())
    }

    #[napi]
    pub fn undo(&self, now_serial: Option<f64>) -> Result<MutationResult> {
        let mut workbook = lock(&self.inner)?;
        let result = workbook
            .undo(CalculationOptions { now_serial })
            .map_err(error)?;
        Ok(map_mutation(&workbook, result))
    }

    #[napi]
    pub fn redo(&self, now_serial: Option<f64>) -> Result<MutationResult> {
        let mut workbook = lock(&self.inner)?;
        let result = workbook
            .redo(CalculationOptions { now_serial })
            .map_err(error)?;
        Ok(map_mutation(&workbook, result))
    }

    #[napi]
    pub fn encode_state_vector(&self) -> Result<Buffer> {
        Ok(lock(&self.inner)?.encode_state_vector_v1().into())
    }

    #[napi]
    pub fn encode_state_as_update(&self) -> Result<Buffer> {
        Ok(lock(&self.inner)?.encode_state_as_update_v1().into())
    }

    #[napi]
    pub fn encode_diff(&self, state_vector: Buffer) -> Result<Buffer> {
        lock(&self.inner)?
            .encode_diff_v1(&state_vector)
            .map(Buffer::from)
            .map_err(error)
    }

    #[napi]
    pub fn apply_update(
        &self,
        update: Buffer,
        now_serial: Option<f64>,
    ) -> Result<CalculationResult> {
        let value = lock(&self.inner)?
            .apply_update_v1(&update, CalculationOptions { now_serial })
            .map_err(error)?;
        let workbook = lock(&self.inner)?;
        Ok(CalculationResult {
            changed: addresses(&workbook, value.changed),
            cycle_cells: addresses(&workbook, value.cycle_cells),
            limited_cells: addresses(&workbook, value.limited_cells),
        })
    }

    #[napi]
    pub fn propose(&self, proposal: ProposalInput, now_serial: Option<f64>) -> Result<Proposal> {
        let edits = proposal
            .edits
            .into_iter()
            .map(|edit| {
                Ok(CoreProposalEditInput {
                    sheet: SheetId(edit.sheet),
                    cell: cell_ref(&edit.address)?,
                    input: edit.input,
                    number_format: None,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let value = lock(&self.inner)?
            .propose(
                ProposalRequest {
                    agent_id: proposal.agent_id,
                    note: proposal.note,
                    edits,
                },
                CalculationOptions { now_serial },
            )
            .map_err(error)?;
        Ok(map_proposal(&value))
    }

    #[napi(getter)]
    pub fn proposals(&self) -> Result<Vec<Proposal>> {
        Ok(lock(&self.inner)?
            .proposals()
            .iter()
            .map(map_proposal)
            .collect())
    }

    #[napi]
    pub fn accept_proposal(
        &self,
        proposal_id: String,
        force: Option<bool>,
        now_serial: Option<f64>,
    ) -> Result<MutationResult> {
        let mut workbook = lock(&self.inner)?;
        let value = workbook
            .accept_proposal(
                &proposal_id,
                force.unwrap_or(false),
                CalculationOptions { now_serial },
            )
            .map_err(error)?;
        Ok(map_mutation(&workbook, value.mutation))
    }

    #[napi]
    pub fn reject_proposal(&self, proposal_id: String) -> Result<bool> {
        Ok(lock(&self.inner)?.reject_proposal(&proposal_id))
    }

    #[napi]
    pub fn merged_ranges(&self, sheet: u32, range: String) -> Result<Vec<String>> {
        Ok(lock(&self.inner)?
            .merged_ranges(SheetId(sheet), cell_range(&range)?)
            .map_err(error)?
            .into_iter()
            .map(|value| value.to_a1())
            .collect())
    }

    #[napi(getter)]
    pub fn last_calculation(&self) -> Result<CalculationResult> {
        let workbook = lock(&self.inner)?;
        let value = workbook.last_calculation();
        Ok(CalculationResult {
            changed: addresses(&workbook, value.changed.clone()),
            cycle_cells: addresses(&workbook, value.cycle_cells.clone()),
            limited_cells: addresses(&workbook, value.limited_cells.clone()),
        })
    }

    #[napi]
    pub fn set_number_format(
        &self,
        sheet: u32,
        range: String,
        format: String,
        now_serial: Option<f64>,
    ) -> Result<MutationResult> {
        let format = match format.to_ascii_lowercase().as_str() {
            "automatic" => NumberFormatMutation::Automatic,
            "text" => NumberFormatMutation::PlainText,
            "number" => NumberFormatMutation::Number,
            "percent" => NumberFormatMutation::Percent,
            "scientific" => NumberFormatMutation::Scientific,
            "currency" => NumberFormatMutation::Currency,
            "date" => NumberFormatMutation::Date,
            "time" => NumberFormatMutation::Time,
            _ => NumberFormatMutation::Custom { pattern: format },
        };
        let mut workbook = lock(&self.inner)?;
        let value = workbook
            .set_range_number_format(
                SheetId(sheet),
                cell_range(&range)?,
                format,
                CalculationOptions { now_serial },
            )
            .map_err(error)?;
        Ok(map_mutation(&workbook, value))
    }

    #[napi]
    pub fn set_style(
        &self,
        sheet: u32,
        range: String,
        style: StyleInput,
        now_serial: Option<f64>,
    ) -> Result<MutationResult> {
        let patch = StylePatch {
            bold: style.bold,
            italic: style.italic,
            strikethrough: style.strikethrough,
            font_family: style.font_family,
            font_size: style.font_size,
            text_color: style.text_color,
            fill_color: style.fill_color,
            border: None,
            horizontal_alignment: style
                .horizontal_alignment
                .as_deref()
                .map(horizontal_alignment)
                .transpose()?,
            vertical_alignment: style
                .vertical_alignment
                .as_deref()
                .map(vertical_alignment)
                .transpose()?,
            text_wrapping: style
                .text_wrapping
                .as_deref()
                .map(text_wrapping)
                .transpose()?,
            clear: Vec::new(),
        };
        let mut workbook = lock(&self.inner)?;
        let value = workbook
            .patch_range_style(
                SheetId(sheet),
                cell_range(&range)?,
                patch,
                CalculationOptions { now_serial },
            )
            .map_err(error)?;
        Ok(map_mutation(&workbook, value))
    }

    #[napi(ts_return_type = "Promise<RenderedSheet>")]
    pub fn render_sheet(
        &self,
        options: Option<RenderSheetOptions>,
    ) -> Result<AsyncTask<RenderTask>> {
        let options = options.unwrap_or(RenderSheetOptions {
            sheet: None,
            range: None,
            scale: None,
            max_width: None,
            max_height: None,
        });
        let sheet = SheetId(options.sheet.unwrap_or(lock(&self.inner)?.active_sheet().0));
        Ok(AsyncTask::new(RenderTask {
            workbook: self.inner.clone(),
            sheet,
            options: RenderOptions {
                range: options.range.as_deref().map(cell_range).transpose()?,
                scale: options.scale.unwrap_or(1.0) as f32,
                max_width: options.max_width,
                max_height: options.max_height,
            },
        }))
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn save(&self) -> AsyncTask<SaveTask> {
        AsyncTask::new(SaveTask {
            workbook: self.inner.clone(),
        })
    }
}

#[napi(ts_return_type = "Promise<Workbook>")]
pub fn open_workbook(data: Buffer, options: Option<OpenWorkbookOptions>) -> AsyncTask<OpenTask> {
    let options = options.unwrap_or(OpenWorkbookOptions {
        client_id: None,
        read_only: None,
        recalculate: None,
        now_serial: None,
    });
    AsyncTask::new(OpenTask {
        bytes: data.to_vec(),
        options,
    })
}
