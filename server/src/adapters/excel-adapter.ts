import { classifyRisk } from "../risk";
import { SharedMemory } from "../shared-memory";
import { AdapterContext, AdapterSuggestion } from "../types";
import { AppAdapter } from "./contract";

/** Map of Excel function names to help text */
const FORMULA_HELP: Record<string, string> = {
  vlookup:
    "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])\nUse FALSE for exact match. Modern alternative: XLOOKUP.",
  xlookup:
    "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])\nMore flexible than VLOOKUP — supports left-lookups and custom defaults.",
  index:
    "=INDEX(array, row_num, [col_num])\nReturns a value from a specific position. Pair with MATCH for dynamic lookups.",
  match:
    "=MATCH(lookup_value, lookup_array, [match_type])\nReturns the position of a value. Use 0 for exact match.",
  sumif:
    "=SUMIF(range, criteria, [sum_range])\nFor multiple criteria: =SUMIFS(sum_range, range1, criteria1, range2, criteria2, ...)",
  countif:
    "=COUNTIF(range, criteria)\nFor multiple criteria: =COUNTIFS(range1, criteria1, range2, criteria2, ...)",
  averageif:
    "=AVERAGEIF(range, criteria, [average_range])\nFor multiple criteria: =AVERAGEIFS(average_range, range1, criteria1, ...)",
  iferror:
    '=IFERROR(value, value_if_error)\nWraps any formula to handle #N/A, #VALUE!, #REF! etc. E.g. =IFERROR(VLOOKUP(...), "Not found")',
  concatenate:
    '=CONCATENATE(text1, text2, ...) or use & operator: =A1 & " " & B1\nModern: =TEXTJOIN(delimiter, ignore_empty, text1, text2, ...)',
  pivot:
    "PivotTable: Select data → Insert → PivotTable. Drag fields to Rows, Columns, Values, Filters. Right-click values to change aggregation (Sum, Count, Average, etc).",
  conditional:
    "Conditional Formatting: Home → Conditional Formatting → New Rule. Use formulas for custom rules: e.g. =A1>100 to highlight cells above 100.",
  array:
    "Dynamic arrays (Excel 365): =SORT(), =FILTER(), =UNIQUE(), =SEQUENCE(). These spill results into adjacent cells automatically.",
  lambda:
    "=LAMBDA(param1, param2, ..., calculation)\nDefine reusable custom functions. Name them via Name Manager for repeated use.",
  let:
    "=LET(name1, value1, name2, value2, ..., calculation)\nAssign intermediate values to names for cleaner, faster formulas.",
};

function findFormulaHelp(message: string): string | null {
  const lower = message.toLowerCase();
  for (const [key, help] of Object.entries(FORMULA_HELP)) {
    if (lower.includes(key)) return help;
  }
  return null;
}

/** Extract file path from message for file operations */
function extractFilePath(message: string): string | null {
  // Match quoted paths or common file patterns
  const quoted = message.match(/["']([^"']+\.(?:xlsx|xls|csv|tsv))["']/i);
  if (quoted) return quoted[1];
  const unquoted = message.match(/(\S+\.(?:xlsx|xls|csv|tsv))/i);
  if (unquoted) return unquoted[1];
  return null;
}

function getExcelSuggestion(context: AdapterContext): AdapterSuggestion {
  const msg = context.message;
  const lower = msg.toLowerCase();

  // Explicit command passthrough
  if (lower.startsWith("cmd:")) {
    return { text: "Command proposal from your request.", command: msg.slice(4).trim() };
  }

  // Formula help (no command needed)
  const formulaHelp = findFormulaHelp(lower);
  if (formulaHelp && (lower.includes("formula") || lower.includes("function") || lower.includes("how to") || lower.includes("help"))) {
    return { text: formulaHelp };
  }

  // CSV conversion
  if (lower.includes("convert") && (lower.includes("csv") || lower.includes("tsv"))) {
    const file = extractFilePath(msg);
    if (file) {
      const outFile = file.replace(/\.xlsx?$/i, ".csv");
      // Use PowerShell + COM automation for real conversion
      const cmd = `powershell -NoProfile -Command "$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $wb = $xl.Workbooks.Open((Resolve-Path '${file}').Path); $wb.SaveAs((Resolve-Path '${file}').Path.Replace('.xlsx','.csv').Replace('.xls','.csv'), 6); $wb.Close($false); $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null; Write-Output 'Converted to ${outFile}'"`;
      return { text: `I'll convert ${file} to CSV format.`, command: cmd };
    }
    return { text: "Provide a file path to convert. Example: convert 'data.xlsx' to csv" };
  }

  // Open file
  if (lower.includes("open") && extractFilePath(msg)) {
    const file = extractFilePath(msg)!;
    return {
      text: `Opening ${file} in Excel.`,
      command: `powershell -NoProfile -Command "Start-Process '${file}'"`,
    };
  }

  // Data analysis — row/column count
  if ((lower.includes("count") || lower.includes("rows") || lower.includes("shape")) && extractFilePath(msg)) {
    const file = extractFilePath(msg)!;
    const cmd = `powershell -NoProfile -Command "$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $wb = $xl.Workbooks.Open((Resolve-Path '${file}').Path); $ws = $wb.Worksheets.Item(1); $rows = $ws.UsedRange.Rows.Count; $cols = $ws.UsedRange.Columns.Count; Write-Output \\"Sheet1: $rows rows x $cols columns\\"; $wb.Close($false); $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null"`;
    return { text: `I'll check the dimensions of ${file}.`, command: cmd };
  }

  // List sheets
  if (lower.includes("sheet") && (lower.includes("list") || lower.includes("show")) && extractFilePath(msg)) {
    const file = extractFilePath(msg)!;
    const cmd = `powershell -NoProfile -Command "$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $wb = $xl.Workbooks.Open((Resolve-Path '${file}').Path); foreach($ws in $wb.Worksheets){ Write-Output $ws.Name }; $wb.Close($false); $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null"`;
    return { text: `Listing sheets in ${file}.`, command: cmd };
  }

  // Create new workbook
  if (lower.includes("create") && (lower.includes("workbook") || lower.includes("spreadsheet"))) {
    const file = extractFilePath(msg) || "new-workbook.xlsx";
    const cmd = `powershell -NoProfile -Command "$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $wb = $xl.Workbooks.Add(); $wb.SaveAs((Join-Path (Get-Location) '${file}')); $wb.Close(); $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null; Write-Output 'Created ${file}'"`;
    return { text: `Creating new workbook: ${file}`, command: cmd };
  }

  // Formula help catch-all
  if (lower.includes("formula") || lower.includes("function")) {
    const knownFunctions = Object.keys(FORMULA_HELP).join(", ");
    return { text: `I can help with Excel formulas. Known functions: ${knownFunctions}. Ask about any specific one.` };
  }

  // Macro/VBA warning
  if (lower.includes("macro") || lower.includes("vba")) {
    return { text: "VBA/Macro support: Alt+F11 opens the editor. I can help write or review macro code. For security, macros from untrusted sources should be reviewed before running." };
  }

  return {
    text: "Excel adapter active. I can: convert files (csv/xlsx), analyze data (row counts, sheet lists), explain formulas (VLOOKUP, SUMIF, INDEX/MATCH, etc.), create workbooks, and run explicit commands with 'cmd:' prefix.",
  };
}

export const excelAdapter: AppAdapter = {
  name: "excel",
  captureContext(context: AdapterContext) {
    return {
      mode: "excel",
      sessionId: context.sessionId,
      cwd: context.cwd,
      hasLastOutput: Boolean(context.lastOutput),
    };
  },
  proposeActions(context: AdapterContext) {
    return getExcelSuggestion(context);
  },
  validateAction(command: string) {
    const lower = command.toLowerCase();
    if (lower.includes("remove-item") || lower.includes("del /") || lower.includes("rm ")) return "blocked";
    if (lower.includes("macro") || lower.includes("vba") || lower.includes("wscript")) return "confirm";
    if (lower.includes("com object") || lower.includes("excel.application")) return "confirm";
    return classifyRisk(command);
  },
  async executeAction() {
    return { supported: false, output: "Execution handled by core policy executor" };
  },
  emitEvents(context: AdapterContext) {
    return [{ type: "excel_context_captured", payload: { sessionId: context.sessionId, cwd: context.cwd } }];
  },
  saveToMemory(context: AdapterContext, memory: SharedMemory) {
    const file = extractFilePath(context.message);
    if (file) {
      memory.set(context.userId, "excel", "last_file", file, context.sessionId);
    }
    memory.set(context.userId, "excel", "last_prompt", context.message.slice(0, 500), context.sessionId);

    // Check if terminal has context we should surface
    const terminalCwd = memory.get(context.userId, "terminal", "last_cwd");
    if (terminalCwd) {
      memory.set(context.userId, "excel", "working_dir_from_terminal", terminalCwd.value, context.sessionId);
    }
  },
};
