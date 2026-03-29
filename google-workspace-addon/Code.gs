/**
 * tsifl — Google Workspace Add-on
 * Works inside Google Sheets, Docs, and Slides.
 * Provides AI-powered assistance via a sidebar connected to the tsifl backend.
 */

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

// ── Homepage Triggers ──────────────────────────────────────────────────────────

function onSheetsHomepage(e) {
  return createSidebarCard_("sheets");
}

function onDocsHomepage(e) {
  return createSidebarCard_("docs");
}

function onSlidesHomepage(e) {
  return createSidebarCard_("slides");
}

function createSidebarCard_(app) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle("tsifl")
    .setSubtitle("AI for Financial Analysts"));

  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextButton()
    .setText("Open tsifl Sidebar")
    .setOnClickAction(CardService.newAction().setFunctionName("openSidebar_" + app)));

  card.addSection(section);
  return card.build();
}

// ── Open Sidebar ───────────────────────────────────────────────────────────────

function openSidebar_sheets() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("tsifl")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openSidebar_docs() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("tsifl")
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}

function openSidebar_slides() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("tsifl")
    .setWidth(360);
  SlidesApp.getUi().showSidebar(html);
}

// ── Also add menu items ────────────────────────────────────────────────────────

function onOpen(e) {
  try {
    SpreadsheetApp.getUi().createMenu("tsifl").addItem("Open tsifl", "openSidebar_sheets").addToUi();
  } catch (_) {}
  try {
    DocumentApp.getUi().createMenu("tsifl").addItem("Open tsifl", "openSidebar_docs").addToUi();
  } catch (_) {}
  try {
    SlidesApp.getUi().createMenu("tsifl").addItem("Open tsifl", "openSidebar_slides").addToUi();
  } catch (_) {}
}

// ── Auth via Supabase REST ─────────────────────────────────────────────────────

function signIn(email, password) {
  var resp = UrlFetchApp.fetch(SUPABASE_URL + "/auth/v1/token?grant_type=password", {
    method: "post",
    contentType: "application/json",
    headers: { "apikey": SUPABASE_ANON_KEY },
    payload: JSON.stringify({ email: email, password: password }),
    muteHttpExceptions: true,
  });
  return JSON.parse(resp.getContentText());
}

function signUp(email, password) {
  var resp = UrlFetchApp.fetch(SUPABASE_URL + "/auth/v1/signup", {
    method: "post",
    contentType: "application/json",
    headers: { "apikey": SUPABASE_ANON_KEY },
    payload: JSON.stringify({ email: email, password: password }),
    muteHttpExceptions: true,
  });
  return JSON.parse(resp.getContentText());
}

function refreshToken(refreshToken) {
  var resp = UrlFetchApp.fetch(SUPABASE_URL + "/auth/v1/token?grant_type=refresh_token", {
    method: "post",
    contentType: "application/json",
    headers: { "apikey": SUPABASE_ANON_KEY },
    payload: JSON.stringify({ refresh_token: refreshToken }),
    muteHttpExceptions: true,
  });
  return JSON.parse(resp.getContentText());
}

// ── Context Capture ────────────────────────────────────────────────────────────

function getSheetsContext() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getActiveRange();
  var dataRange = sheet.getDataRange();

  var context = {
    app: "google_sheets",
    spreadsheet_name: ss.getName(),
    sheet_name: sheet.getName(),
    all_sheets: ss.getSheets().map(function(s) { return s.getName(); }),
    active_cell: range ? range.getA1Notation() : "A1",
    active_value: range ? range.getValue() : "",
    data_range: dataRange.getA1Notation(),
    row_count: dataRange.getNumRows(),
    col_count: dataRange.getNumColumns(),
  };

  // Get visible data (first 50 rows x 26 cols)
  var values = dataRange.getValues();
  var formulas = dataRange.getFormulas();
  context.data = values.slice(0, 50).map(function(r) { return r.slice(0, 26); });
  context.formulas = formulas.slice(0, 50).map(function(r) { return r.slice(0, 26); });

  // Get selection values
  if (range) {
    context.selection_values = range.getValues();
    context.selection_formulas = range.getFormulas();
  }

  return context;
}

function getDocsContext() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var context = {
    app: "google_docs",
    document_name: doc.getName(),
    body_text: body.getText().substring(0, 3000),
    paragraph_count: body.getNumChildren(),
  };

  // Get paragraph details
  var paragraphs = [];
  var numChildren = Math.min(body.getNumChildren(), 50);
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var para = child.asParagraph();
      paragraphs.push({
        text: para.getText().substring(0, 200),
        heading: para.getHeading().toString(),
        alignment: para.getAlignment().toString(),
      });
    } else if (child.getType() === DocumentApp.ElementType.TABLE) {
      var table = child.asTable();
      paragraphs.push({
        type: "table",
        rows: table.getNumRows(),
        cols: table.getRow(0).getNumCells(),
      });
    }
  }
  context.paragraphs = paragraphs;

  // Get cursor position
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    context.cursor_text = element.asText ? element.asText().getText().substring(0, 100) : "";
  }

  // Get selection
  var selection = doc.getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    var selectedText = [];
    for (var j = 0; j < Math.min(elements.length, 5); j++) {
      var el = elements[j].getElement();
      if (el.asText) selectedText.push(el.asText().getText());
    }
    context.selection = selectedText.join("\n").substring(0, 1000);
  }

  return context;
}

function getSlidesContext() {
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();
  var currentSlide = pres.getSelection().getCurrentPage();

  var context = {
    app: "google_slides",
    presentation_name: pres.getName(),
    slide_count: slides.length,
    current_slide_index: 0,
    slides: [],
  };

  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    if (currentSlide && slide.getObjectId() === currentSlide.getObjectId()) {
      context.current_slide_index = i;
    }

    var shapes = slide.getShapes();
    var shapeData = [];
    for (var j = 0; j < Math.min(shapes.length, 10); j++) {
      var shape = shapes[j];
      shapeData.push({
        id: shape.getObjectId(),
        type: shape.getShapeType().toString(),
        text: shape.getText().asString().substring(0, 200),
        left: shape.getLeft(),
        top: shape.getTop(),
        width: shape.getWidth(),
        height: shape.getHeight(),
      });
    }

    context.slides.push({
      index: i,
      id: slide.getObjectId(),
      shapes: shapeData,
      title: shapeData.length > 0 ? shapeData[0].text.substring(0, 80) : "(no title)",
    });
  }

  // Get selected text
  var selection = pres.getSelection();
  if (selection.getSelectionType() === SlidesApp.SelectionType.TEXT) {
    var textRange = selection.getTextRange();
    if (textRange) {
      context.selection = textRange.asString().substring(0, 500);
    }
  }

  return context;
}

// ── Chat Handler ───────────────────────────────────────────────────────────────

function sendChat(userId, message, appType, images) {
  var context;
  if (appType === "sheets") context = getSheetsContext();
  else if (appType === "docs") context = getDocsContext();
  else if (appType === "slides") context = getSlidesContext();
  else context = { app: appType };

  var payload = {
    user_id: userId,
    message: message,
    context: context,
    session_id: appType + "-" + new Date().getTime(),
  };
  if (images && images.length > 0) payload.images = images;

  var resp = UrlFetchApp.fetch(BACKEND_URL + "/chat/", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  var result = JSON.parse(resp.getContentText());

  // Execute actions server-side
  var actions = result.actions && result.actions.length ? result.actions :
    (result.action && result.action.type ? [result.action] : []);

  var actionResults = [];
  for (var i = 0; i < actions.length; i++) {
    try {
      executeAction_(actions[i], appType);
      actionResults.push({ type: actions[i].type, success: true });
    } catch (e) {
      actionResults.push({ type: actions[i].type, success: false, error: e.message });
    }
  }

  result.action_results = actionResults;
  return result;
}

// ── Action Executor ────────────────────────────────────────────────────────────

function executeAction_(action, appType) {
  var type = action.type;
  var p = action.payload;
  if (!p) return;

  if (appType === "sheets") {
    executeSheetsAction_(type, p);
  } else if (appType === "docs") {
    executeDocsAction_(type, p);
  } else if (appType === "slides") {
    executeSlidesAction_(type, p);
  }
}

// ── Sheets Actions ─────────────────────────────────────────────────────────────

function executeSheetsAction_(type, p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = p.sheet ? ss.getSheetByName(p.sheet) : ss.getActiveSheet();
  if (!sheet && p.sheet) {
    sheet = ss.insertSheet(p.sheet);
  }

  switch (type) {
    case "write_cell":
      var cell = sheet.getRange(p.cell);
      if (p.formula) cell.setFormula(p.formula);
      else cell.setValue(p.value);
      if (p.bold) cell.setFontWeight("bold");
      if (p.color) cell.setBackground(p.color);
      if (p.font_color) cell.setFontColor(p.font_color);
      if (p.font_size) cell.setFontSize(p.font_size);
      if (p.number_format) cell.setNumberFormat(p.number_format);
      break;

    case "write_range":
      var range = sheet.getRange(p.range);
      if (p.formulas) {
        // Filter empty formula cells — set only non-empty ones
        range.setFormulas(p.formulas);
      } else if (p.values) {
        range.setValues(p.values);
      }
      if (p.bold) range.setFontWeight("bold");
      if (p.color) range.setBackground(p.color);
      break;

    case "format_range":
      var fmtRange = sheet.getRange(p.range);
      if (p.bold) fmtRange.setFontWeight("bold");
      if (p.italic) fmtRange.setFontStyle("italic");
      if (p.color) fmtRange.setBackground(p.color);
      if (p.font_color) fmtRange.setFontColor(p.font_color);
      if (p.font_size) fmtRange.setFontSize(p.font_size);
      if (p.font_name) fmtRange.setFontFamily(p.font_name);
      if (p.number_format) fmtRange.setNumberFormat(p.number_format);
      if (p.h_align) fmtRange.setHorizontalAlignment(p.h_align);
      if (p.wrap_text) fmtRange.setWrap(true);
      if (p.border) {
        fmtRange.setBorder(true, true, true, true, true, true);
      }
      break;

    case "add_sheet":
      ss.insertSheet(p.name);
      break;

    case "navigate_sheet":
      var targetSheet = ss.getSheetByName(p.sheet);
      if (targetSheet) targetSheet.activate();
      break;

    case "sort_range":
      var sortRange = sheet.getRange(p.range);
      var colNum = p.key_column ? p.key_column.charCodeAt(0) - 64 : 1;
      sortRange.sort({ column: colNum, ascending: p.ascending !== false });
      break;

    case "add_chart":
      var chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType[p.chart_type] || Charts.ChartType.COLUMN)
        .addRange(sheet.getRange(p.data_range))
        .setPosition(p.row || 1, p.col || 5, 0, 0);
      if (p.title) chartBuilder.setOption("title", p.title);
      sheet.insertChart(chartBuilder.build());
      break;

    case "clear_range":
      sheet.getRange(p.range).clear();
      break;

    case "set_number_format":
      sheet.getRange(p.range).setNumberFormat(p.format);
      break;

    case "freeze_panes":
      if (p.rows) sheet.setFrozenRows(p.rows);
      if (p.columns) sheet.setFrozenColumns(p.columns);
      break;

    case "autofit":
      sheet.autoResizeColumns(1, sheet.getLastColumn());
      break;
  }
}

// ── Docs Actions ───────────────────────────────────────────────────────────────

function executeDocsAction_(type, p) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  switch (type) {
    case "insert_text":
      if (p.position === "start") body.insertParagraph(0, p.text);
      else body.appendParagraph(p.text);
      break;

    case "insert_paragraph":
      var para = body.appendParagraph(p.text || "");
      if (p.style) {
        var headingMap = {
          "Heading1": DocumentApp.ParagraphHeading.HEADING1,
          "Heading2": DocumentApp.ParagraphHeading.HEADING2,
          "Heading3": DocumentApp.ParagraphHeading.HEADING3,
          "Title": DocumentApp.ParagraphHeading.TITLE,
          "Subtitle": DocumentApp.ParagraphHeading.SUBTITLE,
          "Normal": DocumentApp.ParagraphHeading.NORMAL,
        };
        if (headingMap[p.style]) para.setHeading(headingMap[p.style]);
      }
      if (p.alignment) {
        var alignMap = {
          "left": DocumentApp.HorizontalAlignment.LEFT,
          "center": DocumentApp.HorizontalAlignment.CENTER,
          "right": DocumentApp.HorizontalAlignment.RIGHT,
          "justify": DocumentApp.HorizontalAlignment.JUSTIFY,
        };
        if (alignMap[p.alignment]) para.setAlignment(alignMap[p.alignment]);
      }
      break;

    case "insert_table":
      if (p.data) {
        body.appendTable(p.data);
      } else {
        body.appendTable(p.rows || 2, p.columns || 2);
      }
      break;

    case "format_text":
      if (p.range_description) {
        var searchResult = body.findText(p.range_description);
        if (searchResult) {
          var el = searchResult.getElement().asText();
          var start = searchResult.getStartOffset();
          var end = searchResult.getEndOffsetInclusive();
          if (p.bold !== undefined) el.setBold(start, end, p.bold);
          if (p.italic !== undefined) el.setItalic(start, end, p.italic);
          if (p.underline !== undefined) el.setUnderline(start, end, p.underline);
          if (p.font_size) el.setFontSize(start, end, p.font_size);
          if (p.font_color) el.setForegroundColor(start, end, p.font_color);
          if (p.font_name) el.setFontFamily(start, end, p.font_name);
        }
      }
      break;

    case "find_and_replace":
      body.replaceText(p.find_text, p.replace_text);
      break;

    case "insert_page_break":
      body.appendPageBreak();
      break;

    case "insert_header":
      var header = doc.getHeader() || doc.addHeader();
      header.appendParagraph(p.text || "");
      break;

    case "insert_footer":
      var footer = doc.getFooter() || doc.addFooter();
      footer.appendParagraph(p.text || "");
      break;

    case "set_page_margins":
      // Not directly available in Apps Script without advanced API
      break;
  }
}

// ── Slides Actions ─────────────────────────────────────────────────────────────

function executeSlidesAction_(type, p) {
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();

  switch (type) {
    case "create_slide":
      var layout = p.layout || "BLANK";
      var layoutMap = {
        "BLANK": SlidesApp.PredefinedLayout.BLANK,
        "Title Slide": SlidesApp.PredefinedLayout.TITLE,
        "Title and Content": SlidesApp.PredefinedLayout.TITLE_AND_BODY,
        "Section Header": SlidesApp.PredefinedLayout.SECTION_HEADER,
        "Title Only": SlidesApp.PredefinedLayout.TITLE_ONLY,
      };
      var newSlide = pres.appendSlide(layoutMap[layout] || SlidesApp.PredefinedLayout.BLANK);
      if (p.title) {
        var titleShape = newSlide.insertTextBox(p.title, 50, 30, 620, 50);
        titleShape.getText().getTextStyle().setFontSize(28).setBold(true);
      }
      if (p.content) {
        newSlide.insertTextBox(p.content, 50, 100, 620, 350);
      }
      break;

    case "add_text_box":
      var idx = p.slide_index || 0;
      if (idx < slides.length) {
        var tb = slides[idx].insertTextBox(p.text || "", p.left || 50, p.top || 50, p.width || 400, p.height || 50);
        if (p.font_size) tb.getText().getTextStyle().setFontSize(p.font_size);
        if (p.bold) tb.getText().getTextStyle().setBold(true);
        if (p.color) tb.getText().getTextStyle().setForegroundColor(p.color);
      }
      break;

    case "add_shape":
      var sIdx = p.slide_index || 0;
      if (sIdx < slides.length) {
        var shapeTypeMap = {
          "Rectangle": SlidesApp.ShapeType.RECTANGLE,
          "RoundedRectangle": SlidesApp.ShapeType.ROUND_RECTANGLE,
          "Oval": SlidesApp.ShapeType.ELLIPSE,
          "Triangle": SlidesApp.ShapeType.TRIANGLE,
        };
        var shape = slides[sIdx].insertShape(
          shapeTypeMap[p.shape_type] || SlidesApp.ShapeType.RECTANGLE,
          p.left || 50, p.top || 50, p.width || 200, p.height || 100
        );
        if (p.fill_color) shape.getFill().setSolidFill(p.fill_color);
        if (p.text) shape.getText().setText(p.text);
      }
      break;

    case "add_table":
      var tIdx = p.slide_index || 0;
      if (tIdx < slides.length && p.data) {
        var table = slides[tIdx].insertTable(p.data.length, p.data[0].length);
        for (var r = 0; r < p.data.length; r++) {
          for (var c = 0; c < p.data[r].length; c++) {
            table.getCell(r, c).getText().setText(String(p.data[r][c] || ""));
          }
        }
      }
      break;

    case "add_image":
      var iIdx = p.slide_index || 0;
      if (iIdx < slides.length && p.image_url) {
        slides[iIdx].insertImage(p.image_url, p.left || 50, p.top || 50, p.width || 400, p.height || 300);
      }
      break;

    case "delete_slide":
      var dIdx = p.slide_index || 0;
      if (dIdx < slides.length) slides[dIdx].remove();
      break;

    case "set_slide_background":
      var bIdx = p.slide_index || 0;
      if (bIdx < slides.length && p.color) {
        slides[bIdx].getBackground().setSolidFill(p.color);
      }
      break;

    case "modify_slide":
      var mIdx = p.slide_index || 0;
      if (mIdx < slides.length && p.changes) {
        var slideShapes = slides[mIdx].getShapes();
        for (var key in p.changes) {
          for (var si = 0; si < slideShapes.length; si++) {
            if (slideShapes[si].getObjectId() === key) {
              if (p.changes[key].text !== undefined) {
                slideShapes[si].getText().setText(p.changes[key].text);
              }
            }
          }
        }
      }
      break;
  }
}
