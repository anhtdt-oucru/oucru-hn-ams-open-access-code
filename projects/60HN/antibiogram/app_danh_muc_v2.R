# =============================================================================
# ỨNG DỤNG 2: Cập nhật Danh mục tham chiếu
# -----------------------------------------------------------------------------
# Mục đích : Phát hiện các mã chưa được chuẩn hóa trong ba danh mục tham chiếu
#            (Vi sinh vật, Tên khoa, Tên bệnh phẩm), cho phép người dùng điền
#            thông tin còn thiếu qua dropdown preset, rồi lưu đè file gốc.
#
# Gồm 3 tab:
#   Tab 1 — Tên VSV      : cột bắt buộc ma_hoa | ten_vsv | loai_vsv | ten_viet_tat
#   Tab 2 — Tên khoa     : cột bắt buộc ma_hoa | ten_khoa | ten_khoa_nhom
#   Tab 3 — Tên bệnh phẩm: cột bắt buộc ma_hoa | ten_benh_pham | ten_benh_pham_nhom
#
# Quy trình mỗi tab:
#   (1) Tải file WHONET gốc + file danh mục tham chiếu
#   (2) Kiểm tra → hiển thị danh sách mã chưa chuẩn hóa
#   (3) Chọn từ dropdown preset (hoặc "Không tìm thấy" → nhập tay) → Áp dụng
#   (4) Lưu đè file danh mục gốc qua hộp thoại Save-As
#
# Thay đổi so với phiên bản cũ (step 3):
#   - Thay bảng DT editable bằng giao diện dropdown + fallback nhập tay
#   - Preset lấy từ file danh_muc.xlsx (3 sheets: ten_vsv, ten_khoa, ten_benh_pham)
#   - Chọn tên VSV từ preset tự động điền loai_vsv và ten_viet_tat
#   - Chọn "Không tìm thấy" mở ô nhập tay cho từng trường
# =============================================================================

suppressPackageStartupMessages({
  library(shiny)     # framework giao diện web
  library(dplyr)     # xử lý data frame
  library(openxlsx)  # đọc / ghi file Excel
  library(stringi)   # chuẩn hóa chuỗi (bỏ dấu, trim)
  library(tibble)    # tạo tibble nhanh khi thêm dòng mới
  library(here)      # đường dẫn tương đối so với thư mục gốc project
})

# =============================================================================
# HẰNG SỐ: Tùy chọn "Không tìm thấy" trong dropdown
# =============================================================================

NOT_FOUND_LABEL <- "— Không tìm thấy (nhập tay) —"

# =============================================================================
# HÀM TIỆN ÍCH DÙNG CHUNG
# =============================================================================

# Chuẩn hóa khóa so khớp: bỏ dấu tiếng Việt, viết thường, trim khoảng trắng
normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- stringi::stri_trans_general(x, "Latin-ASCII")
  tolower(x)
}

# Mở hộp thoại Save-As của Windows thông qua PowerShell
ps_save_dialog <- function(suggested_name, initial_path = NULL) {
  init_dir <- if (!is.null(initial_path) && nchar(initial_path) > 0) {
    normalizePath(dirname(initial_path), winslash = "/", mustWork = FALSE)
  } else {
    normalizePath("~", winslash = "/")
  }

  ps_script <- sprintf(
    "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null;
     $dlg = New-Object System.Windows.Forms.SaveFileDialog;
     $dlg.Title = 'Luu file: %s';
     $dlg.FileName = '%s';
     $dlg.Filter = 'Excel Files (*.xlsx)|*.xlsx';
     $dlg.InitialDirectory = '%s';
     $dlg.OverwritePrompt = $true;
     if ($dlg.ShowDialog() -eq 'OK') { $dlg.FileName } else { '' }",
    suggested_name, suggested_name, init_dir
  )

  result <- tryCatch(
    system2("powershell",
            args = c("-NoProfile", "-NonInteractive", "-Command", ps_script),
            stdout = TRUE, stderr = FALSE),
    error = function(e) character(0)
  )

  result <- trimws(result[nchar(trimws(result)) > 0])
  if (length(result) == 0 || result[length(result)] == "") NA_character_
  else result[length(result)]
}

# Tự động tìm dòng tiêu đề chứa cột "Organism" trong 20 dòng đầu file WHONET
detect_header_row <- function(path) {
  tryCatch({
    probe <- openxlsx::read.xlsx(path, colNames = FALSE, rows = 1:20)
    hit <- which(apply(probe, 1, function(r) {
      any(stringi::stri_trim_both(as.character(r)) == "Organism")
    }))
    if (length(hit) == 0) NULL else hit[1]
  }, error = function(e) NULL)
}

# Đọc file preset danh_muc.xlsx, trả về list gồm 3 data frame
load_preset <- function(path) {
  tryCatch({
    vsv <- openxlsx::read.xlsx(path, sheet = "ten_vsv")
    khoa <- openxlsx::read.xlsx(path, sheet = "ten_khoa")
    bp   <- openxlsx::read.xlsx(path, sheet = "ten_benh_pham")
    list(vsv = vsv, khoa = khoa, bp = bp)
  }, error = function(e) NULL)
}

# =============================================================================
# CSS DÙNG CHUNG
# =============================================================================

app_css <- HTML("
  .drop-zone {
    border: 2px dashed #2c7be5;
    border-radius: 8px;
    padding: 14px;
    text-align: center;
    cursor: pointer;
    background: #f8fbff;
    margin-bottom: 8px;
    font-weight: 600;
    font-size: 13px;
  }
  .drop-zone:hover { background: #eef6ff; }

  .hint-box {
    background: #f5fbff;
    border-left: 4px solid #2c7be5;
    padding: 8px 12px;
    border-radius: 4px;
    font-size: 13px;
    margin-bottom: 10px;
  }

  .notice-box {
    background: #fff8e1;
    border-left: 4px solid #f0a500;
    padding: 10px 14px;
    border-radius: 4px;
    margin-bottom: 8px;
    font-size: 13px;
    color: #7a5000;
  }

  .status-ok   { color: green;   font-weight: bold; margin-top: 10px; }
  .status-warn { color: #cc6600; font-weight: bold; margin-top: 10px; }

  .badge {
    display: inline-block;
    padding: 2px 8px;
    font-size: 12px;
    font-weight: 700;
    color: #fff;
    background: #17a2b8;
    border-radius: 10px;
    margin-left: 8px;
  }

  details summary {
    cursor: pointer;
    font-weight: 700;
    background: #e2f0fb;
    padding: 8px 12px;
    border-radius: 4px;
  }
  details summary::-webkit-details-marker { display: none; }

  /* Card cho mỗi mã cần chuẩn hóa */
  .map-card {
    border: 1px solid #dee2e6;
    border-radius: 6px;
    padding: 12px 16px;
    margin-bottom: 10px;
    background: #fff;
  }
  .map-card .map-code {
    font-weight: 700;
    font-size: 14px;
    color: #1a1a2e;
    margin-bottom: 8px;
    font-family: monospace;
    background: #f0f4f8;
    padding: 3px 8px;
    border-radius: 4px;
    display: inline-block;
  }
  .map-card .map-fields {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    align-items: flex-start;
  }
  .map-card .map-field {
    flex: 1 1 200px;
  }
  .map-card label {
    font-size: 12px;
    font-weight: 600;
    color: #555;
    margin-bottom: 2px;
    display: block;
  }
  .freetext-group {
    margin-top: 6px;
    padding: 8px;
    background: #fffbf0;
    border: 1px dashed #f0a500;
    border-radius: 4px;
  }
  .freetext-group label {
    font-size: 12px;
    color: #7a5000;
  }
")

# =============================================================================
# HÀM TẠO UI MAPPING CHO MỖI TAB
# Tham số:
#   ns_prefix : tiền tố namespace (vd "v", "k", "b")
#   missing_df: data frame các mã chưa chuẩn hóa
#   preset    : list preset từ danh_muc.xlsx
#   tab_type  : "vsv" | "khoa" | "bp"
# =============================================================================

build_mapping_ui <- function(ns_prefix, missing_df, preset, tab_type) {
  if (is.null(missing_df) || nrow(missing_df) == 0) {
    return(div(
      class = "hint-box",
      "✅ Không có mã nào cần chuẩn hóa."
    ))
  }

  cards <- lapply(seq_len(nrow(missing_df)), function(i) {
    code <- missing_df$ma_hoa[i]
    id   <- paste0(ns_prefix, "_row_", i)

    if (tab_type == "vsv") {
      # Dropdown chọn tên VSV — tự động điền loai_vsv và ten_viet_tat
      vsv_choices <- c("", NOT_FOUND_LABEL, preset$vsv$ten_vsv)
      current_vsv <- missing_df$ten_vsv[i]

      div(
        class = "map-card",
        div(class = "map-code", code),
        div(
          class = "map-fields",
          div(
            class = "map-field",
            tags$label("Tên vi sinh vật (chọn từ danh sách)"),
            selectInput(
              inputId  = paste0(id, "_sel"),
              label    = NULL,
              choices  = vsv_choices,
              selected = if (current_vsv %in% vsv_choices) current_vsv else "",
              width    = "100%"
            )
          )
        ),
        # Vùng nhập tay — chỉ hiện khi chọn "Không tìm thấy"
        conditionalPanel(
          condition = sprintf(
            "input['%s'] == '%s'",
            paste0(id, "_sel"), NOT_FOUND_LABEL
          ),
          div(
            class = "freetext-group",
            fluidRow(
              column(4,
                     tags$label("Tên vi sinh vật"),
                     textInput(paste0(id, "_ten_vsv"),      NULL,
                               value = missing_df$ten_vsv[i],      width = "100%")),
              column(4,
                     tags$label("Loại vi sinh vật"),
                     textInput(paste0(id, "_loai_vsv"),     NULL,
                               value = missing_df$loai_vsv[i],     width = "100%")),
              column(4,
                     tags$label("Tên viết tắt"),
                     textInput(paste0(id, "_ten_viet_tat"), NULL,
                               value = missing_df$ten_viet_tat[i], width = "100%"))
            )
          )
        )
      )

    } else if (tab_type == "khoa") {
      # Dropdown chọn nhóm khoa
      khoa_choices  <- c("", NOT_FOUND_LABEL, preset$khoa$nhom_khoa)
      current_nhom  <- missing_df$ten_khoa_nhom[i]

      div(
        class = "map-card",
        div(class = "map-code", code),
        div(
          class = "map-fields",
          div(
            class = "map-field",
            tags$label("Tên khoa (nhập tay)"),
            textInput(paste0(id, "_ten_khoa"), NULL,
                      value = missing_df$ten_khoa[i], width = "100%")
          ),
          div(
            class = "map-field",
            tags$label("Nhóm khoa (chọn từ danh sách)"),
            selectInput(
              inputId  = paste0(id, "_nhom_sel"),
              label    = NULL,
              choices  = khoa_choices,
              selected = if (current_nhom %in% khoa_choices) current_nhom else "",
              width    = "100%"
            )
          )
        ),
        conditionalPanel(
          condition = sprintf(
            "input['%s'] == '%s'",
            paste0(id, "_nhom_sel"), NOT_FOUND_LABEL
          ),
          div(
            class = "freetext-group",
            tags$label("Nhóm khoa (nhập tay)"),
            textInput(paste0(id, "_nhom_free"), NULL,
                      value = "", width = "100%")
          )
        )
      )

    } else { # tab_type == "bp"
      # Dropdown chọn nhóm bệnh phẩm
      bp_choices   <- c("", NOT_FOUND_LABEL, preset$bp$nhom_benh_pham)
      current_nhom <- missing_df$ten_benh_pham_nhom[i]

      div(
        class = "map-card",
        div(class = "map-code", code),
        div(
          class = "map-fields",
          div(
            class = "map-field",
            tags$label("Tên bệnh phẩm (nhập tay)"),
            textInput(paste0(id, "_ten_bp"), NULL,
                      value = missing_df$ten_benh_pham[i], width = "100%")
          ),
          div(
            class = "map-field",
            tags$label("Nhóm bệnh phẩm (chọn từ danh sách)"),
            selectInput(
              inputId  = paste0(id, "_nhom_sel"),
              label    = NULL,
              choices  = bp_choices,
              selected = if (current_nhom %in% bp_choices) current_nhom else "",
              width    = "100%"
            )
          )
        ),
        conditionalPanel(
          condition = sprintf(
            "input['%s'] == '%s'",
            paste0(id, "_nhom_sel"), NOT_FOUND_LABEL
          ),
          div(
            class = "freetext-group",
            tags$label("Nhóm bệnh phẩm (nhập tay)"),
            textInput(paste0(id, "_nhom_free"), NULL,
                      value = "", width = "100%")
          )
        )
      )
    }
  })

  tagList(cards)
}

# =============================================================================
# HÀM THU GOM GIÁ TRỊ TỪ CÁC INPUT ĐỘNG (mapping cards)
# Đọc toàn bộ input hiện tại, trả về data frame với các cột đã điền
# =============================================================================

collect_vsv_inputs <- function(input, missing_df, preset_vsv) {
  n <- nrow(missing_df)
  if (n == 0) return(missing_df)

  result <- missing_df

  for (i in seq_len(n)) {
    id  <- paste0("v_row_", i)
    sel <- input[[paste0(id, "_sel")]]

    if (!is.null(sel) && sel != "" && sel != NOT_FOUND_LABEL) {
      # Lấy từ preset: tự điền loai_vsv và ten_viet_tat
      row <- preset_vsv %>% filter(ten_vsv == sel)
      if (nrow(row) > 0) {
        result$ten_vsv[i]      <- row$ten_vsv[1]
        result$loai_vsv[i]     <- row$loai_vsv[1]
        result$ten_viet_tat[i] <- row$ten_viet_tat[1]
      }
    } else if (!is.null(sel) && sel == NOT_FOUND_LABEL) {
      # Lấy giá trị nhập tay
      result$ten_vsv[i]      <- input[[paste0(id, "_ten_vsv")]]      %||% ""
      result$loai_vsv[i]     <- input[[paste0(id, "_loai_vsv")]]     %||% ""
      result$ten_viet_tat[i] <- input[[paste0(id, "_ten_viet_tat")]] %||% ""
    }
  }

  result
}

collect_khoa_inputs <- function(input, missing_df) {
  n <- nrow(missing_df)
  if (n == 0) return(missing_df)

  result <- missing_df

  for (i in seq_len(n)) {
    id       <- paste0("k_row_", i)
    ten_khoa <- input[[paste0(id, "_ten_khoa")]] %||% ""
    nhom_sel <- input[[paste0(id, "_nhom_sel")]]

    nhom_val <- if (!is.null(nhom_sel) && nhom_sel == NOT_FOUND_LABEL) {
      input[[paste0(id, "_nhom_free")]] %||% ""
    } else if (!is.null(nhom_sel) && nhom_sel != "") {
      nhom_sel
    } else {
      ""
    }

    result$ten_khoa[i]      <- ten_khoa
    result$ten_khoa_nhom[i] <- nhom_val
  }

  result
}

collect_bp_inputs <- function(input, missing_df) {
  n <- nrow(missing_df)
  if (n == 0) return(missing_df)

  result <- missing_df

  for (i in seq_len(n)) {
    id      <- paste0("b_row_", i)
    ten_bp  <- input[[paste0(id, "_ten_bp")]] %||% ""
    nhom_sel <- input[[paste0(id, "_nhom_sel")]]

    nhom_val <- if (!is.null(nhom_sel) && nhom_sel == NOT_FOUND_LABEL) {
      input[[paste0(id, "_nhom_free")]] %||% ""
    } else if (!is.null(nhom_sel) && nhom_sel != "") {
      nhom_sel
    } else {
      ""
    }

    result$ten_benh_pham[i]      <- ten_bp
    result$ten_benh_pham_nhom[i] <- nhom_val
  }

  result
}

# Toán tử tiện ích: thay thế NULL bằng giá trị mặc định
`%||%` <- function(a, b) if (is.null(a)) b else a

# =============================================================================
# GIAO DIỆN NGƯỜI DÙNG (UI)
# =============================================================================

ui <- fluidPage(

  titlePanel("Cập nhật Danh mục tham chiếu"),
  tags$style(app_css),

  tabsetPanel(
    id = "main_tabs",

    # =========================================================================
    # TAB 1 — TÊN VI SINH VẬT
    # =========================================================================
    tabPanel(
      "Tên VSV",
      br(),
      sidebarLayout(
        sidebarPanel(

          uiOutput("preset_status"),

          hr(),
          h4("1) Chọn file đầu vào"),

          div(class = "drop-zone",
              onclick = "document.getElementById('v_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("v_raw_files", label = NULL, multiple = TRUE, accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('v_ref_file').click()",
              "📘 Chọn Danh mục VSV (.xlsx)"),
          fileInput("v_ref_file", label = NULL, multiple = FALSE, accept = ".xlsx"),

          hr(),
          h4("2) Kiểm tra"),
          actionButton("v_check_missing", "🔍 Kiểm tra mã chưa chuẩn hóa",
                       class = "btn-primary", width = "100%"),

          hr(),
          h4("3) Áp dụng cập nhật"),
          div(class = "hint-box",
              "Điền thông tin cho các mã còn thiếu ở bảng bên phải, sau đó nhấn Áp dụng."),
          actionButton("v_apply_updates", "✔ Áp dụng cập nhật",
                       class = "btn-success", width = "100%"),

          br(), br(),
          h4("4) Lưu file"),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu đè:</b> Ghi đè lên file Danh mục VSV gốc.
                   Chọn <b>đúng tên file gốc</b> trong hộp thoại.")),
          actionButton("v_save_as_btn", "💾 Lưu đè file Danh mục VSV...",
                       class = "btn-warning", width = "100%"),

          br(), br(),
          uiOutput("v_save_status")
        ),

        mainPanel(
          h4("Tổng quan"),
          verbatimTextOutput("v_summary"),
          hr(),
          h4("Mã VSV cần chuẩn hóa"),
          div(class = "hint-box", HTML(
            "Với mỗi mã, chọn tên VSV từ danh sách preset.<br>
             Nếu không tìm thấy, chọn <b>\"Không tìm thấy\"</b> để nhập tay."
          )),
          uiOutput("v_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DT::DTOutput("v_ref_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 2 — TÊN KHOA
    # =========================================================================
    tabPanel(
      "Tên khoa",
      br(),
      sidebarLayout(
        sidebarPanel(

          h4("1) Chọn file đầu vào"),

          div(class = "drop-zone",
              onclick = "document.getElementById('k_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("k_raw_files", label = NULL, multiple = TRUE, accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('k_ref_file').click()",
              "📘 Chọn Danh mục Tên khoa (.xlsx)"),
          fileInput("k_ref_file", label = NULL, multiple = FALSE, accept = ".xlsx"),

          hr(),
          h4("2) Kiểm tra"),
          actionButton("k_check_missing", "🔍 Kiểm tra Tên khoa chưa chuẩn hóa",
                       class = "btn-primary", width = "100%"),

          hr(),
          h4("3) Áp dụng cập nhật"),
          div(class = "hint-box",
              "Điền thông tin cho các khoa còn thiếu ở bảng bên phải, sau đó nhấn Áp dụng."),
          actionButton("k_apply_updates", "✔ Áp dụng cập nhật",
                       class = "btn-success", width = "100%"),

          br(), br(),
          h4("4) Lưu file"),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu đè:</b> Ghi đè lên file Danh mục Tên khoa gốc.
                   Chọn <b>đúng tên file gốc</b> trong hộp thoại.")),
          actionButton("k_save_as_btn", "💾 Lưu đè file Danh mục Tên khoa...",
                       class = "btn-warning", width = "100%"),

          br(), br(),
          uiOutput("k_save_status")
        ),

        mainPanel(
          h4("Tổng quan"),
          verbatimTextOutput("k_summary"),
          hr(),
          h4("Tên khoa cần chuẩn hóa"),
          div(class = "hint-box", HTML(
            "Nhập <b>Tên khoa</b> và chọn <b>Nhóm khoa</b> từ danh sách preset.<br>
             Nếu không tìm thấy nhóm phù hợp, chọn <b>\"Không tìm thấy\"</b> để nhập tay."
          )),
          uiOutput("k_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DT::DTOutput("k_ref_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 3 — TÊN BỆNH PHẨM
    # =========================================================================
    tabPanel(
      "Tên bệnh phẩm",
      br(),
      sidebarLayout(
        sidebarPanel(

          h4("1) Chọn file đầu vào"),

          div(class = "drop-zone",
              onclick = "document.getElementById('b_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("b_raw_files", label = NULL, multiple = TRUE, accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('b_ref_file').click()",
              "📘 Chọn Danh mục Tên bệnh phẩm (.xlsx)"),
          fileInput("b_ref_file", label = NULL, multiple = FALSE, accept = ".xlsx"),

          hr(),
          h4("2) Kiểm tra"),
          actionButton("b_check_missing",
                       "🔍 Kiểm tra Tên bệnh phẩm chưa chuẩn hóa",
                       class = "btn-primary", width = "100%"),

          hr(),
          h4("3) Áp dụng cập nhật"),
          div(class = "hint-box",
              "Điền thông tin cho các bệnh phẩm còn thiếu ở bảng bên phải, sau đó nhấn Áp dụng."),
          actionButton("b_apply_updates", "✔ Áp dụng cập nhật",
                       class = "btn-success", width = "100%"),

          br(), br(),
          h4("4) Lưu file"),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu đè:</b> Ghi đè lên file Danh mục Tên bệnh phẩm gốc.
                   Chọn <b>đúng tên file gốc</b> trong hộp thoại.")),
          actionButton("b_save_as_btn",
                       "💾 Lưu đè file Danh mục Tên bệnh phẩm...",
                       class = "btn-warning", width = "100%"),

          br(), br(),
          uiOutput("b_save_status")
        ),

        mainPanel(
          h4("Tổng quan"),
          verbatimTextOutput("b_summary"),
          hr(),
          h4("Tên bệnh phẩm cần chuẩn hóa"),
          div(class = "hint-box", HTML(
            "Nhập <b>Tên bệnh phẩm</b> và chọn <b>Nhóm bệnh phẩm</b> từ danh sách preset.<br>
             Nếu không tìm thấy nhóm phù hợp, chọn <b>\"Không tìm thấy\"</b> để nhập tay."
          )),
          uiOutput("b_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DT::DTOutput("b_ref_preview")
        )
      )
    )
  )
)

# =============================================================================
# LOGIC XỬ LÝ (SERVER)
# =============================================================================

server <- function(input, output, session) {

  # ---------------------------------------------------------------------------
  # PRESET — dùng chung cho tất cả tab
  # ---------------------------------------------------------------------------

  # Đọc preset một lần khi app khởi động từ đường dẫn tương đối
  # File danh_muc.xlsx phải nằm cùng thư mục gốc project (cạnh app_danh_muc.R)
  preset_data <- local({
    path <- here::here("danh_muc.xlsx")
    p    <- load_preset(path)
    if (is.null(p)) stop(sprintf("Không tìm thấy hoặc không đọc được file preset: %s", path))
    p
  })

  output$preset_status <- renderUI({
    div(class = "status-ok",
        sprintf("✅ Preset đã tải: %d VSV | %d nhóm khoa | %d nhóm bệnh phẩm",
                nrow(preset_data$vsv),
                nrow(preset_data$khoa),
                nrow(preset_data$bp)))
  })

  # ===========================================================================
  # TAB 1 — TÊN VI SINH VẬT
  # ===========================================================================

  v_missing_df   <- reactiveVal(NULL)
  v_updated_ref  <- reactiveVal(NULL)
  v_save_message <- reactiveVal("")

  v_ref_df <- reactive({
    req(input$v_ref_file)
    df <- tryCatch(openxlsx::read.xlsx(input$v_ref_file$datapath),
                   error = function(e) NULL)
    validate(need(!is.null(df), "Không đọc được file Danh mục VSV."))
    names(df) <- stringi::stri_trim_both(names(df))
    validate(
      need("ma_hoa"       %in% names(df), "Thiếu cột: ma_hoa"),
      need("ten_vsv"      %in% names(df), "Thiếu cột: ten_vsv"),
      need("loai_vsv"     %in% names(df), "Thiếu cột: loai_vsv"),
      need("ten_viet_tat" %in% names(df), "Thiếu cột: ten_viet_tat")
    )
    df %>% transmute(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat,
                     .key = normalize_key(ma_hoa))
  })

  v_raw_df <- reactive({
    req(input$v_raw_files)
    files <- input$v_raw_files$datapath

    all_org <- lapply(seq_along(files), function(i) {
      f          <- files[i]
      header_row <- detect_header_row(f)
      if (is.null(header_row)) return(NULL)
      df <- tryCatch(openxlsx::read.xlsx(f, startRow = header_row),
                     error = function(e) NULL)
      if (is.null(df)) return(NULL)
      names(df) <- stringi::stri_trim_both(names(df))
      if (!"Organism" %in% names(df)) return(NULL)
      df %>% transmute(ma_hoa = as.character(Organism))
    })

    all_org <- all_org[!sapply(all_org, is.null)]
    validate(need(length(all_org) > 0, "Không file nào có cột Organism."))

    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })

  output$v_summary <- renderText({
    joined <- v_raw_df() %>% left_join(v_ref_df(), by = ".key")
    total  <- nrow(joined)
    done   <- sum(
      !is.na(joined$ten_vsv)      & trimws(joined$ten_vsv)      != "" &
        !is.na(joined$loai_vsv)     & trimws(joined$loai_vsv)     != "" &
        !is.na(joined$ten_viet_tat) & trimws(joined$ten_viet_tat) != ""
    )
    paste0(
      "Tổng số mã VSV trong WHONET    : ", total, "\n",
      "Đã chuẩn hóa đầy đủ            : ", done,  "\n",
      "Chưa chuẩn hóa / thiếu thông tin: ", total - done
    )
  })

  observeEvent(input$v_check_missing, {
    joined <- v_raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(v_ref_df(), by = ".key")

    miss <- joined %>%
      filter(
        is.na(ten_vsv) | trimws(ten_vsv) == "" |
          is.na(loai_vsv) | trimws(loai_vsv) == "" |
          is.na(ten_viet_tat) | trimws(ten_viet_tat) == ""
      ) %>%
      transmute(
        ma_hoa       = ma_hoa_raw,
        ten_vsv      = ifelse(is.na(ten_vsv),      "", ten_vsv),
        loai_vsv     = ifelse(is.na(loai_vsv),     "", loai_vsv),
        ten_viet_tat = ifelse(is.na(ten_viet_tat), "", ten_viet_tat)
      ) %>%
      arrange(ma_hoa)

    v_missing_df(miss)
    v_updated_ref(NULL)
  })

  output$v_missing_folded_ui <- renderUI({
    df <- v_missing_df()
    if (is.null(df)) return(NULL)

    p <- preset_data

    tags$details(
      open = TRUE,
      tags$summary(
        "Danh sách cần chuẩn hóa",
        span(class = "badge", nrow(df))
      ),
      br(),
      build_mapping_ui("v", df, p, "vsv")
    )
  })

  v_apply_updates_core <- function() {
    work <- v_updated_ref()
    if (is.null(work)) work <- v_ref_df()

    miss <- v_missing_df()
    if (is.null(miss) || nrow(miss) == 0) return(work)

    p <- preset_data
    to_apply  <- collect_vsv_inputs(input, miss, p$vsv) %>%
      filter(trimws(ten_vsv) != "")

    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)

      if (length(idx) > 0) {
        work$ten_vsv[idx]      <- to_apply$ten_vsv[i]
        work$loai_vsv[idx]     <- to_apply$loai_vsv[i]
        work$ten_viet_tat[idx] <- to_apply$ten_viet_tat[i]
      } else {
        work <- bind_rows(work, tibble(
          ma_hoa       = to_apply$ma_hoa[i],
          ten_vsv      = to_apply$ten_vsv[i],
          loai_vsv     = to_apply$loai_vsv[i],
          ten_viet_tat = to_apply$ten_viet_tat[i],
          .key         = key
        ))
      }
    }

    v_updated_ref(work)
    work
  }

  observeEvent(input$v_apply_updates, {
    v_apply_updates_core()
    v_save_message("Đã áp dụng cập nhật vào dữ liệu tạm. Nhấn Lưu đè để ghi file.")
  })

  observeEvent(input$v_save_as_btn, {
    if (is.null(input$v_ref_file)) {
      v_save_message("⚠️ Chưa chọn file Danh mục VSV.")
      return()
    }
    out <- tryCatch(v_apply_updates_core(), error = function(e) {
      v_save_message(paste("Lỗi khi chuẩn bị dữ liệu:", e$message))
      NULL
    })
    if (is.null(out)) return()

    ref_filename <- input$v_ref_file$name
    path <- ps_save_dialog(ref_filename, input$v_ref_file$datapath)

    if (is.na(path) || trimws(path) == "") {
      v_save_message("ℹ️ Đã hủy: không chọn đường dẫn lưu.")
      return()
    }
    if (!identical(basename(path), ref_filename)) {
      v_save_message(paste0("⚠️ Tên file không khớp với file gốc (", ref_filename, ")."))
      return()
    }

    tryCatch({
      out_clean <- out %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat)
      openxlsx::write.xlsx(out_clean, file = path, overwrite = TRUE)
      v_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      v_save_message(paste0("❌ Lỗi khi ghi file: ", e$message))
    })
  })

  output$v_save_status <- renderUI({
    msg <- v_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  output$v_ref_preview <- DT::renderDT({
    req(v_updated_ref())
    DT::datatable(
      v_updated_ref() %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat),
      rownames = FALSE
    )
  })

  # ===========================================================================
  # TAB 2 — TÊN KHOA
  # ===========================================================================

  k_missing_df   <- reactiveVal(NULL)
  k_updated_ref  <- reactiveVal(NULL)
  k_save_message <- reactiveVal("")

  k_ref_df <- reactive({
    req(input$k_ref_file)
    df <- tryCatch(openxlsx::read.xlsx(input$k_ref_file$datapath),
                   error = function(e) NULL)
    validate(need(!is.null(df), "Không đọc được file Danh mục Tên khoa."))
    names(df) <- stringi::stri_trim_both(names(df))
    validate(
      need("ma_hoa"        %in% names(df), "Thiếu cột: ma_hoa"),
      need("ten_khoa"      %in% names(df), "Thiếu cột: ten_khoa"),
      need("ten_khoa_nhom" %in% names(df), "Thiếu cột: ten_khoa_nhom")
    )
    df %>% transmute(ma_hoa, ten_khoa, ten_khoa_nhom,
                     .key = normalize_key(ma_hoa))
  })

  k_raw_df <- reactive({
    req(input$k_raw_files)
    files <- input$k_raw_files$datapath

    all_org <- lapply(seq_along(files), function(i) {
      f          <- files[i]
      header_row <- detect_header_row(f)
      if (is.null(header_row)) return(NULL)
      df <- tryCatch(openxlsx::read.xlsx(f, startRow = header_row),
                     error = function(e) NULL)
      if (is.null(df)) return(NULL)
      names(df) <- stringi::stri_trim_both(names(df))
      if (!"Location" %in% names(df)) return(NULL)
      df %>% transmute(ma_hoa = as.character(Location))
    })

    all_org <- all_org[!sapply(all_org, is.null)]
    validate(need(length(all_org) > 0, "Không file nào có cột Location."))

    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })

  output$k_summary <- renderText({
    joined <- k_raw_df() %>% left_join(k_ref_df(), by = ".key")
    total  <- nrow(joined)
    done   <- sum(
      !is.na(joined$ten_khoa)      & trimws(joined$ten_khoa)      != "" &
        !is.na(joined$ten_khoa_nhom) & trimws(joined$ten_khoa_nhom) != ""
    )
    paste0(
      "Tổng số tên khoa trong WHONET   : ", total, "\n",
      "Đã chuẩn hóa đầy đủ             : ", done,  "\n",
      "Chưa chuẩn hóa / thiếu thông tin: ", total - done
    )
  })

  observeEvent(input$k_check_missing, {
    joined <- k_raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(k_ref_df(), by = ".key")

    miss <- joined %>%
      filter(
        is.na(ten_khoa) | trimws(ten_khoa) == "" |
          is.na(ten_khoa_nhom) | trimws(ten_khoa_nhom) == ""
      ) %>%
      transmute(
        ma_hoa        = ma_hoa_raw,
        ten_khoa      = ifelse(is.na(ten_khoa),      "", ten_khoa),
        ten_khoa_nhom = ifelse(is.na(ten_khoa_nhom), "", ten_khoa_nhom)
      ) %>%
      arrange(ma_hoa)

    k_missing_df(miss)
    k_updated_ref(NULL)
  })

  output$k_missing_folded_ui <- renderUI({
    df <- k_missing_df()
    if (is.null(df)) return(NULL)

    p <- preset_data

    tags$details(
      open = TRUE,
      tags$summary(
        "Danh sách cần chuẩn hóa",
        span(class = "badge", nrow(df))
      ),
      br(),
      build_mapping_ui("k", df, p, "khoa")
    )
  })

  k_apply_updates_core <- function() {
    work <- k_updated_ref()
    if (is.null(work)) work <- k_ref_df()

    miss <- k_missing_df()
    if (is.null(miss) || nrow(miss) == 0) return(work)

    to_apply <- collect_khoa_inputs(input, miss) %>%
      filter(trimws(ten_khoa) != "" | trimws(ten_khoa_nhom) != "")

    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)

      if (length(idx) > 0) {
        work$ten_khoa[idx]      <- to_apply$ten_khoa[i]
        work$ten_khoa_nhom[idx] <- to_apply$ten_khoa_nhom[i]
      } else {
        work <- bind_rows(work, tibble(
          ma_hoa        = to_apply$ma_hoa[i],
          ten_khoa      = to_apply$ten_khoa[i],
          ten_khoa_nhom = to_apply$ten_khoa_nhom[i],
          .key          = key
        ))
      }
    }

    k_updated_ref(work)
    work
  }

  observeEvent(input$k_apply_updates, {
    k_apply_updates_core()
    k_save_message("Đã áp dụng cập nhật vào dữ liệu tạm. Nhấn Lưu đè để ghi file.")
  })

  observeEvent(input$k_save_as_btn, {
    if (is.null(input$k_ref_file)) {
      k_save_message("⚠️ Chưa chọn file Danh mục Tên khoa.")
      return()
    }
    out <- tryCatch(k_apply_updates_core(), error = function(e) {
      k_save_message(paste("Lỗi:", e$message))
      NULL
    })
    if (is.null(out)) return()

    ref_filename <- input$k_ref_file$name
    path <- ps_save_dialog(ref_filename, input$k_ref_file$datapath)

    if (is.na(path) || trimws(path) == "") {
      k_save_message("ℹ️ Đã hủy: không chọn đường dẫn lưu.")
      return()
    }
    if (!identical(basename(path), ref_filename)) {
      k_save_message(paste0("⚠️ Tên file không khớp với file gốc (", ref_filename, ")."))
      return()
    }

    tryCatch({
      out_clean <- out %>% select(ma_hoa, ten_khoa, ten_khoa_nhom)
      openxlsx::write.xlsx(out_clean, file = path, overwrite = TRUE)
      k_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      k_save_message(paste0("❌ Lỗi khi ghi file: ", e$message))
    })
  })

  output$k_save_status <- renderUI({
    msg <- k_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  output$k_ref_preview <- DT::renderDT({
    req(k_updated_ref())
    DT::datatable(
      k_updated_ref() %>% select(ma_hoa, ten_khoa, ten_khoa_nhom),
      rownames = FALSE
    )
  })

  # ===========================================================================
  # TAB 3 — TÊN BỆNH PHẨM
  # ===========================================================================

  b_missing_df   <- reactiveVal(NULL)
  b_updated_ref  <- reactiveVal(NULL)
  b_save_message <- reactiveVal("")

  b_ref_df <- reactive({
    req(input$b_ref_file)
    df <- tryCatch(openxlsx::read.xlsx(input$b_ref_file$datapath),
                   error = function(e) NULL)
    validate(need(!is.null(df), "Không đọc được file Danh mục Tên bệnh phẩm."))
    names(df) <- stringi::stri_trim_both(names(df))
    validate(
      need("ma_hoa"             %in% names(df), "Thiếu cột: ma_hoa"),
      need("ten_benh_pham"      %in% names(df), "Thiếu cột: ten_benh_pham"),
      need("ten_benh_pham_nhom" %in% names(df), "Thiếu cột: ten_benh_pham_nhom")
    )
    df %>% transmute(ma_hoa, ten_benh_pham, ten_benh_pham_nhom,
                     .key = normalize_key(ma_hoa))
  })

  b_raw_df <- reactive({
    req(input$b_raw_files)
    files <- input$b_raw_files$datapath

    all_org <- lapply(seq_along(files), function(i) {
      f          <- files[i]
      header_row <- detect_header_row(f)
      if (is.null(header_row)) return(NULL)
      df <- tryCatch(openxlsx::read.xlsx(f, startRow = header_row),
                     error = function(e) NULL)
      if (is.null(df)) return(NULL)
      names(df) <- stringi::stri_trim_both(names(df))
      if (!"Specimen type" %in% names(df)) return(NULL)
      df %>% transmute(ma_hoa = as.character(`Specimen type`))
    })

    all_org <- all_org[!sapply(all_org, is.null)]
    validate(need(length(all_org) > 0, "Không file nào có cột 'Specimen type'."))

    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })

  output$b_summary <- renderText({
    joined <- b_raw_df() %>% left_join(b_ref_df(), by = ".key")
    total  <- nrow(joined)
    done   <- sum(
      !is.na(joined$ten_benh_pham)      & trimws(joined$ten_benh_pham)      != "" &
        !is.na(joined$ten_benh_pham_nhom) & trimws(joined$ten_benh_pham_nhom) != ""
    )
    paste0(
      "Tổng số tên bệnh phẩm trong WHONET: ", total, "\n",
      "Đã chuẩn hóa đầy đủ               : ", done,  "\n",
      "Chưa chuẩn hóa / thiếu thông tin  : ", total - done
    )
  })

  observeEvent(input$b_check_missing, {
    joined <- b_raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(b_ref_df(), by = ".key")

    miss <- joined %>%
      filter(
        is.na(ten_benh_pham) | trimws(ten_benh_pham) == "" |
          is.na(ten_benh_pham_nhom) | trimws(ten_benh_pham_nhom) == ""
      ) %>%
      transmute(
        ma_hoa             = ma_hoa_raw,
        ten_benh_pham      = ifelse(is.na(ten_benh_pham),      "", ten_benh_pham),
        ten_benh_pham_nhom = ifelse(is.na(ten_benh_pham_nhom), "", ten_benh_pham_nhom)
      ) %>%
      arrange(ma_hoa)

    b_missing_df(miss)
    b_updated_ref(NULL)
  })

  output$b_missing_folded_ui <- renderUI({
    df <- b_missing_df()
    if (is.null(df)) return(NULL)

    p <- preset_data

    tags$details(
      open = TRUE,
      tags$summary(
        "Danh sách cần chuẩn hóa",
        span(class = "badge", nrow(df))
      ),
      br(),
      build_mapping_ui("b", df, p, "bp")
    )
  })

  b_apply_updates_core <- function() {
    work <- b_updated_ref()
    if (is.null(work)) work <- b_ref_df()

    miss <- b_missing_df()
    if (is.null(miss) || nrow(miss) == 0) return(work)

    to_apply <- collect_bp_inputs(input, miss) %>%
      filter(trimws(ten_benh_pham) != "" | trimws(ten_benh_pham_nhom) != "")

    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)

      if (length(idx) > 0) {
        work$ten_benh_pham[idx]      <- to_apply$ten_benh_pham[i]
        work$ten_benh_pham_nhom[idx] <- to_apply$ten_benh_pham_nhom[i]
      } else {
        work <- bind_rows(work, tibble(
          ma_hoa             = to_apply$ma_hoa[i],
          ten_benh_pham      = to_apply$ten_benh_pham[i],
          ten_benh_pham_nhom = to_apply$ten_benh_pham_nhom[i],
          .key               = key
        ))
      }
    }

    b_updated_ref(work)
    work
  }

  observeEvent(input$b_apply_updates, {
    b_apply_updates_core()
    b_save_message("Đã áp dụng cập nhật vào dữ liệu tạm. Nhấn Lưu đè để ghi file.")
  })

  observeEvent(input$b_save_as_btn, {
    if (is.null(input$b_ref_file)) {
      b_save_message("⚠️ Chưa chọn file Danh mục Tên bệnh phẩm.")
      return()
    }
    out <- tryCatch(b_apply_updates_core(), error = function(e) {
      b_save_message(paste("Lỗi:", e$message))
      NULL
    })
    if (is.null(out)) return()

    ref_filename <- input$b_ref_file$name
    path <- ps_save_dialog(ref_filename, input$b_ref_file$datapath)

    if (is.na(path) || trimws(path) == "") {
      b_save_message("ℹ️ Đã hủy: không chọn đường dẫn lưu.")
      return()
    }
    if (!identical(basename(path), ref_filename)) {
      b_save_message(paste0("⚠️ Tên file không khớp với file gốc (", ref_filename, ")."))
      return()
    }

    tryCatch({
      out_clean <- out %>% select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom)
      openxlsx::write.xlsx(out_clean, file = path, overwrite = TRUE)
      b_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      b_save_message(paste0("❌ Lỗi khi ghi file: ", e$message))
    })
  })

  output$b_save_status <- renderUI({
    msg <- b_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  output$b_ref_preview <- DT::renderDT({
    req(b_updated_ref())
    DT::datatable(
      b_updated_ref() %>% select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom),
      rownames = FALSE
    )
  })

}

# =============================================================================
# KHỞI CHẠY ỨNG DỤNG
# =============================================================================

shinyApp(ui, server)
