# =============================================================================
# Công cụ làm sạch — Consolidated App
# Tabs: WHONET phiên giải | Tên VSV | Tên khoa | Tên bệnh phẩm
# All tabs: file upload input + PowerShell Save-As dialog for output
# =============================================================================

suppressPackageStartupMessages({
  library(shiny)
  library(DT)
  library(dplyr)
  library(tidyr)
  library(openxlsx)
  library(stringi)
  library(stringr)
  library(tibble)
  library(janitor)
  library(writexl)
  library(svDialogs)
})

# =============================================================================
# SHARED HELPERS
# =============================================================================

normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- stringi::stri_trans_general(x, "Latin-ASCII")
  tolower(x)
}

# PowerShell Save-As dialog — returns chosen path or NA
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

# Shared CSS
shared_css <- HTML("
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
")

# =============================================================================
# UI
# =============================================================================

ui <- fluidPage(

  titlePanel("Công cụ làm sạch — iSharp AMS"),
  tags$style(shared_css),

  tabsetPanel(
    id = "main_tabs",

    # =========================================================================
    # TAB 1 — WHONET PHIÊN GIẢI
    # =========================================================================
    tabPanel(
      "Dữ liệu WHONET phiên giải",
      br(),
      sidebarLayout(
        sidebarPanel(
          h4("1) Nguồn dữ liệu"),

          div(class = "drop-zone",
              onclick = "document.getElementById('w_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("w_raw_files", label = NULL, multiple = TRUE,
                    accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('w_ref_org').click()",
              "📘 Chọn Danh mục VSV (.xlsx)"),
          fileInput("w_ref_org", label = NULL, multiple = FALSE,
                    accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('w_ref_resis').click()",
              "📙 Chọn Danh mục cơ chế kháng (.xlsx)"),
          fileInput("w_ref_resis", label = NULL, multiple = FALSE,
                    accept = ".xlsx"),

          hr(),
          h4("2) Xử lý & Xuất"),
          actionButton("w_process", "Xử lý dữ liệu",
                       class = "btn-primary", width = "100%"),
          br(), br(),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu ý:</b> Nhấn <b>Xuất</b> sau khi đã xử lý
                   thành công để chọn nơi lưu file.")),
          actionButton("w_export", "💾 Xuất file đã làm sạch...",
                       class = "btn-warning", width = "100%"),
          br(), br(),
          uiOutput("w_save_status")
        ),

        mainPanel(
          h4("Tổng quan"),
          verbatimTextOutput("w_summary"),
          hr(),
          h4("Xem trước dữ liệu (100 dòng đầu)"),
          DTOutput("w_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 2 — TÊN VSV
    # =========================================================================
    tabPanel(
      "Tên VSV",
      br(),
      sidebarLayout(
        sidebarPanel(
          h4("1) Nguồn dữ liệu"),

          div(class = "drop-zone",
              onclick = "document.getElementById('v_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("v_raw_files", label = NULL, multiple = TRUE,
                    accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('v_ref_file').click()",
              "📘 Chọn Danh mục VSV (.xlsx)"),
          fileInput("v_ref_file", label = NULL, multiple = FALSE,
                    accept = ".xlsx"),

          hr(),
          h4("2) Kiểm tra"),
          actionButton("v_check_missing", "Kiểm tra mã chưa chuẩn hóa",
                       class = "btn-primary", width = "100%"),
          hr(),
          h4("3) Áp dụng"),
          actionButton("v_apply_updates", "Áp dụng cập nhật",
                       class = "btn-success", width = "100%"),
          br(), br(),
          h4("4) Lưu file"),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu ý:</b> Chức năng <b>Lưu đè</b> sẽ ghi đè lên
                   file Danh mục VSV gốc. Vui lòng chọn <b>đúng tên file gốc</b>.")),
          actionButton("v_save_as_btn", "💾 Lưu đè file Danh mục VSV...",
                       class = "btn-warning", width = "100%"),
          br(), br(),
          uiOutput("v_save_status")
        ),

        mainPanel(
          h4("Tổng quan"),
          verbatimTextOutput("v_summary"),
          hr(),
          h4("Mã vi sinh vật cần chuẩn hóa"),
          div(class = "hint-box",
              "Điền Tên vi sinh vật, Loại vi sinh vật, Tên viết tắt."),
          uiOutput("v_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DTOutput("v_ref_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 3 — TÊN KHOA
    # =========================================================================
    tabPanel(
      "Tên khoa",
      br(),
      sidebarLayout(
        sidebarPanel(
          h4("1) Nguồn dữ liệu"),

          div(class = "drop-zone",
              onclick = "document.getElementById('k_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("k_raw_files", label = NULL, multiple = TRUE,
                    accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('k_ref_file').click()",
              "📘 Chọn Danh mục Tên khoa (.xlsx)"),
          fileInput("k_ref_file", label = NULL, multiple = FALSE,
                    accept = ".xlsx"),

          hr(),
          h4("2) Kiểm tra"),
          actionButton("k_check_missing", "Kiểm tra Tên khoa chưa chuẩn hóa",
                       class = "btn-primary", width = "100%"),
          hr(),
          h4("3) Áp dụng"),
          actionButton("k_apply_updates", "Áp dụng cập nhật",
                       class = "btn-success", width = "100%"),
          br(), br(),
          h4("4) Lưu file"),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu ý:</b> Chức năng <b>Lưu đè</b> sẽ ghi đè lên
                   file Danh mục Tên khoa gốc. Vui lòng chọn <b>đúng tên file gốc</b>.")),
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
          div(class = "hint-box",
              "Điền Tên khoa và/hoặc Nhóm khoa."),
          uiOutput("k_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DTOutput("k_ref_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 4 — TÊN BỆNH PHẨM
    # =========================================================================
    tabPanel(
      "Tên bệnh phẩm",
      br(),
      sidebarLayout(
        sidebarPanel(
          h4("1) Nguồn dữ liệu"),

          div(class = "drop-zone",
              onclick = "document.getElementById('b_raw_files').click()",
              "📂 Chọn file(s) WHONET (.xlsx)"),
          fileInput("b_raw_files", label = NULL, multiple = TRUE,
                    accept = ".xlsx"),

          div(class = "drop-zone",
              onclick = "document.getElementById('b_ref_file').click()",
              "📘 Chọn Danh mục Tên bệnh phẩm (.xlsx)"),
          fileInput("b_ref_file", label = NULL, multiple = FALSE,
                    accept = ".xlsx"),

          hr(),
          h4("2) Kiểm tra"),
          actionButton("b_check_missing",
                       "Kiểm tra Tên bệnh phẩm chưa chuẩn hóa",
                       class = "btn-primary", width = "100%"),
          hr(),
          h4("3) Áp dụng"),
          actionButton("b_apply_updates", "Áp dụng cập nhật",
                       class = "btn-success", width = "100%"),
          br(), br(),
          h4("4) Lưu file"),
          div(class = "notice-box",
              HTML("⚠️ <b>Lưu ý:</b> Chức năng <b>Lưu đè</b> sẽ ghi đè lên
                   file Danh mục Tên bệnh phẩm gốc. Vui lòng chọn <b>đúng tên file gốc</b>.")),
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
          div(class = "hint-box",
              "Điền Tên bệnh phẩm và/hoặc Nhóm bệnh phẩm."),
          uiOutput("b_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DTOutput("b_ref_preview")
        )
      )
    )
  )
)

# =============================================================================
# SERVER
# =============================================================================

server <- function(input, output, session) {

  # ===========================================================================
  # TAB 1 — WHONET PHIÊN GIẢI
  # ===========================================================================

  w_processed <- reactiveVal(NULL)
  w_save_message <- reactiveVal("")

  # Helper: detect header row containing "Organism"
  detect_header_row <- function(path) {
    tryCatch({
      probe <- openxlsx::read.xlsx(path, colNames = FALSE, rows = 1:20)
      hit <- which(apply(probe, 1, function(r) {
        any(stringi::stri_trim_both(as.character(r)) == "Organism")
      }))
      if (length(hit) == 0) NULL else hit[1]
    }, error = function(e) NULL)
  }

  # Column rename map Vietnamese -> English
  col_mapping_whonet <- c(
    "identification_number" = "so_benh_an",
    "last_name"             = "ho",
    "first_name"            = "ten",
    "sex"                   = "gioi_tinh",
    "age"                   = "tuoi",
    "location"              = "khoa",
    "date_of_admission"     = "ngay_nhap_vien",
    "specimen_number"       = "so_benh_pham",
    "specimen_date"         = "ngay_lay_benh_pham",
    "specimen_type"         = "loai_benh_pham",
    "organism"              = "vi_khuan"
  )

  cols_remove_whonet <- c(
    "macro_name", "ten_macro", "country", "quoc_gia", "laboratory",
    "phong_xet_nghiem", "origin", "nguon_goc", "date_of_birth", "ngay_sinh",
    "age_category", "loai_tuoi", "ward", "vung", "institution", "vien",
    "department", "location_type", "loai_vung", "local_specimen_code",
    "vung_ma_benh_pham", "specimen_type_numeric", "loai_benh_pham_dang_ma_so",
    "reason", "ly_do", "isolate_number", "so_phan_lap",
    "local_organism_code", "vung_ma_vi_khuan", "organism_type",
    "loai_vi_khuan", "serotype", "kieu_huyet_thanh", "mrsa", "mu_hon",
    "vre", "vang", "beta_lactamase", "esbl", "carbapenem_resistance",
    "khang_carbapenem", "mrsa_screening_test", "kiem_tra_khang_mrsa",
    "inducible_clindamycin_resistance", "ket_luan_khang_clindamycin",
    "comment", "ghi_chu", "date_of_data_entry", "ngay_vao_du_lieu",
    "ngay_tra_ket_qua"
  )

  process_whonet_file <- function(file_path, ref_org, ref_resis) {
    header_row <- detect_header_row(file_path)
    if (is.null(header_row)) return(NULL)

    df <- tryCatch(
      openxlsx::read.xlsx(file_path, startRow = header_row),
      error = function(e) NULL
    )
    if (is.null(df)) return(NULL)

    # Normalize column names
    names(df) <- stringi::stri_trim_both(names(df))
    df <- janitor::clean_names(df)

    # Remove unwanted columns
    df <- df %>% select(-any_of(cols_remove_whonet))

    # Rename VN cols -> EN where present
    for (en in names(col_mapping_whonet)) {
      vn <- col_mapping_whonet[[en]]
      if (vn %in% names(df) && !en %in% names(df)) {
        names(df)[names(df) == vn] <- en
      }
    }

    if (!"organism" %in% names(df)) return(NULL)

    # Fix dates
    fix_date <- function(x) {
      if (is.numeric(x)) as.Date(x, origin = "1899-12-30")
      else as.Date(as.character(x), format = "%m/%d/%Y")
    }

    if ("specimen_date" %in% names(df)) {
      df <- df %>% mutate(specimen_date = fix_date(specimen_date))
    }
    if ("date_of_admission" %in% names(df)) {
      df <- df %>% mutate(date_of_admission = fix_date(date_of_admission))
    }

    if ("specimen_date" %in% names(df)) {
      df <- df %>%
        filter(!is.na(specimen_date)) %>%
        mutate(nam_nuoi_cay = format(specimen_date, "%Y"))
    }

    # Key columns
    key_cols <- intersect(
      c("identification_number", "first_name", "last_name", "sex", "age",
        "location", "date_of_admission", "specimen_date", "nam_nuoi_cay",
        "specimen_type", "specimen_number", "organism"),
      names(df)
    )

    # Dedup: keep earliest specimen date per isolate
    group_vars <- intersect(
      c("identification_number", "last_name", "first_name", "sex", "age",
        "nam_nuoi_cay", "specimen_type", "organism"),
      names(df)
    )

    if ("specimen_date" %in% names(df) && length(group_vars) > 0) {
      df <- df %>%
        group_by(across(all_of(group_vars))) %>%
        slice_min(specimen_date, with_ties = FALSE) %>%
        ungroup()
    }

    # Pivot to long
    antibiotic_cols <- setdiff(names(df), key_cols)
    if (length(antibiotic_cols) == 0) return(NULL)

    df_long <- df %>%
      pivot_longer(
        cols = all_of(antibiotic_cols),
        names_to = "khang_sinh",
        values_to = "kq_ksd"
      )

    # Standardise
    rename_map <- c(
      ma_bn          = "identification_number",
      ho_dem         = "last_name",
      ten_bn         = "first_name",
      gioi_tinh      = "sex",
      tuoi           = "age",
      ma_khoa        = "location",
      ngay_nhap_vien = "date_of_admission",
      ma_benh_pham   = "specimen_number",
      ngay_nuoi_cay  = "specimen_date",
      ten_benh_pham  = "specimen_type",
      ma_vsv         = "organism"
    )
    for (new_nm in names(rename_map)) {
      old_nm <- rename_map[[new_nm]]
      if (old_nm %in% names(df_long) && !new_nm %in% names(df_long)) {
        names(df_long)[names(df_long) == old_nm] <- new_nm
      }
    }

    df_long <- df_long %>%
      mutate(
        khang_sinh = str_to_upper(str_trim(str_split_fixed(khang_sinh, "_", 2)[, 1])),
        kq_ksd     = str_to_upper(str_trim(as.character(kq_ksd)))
      )

    if ("ma_vsv" %in% names(df_long)) {
      df_long <- df_long %>%
        mutate(ma_vsv = str_to_lower(str_trim(as.character(ma_vsv))))
    }
    if ("ma_bn" %in% names(df_long)) {
      df_long <- df_long %>% mutate(ma_bn = as.character(ma_bn))
    }

    # Join reference tables
    if (!is.null(ref_org) && "ma_vsv" %in% names(df_long)) {
      ref_org_clean <- ref_org %>%
        select(ma_hoa, ten_vsv) %>%
        distinct() %>%
        mutate(ma_hoa_key = normalize_key(ma_hoa))

      df_long <- df_long %>%
        mutate(.join_key = normalize_key(ma_vsv)) %>%
        left_join(ref_org_clean, by = c(".join_key" = "ma_hoa_key")) %>%
        select(-.join_key, -ma_hoa)
    }

    if (!is.null(ref_resis) && "ten_vsv" %in% names(df_long)) {
      df_long <- df_long %>%
        left_join(
          ref_resis,
          by = c("ten_vsv", "khang_sinh", "kq_ksd" = "ket_qua_ksd")
        )
    }

    df_long
  }

  # Reactive: load ref tables
  w_ref_org <- reactive({
    req(input$w_ref_org)
    tryCatch(openxlsx::read.xlsx(input$w_ref_org$datapath), error = function(e) NULL)
  })

  w_ref_resis <- reactive({
    req(input$w_ref_resis)
    tryCatch(openxlsx::read.xlsx(input$w_ref_resis$datapath), error = function(e) NULL)
  })

  # Process button
  observeEvent(input$w_process, {
    req(input$w_raw_files)

    withProgress(message = "Đang xử lý dữ liệu WHONET...", value = 0, {
      files <- input$w_raw_files$datapath
      ref_org   <- w_ref_org()
      ref_resis <- w_ref_resis()

      all_data <- lapply(seq_along(files), function(i) {
        incProgress(1 / length(files),
                    detail = input$w_raw_files$name[i])
        df <- process_whonet_file(files[i], ref_org, ref_resis)
        if (!is.null(df)) df$source_file <- input$w_raw_files$name[i]
        df
      })

      all_data <- all_data[!sapply(all_data, is.null)]

      if (length(all_data) == 0) {
        showNotification("Không file nào xử lý được — kiểm tra định dạng WHONET.",
                         type = "error")
        w_processed(NULL)
      } else {
        w_processed(bind_rows(all_data))
        w_save_message("")
      }
    })
  })

  output$w_summary <- renderText({
    df <- w_processed()
    if (is.null(df)) return("Chưa xử lý dữ liệu.")
    paste0(
      "Số dòng: ", nrow(df), "\n",
      "Số cột: ", ncol(df), "\n",
      if ("ngay_nuoi_cay" %in% names(df) && !all(is.na(df$ngay_nuoi_cay))) {
        yrs <- na.omit(format(df$ngay_nuoi_cay, "%Y"))
        paste0("Năm: ", min(yrs), " – ", max(yrs))
      } else ""
    )
  })

  output$w_preview <- renderDT({
    req(w_processed())
    datatable(head(w_processed(), 100), rownames = FALSE,
              options = list(scrollX = TRUE, pageLength = 10))
  })

  # Export via PowerShell Save-As
  observeEvent(input$w_export, {
    df <- w_processed()
    if (is.null(df)) {
      w_save_message("⚠️ Chưa có dữ liệu. Hãy xử lý trước khi xuất.")
      return()
    }

    yrs <- if ("ngay_nuoi_cay" %in% names(df) && !all(is.na(df$ngay_nuoi_cay))) {
      yrs_vec <- na.omit(format(df$ngay_nuoi_cay, "%Y"))
      paste0("_", min(yrs_vec), "_", max(yrs_vec))
    } else ""

    suggested <- paste0("whonet_phien_giai_da_lam_sach", yrs, ".xlsx")
    path <- ps_save_dialog(suggested)

    if (is.na(path) || trimws(path) == "") {
      w_save_message("ℹ️ Đã hủy: không chọn đường dẫn lưu.")
      return()
    }

    tryCatch({
      writexl::write_xlsx(df, path)
      w_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      w_save_message(paste0("❌ Lỗi khi lưu: ", e$message))
    })
  })

  output$w_save_status <- renderUI({
    msg <- w_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  # ===========================================================================
  # TAB 2 — TÊN VSV  (mirrors original app.R exactly)
  # ===========================================================================

  v_missing_editable <- reactiveVal(NULL)
  v_updated_ref      <- reactiveVal(NULL)
  v_save_message     <- reactiveVal("")

  v_ref_df <- reactive({
    req(input$v_ref_file)
    df <- tryCatch(
      openxlsx::read.xlsx(input$v_ref_file$datapath),
      error = function(e) NULL
    )
    validate(need(!is.null(df), "Không đọc được file Danh mục VSV."))
    names(df) <- stringi::stri_trim_both(names(df))
    validate(
      need("ma_hoa"       %in% names(df), "Thiếu cột ma_hoa"),
      need("ten_vsv"      %in% names(df), "Thiếu cột ten_vsv"),
      need("loai_vsv"     %in% names(df), "Thiếu cột loai_vsv"),
      need("ten_viet_tat" %in% names(df), "Thiếu cột ten_viet_tat")
    )
    df %>% transmute(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat,
                     .key = normalize_key(ma_hoa))
  })

  v_raw_df <- reactive({
    req(input$v_raw_files)
    files <- input$v_raw_files$datapath

    all_org <- lapply(seq_along(files), function(i) {
      f <- files[i]
      header_row <- tryCatch({
        probe <- openxlsx::read.xlsx(f, colNames = FALSE, rows = 1:20)
        hit <- which(apply(probe, 1, function(r) {
          any(stringi::stri_trim_both(as.character(r)) == "Organism")
        }))
        if (length(hit) == 0) return(NULL)
        hit[1]
      }, error = function(e) NULL)

      if (is.null(header_row)) return(NULL)
      df <- tryCatch(
        openxlsx::read.xlsx(f, startRow = header_row),
        error = function(e) NULL
      )
      if (is.null(df)) return(NULL)
      names(df) <- stringi::stri_trim_both(names(df))
      if (!"Organism" %in% names(df)) return(NULL)
      df %>% transmute(ma_hoa = as.character(Organism))
    })

    all_org <- all_org[!sapply(all_org, is.null)]
    validate(need(length(all_org) > 0, "Không file nào có cột Organism"))

    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })

  output$v_summary <- renderText({
    joined <- v_raw_df() %>% left_join(v_ref_df(), by = ".key")
    total <- nrow(joined)
    done <- sum(
      !is.na(joined$ten_vsv)      & trimws(joined$ten_vsv)      != "" &
        !is.na(joined$loai_vsv)     & trimws(joined$loai_vsv)     != "" &
        !is.na(joined$ten_viet_tat) & trimws(joined$ten_viet_tat) != ""
    )
    paste0("Tổng số mã VSV: ", total, "\n",
           "Đã chuẩn hóa đầy đủ: ", done, "\n",
           "Chưa chuẩn hóa / thiếu: ", total - done)
  })

  observeEvent(input$v_check_missing, {
    joined <- v_raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(v_ref_df(), by = ".key")

    miss <- joined %>%
      filter(is.na(ten_vsv) | trimws(ten_vsv) == "" |
               is.na(loai_vsv) | trimws(loai_vsv) == "" |
               is.na(ten_viet_tat) | trimws(ten_viet_tat) == "") %>%
      transmute(
        ma_hoa       = ma_hoa_raw,
        ten_vsv      = ifelse(is.na(ten_vsv), "", ten_vsv),
        loai_vsv     = ifelse(is.na(loai_vsv), "", loai_vsv),
        ten_viet_tat = ifelse(is.na(ten_viet_tat), "", ten_viet_tat)
      ) %>%
      arrange(ma_hoa)

    v_missing_editable(miss)
    v_updated_ref(NULL)
  })

  output$v_missing_folded_ui <- renderUI({
    df <- v_missing_editable()
    if (is.null(df)) return(NULL)
    tags$details(
      tags$summary("Danh sách cần chuẩn hóa", span(class = "badge", nrow(df))),
      DTOutput("v_missing_tbl")
    )
  })

  output$v_missing_tbl <- renderDT({
    datatable(
      v_missing_editable(), rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options = list(dom = "t", paging = FALSE, scrollY = "45vh",
                     language = list(emptyTable = "Không có dữ liệu"))
    )
  })

  observeEvent(input$v_missing_tbl_cell_edit, {
    info <- input$v_missing_tbl_cell_edit
    df   <- v_missing_editable()
    df[info$row, info$col + 1] <- info$value
    v_missing_editable(df)
  })

  v_apply_updates_core <- function() {
    work <- v_updated_ref()
    if (is.null(work)) work <- v_ref_df()
    miss <- v_missing_editable()
    if (is.null(miss)) return(work)

    to_apply <- miss %>%
      filter(trimws(ten_vsv) != "" | trimws(loai_vsv) != "" |
               trimws(ten_viet_tat) != "")

    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)
      if (length(idx) > 0) {
        work$ten_vsv[idx]      <- to_apply$ten_vsv[i]
        work$loai_vsv[idx]     <- to_apply$loai_vsv[i]
        work$ten_viet_tat[idx] <- to_apply$ten_viet_tat[i]
      } else {
        work <- bind_rows(work, tibble(
          ma_hoa = to_apply$ma_hoa[i], ten_vsv = to_apply$ten_vsv[i],
          loai_vsv = to_apply$loai_vsv[i],
          ten_viet_tat = to_apply$ten_viet_tat[i], .key = key
        ))
      }
    }
    v_updated_ref(work)
    work
  }

  observeEvent(input$v_apply_updates, {
    v_apply_updates_core()
    v_save_message("Đã áp dụng cập nhật vào dữ liệu tạm.")
  })

  observeEvent(input$v_save_as_btn, {
    if (is.null(input$v_ref_file)) {
      v_save_message("⚠️ Chưa chọn file Danh mục VSV.")
      return()
    }
    out <- tryCatch(v_apply_updates_core(), error = function(e) {
      v_save_message(paste("Lỗi:", e$message)); NULL
    })
    if (is.null(out)) return()

    ref_filename <- input$v_ref_file$name
    path <- ps_save_dialog(ref_filename, input$v_ref_file$datapath)

    if (is.na(path) || trimws(path) == "") {
      v_save_message("ℹ️ Đã hủy lưu file.")
      return()
    }
    if (!identical(basename(path), ref_filename)) {
      v_save_message(paste0("⚠️ Tên file không khớp với file gốc (",
                            ref_filename, "). Vui lòng đặt đúng tên."))
      return()
    }
    tryCatch({
      out_clean <- out %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat)
      openxlsx::write.xlsx(out_clean, file = path, overwrite = TRUE)
      v_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      v_save_message(paste0("❌ Lỗi khi lưu: ", e$message))
    })
  })

  output$v_save_status <- renderUI({
    msg <- v_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  output$v_ref_preview <- renderDT({
    req(v_updated_ref())
    datatable(
      v_updated_ref() %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat),
      rownames = FALSE
    )
  })

  # ===========================================================================
  # TAB 3 — TÊN KHOA
  # ===========================================================================

  k_missing_editable <- reactiveVal(NULL)
  k_updated_ref      <- reactiveVal(NULL)
  k_save_message     <- reactiveVal("")

  k_ref_df <- reactive({
    req(input$k_ref_file)
    df <- tryCatch(
      openxlsx::read.xlsx(input$k_ref_file$datapath),
      error = function(e) NULL
    )
    validate(need(!is.null(df), "Không đọc được file Danh mục Tên khoa."))
    names(df) <- stringi::stri_trim_both(names(df))
    validate(
      need("ma_hoa"       %in% names(df), "Thiếu cột ma_hoa"),
      need("ten_khoa"     %in% names(df), "Thiếu cột ten_khoa"),
      need("ten_khoa_nhom" %in% names(df), "Thiếu cột ten_khoa_nhom")
    )
    df %>% transmute(ma_hoa, ten_khoa, ten_khoa_nhom,
                     .key = normalize_key(ma_hoa))
  })

  k_raw_df <- reactive({
    req(input$k_raw_files)
    files <- input$k_raw_files$datapath

    all_org <- lapply(seq_along(files), function(i) {
      f <- files[i]
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
    validate(need(length(all_org) > 0, "Không file nào có cột Location"))

    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })

  output$k_summary <- renderText({
    joined <- k_raw_df() %>% left_join(k_ref_df(), by = ".key")
    total <- nrow(joined)
    done  <- sum(!is.na(joined$ten_khoa) & trimws(joined$ten_khoa) != "" &
                   !is.na(joined$ten_khoa_nhom) & trimws(joined$ten_khoa_nhom) != "")
    paste0("Tổng số tên khoa: ", total, "\n",
           "Đã chuẩn hóa đầy đủ: ", done, "\n",
           "Chưa chuẩn hóa / thiếu: ", total - done)
  })

  observeEvent(input$k_check_missing, {
    joined <- k_raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(k_ref_df(), by = ".key")

    miss <- joined %>%
      filter(is.na(ten_khoa) | trimws(ten_khoa) == "" |
               is.na(ten_khoa_nhom) | trimws(ten_khoa_nhom) == "") %>%
      transmute(
        ma_hoa        = ma_hoa_raw,
        ten_khoa      = ifelse(is.na(ten_khoa), "", ten_khoa),
        ten_khoa_nhom = ifelse(is.na(ten_khoa_nhom), "", ten_khoa_nhom)
      ) %>%
      arrange(ma_hoa)

    k_missing_editable(miss)
    k_updated_ref(NULL)
  })

  output$k_missing_folded_ui <- renderUI({
    df <- k_missing_editable()
    if (is.null(df)) return(NULL)
    tags$details(
      tags$summary("Danh sách cần chuẩn hóa", span(class = "badge", nrow(df))),
      DTOutput("k_missing_tbl")
    )
  })

  output$k_missing_tbl <- renderDT({
    datatable(
      k_missing_editable(), rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options = list(dom = "t", paging = FALSE, scrollY = "45vh",
                     language = list(emptyTable = "Không có dữ liệu"))
    )
  })

  observeEvent(input$k_missing_tbl_cell_edit, {
    info <- input$k_missing_tbl_cell_edit
    df   <- k_missing_editable()
    df[info$row, info$col + 1] <- info$value
    k_missing_editable(df)
  })

  k_apply_updates_core <- function() {
    work <- k_updated_ref()
    if (is.null(work)) work <- k_ref_df()
    miss <- k_missing_editable()
    if (is.null(miss)) return(work)

    to_apply <- miss %>%
      filter(trimws(ten_khoa) != "" | trimws(ten_khoa_nhom) != "")

    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)
      if (length(idx) > 0) {
        work$ten_khoa[idx]      <- to_apply$ten_khoa[i]
        work$ten_khoa_nhom[idx] <- to_apply$ten_khoa_nhom[i]
      } else {
        work <- bind_rows(work, tibble(
          ma_hoa = to_apply$ma_hoa[i], ten_khoa = to_apply$ten_khoa[i],
          ten_khoa_nhom = to_apply$ten_khoa_nhom[i], .key = key
        ))
      }
    }
    k_updated_ref(work)
    work
  }

  observeEvent(input$k_apply_updates, {
    k_apply_updates_core()
    k_save_message("Đã áp dụng cập nhật vào dữ liệu tạm.")
  })

  observeEvent(input$k_save_as_btn, {
    if (is.null(input$k_ref_file)) {
      k_save_message("⚠️ Chưa chọn file Danh mục Tên khoa.")
      return()
    }
    out <- tryCatch(k_apply_updates_core(), error = function(e) {
      k_save_message(paste("Lỗi:", e$message)); NULL
    })
    if (is.null(out)) return()

    ref_filename <- input$k_ref_file$name
    path <- ps_save_dialog(ref_filename, input$k_ref_file$datapath)

    if (is.na(path) || trimws(path) == "") {
      k_save_message("ℹ️ Đã hủy lưu file.")
      return()
    }
    if (!identical(basename(path), ref_filename)) {
      k_save_message(paste0("⚠️ Tên file không khớp với file gốc (",
                            ref_filename, "). Vui lòng đặt đúng tên."))
      return()
    }
    tryCatch({
      out_clean <- out %>% select(ma_hoa, ten_khoa, ten_khoa_nhom)
      openxlsx::write.xlsx(out_clean, file = path, overwrite = TRUE)
      k_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      k_save_message(paste0("❌ Lỗi khi lưu: ", e$message))
    })
  })

  output$k_save_status <- renderUI({
    msg <- k_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  output$k_ref_preview <- renderDT({
    req(k_updated_ref())
    datatable(
      k_updated_ref() %>% select(ma_hoa, ten_khoa, ten_khoa_nhom),
      rownames = FALSE
    )
  })

  # ===========================================================================
  # TAB 4 — TÊN BỆNH PHẨM
  # ===========================================================================

  b_missing_editable <- reactiveVal(NULL)
  b_updated_ref      <- reactiveVal(NULL)
  b_save_message     <- reactiveVal("")

  b_ref_df <- reactive({
    req(input$b_ref_file)
    df <- tryCatch(
      openxlsx::read.xlsx(input$b_ref_file$datapath),
      error = function(e) NULL
    )
    validate(need(!is.null(df), "Không đọc được file Danh mục Tên bệnh phẩm."))
    names(df) <- stringi::stri_trim_both(names(df))
    validate(
      need("ma_hoa"            %in% names(df), "Thiếu cột ma_hoa"),
      need("ten_benh_pham"     %in% names(df), "Thiếu cột ten_benh_pham"),
      need("ten_benh_pham_nhom" %in% names(df), "Thiếu cột ten_benh_pham_nhom")
    )
    df %>% transmute(ma_hoa, ten_benh_pham, ten_benh_pham_nhom,
                     .key = normalize_key(ma_hoa))
  })

  b_raw_df <- reactive({
    req(input$b_raw_files)
    files <- input$b_raw_files$datapath

    all_org <- lapply(seq_along(files), function(i) {
      f <- files[i]
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
    validate(need(length(all_org) > 0, "Không file nào có cột 'Specimen type'"))

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
    paste0("Tổng số tên bệnh phẩm: ", total, "\n",
           "Đã chuẩn hóa đầy đủ: ", done, "\n",
           "Chưa chuẩn hóa / thiếu: ", total - done)
  })

  observeEvent(input$b_check_missing, {
    joined <- b_raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(b_ref_df(), by = ".key")

    miss <- joined %>%
      filter(is.na(ten_benh_pham) | trimws(ten_benh_pham) == "" |
               is.na(ten_benh_pham_nhom) | trimws(ten_benh_pham_nhom) == "") %>%
      transmute(
        ma_hoa             = ma_hoa_raw,
        ten_benh_pham      = ifelse(is.na(ten_benh_pham), "", ten_benh_pham),
        ten_benh_pham_nhom = ifelse(is.na(ten_benh_pham_nhom), "",
                                    ten_benh_pham_nhom)
      ) %>%
      arrange(ma_hoa)

    b_missing_editable(miss)
    b_updated_ref(NULL)
  })

  output$b_missing_folded_ui <- renderUI({
    df <- b_missing_editable()
    if (is.null(df)) return(NULL)
    tags$details(
      tags$summary("Danh sách cần chuẩn hóa", span(class = "badge", nrow(df))),
      DTOutput("b_missing_tbl")
    )
  })

  output$b_missing_tbl <- renderDT({
    datatable(
      b_missing_editable(), rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options = list(dom = "t", paging = FALSE, scrollY = "45vh",
                     language = list(emptyTable = "Không có dữ liệu"))
    )
  })

  observeEvent(input$b_missing_tbl_cell_edit, {
    info <- input$b_missing_tbl_cell_edit
    df   <- b_missing_editable()
    df[info$row, info$col + 1] <- info$value
    b_missing_editable(df)
  })

  b_apply_updates_core <- function() {
    work <- b_updated_ref()
    if (is.null(work)) work <- b_ref_df()
    miss <- b_missing_editable()
    if (is.null(miss)) return(work)

    to_apply <- miss %>%
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
    b_save_message("Đã áp dụng cập nhật vào dữ liệu tạm.")
  })

  observeEvent(input$b_save_as_btn, {
    if (is.null(input$b_ref_file)) {
      b_save_message("⚠️ Chưa chọn file Danh mục Tên bệnh phẩm.")
      return()
    }
    out <- tryCatch(b_apply_updates_core(), error = function(e) {
      b_save_message(paste("Lỗi:", e$message)); NULL
    })
    if (is.null(out)) return()

    ref_filename <- input$b_ref_file$name
    path <- ps_save_dialog(ref_filename, input$b_ref_file$datapath)

    if (is.na(path) || trimws(path) == "") {
      b_save_message("ℹ️ Đã hủy lưu file.")
      return()
    }
    if (!identical(basename(path), ref_filename)) {
      b_save_message(paste0("⚠️ Tên file không khớp với file gốc (",
                            ref_filename, "). Vui lòng đặt đúng tên."))
      return()
    }
    tryCatch({
      out_clean <- out %>% select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom)
      openxlsx::write.xlsx(out_clean, file = path, overwrite = TRUE)
      b_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      b_save_message(paste0("❌ Lỗi khi lưu: ", e$message))
    })
  })

  output$b_save_status <- renderUI({
    msg <- b_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })

  output$b_ref_preview <- renderDT({
    req(b_updated_ref())
    datatable(
      b_updated_ref() %>% select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom),
      rownames = FALSE
    )
  })

}

# =============================================================================
# RUN APP
# =============================================================================

shinyApp(ui, server)
