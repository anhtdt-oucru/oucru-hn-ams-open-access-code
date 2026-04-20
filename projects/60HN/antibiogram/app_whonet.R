# =============================================================================
# ỨNG DỤNG 1: Làm sạch dữ liệu WHONET phiên giải
# -----------------------------------------------------------------------------
# Mục đích : Đọc nhiều file WHONET xuất từ phần mềm WHONET (.xlsx),
#            ghép với danh mục VSV và danh mục cơ chế kháng, sau đó
#            xuất ra một file dài (long-format) đã chuẩn hóa.
# Đầu vào  : (1) Một hoặc nhiều file WHONET (.xlsx)
#            (2) Danh mục VSV (.xlsx) — cột: ma_hoa, ten_vsv
#            (3) Danh mục cơ chế kháng (.xlsx) — cột: ten_vsv, khang_sinh,
#                ket_qua_ksd, ...
# Đầu ra   : File Excel long-format đã làm sạch, lưu qua hộp thoại Save-As
# =============================================================================

suppressPackageStartupMessages({
  library(shiny)      # framework giao diện web
  library(DT)         # bảng dữ liệu tương tác
  library(dplyr)      # xử lý data frame
  library(tidyr)      # pivot_longer
  library(openxlsx)   # đọc file Excel
  library(stringi)    # chuẩn hóa chuỗi (bỏ dấu, trim)
  library(stringr)    # xử lý chuỗi bổ sung
  library(tibble)     # tạo tibble nhanh
  library(janitor)    # clean_names — chuẩn hóa tên cột
  library(writexl)    # ghi file Excel nhẹ hơn openxlsx
})

# =============================================================================
# HÀM TIỆN ÍCH DÙNG CHUNG
# =============================================================================

# Chuẩn hóa khóa so khớp: bỏ dấu, viết thường, trim khoảng trắng
# Dùng để so khớp ma_hoa giữa dữ liệu thô và danh mục tham chiếu
normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- stringi::stri_trans_general(x, "Latin-ASCII") # bỏ dấu tiếng Việt
  tolower(x)
}

# Mở hộp thoại Save-As của Windows thông qua PowerShell
# Trả về đường dẫn file được chọn, hoặc NA nếu người dùng hủy
ps_save_dialog <- function(suggested_name, initial_path = NULL) {

  # Xác định thư mục mở mặc định của hộp thoại
  init_dir <- if (!is.null(initial_path) && nchar(initial_path) > 0) {
    normalizePath(dirname(initial_path), winslash = "/", mustWork = FALSE)
  } else {
    normalizePath("~", winslash = "/")
  }

  # Script PowerShell gọi SaveFileDialog của Windows Forms
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

  # Lọc dòng trống, lấy dòng cuối cùng là đường dẫn được chọn
  result <- trimws(result[nchar(trimws(result)) > 0])
  if (length(result) == 0 || result[length(result)] == "") NA_character_
  else result[length(result)]
}

# CSS dùng chung cho giao diện
app_css <- HTML("
  /* Vùng kéo thả / nhấn để chọn file */
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

  /* Hộp gợi ý màu xanh */
  .hint-box {
    background: #f5fbff;
    border-left: 4px solid #2c7be5;
    padding: 8px 12px;
    border-radius: 4px;
    font-size: 13px;
    margin-bottom: 10px;
  }

  /* Hộp cảnh báo màu vàng */
  .notice-box {
    background: #fff8e1;
    border-left: 4px solid #f0a500;
    padding: 10px 14px;
    border-radius: 4px;
    margin-bottom: 8px;
    font-size: 13px;
    color: #7a5000;
  }

  /* Trạng thái lưu file */
  .status-ok   { color: green;   font-weight: bold; margin-top: 10px; }
  .status-warn { color: #cc6600; font-weight: bold; margin-top: 10px; }
")

# =============================================================================
# HÀM XỬ LÝ DỮ LIỆU WHONET
# =============================================================================

# Tự động tìm dòng tiêu đề chứa cột "Organism" trong 20 dòng đầu file
# WHONET thường có nhiều dòng metadata trước khi bắt đầu bảng dữ liệu
detect_header_row <- function(path) {
  tryCatch({
    probe <- openxlsx::read.xlsx(path, colNames = FALSE, rows = 1:20)
    hit <- which(apply(probe, 1, function(r) {
      any(stringi::stri_trim_both(as.character(r)) == "Organism")
    }))
    if (length(hit) == 0) NULL else hit[1]
  }, error = function(e) NULL)
}

# Các cột cần xóa khỏi file WHONET — không cần thiết cho phân tích
cols_remove_whonet <- c(
  "macro_name", "ten_macro", "country", "quoc_gia",
  "laboratory", "phong_xet_nghiem", "origin", "nguon_goc",
  "date_of_birth", "ngay_sinh", "age_category", "loai_tuoi",
  "ward", "vung", "institution", "vien", "department",
  "location_type", "loai_vung", "local_specimen_code",
  "vung_ma_benh_pham", "specimen_type_numeric",
  "loai_benh_pham_dang_ma_so", "reason", "ly_do",
  "isolate_number", "so_phan_lap", "local_organism_code",
  "vung_ma_vi_khuan", "organism_type", "loai_vi_khuan",
  "serotype", "kieu_huyet_thanh", "mrsa", "mu_hon",
  "vre", "vang", "beta_lactamase", "esbl",
  "carbapenem_resistance", "khang_carbapenem",
  "mrsa_screening_test", "kiem_tra_khang_mrsa",
  "inducible_clindamycin_resistance", "ket_luan_khang_clindamycin",
  "comment", "ghi_chu", "date_of_data_entry",
  "ngay_vao_du_lieu", "ngay_tra_ket_qua"
)

# Bảng ánh xạ tên cột tiếng Việt -> tiếng Anh chuẩn
# WHONET có thể xuất file với tên cột tiếng Việt hoặc tiếng Anh
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

# Xử lý một file WHONET đơn lẻ
# Trả về data frame long-format hoặc NULL nếu file không hợp lệ
process_whonet_file <- function(file_path, ref_org, ref_resis) {

  # Bước 1: Tìm dòng tiêu đề thực sự của bảng
  header_row <- detect_header_row(file_path)
  if (is.null(header_row)) return(NULL)

  # Bước 2: Đọc file từ dòng tiêu đề
  df <- tryCatch(
    openxlsx::read.xlsx(file_path, startRow = header_row),
    error = function(e) NULL
  )
  if (is.null(df)) return(NULL)

  # Bước 3: Chuẩn hóa tên cột — trim khoảng trắng, chuyển về snake_case
  names(df) <- stringi::stri_trim_both(names(df))
  df <- janitor::clean_names(df)

  # Bước 4: Xóa các cột không cần thiết
  df <- df %>% select(-any_of(cols_remove_whonet))

  # Bước 5: Đổi tên cột tiếng Việt sang tiếng Anh (nếu có)
  for (en in names(col_mapping_whonet)) {
    vn <- col_mapping_whonet[[en]]
    if (vn %in% names(df) && !en %in% names(df)) {
      names(df)[names(df) == vn] <- en
    }
  }

  # Bắt buộc phải có cột organism
  if (!"organism" %in% names(df)) return(NULL)

  # Bước 6: Chuẩn hóa ngày tháng
  # WHONET có thể xuất ngày dưới dạng số serial Excel hoặc chuỗi mm/dd/yyyy
  fix_date <- function(x) {
    if (is.numeric(x)) as.Date(x, origin = "1899-12-30")
    else as.Date(as.character(x), format = "%m/%d/%Y")
  }
  if ("specimen_date"     %in% names(df)) df <- df %>% mutate(specimen_date     = fix_date(specimen_date))
  if ("date_of_admission" %in% names(df)) df <- df %>% mutate(date_of_admission = fix_date(date_of_admission))

  # Bước 7: Bỏ dòng không có ngày nuôi cấy, thêm cột năm
  if ("specimen_date" %in% names(df)) {
    df <- df %>%
      filter(!is.na(specimen_date)) %>%
      mutate(nam_nuoi_cay = format(specimen_date, "%Y"))
  }

  # Bước 8: Loại phân lập trùng lặp — giữ lần nuôi cấy sớm nhất
  # theo nhóm: bệnh nhân + năm + loại bệnh phẩm + vi khuẩn
  group_vars <- intersect(
    c("identification_number", "last_name", "first_name", "sex",
      "age", "nam_nuoi_cay", "specimen_type", "organism"),
    names(df)
  )
  if ("specimen_date" %in% names(df) && length(group_vars) > 0) {
    df <- df %>%
      group_by(across(all_of(group_vars))) %>%
      slice_min(specimen_date, with_ties = FALSE) %>%
      ungroup()
  }

  # Bước 9: Xác định các cột khóa và cột kháng sinh
  key_cols <- intersect(
    c("identification_number", "first_name", "last_name", "sex",
      "age", "location", "date_of_admission", "specimen_date",
      "nam_nuoi_cay", "specimen_type", "specimen_number", "organism"),
    names(df)
  )
  antibiotic_cols <- setdiff(names(df), key_cols)
  if (length(antibiotic_cols) == 0) return(NULL)

  # Bước 10: Ép kiểu tất cả cột kháng sinh về character trước khi pivot
  # WHONET trộn lẫn kết quả SIR (character: "S","R","I") với MIC (double: 0.5, 32)
  # trong cùng một file -> pivot_longer sẽ báo lỗi nếu không ép kiểu trước
  df <- df %>%
    mutate(across(all_of(antibiotic_cols), as.character))

  # Bước 11: Pivot sang dạng long (mỗi dòng = một kết quả kháng sinh đơn)
  df_long <- df %>%
    pivot_longer(
      cols      = all_of(antibiotic_cols),
      names_to  = "khang_sinh",
      values_to = "kq_ksd"
    )

  # Bước 12: Đổi tên cột sang tiếng Việt chuẩn của hệ thống
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

  # Bước 13: Chuẩn hóa giá trị
  # - Tên kháng sinh: lấy phần trước dấu "_" (bỏ hậu tố _nd30, _e, v.v.), viết hoa
  # - Kết quả KSD: viết hoa, trim
  # - Mã VSV: viết thường, trim
  df_long <- df_long %>%
    mutate(
      khang_sinh = str_to_upper(str_trim(str_split_fixed(khang_sinh, "_", 2)[, 1])),
      kq_ksd     = str_to_upper(str_trim(as.character(kq_ksd)))
    )
  if ("ma_vsv" %in% names(df_long)) {
    df_long <- df_long %>% mutate(ma_vsv = str_to_lower(str_trim(as.character(ma_vsv))))
  }
  if ("ma_bn" %in% names(df_long)) {
    df_long <- df_long %>% mutate(ma_bn = as.character(ma_bn))
  }

  # Bước 14: Ghép tên vi sinh vật từ danh mục VSV (nếu có)
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

  # Bước 15: Ghép thông tin cơ chế kháng từ danh mục (nếu có)
  if (!is.null(ref_resis) && "ten_vsv" %in% names(df_long)) {
    df_long <- df_long %>%
      left_join(
        ref_resis,
        by = c("ten_vsv", "khang_sinh", "kq_ksd" = "ket_qua_ksd")
      )
  }

  df_long
}

# =============================================================================
# GIAO DIỆN NGƯỜI DÙNG (UI)
# =============================================================================

ui <- fluidPage(

  titlePanel("Làm sạch dữ liệu WHONET phiên giải"),
  tags$style(app_css),

  sidebarLayout(

    # Cột trái: các bước thao tác
    sidebarPanel(

      h4("1) Chọn file đầu vào"),

      # Chọn file WHONET (cho phép chọn nhiều file cùng lúc)
      div(class = "drop-zone",
          onclick = "document.getElementById('w_raw_files').click()",
          "📂 Chọn file(s) WHONET (.xlsx)"),
      fileInput("w_raw_files", label = NULL, multiple = TRUE, accept = ".xlsx"),

      # Chọn danh mục VSV để ghép tên vi sinh vật
      div(class = "drop-zone",
          onclick = "document.getElementById('w_ref_org').click()",
          "📘 Chọn Danh mục VSV (.xlsx)"),
      fileInput("w_ref_org", label = NULL, multiple = FALSE, accept = ".xlsx"),

      # Chọn danh mục cơ chế kháng để ghép phân loại kết quả KSD
      div(class = "drop-zone",
          onclick = "document.getElementById('w_ref_resis').click()",
          "📙 Chọn Danh mục cơ chế kháng (.xlsx)"),
      fileInput("w_ref_resis", label = NULL, multiple = FALSE, accept = ".xlsx"),

      hr(),
      h4("2) Xử lý & Xuất"),

      # Nút xử lý: đọc tất cả file, ghép danh mục, tạo dữ liệu long-format
      actionButton("w_process", "▶ Xử lý dữ liệu",
                   class = "btn-primary", width = "100%"),

      br(), br(),

      div(class = "notice-box",
          HTML("⚠️ <b>Lưu ý:</b> Nhấn <b>Xuất</b> sau khi xử lý thành công
               để mở hộp thoại chọn nơi lưu file.")),

      # Nút xuất: mở hộp thoại Save-As của Windows
      actionButton("w_export", "💾 Xuất file đã làm sạch...",
                   class = "btn-warning", width = "100%"),

      br(), br(),

      # Hiển thị trạng thái lưu file (thành công / lỗi)
      uiOutput("w_save_status")
    ),

    # Cột phải: kết quả xử lý
    mainPanel(
      h4("Tổng quan"),
      verbatimTextOutput("w_summary"),   # Số dòng, số cột, khoảng năm
      hr(),
      h4("Xem trước dữ liệu (100 dòng đầu)"),
      DTOutput("w_preview")              # Bảng xem trước có thể cuộn ngang
    )
  )
)

# =============================================================================
# LOGIC XỬ LÝ (SERVER)
# =============================================================================

server <- function(input, output, session) {

  # --- Reactive: dữ liệu đã xử lý (NULL cho đến khi nhấn nút Xử lý) ---
  w_processed   <- reactiveVal(NULL)
  w_save_message <- reactiveVal("")

  # --- Đọc danh mục VSV khi file được tải lên ---
  w_ref_org <- reactive({
    req(input$w_ref_org)
    tryCatch(openxlsx::read.xlsx(input$w_ref_org$datapath), error = function(e) NULL)
  })

  # --- Đọc danh mục cơ chế kháng khi file được tải lên ---
  w_ref_resis <- reactive({
    req(input$w_ref_resis)
    tryCatch(openxlsx::read.xlsx(input$w_ref_resis$datapath), error = function(e) NULL)
  })

  # --- Xử lý khi nhấn nút "Xử lý dữ liệu" ---
  observeEvent(input$w_process, {
    req(input$w_raw_files)

    withProgress(message = "Đang xử lý dữ liệu WHONET...", value = 0, {

      files     <- input$w_raw_files$datapath
      ref_org   <- w_ref_org()
      ref_resis <- w_ref_resis()

      # Xử lý từng file, cập nhật thanh tiến trình
      all_data <- lapply(seq_along(files), function(i) {
        incProgress(1 / length(files), detail = input$w_raw_files$name[i])
        df <- process_whonet_file(files[i], ref_org, ref_resis)
        if (!is.null(df)) df$source_file <- input$w_raw_files$name[i]
        df
      })

      # Loại các file không xử lý được
      all_data <- all_data[!sapply(all_data, is.null)]

      if (length(all_data) == 0) {
        showNotification(
          "Không file nào xử lý được — kiểm tra lại định dạng WHONET.",
          type = "error"
        )
        w_processed(NULL)
      } else {
        # Gộp tất cả file thành một data frame duy nhất
        w_processed(bind_rows(all_data))
        w_save_message("")
      }
    })
  })

  # --- Hiển thị tổng quan dữ liệu đã xử lý ---
  output$w_summary <- renderText({
    df <- w_processed()
    if (is.null(df)) return("Chưa xử lý dữ liệu. Hãy chọn file và nhấn Xử lý.")

    # Tính khoảng năm từ cột ngày nuôi cấy
    nam_range <- if ("ngay_nuoi_cay" %in% names(df) && !all(is.na(df$ngay_nuoi_cay))) {
      yrs <- na.omit(format(df$ngay_nuoi_cay, "%Y"))
      paste0("Khoảng năm  : ", min(yrs), " – ", max(yrs), "\n")
    } else ""

    paste0(
      "Số dòng     : ", nrow(df), "\n",
      "Số cột      : ", ncol(df), "\n",
      nam_range,
      "File nguồn  : ", length(unique(df$source_file)), " file"
    )
  })

  # --- Bảng xem trước 100 dòng đầu ---
  output$w_preview <- renderDT({
    req(w_processed())
    datatable(
      head(w_processed(), 100),
      rownames = FALSE,
      options  = list(scrollX = TRUE, pageLength = 10)
    )
  })

  # --- Xuất file qua hộp thoại Save-As của Windows ---
  observeEvent(input$w_export, {

    df <- w_processed()
    if (is.null(df)) {
      w_save_message("⚠️ Chưa có dữ liệu. Hãy xử lý trước khi xuất.")
      return()
    }

    # Tạo tên file gợi ý: whonet_phien_giai_da_lam_sach_YYYY_YYYY.xlsx
    yrs_suffix <- if ("ngay_nuoi_cay" %in% names(df) && !all(is.na(df$ngay_nuoi_cay))) {
      yrs <- na.omit(format(df$ngay_nuoi_cay, "%Y"))
      paste0("_", min(yrs), "_", max(yrs))
    } else ""
    suggested <- paste0("whonet_phien_giai_da_lam_sach", yrs_suffix, ".xlsx")

    # Mở hộp thoại chọn nơi lưu
    path <- ps_save_dialog(suggested)

    if (is.na(path) || trimws(path) == "") {
      w_save_message("ℹ️ Đã hủy: không chọn đường dẫn lưu.")
      return()
    }

    tryCatch({
      writexl::write_xlsx(df, path)
      w_save_message(paste0("✅ Đã lưu thành công: ", path))
    }, error = function(e) {
      w_save_message(paste0("❌ Lỗi khi lưu file: ", e$message))
    })
  })

  # --- Hiển thị thông báo trạng thái lưu ---
  output$w_save_status <- renderUI({
    msg <- w_save_message()
    if (msg == "") return(NULL)
    cls <- if (grepl("^✅", msg)) "status-ok" else "status-warn"
    div(class = cls, msg)
  })
}

# =============================================================================
# KHỞI CHẠY ỨNG DỤNG
# =============================================================================

shinyApp(ui, server)
