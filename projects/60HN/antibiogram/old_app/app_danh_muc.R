# =============================================================================
# ỨNG DỤNG 2: Cập nhật Danh mục tham chiếu
# -----------------------------------------------------------------------------
# Mục đích : Phát hiện các mã chưa được chuẩn hóa trong ba danh mục tham chiếu
#            (Vi sinh vật, Tên khoa, Tên bệnh phẩm), cho phép người dùng điền
#            thông tin còn thiếu trực tiếp trên giao diện, rồi lưu đè file gốc.
#
# Gồm 3 tab:
#   Tab 1 — Tên VSV      : cột bắt buộc ma_hoa | ten_vsv | loai_vsv | ten_viet_tat
#   Tab 2 — Tên khoa     : cột bắt buộc ma_hoa | ten_khoa | ten_khoa_nhom
#   Tab 3 — Tên bệnh phẩm: cột bắt buộc ma_hoa | ten_benh_pham | ten_benh_pham_nhom
#
# Quy trình mỗi tab:
#   (1) Tải file WHONET gốc + file danh mục tham chiếu
#   (2) Kiểm tra → hiển thị danh sách mã chưa chuẩn hóa (bảng có thể sửa)
#   (3) Sửa trực tiếp trên bảng → Áp dụng cập nhật
#   (4) Lưu đè file danh mục gốc qua hộp thoại Save-As
# =============================================================================

suppressPackageStartupMessages({
  library(shiny)    # framework giao diện web
  library(DT)       # bảng dữ liệu tương tác, hỗ trợ sửa trực tiếp (editable)
  library(dplyr)    # xử lý data frame
  library(openxlsx) # đọc / ghi file Excel
  library(stringi)  # chuẩn hóa chuỗi (bỏ dấu, trim)
  library(tibble)   # tạo tibble nhanh khi thêm dòng mới
})

# =============================================================================
# HÀM TIỆN ÍCH DÙNG CHUNG
# =============================================================================

# Chuẩn hóa khóa so khớp: bỏ dấu tiếng Việt, viết thường, trim khoảng trắng
# Đảm bảo so khớp chính xác dù người dùng nhập có dấu hay không
normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- stringi::stri_trans_general(x, "Latin-ASCII")
  tolower(x)
}

# Mở hộp thoại Save-As của Windows thông qua PowerShell
# Trả về đường dẫn đầy đủ được chọn, hoặc NA nếu người dùng bấm Cancel
ps_save_dialog <- function(suggested_name, initial_path = NULL) {

  # Thư mục mặc định: lấy từ đường dẫn file đã tải lên (nếu có)
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

# CSS dùng chung cho giao diện cả 3 tab
app_css <- HTML("
  /* Vùng nhấn để chọn file */
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

  /* Badge số lượng mã cần chuẩn hóa */
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

  /* Khối có thể thu gọn chứa bảng cần chuẩn hóa */
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
# GIAO DIỆN NGƯỜI DÙNG (UI)
# =============================================================================

ui <- fluidPage(

  titlePanel("Cập nhật Danh mục tham chiếu"),
  tags$style(app_css),

  tabsetPanel(
    id = "main_tabs",

    # =========================================================================
    # TAB 1 — TÊN VI SINH VẬT
    # Danh mục cần có cột: ma_hoa | ten_vsv | loai_vsv | ten_viet_tat
    # =========================================================================
    tabPanel(
      "Tên VSV",
      br(),
      sidebarLayout(
        sidebarPanel(

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
          div(class = "hint-box",
              "Nhấn vào ô trong bảng để sửa trực tiếp. Các cột cần điền:
               Tên vi sinh vật, Loại vi sinh vật, Tên viết tắt."),
          uiOutput("v_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DTOutput("v_ref_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 2 — TÊN KHOA
    # Danh mục cần có cột: ma_hoa | ten_khoa | ten_khoa_nhom
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
          div(class = "hint-box",
              "Nhấn vào ô trong bảng để sửa trực tiếp. Các cột cần điền:
               Tên khoa, Nhóm khoa."),
          uiOutput("k_missing_folded_ui"),
          hr(),
          h4("Xem trước danh mục sau cập nhật"),
          DTOutput("k_ref_preview")
        )
      )
    ),

    # =========================================================================
    # TAB 3 — TÊN BỆNH PHẨM
    # Danh mục cần có cột: ma_hoa | ten_benh_pham | ten_benh_pham_nhom
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
          div(class = "hint-box",
              "Nhấn vào ô trong bảng để sửa trực tiếp. Các cột cần điền:
               Tên bệnh phẩm, Nhóm bệnh phẩm."),
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
# LOGIC XỬ LÝ (SERVER)
# =============================================================================

server <- function(input, output, session) {

  # ===========================================================================
  # TAB 1 — TÊN VI SINH VẬT
  # ===========================================================================

  # Trạng thái nội bộ của tab: bảng cần chuẩn hóa và danh mục đã cập nhật
  v_missing_editable <- reactiveVal(NULL)  # bảng mã chưa chuẩn hóa (có thể sửa)
  v_updated_ref      <- reactiveVal(NULL)  # danh mục sau khi áp dụng cập nhật
  v_save_message     <- reactiveVal("")    # thông báo trạng thái lưu

  # Đọc và kiểm tra file danh mục VSV
  v_ref_df <- reactive({
    req(input$v_ref_file)
    df <- tryCatch(
      openxlsx::read.xlsx(input$v_ref_file$datapath),
      error = function(e) NULL
    )
    validate(need(!is.null(df), "Không đọc được file Danh mục VSV."))
    names(df) <- stringi::stri_trim_both(names(df))

    # Kiểm tra các cột bắt buộc
    validate(
      need("ma_hoa"       %in% names(df), "Thiếu cột: ma_hoa"),
      need("ten_vsv"      %in% names(df), "Thiếu cột: ten_vsv"),
      need("loai_vsv"     %in% names(df), "Thiếu cột: loai_vsv"),
      need("ten_viet_tat" %in% names(df), "Thiếu cột: ten_viet_tat")
    )

    # Thêm cột khóa chuẩn hóa để so khớp không phân biệt dấu/hoa thường
    df %>% transmute(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat,
                     .key = normalize_key(ma_hoa))
  })

  # Đọc tất cả file WHONET, trích xuất cột Organism (mã vi sinh vật)
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

    # Gộp, lọc trống, loại trùng, thêm khóa chuẩn hóa
    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })

  # Tổng quan: số mã đã / chưa chuẩn hóa
  output$v_summary <- renderText({
    joined <- v_raw_df() %>% left_join(v_ref_df(), by = ".key")
    total  <- nrow(joined)
    done   <- sum(
      !is.na(joined$ten_vsv)      & trimws(joined$ten_vsv)      != "" &
        !is.na(joined$loai_vsv)     & trimws(joined$loai_vsv)     != "" &
        !is.na(joined$ten_viet_tat) & trimws(joined$ten_viet_tat) != ""
    )
    paste0(
      "Tổng số mã VSV trong WHONET : ", total, "\n",
      "Đã chuẩn hóa đầy đủ         : ", done, "\n",
      "Chưa chuẩn hóa / thiếu thông tin: ", total - done
    )
  })

  # Tìm và hiển thị danh sách mã chưa chuẩn hóa khi nhấn Kiểm tra
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

    v_missing_editable(miss)
    v_updated_ref(NULL) # reset bản nháp khi kiểm tra lại
  })

  # Render bảng danh sách cần chuẩn hóa trong khối thu gọn (details/summary)
  output$v_missing_folded_ui <- renderUI({
    df <- v_missing_editable()
    if (is.null(df)) return(NULL)
    tags$details(
      tags$summary("Danh sách cần chuẩn hóa", span(class = "badge", nrow(df))),
      DTOutput("v_missing_tbl")
    )
  })

  # Bảng có thể sửa trực tiếp — cột đầu (ma_hoa) bị khóa, các cột còn lại được sửa
  output$v_missing_tbl <- renderDT({
    datatable(
      v_missing_editable(), rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options  = list(dom = "t", paging = FALSE, scrollY = "45vh",
                      language = list(emptyTable = "Không có dữ liệu cần chuẩn hóa"))
    )
  })

  # Cập nhật reactive khi người dùng sửa một ô trong bảng
  observeEvent(input$v_missing_tbl_cell_edit, {
    info <- input$v_missing_tbl_cell_edit
    df   <- v_missing_editable()
    df[info$row, info$col + 1] <- info$value # col + 1 vì rownames = FALSE
    v_missing_editable(df)
  })

  # Hàm nội bộ: áp dụng các thay đổi từ bảng vào danh mục tham chiếu
  # - Cập nhật dòng đã có (khớp .key)
  # - Thêm dòng mới nếu mã chưa tồn tại trong danh mục
  v_apply_updates_core <- function() {
    work <- v_updated_ref()
    if (is.null(work)) work <- v_ref_df() # bắt đầu từ danh mục hiện tại

    miss <- v_missing_editable()
    if (is.null(miss)) return(work)

    # Chỉ áp dụng những dòng có ít nhất một trường đã được điền
    to_apply <- miss %>%
      filter(trimws(ten_vsv) != "" | trimws(loai_vsv) != "" | trimws(ten_viet_tat) != "")

    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)

      if (length(idx) > 0) {
        # Cập nhật dòng đã có
        work$ten_vsv[idx]      <- to_apply$ten_vsv[i]
        work$loai_vsv[idx]     <- to_apply$loai_vsv[i]
        work$ten_viet_tat[idx] <- to_apply$ten_viet_tat[i]
      } else {
        # Thêm dòng mới
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

  # Nút Áp dụng: cập nhật danh mục nháp, báo người dùng
  observeEvent(input$v_apply_updates, {
    v_apply_updates_core()
    v_save_message("Đã áp dụng cập nhật vào dữ liệu tạm. Nhấn Lưu đè để ghi file.")
  })

  # Nút Lưu đè: mở hộp thoại Save-As, kiểm tra tên file, ghi Excel
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

    # Bảo vệ: cảnh báo nếu người dùng đổi tên file (có thể ghi nhầm)
    if (!identical(basename(path), ref_filename)) {
      v_save_message(paste0(
        "⚠️ Tên file không khớp với file gốc (", ref_filename,
        "). Vui lòng đặt đúng tên file gốc."
      ))
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

  # Bảng xem trước danh mục sau khi áp dụng cập nhật
  output$v_ref_preview <- renderDT({
    req(v_updated_ref())
    datatable(
      v_updated_ref() %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat),
      rownames = FALSE
    )
  })

  # ===========================================================================
  # TAB 2 — TÊN KHOA
  # Cấu trúc giống hệt Tab 1, chỉ khác tên cột và tên file danh mục
  # ===========================================================================

  k_missing_editable <- reactiveVal(NULL)
  k_updated_ref      <- reactiveVal(NULL)
  k_save_message     <- reactiveVal("")

  # Đọc và kiểm tra file danh mục Tên khoa
  k_ref_df <- reactive({
    req(input$k_ref_file)
    df <- tryCatch(
      openxlsx::read.xlsx(input$k_ref_file$datapath),
      error = function(e) NULL
    )
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

  # Trích xuất cột Location (tên khoa) từ các file WHONET
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
      "Tổng số tên khoa trong WHONET: ", total, "\n",
      "Đã chuẩn hóa đầy đủ          : ", done, "\n",
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
      options  = list(dom = "t", paging = FALSE, scrollY = "45vh",
                      language = list(emptyTable = "Không có dữ liệu cần chuẩn hóa"))
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
      k_save_message(paste0(
        "⚠️ Tên file không khớp với file gốc (", ref_filename, ")."
      ))
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

  output$k_ref_preview <- renderDT({
    req(k_updated_ref())
    datatable(
      k_updated_ref() %>% select(ma_hoa, ten_khoa, ten_khoa_nhom),
      rownames = FALSE
    )
  })

  # ===========================================================================
  # TAB 3 — TÊN BỆNH PHẨM
  # Cấu trúc giống hệt Tab 1 & 2, khác cột: ten_benh_pham, ten_benh_pham_nhom
  # Trích xuất từ cột "Specimen type" trong file WHONET
  # ===========================================================================

  b_missing_editable <- reactiveVal(NULL)
  b_updated_ref      <- reactiveVal(NULL)
  b_save_message     <- reactiveVal("")

  # Đọc và kiểm tra file danh mục Tên bệnh phẩm
  b_ref_df <- reactive({
    req(input$b_ref_file)
    df <- tryCatch(
      openxlsx::read.xlsx(input$b_ref_file$datapath),
      error = function(e) NULL
    )
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

  # Trích xuất cột "Specimen type" (loại bệnh phẩm) từ các file WHONET
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
      "Đã chuẩn hóa đầy đủ               : ", done, "\n",
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
      options  = list(dom = "t", paging = FALSE, scrollY = "45vh",
                      language = list(emptyTable = "Không có dữ liệu cần chuẩn hóa"))
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
      b_save_message(paste0(
        "⚠️ Tên file không khớp với file gốc (", ref_filename, ")."
      ))
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

  output$b_ref_preview <- renderDT({
    req(b_updated_ref())
    datatable(
      b_updated_ref() %>% select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom),
      rownames = FALSE
    )
  })

}

# =============================================================================
# KHỞI CHẠY ỨNG DỤNG
# =============================================================================

shinyApp(ui, server)
