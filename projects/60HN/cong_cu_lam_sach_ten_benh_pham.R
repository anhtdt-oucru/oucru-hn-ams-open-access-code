
# =========================
# Chuẩn hóa Danh mục Tên bệnh phẩm
# =========================

suppressPackageStartupMessages({
  library(shiny)
  library(DT)
  library(dplyr)
  library(readxl)
  library(stringi)
  library(writexl)
  library(tibble)
})

# ---------- String normalization ----------
normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- stringi::stri_trans_general(x, "Latin-ASCII")
  tolower(x)
}

# ---------- Paths ----------
RAW_DIR <- "G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/DTH/3.du_lieu_whonet_phien_giai"

REF_FILE <- "G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/DTH/1.danh_muc/DTH_danh_muc_benh_pham.xlsx"

# =========================
# UI
# =========================
ui <- fluidPage(
  titlePanel("Chuẩn hóa Danh mục tên bệnh phẩm"),
  
  tags$style(HTML("
    #missing_tbl table thead th:nth-child(2),
    #missing_tbl table thead th:nth-child(3) {
      background-color: #FFE999 !important;
      font-weight: 700;
    }
    #missing_tbl table tbody td:nth-child(2),
    #missing_tbl table tbody td:nth-child(3) {
      background-color: #FFF7CC !important;
    }
    .hint-box {
      background: #f5fbff;
      border-left: 4px solid #2c7be5;
      padding: 8px 12px;
      border-radius: 4px;
    }
    details summary {
      cursor: pointer;
      font-weight: 700;
      background: #e2f0fb;
      padding: 8px 12px;
      border-radius: 4px;
    }
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
    details summary::-webkit-details-marker { display:none; }
  ")),
  
  sidebarLayout(
    sidebarPanel(
      h4("1) Nguồn dữ liệu"),
      p("Dữ liệu gốc: nhiều file WHONET (cột Specimen type)"),
      p("Danh mục chuẩn hóa: Danh mục bệnh phẩm"),
      actionButton("reload_files", "Tải lại dữ liệu"),
      hr(),
      
      h4("2) Kiểm tra"),
      actionButton(
        "check_missing",
        "Kiểm tra tên bệnh phẩm chưa được chuẩn hóa",
        class = "btn-primary"
      ),
      hr(),
      
      h4("3) Áp dụng & Xuất"),
      actionButton(
        "apply_updates",
        "Áp dụng cập nhật",
        class = "btn-success",
        onclick = "document.activeElement && document.activeElement.blur();"
      ),
      br(), br(),
      actionButton(
        "export_new",
        "Xuất (ghi đè danh mục)",
        class = "btn-warning",
        onclick = "document.activeElement && document.activeElement.blur();"
      )
    ),
    
    mainPanel(
      h4("Tổng quan"),
      verbatimTextOutput("summary"),
      hr(),
      
      h4("Tên bệnh phẩm cần chuẩn hóa"),
      div(
        class = "hint-box",
        "Điền Tên bệnh phẩm và/hoặc Nhóm bệnh phẩm. Có thể điền từng phần, không cần đủ mới áp dụng."
      ),
      uiOutput("missing_folded_ui"),
      
      hr(),
      h4("Xem trước danh mục sau cập nhật"),
      DTOutput("ref_preview")
    )
  )
)

# =========================
# SERVER
# =========================
server <- function(input, output, session) {
  
  # ---- reload trigger ----
  files_version <- reactiveVal(Sys.time())
  observeEvent(input$reload_files, {
    files_version(Sys.time())
  })
  
  # ---------- RAW ----------
  raw_df <- reactive({
    files_version()
    
    validate(
      need(dir.exists(RAW_DIR), paste("Không tìm thấy thư mục:", RAW_DIR))
    )
    
    files <- list.files(RAW_DIR, pattern = "\\.xlsx$", full.names = TRUE)
    validate(need(length(files) > 0, "Không có file Excel trong thư mục RAW"))
    
    all_specimen <- lapply(files, function(f) {
      df <- readxl::read_excel(f, skip = 7)
      if (!"Specimen type" %in% names(df)) return(NULL)
      
      df %>%
        transmute(ma_hoa = as.character(`Specimen type`))
    })
    
    bind_rows(all_specimen) %>%
      filter(!is.na(ma_hoa), trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })
  
  # ---------- REF ----------
  ref_df <- reactive({
    files_version()
    
    validate(
      need(file.exists(REF_FILE), paste("Không tìm thấy file:", REF_FILE))
    )
    
    df <- readxl::read_excel(REF_FILE)
    
    validate(
      need("ma_hoa" %in% names(df), "Thiếu cột ma_hoa"),
      need("ten_benh_pham" %in% names(df), "Thiếu cột ten_benh_pham"),
      need("ten_benh_pham_nhom" %in% names(df), "Thiếu cột ten_benh_pham_nhom")
    )
    
    df %>%
      transmute(
        ma_hoa,
        ten_benh_pham,
        ten_benh_pham_nhom,
        .key = normalize_key(ma_hoa)
      )
  })
  
  # ---------- State ----------
  missing_editable <- reactiveVal(NULL)
  updated_ref <- reactiveVal(NULL)
  
  # ---------- Summary ----------
  output$summary <- renderText({
    joined <- raw_df() %>%
      left_join(ref_df(), by = ".key")
    
    total <- nrow(joined)
    
    done <- sum(
      !is.na(joined$ten_benh_pham) &
        trimws(joined$ten_benh_pham) != "" &
        !is.na(joined$ten_benh_pham_nhom) &
        trimws(joined$ten_benh_pham_nhom) != ""
    )
    
    paste0(
      "Tổng số tên bệnh phẩm từ WHONET: ", total, "\n",
      "Đã chuẩn hóa đầy đủ: ", done, "\n",
      "Chưa chuẩn hóa / thiếu thông tin: ", total - done
    )
  })
  
  # ---------- Check missing ----------
  observeEvent(input$check_missing, {
    joined <- raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(ref_df(), by = ".key")
    
    miss <- joined %>%
      filter(
        is.na(ten_benh_pham) | trimws(ten_benh_pham) == "" |
          is.na(ten_benh_pham_nhom) | trimws(ten_benh_pham_nhom) == ""
      ) %>%
      transmute(
        ma_hoa = ma_hoa_raw,
        ten_benh_pham = ifelse(is.na(ten_benh_pham), "", ten_benh_pham),
        ten_benh_pham_nhom = ifelse(is.na(ten_benh_pham_nhom), "", ten_benh_pham_nhom)
      ) %>%
      arrange(ma_hoa)
    
    missing_editable(miss)
    updated_ref(NULL)
  })
  
  # ---------- UI folded ----------
  output$missing_folded_ui <- renderUI({
    df <- missing_editable()
    if (is.null(df)) return(NULL)
    
    tags$details(
      open = TRUE,
      tags$summary(
        "Danh sách cần chuẩn hóa",
        span(class = "badge", nrow(df))
      ),
      DTOutput("missing_tbl")
    )
  })
  
  # ---------- Editable table ----------
  output$missing_tbl <- renderDT({
    datatable(
      missing_editable(),
      rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options = list(
        dom = "t",
        paging = FALSE,
        scrollY = "45vh",
        language = list(
          emptyTable = "Không có dữ liệu để hiển thị"
        )
      )
    )
  })
  
  observeEvent(input$missing_tbl_cell_edit, {
    info <- input$missing_tbl_cell_edit
    df <- missing_editable()
    df[info$row, info$col + 1] <- info$value
    missing_editable(df)
  })
  
  # ---------- Apply updates ----------
  apply_updates_core <- function() {
    work <- updated_ref()
    if (is.null(work)) work <- ref_df()
    
    miss <- missing_editable()
    if (is.null(miss)) return()
    
    to_apply <- miss %>%
      filter(trimws(ten_benh_pham) != "")
    
    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)
      
      if (length(idx) > 0) {
        work$ten_benh_pham[idx] <- to_apply$ten_benh_pham[i]
        work$ten_benh_pham_nhom[idx] <- to_apply$ten_benh_pham_nhom[i]
      } else {
        work <- bind_rows(
          work,
          tibble(
            ma_hoa = to_apply$ma_hoa[i],
            ten_benh_pham = to_apply$ten_benh_pham[i],
            ten_benh_pham_nhom = to_apply$ten_benh_pham_nhom[i],
            .key = key
          )
        )
      }
    }
    
    updated_ref(work)
    
    missing_editable(
      miss %>%
        filter(
          trimws(ten_benh_pham) == "" |
            trimws(ten_benh_pham_nhom) == ""
        )
    )
  }
  
  observeEvent(input$apply_updates, {
    apply_updates_core()
  })
  
  # ---------- Export ----------
  observeEvent(input$export_new, {
    apply_updates_core()
    
    out <- updated_ref()
    if (is.null(out)) out <- ref_df()
    
    writexl::write_xlsx(
      out %>% select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom),
      REF_FILE
    )
    
    showModal(modalDialog(
      title = "Hoàn tất",
      paste("Đã ghi đè file:", REF_FILE),
      easyClose = TRUE
    ))
  })
  
  # ---------- Preview ----------
  output$ref_preview <- renderDT({
    req(updated_ref())
    datatable(
      updated_ref() %>%
        select(ma_hoa, ten_benh_pham, ten_benh_pham_nhom),
      rownames = FALSE
    )
  })
}

shinyApp(ui, server)
