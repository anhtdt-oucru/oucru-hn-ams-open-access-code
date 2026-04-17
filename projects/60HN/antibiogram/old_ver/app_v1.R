# =========================
# Chuẩn hóa Danh mục Vi sinh vật
# Full version with Browse Upload Logic
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

# =========================
# UI
# =========================
ui <- fluidPage(
  titlePanel("Chuẩn hóa Danh mục Vi sinh vật"),
  
  tags$style(HTML("
    .drop-zone {
      border: 2px dashed #2c7be5;
      border-radius: 8px;
      padding: 16px;
      text-align: center;
      cursor: pointer;
      background: #f8fbff;
      margin-bottom: 10px;
      font-weight: 600;
    }
    .drop-zone:hover {
      background: #eef6ff;
    }
    #missing_tbl table thead th:nth-child(2),
    #missing_tbl table thead th:nth-child(3),
    #missing_tbl table thead th:nth-child(4) {
      background-color: #FFE999 !important;
      font-weight: 700;
    }
    #missing_tbl table tbody td:nth-child(2),
    #missing_tbl table tbody td:nth-child(3),
    #missing_tbl table tbody td:nth-child(4) {
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
    details summary::-webkit-details-marker {
      display:none;
    }
  ")),
  
  sidebarLayout(
    sidebarPanel(
      
      h4("1) Nguồn dữ liệu"),
      
      div(
        class = "drop-zone",
        onclick = "document.getElementById('raw_files').click()",
        "📂 Chọn file WHONET (.xlsx)"
      ),
      fileInput(
        "raw_files",
        label = NULL,
        multiple = TRUE,
        accept = c(".xlsx")
      ),
      
      br(),
      
      div(
        class = "drop-zone",
        onclick = "document.getElementById('ref_file').click()",
        "📘 Chọn file Danh mục VSV"
      ),
      fileInput(
        "ref_file",
        label = NULL,
        multiple = FALSE,
        accept = c(".xlsx")
      ),
      
      hr(),
      
      h4("2) Kiểm tra"),
      actionButton(
        "check_missing",
        "Kiểm tra mã vi sinh vật chưa chuẩn hóa",
        class = "btn-primary"
      ),
      
      hr(),
      
      h4("3) Áp dụng & Xuất"),
      actionButton(
        "apply_updates",
        "Áp dụng cập nhật",
        class = "btn-success"
      ),
      
      br(), br(),
      
      downloadButton(
        "export_new",
        "Xuất danh mục cập nhật",
        class = "btn-warning"
      )
    ),
    
    mainPanel(
      h4("Tổng quan"),
      verbatimTextOutput("summary"),
      
      hr(),
      
      h4("Mã vi sinh vật cần chuẩn hóa"),
      div(
        class = "hint-box",
        "Điền Tên vi sinh vật, Loại vi sinh vật, Tên viết tắt. Có thể điền dần, không cần đủ mới áp dụng."
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
  
  # ---------- REF ----------
  ref_df <- reactive({
    req(input$ref_file)
    
    df <- readxl::read_excel(input$ref_file$datapath)
    
    validate(
      need("ma_hoa" %in% names(df), "Thiếu cột ma_hoa"),
      need("ten_vsv" %in% names(df), "Thiếu cột ten_vsv"),
      need("loai_vsv" %in% names(df), "Thiếu cột loai_vsv"),
      need("ten_viet_tat" %in% names(df), "Thiếu cột ten_viet_tat")
    )
    
    df %>%
      transmute(
        ma_hoa,
        ten_vsv,
        loai_vsv,
        ten_viet_tat,
        .key = normalize_key(ma_hoa)
      )
  })
  
  # ---------- RAW ----------
  raw_df <- reactive({
    req(input$raw_files)
    
    files <- input$raw_files$datapath
    
    all_org <- lapply(files, function(f) {
      df <- tryCatch(
        readxl::read_excel(f, skip = 7),
        error = function(e) return(NULL)
      )
      
      if (is.null(df)) return(NULL)
      
      names(df) <- trimws(names(df))
      
      if (!"Organism" %in% names(df)) return(NULL)
      
      df %>%
        transmute(ma_hoa = as.character(Organism))
    })
    
    all_org <- all_org[!sapply(all_org, is.null)]
    
    validate(
      need(length(all_org) > 0, "Không file nào có cột Organism")
    )
    
    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa) & trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
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
      !is.na(joined$ten_vsv) & trimws(joined$ten_vsv) != "" &
        !is.na(joined$loai_vsv) & trimws(joined$loai_vsv) != "" &
        !is.na(joined$ten_viet_tat) & trimws(joined$ten_viet_tat) != ""
    )
    
    paste0(
      "Tổng số mã vi sinh vật: ", total, "\n",
      "Đã chuẩn hóa đầy đủ: ", done, "\n",
      "Chưa chuẩn hóa / thiếu thông tin: ", total - done
    )
  })
  
  # ---------- Check Missing ----------
  observeEvent(input$check_missing, {
    
    joined <- raw_df() %>%
      select(ma_hoa_raw = ma_hoa, .key) %>%
      left_join(ref_df(), by = ".key")
    
    miss <- joined %>%
      filter(
        is.na(ten_vsv) | trimws(ten_vsv) == "" |
          is.na(loai_vsv) | trimws(loai_vsv) == "" |
          is.na(ten_viet_tat) | trimws(ten_viet_tat) == ""
      ) %>%
      transmute(
        ma_hoa = ma_hoa_raw,
        ten_vsv = ifelse(is.na(ten_vsv), "", ten_vsv),
        loai_vsv = ifelse(is.na(loai_vsv), "", loai_vsv),
        ten_viet_tat = ifelse(is.na(ten_viet_tat), "", ten_viet_tat)
      ) %>%
      arrange(ma_hoa)
    
    missing_editable(miss)
    updated_ref(NULL)
  })
  
  # ---------- Folded UI ----------
  output$missing_folded_ui <- renderUI({
    df <- missing_editable()
    if (is.null(df)) return(NULL)
    
    tags$details(
      tags$summary(
        "Danh sách cần chuẩn hóa",
        span(class = "badge", nrow(df))
      ),
      DTOutput("missing_tbl")
    )
  })
  
  # ---------- Editable Table ----------
  output$missing_tbl <- renderDT({
    datatable(
      missing_editable(),
      rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options = list(
        dom = "t",
        paging = FALSE,
        scrollY = "45vh"
      )
    )
  })
  
  observeEvent(input$missing_tbl_cell_edit, {
    info <- input$missing_tbl_cell_edit
    df <- missing_editable()
    df[info$row, info$col + 1] <- info$value
    missing_editable(df)
  })
  
  # ---------- Apply Updates ----------
  apply_updates_core <- function() {
    work <- updated_ref()
    if (is.null(work)) work <- ref_df()
    
    miss <- missing_editable()
    if (is.null(miss)) return()
    
    to_apply <- miss %>%
      filter(
        trimws(ten_vsv) != "" |
          trimws(loai_vsv) != "" |
          trimws(ten_viet_tat) != ""
      )
    
    for (i in seq_len(nrow(to_apply))) {
      key <- normalize_key(to_apply$ma_hoa[i])
      idx <- which(work$.key == key)
      
      if (length(idx) > 0) {
        work$ten_vsv[idx] <- to_apply$ten_vsv[i]
        work$loai_vsv[idx] <- to_apply$loai_vsv[i]
        work$ten_viet_tat[idx] <- to_apply$ten_viet_tat[i]
      } else {
        work <- bind_rows(
          work,
          tibble(
            ma_hoa = to_apply$ma_hoa[i],
            ten_vsv = to_apply$ten_vsv[i],
            loai_vsv = to_apply$loai_vsv[i],
            ten_viet_tat = to_apply$ten_viet_tat[i],
            .key = key
          )
        )
      }
    }
    
    updated_ref(work)
    
    missing_editable(
      miss %>%
        filter(
          trimws(ten_vsv) == "" |
            trimws(loai_vsv) == "" |
            trimws(ten_viet_tat) == ""
        )
    )
  }
  
  observeEvent(input$apply_updates, {
    apply_updates_core()
  })
  
  # ---------- Export ----------
  output$export_new <- downloadHandler(
    filename = function() {
      "DTH_danh_muc_vsv_updated.xlsx"
    },
    content = function(file) {
      apply_updates_core()
      
      out <- updated_ref()
      if (is.null(out)) out <- ref_df()
      
      writexl::write_xlsx(
        out %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat),
        file
      )
    }
  )
  
  # ---------- Preview ----------
  output$ref_preview <- renderDT({
    req(updated_ref())
    
    datatable(
      updated_ref() %>%
        select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat),
      rownames = FALSE
    )
  })
}

# =========================
# RUN APP
# =========================
shinyApp(ui, server, options = list(launch.browser = TRUE))