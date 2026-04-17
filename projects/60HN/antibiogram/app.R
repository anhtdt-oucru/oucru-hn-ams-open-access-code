# =========================
# Chuẩn hóa Danh mục Vi sinh vật
# =========================

renv::restore()

suppressPackageStartupMessages({
  library(shiny)
  library(DT)
  library(dplyr)
  library(readxl)     # FIXED (was openxlsx)
  library(stringi)
  library(writexl)
  library(tibble)
  library(svDialogs)
})

# ---------- String normalization ----------
normalize_key <- function(x) {
  x <- trimws(as.character(x))
  x <- stringi::stri_trans_general(x, "Latin-ASCII")
  tolower(x)
}

# ---------- Flexible column finder ----------
find_col <- function(df, candidates) {
  cols <- names(df)
  norm_cols <- normalize_key(cols)
  
  for (cand in candidates) {
    idx <- which(norm_cols == normalize_key(cand))
    if (length(idx) > 0) return(cols[idx[1]])
  }
  return(NULL)
}

# =========================
# UI
# =========================
ui <- fluidPage(
  titlePanel("Chuẩn hóa Danh mục Vi sinh vật (V3 FIXED)"),
  
  sidebarLayout(
    sidebarPanel(
      
      h4("1) Nguồn dữ liệu"),
      
      fileInput("raw_files", "Chọn file WHONET (.xlsx)", multiple = TRUE),
      fileInput("ref_file", "Chọn file Danh mục VSV", multiple = FALSE),
      
      hr(),
      
      h4("2) Kiểm tra"),
      actionButton("check_missing", "Kiểm tra mã chưa chuẩn hóa", class = "btn-primary"),
      
      hr(),
      
      h4("3) Áp dụng"),
      actionButton("apply_updates", "Áp dụng cập nhật", class = "btn-success"),
      
      br(), br(),
      
      h4("4) Lưu file"),
      actionButton("save_as_btn", "💾 Lưu đè file Danh mục VSV...", class = "btn-warning"),
      
      br(), br(),
      uiOutput("save_status")
    ),
    
    mainPanel(
      h4("Tổng quan"),
      verbatimTextOutput("summary"),
      
      hr(),
      
      h4("Mã vi sinh vật cần chuẩn hóa"),
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
  
  missing_editable <- reactiveVal(NULL)
  updated_ref      <- reactiveVal(NULL)
  save_message     <- reactiveVal("")
  
  # ---------- REF ----------
  ref_df <- reactive({
    req(input$ref_file)
    
    df <- readxl::read_excel(input$ref_file$datapath)
    
    # Detect columns (EN + VN)
    ma_hoa_col <- find_col(df, c("ma_hoa", "ma hoa", "Mã hóa"))
    ten_vsv_col <- find_col(df, c("ten_vsv", "ten vi sinh vat", "Tên vi sinh vật"))
    loai_vsv_col <- find_col(df, c("loai_vsv", "loai", "Loại vi sinh vật"))
    viet_tat_col <- find_col(df, c("ten_viet_tat", "viet tat", "Tên viết tắt"))
    
    validate(
      need(!is.null(ma_hoa_col), "Thiếu cột mã hóa"),
      need(!is.null(ten_vsv_col), "Thiếu cột tên VSV"),
      need(!is.null(loai_vsv_col), "Thiếu cột loại VSV"),
      need(!is.null(viet_tat_col), "Thiếu cột viết tắt")
    )
    
    df %>%
      transmute(
        ma_hoa       = .data[[ma_hoa_col]],
        ten_vsv      = .data[[ten_vsv_col]],
        loai_vsv     = .data[[loai_vsv_col]],
        ten_viet_tat = .data[[viet_tat_col]],
        .key         = normalize_key(.data[[ma_hoa_col]])
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
      
      # Detect organism column (EN + VN)
      org_col <- find_col(df, c(
        "Organism",
        "Vi sinh vật",
        "Vi sinh vat",
        "Ten vi sinh vat"
      ))
      
      if (is.null(org_col)) return(NULL)
      
      df %>%
        transmute(ma_hoa = as.character(.data[[org_col]]))
    })
    
    all_org <- all_org[!sapply(all_org, is.null)]
    
    validate(
      need(length(all_org) > 0, "Không file nào có cột Organism / Vi sinh vật")
    )
    
    bind_rows(all_org) %>%
      filter(!is.na(ma_hoa) & trimws(ma_hoa) != "") %>%
      distinct() %>%
      mutate(.key = normalize_key(ma_hoa))
  })
  
  # ---------- Summary ----------
  output$summary <- renderText({
    joined <- raw_df() %>%
      left_join(ref_df(), by = ".key")
    
    total <- nrow(joined)
    
    done <- sum(
      !is.na(joined$ten_vsv)      & trimws(joined$ten_vsv)      != "" &
        !is.na(joined$loai_vsv)     & trimws(joined$loai_vsv)     != "" &
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
        ma_hoa       = ma_hoa_raw,
        ten_vsv      = ifelse(is.na(ten_vsv), "", ten_vsv),
        loai_vsv     = ifelse(is.na(loai_vsv), "", loai_vsv),
        ten_viet_tat = ifelse(is.na(ten_viet_tat), "", ten_viet_tat)
      ) %>%
      arrange(ma_hoa)
    
    missing_editable(miss)
    updated_ref(NULL)
  })
  
  # ---------- UI ----------
  output$missing_folded_ui <- renderUI({
    df <- missing_editable()
    if (is.null(df)) return(NULL)
    
    tags$details(
      tags$summary("Danh sách cần chuẩn hóa"),
      DTOutput("missing_tbl")
    )
  })
  
  output$missing_tbl <- renderDT({
    datatable(
      missing_editable(),
      rownames = FALSE,
      editable = list(target = "cell", disable = list(columns = 0)),
      options  = list(dom = "t", paging = FALSE, scrollY = "45vh")
    )
  })
  
  observeEvent(input$missing_tbl_cell_edit, {
    info <- input$missing_tbl_cell_edit
    df   <- missing_editable()
    df[info$row, info$col + 1] <- info$value
    missing_editable(df)
  })
  
  # ---------- Apply ----------
  apply_updates_core <- function() {
    work <- updated_ref()
    if (is.null(work)) work <- ref_df()
    
    miss <- missing_editable()
    if (is.null(miss)) return(work)
    
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
        work$ten_vsv[idx]      <- to_apply$ten_vsv[i]
        work$loai_vsv[idx]     <- to_apply$loai_vsv[i]
        work$ten_viet_tat[idx] <- to_apply$ten_viet_tat[i]
      } else {
        work <- bind_rows(
          work,
          tibble(
            ma_hoa       = to_apply$ma_hoa[i],
            ten_vsv      = to_apply$ten_vsv[i],
            loai_vsv     = to_apply$loai_vsv[i],
            ten_viet_tat = to_apply$ten_viet_tat[i],
            .key         = key
          )
        )
      }
    }
    
    updated_ref(work)
    return(work)
  }
  
  observeEvent(input$apply_updates, {
    apply_updates_core()
    save_message("Đã áp dụng cập nhật.")
  })
  
  # ---------- Save ----------
  observeEvent(input$save_as_btn, {
    
    if (is.null(input$ref_file)) {
      save_message("⚠️ Chưa chọn file danh mục.")
      return()
    }
    
    out <- apply_updates_core()
    
    path <- svDialogs::dlgSave(default = input$ref_file$name)$res
    
    if (is.null(path) || path == "") {
      save_message("Đã hủy lưu file.")
      return()
    }
    
    tryCatch({
      writexl::write_xlsx(
        out %>% select(ma_hoa, ten_vsv, loai_vsv, ten_viet_tat),
        path
      )
      save_message(paste("✅ Đã lưu:", path))
    }, error = function(e) {
      save_message(paste("❌ Lỗi:", e$message))
    })
  })
  
  output$save_status <- renderUI({
    msg <- save_message()
    if (msg == "") return(NULL)
    div(msg)
  })
  
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

shinyApp(ui, server)