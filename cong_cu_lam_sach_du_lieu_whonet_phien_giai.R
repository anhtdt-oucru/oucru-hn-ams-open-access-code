
# =========================
# Làm sạch dữ liệu WHONET
# =========================
suppressPackageStartupMessages({
  library(shiny)
  library(dplyr)
  library(readxl)
  library(tidyr)
  library(writexl)
  library(stringr)
})

# ====== PATH ======
RAW_DAT <- "G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/DTH/3.du_lieu_whonet_phien_giai" 

OUTPUT_DAT <- "G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/DTH/4.du_lieu_whonet_phien_giai_da_lam_sach"

REF_ORG <- read_xlsx("G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/DTH/1.danh_muc/DTH_danh_muc_vsv.xlsx")

REF_RESIS <- read_xlsx("G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/DTH/1.danh_muc/DTH_danh_muc_co_che_khang.xlsx")

# ====== PROCESS FUNCTION ======
process_file <- function(file_path) {
  
  df <- read_excel(file_path, skip = 7)
  
  # ===== REMOVE COLUMNS =====
  cols_remove <- c("Macro name", "Country", "Laboratory", "Origin", "Date of birth", "Age category", "Institution", 
                   "Department", "Location type", "Local specimen code", "Specimen type (Numeric)", "Reason", 
                   "Isolate number", "Local organism code", "Organism type", "Serotype", 
                   "MRSA", "VRE", "Beta-lactamase", "ESBL", "Carbapenem resistance", "MRSA screening test", 
                   "Inducible clindamycin resistance", "Comment", "Date of data entry")
  
  df <- df %>% select(-any_of(cols_remove))
  
  # ===== FIX DATE =====
  df <- df %>%
    mutate(
      `Specimen date` = case_when(
        is.numeric(`Specimen date`) ~ as.Date(`Specimen date`, origin = "1899-12-30"),
        TRUE ~ as.Date(`Specimen date`, format = "%m/%d/%Y")
      ),
      `Date of admission` = case_when(
        is.numeric(`Date of admission`) ~ as.Date(`Date of admission`, origin = "1899-12-30"),
        TRUE ~ as.Date(`Date of admission`, format = "%m/%d/%Y")
      )
    )
  
  # ===== REMOVE NA DATE =====
  df <- df %>% filter(!is.na(`Specimen date`))
  
  # ===== ADD YEAR =====
  df <- df %>%
    mutate(nam_nuoi_cay = format(`Specimen date`, "%Y"))
  
  # ===== GROUP ISOLATE + MIN DATE =====
  df <- df %>%
    mutate(
      id_up = toupper(`Identification number`),
      last_up = toupper(`Last name`),
      first_up = toupper(`First name`),
      sex_up = toupper(`Sex`),
      specimen_up = toupper(`Specimen type`),
      org_up = toupper(`Organism`)
    ) %>%
    group_by(
      id_up, last_up, first_up, sex_up,
      Age, nam_nuoi_cay,
      specimen_up, org_up
    ) %>%
    slice_min(`Specimen date`, with_ties = FALSE) %>%
    ungroup() %>%
    select(-ends_with("_up"))
  
  # ===== KEY COLUMNS =====
  key_cols <- c("Identification number", "First name", "Last name", "Sex",
                "Age", "Location", "Date of admission", "Specimen date", "nam_nuoi_cay", "Specimen type", 
                "Specimen number", "Organism")
  
  key_cols <- intersect(key_cols, names(df))
  
  # ===== PIVOT =====
  df_long <- df %>%
    pivot_longer(
      cols = -all_of(key_cols),
      names_to = "khang_sinh",
      values_to = "kq_ksd"
    )
  
  # ===== STANDARDIZE =====
  df_long <- df_long %>%
    rename(
      ma_bn = `Identification number`,
      ho_dem = `Last name`,
      ten_bn = `First name`,
      gioi_tinh = Sex,
      tuoi = Age,
      ma_khoa = Location,
      ngay_nhap_vien = `Date of admission`,
      ma_benh_pham = `Specimen number`,
      ngay_nuoi_cay = `Specimen date`,
      ten_benh_pham = `Specimen type`,
      ma_vsv = Organism
    ) %>%
    mutate(
      khang_sinh = str_split_fixed(khang_sinh, "_", 2)[,1],
      ma_bn = as.character(ma_bn),
      ma_vsv = str_to_lower(str_trim(ma_vsv)),
      khang_sinh = str_to_upper(str_trim(khang_sinh)),
      kq_ksd = str_to_upper(str_trim(kq_ksd))
    )
  
  # ===== JOIN REF =====
  df_long <- df_long %>%
    left_join(
      REF_ORG %>% select(ma_hoa, ten_vsv) %>% distinct(),
      by = c("ma_vsv" = "ma_hoa")
    ) %>%
    left_join(
      REF_RESIS,
      by = c(
        "ten_vsv",
        "khang_sinh",
        "kq_ksd" = "ket_qua_ksd"
      )
    )
  
  return(df_long)
}

# ====== LOAD ALL FILES ======
load_data <- function() {
  
  files <- list.files(RAW_DAT, pattern = "\\.xlsx$", full.names = TRUE)
  
  all_data <- list()
  
  for (f in files) {
    try({
      df <- process_file(f)
      df$source_file <- basename(f)
      all_data[[length(all_data) + 1]] <- df
    }, silent = TRUE)
  }
  
  bind_rows(all_data)
}

# ====== UI ======
ui <- fluidPage(
  titlePanel("Làm sạch dữ liệu WHONET phiên giải"),
  
  sidebarLayout(
    sidebarPanel(
      actionButton("reload", "Tải lại dữ liệu"),
      br(), br(),
      actionButton("export", "Xuất dữ liệu đã làm sạch")
    ),
    
    mainPanel(
      verbatimTextOutput("status")
    )
  )
)

# ====== SERVER ======
server <- function(input, output, session) {
  
  data_all <- reactiveVal(NULL)
  
  # AUTO LOAD
  observeEvent(TRUE, {
    withProgress(message = "Đang load dữ liệu...", value = 0, {
      data_all(load_data())
    })
  }, once = TRUE)
  
  # RELOAD
  observeEvent(input$reload, {
    withProgress(message = "Đang tải lại dữ liệu...", value = 0, {
      data_all(load_data())
    })
  })
  
  # STATUS
  output$status <- renderText({
    df <- data_all()
    if (is.null(df)) return("Chưa load dữ liệu")
    
    paste0(
      "Số dòng: ", nrow(df), "\n",
      "Số cột: ", ncol(df)
    )
  })
  
  # EXPORT
  observeEvent(input$export, {
    
    df <- data_all()
    req(df)
    
    years <- na.omit(format(df$ngay_nuoi_cay, "%Y"))
    
    file_name <- paste0(
      "whonet_phien_giai_da_lam_sach_",
      min(years), "_", max(years), ".xlsx"
    )
    
    file_path <- file.path(OUTPUT_DAT, file_name)
    
    write_xlsx(df, file_path)
    
    showModal(modalDialog(
      title = "Hoàn tất",
      paste("Đã lưu file:", file_path),
      easyClose = TRUE
    ))
  })
}

shinyApp(ui, server)