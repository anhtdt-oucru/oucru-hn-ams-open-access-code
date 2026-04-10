
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
  library(janitor)
})

# ====== PATH ======
RAW_DAT <- "G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/PTH/3.du_lieu_whonet_phien_giai" 

OUTPUT_DAT <- "G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/PTH/4.du_lieu_whonet_phien_giai_da_lam_sach"

REF_ORG <- read_xlsx("G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/PTH/1.danh_muc/danh_muc_vsv.xlsx")

REF_RESIS <- read_xlsx("G:/Group Folders/ACADEMIC/Controlled documents/60HN - ASPARNet/7. DATA/ANTIBIOGRAM/PTH/1.danh_muc/danh_muc_co_che_khang.xlsx")

# ====== PROCESS FUNCTION ======
process_file <- function(file_path) {
  
  df <- read_excel(file_path, skip = 7) %>% clean_names()
  
  # ===== REMOVE COLUMNS =====
  cols_remove <- c("macro_name", "ten_macro", "country", "quoc_gia", "laboratory", "phong_xet_nghiem", "origin", "nguon_goc", "date_of_birth", "ngay_sinh", "age_category", "loai_tuoi", "ward", "vung", "institution", "vien", "department", "location_type", "loai_vung", "local_specimen_code", "vung_ma_benh_pham", "specimen_type_numeric", "loai_benh_pham_dang_ma_so", "reason", "ly_do", "isolate_number", "so_phan_lap", "local_organism_code", "vung_ma_vi_khuan", "organism_type", "loai_vi_khuan", "serotype", "kieu_huyet_thanh", "mrsa", "mu_hon", "vre", "vang", "beta_lactamase", "esbl", "carbapenem_resistance", "khang_carbapenem", "mrsa_screening_test", "kiem_tra_khang_mrsa", "inducible_clindamycin_resistance", "ket_luan_khang_clindamycin", "comment", "ghi_chu", "date_of_data_entry", "ngay_vao_du_lieu", "ngay_tra_ket_qua")
  
  df <- df %>% select(-any_of(cols_remove))
  
  # ===== Mapping tên cột VN -> EN =====
  col_mapping <- c(
    "so_benh_an"         = "identification_number",
    "ho"                 = "last_name",
    "ten"                = "first_name",
    "gioi_tinh"          = "sex",
    "tuoi"               = "age",
    "khoa"               = "location",
    "khoa"               = "department",
    "ngay_nhap_vien"     = "date_of_admission",
    "so_benh_pham"       = "specimen_number",
    "ngay_lay_benh_pham" = "specimen_date",
    "loai_benh_pham"     = "specimen_type",
    "vi_khuan"           = "organism"
  )
  
  # ===== HÀM chuẩn hóa tên cột =====
  standardize_colnames <- function(df) {
    current_names <- names(df)

    new_names <- ifelse(current_names %in% names(col_mapping),
                        col_mapping[current_names],
                        current_names)
    
    names(df) <- new_names
    return(df)
  }
  

  df <- standardize_colnames(df)
  
  # ===== FIX DATE =====
  df <- df %>%
    mutate(
      specimen_date = case_when(
        is.numeric(specimen_date) ~ as.Date(specimen_date, origin = "1899-12-30"),
        TRUE ~ as.Date(specimen_date, format = "%m/%d/%Y")
      ),
      date_of_admission = case_when(
        is.numeric(date_of_admission) ~ as.Date(date_of_admission, origin = "1899-12-30"),
        TRUE ~ as.Date(date_of_admission, format = "%m/%d/%Y")
      )
    )
  
  # ===== REMOVE NA DATE =====
  df <- df %>% filter(!is.na(specimen_date))
  
  # ===== ADD YEAR =====
  df <- df %>%
    mutate(nam_nuoi_cay = format(specimen_date, "%Y"))
  
  # ===== GROUP ISOLATE + MIN DATE =====
  df <- df %>%
    group_by(
      identification_number, last_name, first_name, sex,
      age, nam_nuoi_cay, specimen_type, organism) %>%
    slice_min(specimen_date, with_ties = FALSE) %>%
    ungroup()
  
  # ===== KEY COLUMNS =====
  key_cols <- c("identification_number", "first_name", "last_name", "sex",
                "age", "location", "date_of_admission", "specimen_date", "nam_nuoi_cay", "specimen_type", 
                "specimen_number", "organism")
  
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
      ma_bn = identification_number,
      ho_dem = last_name,
      ten_bn = first_name,
      gioi_tinh = sex,
      tuoi = age,
      ma_khoa = location,
      ngay_nhap_vien = date_of_admission,
      ma_benh_pham = specimen_number,
      ngay_nuoi_cay = specimen_date,
      ten_benh_pham = specimen_type,
      ma_vsv = organism
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