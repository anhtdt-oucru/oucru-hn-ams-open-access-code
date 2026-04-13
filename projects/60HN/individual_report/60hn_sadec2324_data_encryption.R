# Import libraries
library(RODBC)
library(rlang)
library(dplyr)
library(purrr)
library(stringr)
library(readxl)

#### Import raw data
read_data <- function(folder_path, sheet, skip, guess_max = 99999) {
  library(readxl)
  library(dplyr)
  
  # Get list of Excel files
  files <- list.files(path = folder_path, pattern = "\\.xlsx$", full.names = TRUE)
  
  # Read and process all Excel files
  combined_data <- files %>%
    lapply(function(file) {
      read_excel(file, sheet = sheet, skip = skip, guess_max = guess_max, col_types = "text")
    }) %>%
    bind_rows()
  
  return(combined_data)
}

## function to save RDS
saveRDS2 <- function(x){
  name <-  deparse(substitute(x))
  saveRDS(x, 
          paste0("C:\\Users\\tranglq\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Encrypted\\",
                 name, "_", Sys.Date(),".rds"))
}

# --- ADMIN ---
admin_23 <- read_data("C:/Users/tranglnh/OneDrive - Oxford University Clinical Research Unit/Hospital AMS Team - 60HN/60HN-DOCUMENTS/Data/Dong Thap/Sadec/Raw/2023/Admin", 
                      sheet = 1, skip = 0) %>% select(-c(1,2,4)) %>% distinct() %>% 
  dplyr::mutate(across(where(~ !is.character(.x)), as.character))

admin_24 <- read_data("C:/Users/tranglnh/OneDrive - Oxford University Clinical Research Unit/Hospital AMS Team - 60HN/60HN-DOCUMENTS/Data/Dong Thap/Sadec/Raw/2024/Admin", 
                      sheet = 1, skip = 0) %>% select(-c(1,2,4)) %>% distinct() %>% 
  dplyr::mutate(across(where(~ !is.character(.x)), as.character))

# --- 
common_cols <- intersect(names(admin_23), names(admin_24))
# ---
admin_2324 <- bind_rows(
                 admin_23 %>% select(all_of(common_cols)),
                 admin_24 %>% select(all_of(common_cols))) %>% distinct()
admin_2324 <- admin_2324 %>% mutate(ngay_vao = ymd(substr(ngay_vao,1,8)),
                                    ngay_ra = ymd(substr(ngay_ra,1,8)),
                                    ngay_sinh = ymd(substr(ngay_sinh,1,8)))

# --- BED ---
bed_23 <- read_data("C:\\Users\\tranglnh\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Raw\\2023\\Bed",
                     sheet = 1, skip = 0) %>% select(-c(2)) %>% distinct() %>% 
           dplyr::mutate(across(where(~ !is.character(.x)), as.character))

bed_24 <- read_data("C:\\Users\\tranglnh\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Raw\\2024\\Bed",
                     sheet = 1, skip = 0) %>% select(-c(2)) %>% distinct() %>% 
  dplyr::mutate(across(where(~ !is.character(.x)), as.character))
# ---
common_cols <- intersect(names(bed_23), names(bed_24))
# ---
bed_2324 <- bind_rows(
               bed_23 %>% select(all_of(common_cols)),
               bed_24 %>% select(all_of(common_cols))) %>% distinct()

bed_2324 <- bed_2324 %>% 
               mutate(ngay_vaokhoa = ymd(substr(ngay_yl,1,8)),
                      ngay_rakhoa = ymd(substr(ngay_kq,1,8))) %>% 
               select(-c(ngay_yl, ngay_kq)) %>% distinct()


# --- DRUG ---
drug_23 <- read_data("C:/Users/tranglnh/OneDrive - Oxford University Clinical Research Unit/Hospital AMS Team - 60HN/60HN-DOCUMENTS/Data/Dong Thap/Sadec/Raw/2023/Drug", 
                      sheet = 1, skip = 0) %>% select(-c(2)) %>% distinct() %>% 
            dplyr::mutate(across(where(~ !is.character(.x)), as.character))

drug_24 <- read_data("C:/Users/tranglnh/OneDrive - Oxford University Clinical Research Unit/Hospital AMS Team - 60HN/60HN-DOCUMENTS/Data/Dong Thap/Sadec/Raw/2024/Drug", sheet = 1, skip = 0) %>% select(-2) %>% 
  distinct() %>% 
  mutate(across(where(~ !is.character(.x)), as.character))
# ---
common_cols <- intersect(names(drug_23), names(drug_24))
# ---
drug_2324 <- bind_rows(
                drug_23 %>% select(all_of(common_cols)),
                drug_24 %>% select(all_of(common_cols))) %>% distinct()

drug_2324 <- drug_2324 %>% 
                mutate(ngay_yl = ymd(substr(ngay_yl,1,8))) %>% filter(ngay_yl >= ymd(20230101) & ngay_yl <= ymd(20241231))

##Bổ sung ma_bn vào bed/drug
ma_bn <- admin_2324 %>% distinct(ma_lk, ma_bn) 
bed_2324 <- bed_2324 %>% left_join(., ma_bn, by = "ma_lk")
drug_2324 <- drug_2324 %>% left_join(., ma_bn, by = "ma_lk")

#### Encrypt dataset
## function to get initials
name_initial <-  function(x){
                 paste(substr(strsplit(x, " ")[[1]], 1, 1), collapse="")}


## Raw ID
sd_id_dat_2324 <- admin_2324 %>% 
                  select(ma_lk, ma_bn, ho_ten) %>%  distinct()


## Create/encrypt ID
sd_id_dat_2324 <- sd_id_dat_2324 %>% 
                  mutate(id_patient = substring(map_chr(ma_bn, hash), 1, 10),
                         id_link=substring(map_chr(ma_lk, hash), 1, 10),
                         name_patient = map_chr(ho_ten, name_initial))


## export master id data
saveRDS(sd_id_dat_2324, "C:\\Users\\tranglnh\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Raw\\2023\\sd_id_dat_2324.rds")

## encrypt data
# admin
sd_admin_2324_enc <- sd_id_dat_2324 %>% 
                        select(ma_lk,id_patient, id_link, name_patient) %>% 
                        right_join(., admin_2324, by=c("ma_lk"))%>% 
                        select(-c(ma_lk, ma_bn, ho_ten, ma_the, dia_chi))

# bed
sd_bed_2324_enc <- sd_id_dat_2324 %>% 
                      select(ma_lk, id_patient, id_link, name_patient) %>% 
                      right_join(., bed_2324, by=c("ma_lk")) %>% select(-c(ma_lk, ma_bn))

# ---- Extract Inpatients
#
sd_bed_2324_enc <-  sd_bed_2324_enc %>% filter(!is.na(ma_giuong))

#sd_bed_2324_enc <- sd_bed_2324_enc %>% filter(substr(ten_dich_vu,1,6) == "Giường") %>% distinct() 
#match_result_all <- str_match(sd_bed_inp_2324_enc$ten_dich_vu, "(.*?) - (.*)")
#sd_bed_inp_2324_enc$to_del <- str_trim(match_result_all[, 2])
#sd_bed_inp_2324_enc$ward <- str_trim(match_result_all[, 3])
#sd_bed_inp_2324_enc <- sd_bed_inp_2324_enc %>% select(-c(to_del))
#sd_bed_inp_2324_enc$ward[sd_bed_inp_2324_enc$ward=="Khoa nội tổng hợp"] <- "Khoa Nội tổng hợp"


# drug
sd_drug_2324_enc <- sd_id_dat_2324 %>% 
                       select(ma_lk, id_patient, id_link, name_patient) %>% 
                       right_join(., drug_2324, by=c("ma_lk")) %>% distinct() %>% select(-c(ma_lk, ma_bn))


### save encrypted files
saveRDS(sd_admin_2324_enc, "C:\\Users\\tranglnh\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Encrypted\\sd_admin_2324_enc.rds")
saveRDS(sd_bed_2324_enc, "C:\\Users\\tranglnh\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Encrypted\\sd_bed_2324_enc.rds")
saveRDS(sd_drug_2324_enc, "C:\\Users\\tranglnh\\OneDrive - Oxford University Clinical Research Unit\\Hospital AMS Team - 60HN\\60HN-DOCUMENTS\\Data\\Dong Thap\\Sadec\\Encrypted\\sd_drug_2324_enc.rds")








