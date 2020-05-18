# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Run
# programmer:   Zhe Liu
# Date:         2020-05-12
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


##---- Readin ----
raw.data <- read.xlsx("02_Inputs/乙肝MAX2016-2019.xlsx")

form.table <- read.xlsx("03_Outputs/Form_table.xlsx")


##---- Period ----
latest.date <- sort(unique(raw.data$Date), decreasing = TRUE)

mat.date <- data.frame(Date = sort(unique(raw.data$Date))) %>% 
  mutate(ymd = ymd(stri_paste(Date, "01")),
         MAT = case_when(
           ymd <= max(ymd) & ymd > max(ymd)-months(x = 12) ~ 
             stri_paste(stri_sub(latest.date[1], 3, 4), "M", stri_sub(latest.date[1], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 12) & ymd > max(ymd)-months(x = 24) ~
             stri_paste(stri_sub(latest.date[13], 3, 4), "M", stri_sub(latest.date[13], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 24) & ymd > max(ymd)-months(x = 36) ~
             stri_paste(stri_sub(latest.date[25], 3, 4), "M", stri_sub(latest.date[25], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 36) & ymd > max(ymd)-months(x = 48) ~
             stri_paste(stri_sub(latest.date[37], 3, 4), "M", stri_sub(latest.date[37], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 48) & ymd > max(ymd)-months(x = 60) ~
             stri_paste(stri_sub(latest.date[49], 3, 4), "M", stri_sub(latest.date[49], 5, 6), " MAT"),
           TRUE ~ NA_character_
         )) %>% 
  select(Date, MAT)

ytd.date <- data.frame(Date = sort(unique(raw.data$Date))) %>% 
  mutate(ymd = ymd(stri_paste(Date, "01")),
         YTD = case_when(
           ymd <= max(ymd) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1], 1, 4) ~ 
             stri_paste(stri_sub(latest.date[1], 3, 4), "M", stri_sub(latest.date[1], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 12) & stri_sub(Date, 1, 4) == stri_sub(latest.date[13], 1, 4) ~
             stri_paste(stri_sub(latest.date[13], 3, 4), "M", stri_sub(latest.date[13], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 24) & stri_sub(Date, 1, 4) == stri_sub(latest.date[25], 1, 4) ~
             stri_paste(stri_sub(latest.date[25], 3, 4), "M", stri_sub(latest.date[25], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 36) & stri_sub(Date, 1, 4) == stri_sub(latest.date[37], 1, 4) ~
             stri_paste(stri_sub(latest.date[37], 3, 4), "M", stri_sub(latest.date[37], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 48) & stri_sub(Date, 1, 4) == stri_sub(latest.date[49], 1, 4) ~
             stri_paste(stri_sub(latest.date[49], 3, 4), "M", stri_sub(latest.date[49], 5, 6), " YTD"),
           TRUE ~ NA_character_
         )) %>% 
  select(Date, YTD)

mth.date <- data.frame(Date = sort(unique(raw.data$Date))) %>% 
  mutate(MTH = stri_paste(stri_sub(Date, 3, 4), "M", stri_sub(Date, 5, 6))) %>% 
  select(Date, MTH)

data <- raw.data %>% 
  left_join(mat.date, by = "Date") %>% 
  left_join(ytd.date, by = "Date") %>% 
  left_join(mth.date, by = "Date") %>% 
  filter(!is.na(MAT), !is.na(YTD))











