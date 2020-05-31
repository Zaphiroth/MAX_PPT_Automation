# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Run
# programmer:   Zhe Liu
# Date:         2020-05-12
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


##---- Readin ----
raw.data <- data.frame(read_excel("02_Inputs/Trend_Mylan_ELID_delivery_201801-202003_v0527.xlsx"), 
                       check.names = FALSE)

form.table <- data.frame(read_excel("03_Outputs/Form_ELID trend.xlsx"), 
                         check.names = FALSE)


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
  inner_join(raw.data, by = "Date") %>% 
  distinct(Date, MAT) %>% 
  add_count(MAT, name = "n") %>% 
  filter(n == 12) %>% 
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
  distinct() %>% 
  left_join(mat.date, by = "Date") %>% 
  left_join(ytd.date, by = "Date") %>% 
  left_join(mth.date, by = "Date")


##---- Generation function ----
source("04_Codes/Mylan/02_MarketSize.R", encoding = "UTF-8")
source("04_Codes/Mylan/03_SubMarketShare.R", encoding = "UTF-8")
source("04_Codes/Mylan/04_SubMarketGrowth.R", encoding = "UTF-8")
source("04_Codes/Mylan/05_Performance.R", encoding = "UTF-8")
source("04_Codes/Mylan/06_Trend.R", encoding = "UTF-8")
source("04_Codes/Mylan/07_ShareTrend.R", encoding = "UTF-8")
source("04_Codes/Mylan/08_Ranking.R", encoding = "UTF-8")
source("04_Codes/Mylan/09_RegionPerformance.R", encoding = "UTF-8")
source("04_Codes/Mylan/10_RegionTrend.R", encoding = "UTF-8")
source("04_Codes/Mylan/11_CompetitorPerformance.R", encoding = "UTF-8")
source("04_Codes/Mylan/12_GrowthTrend.R", encoding = "UTF-8")
source("04_Codes/Mylan/13_GrowthTrendMAT.R", encoding = "UTF-8")

GenerateFile <- function(page,
                         data,
                         form.table,
                         directory) {
  ##---- Identify table type ----
  form <- form.table %>% 
    filter(Page == page)
  
  type <- form %>% 
    distinct(Type) %>% 
    unlist()
  
  if (!is.na(unique(form$Restriction))) {
    restriction <- stri_split_fixed(unique(form$Restriction), ",", simplify = TRUE) %>% 
      as.character()
    
    for (i in restriction) {
      rst <- stri_split_fixed(i, ":", simplify = TRUE) %>% 
        as.character()
      
      data <- data %>% 
        filter(!!sym(rst[1]) == rst[2])
    }
  }
  
  if (is.na(unique(form$Digit))) {
    digit <- 1
    
  } else if (unique(form$Digit) == "K") {
    digit <- 1000
    
  } else if (unique(form$Digit) == "Mn") {
    digit <- 1000000
    
  } else {
    digit <- 1
  }
  
  ##---- Calculate function ----
  if (type == "MarketSize") {
    print(page)
    MarketSize(data, form, page, digit, directory)
    
  } else if (type == "SubMarketShare") {
    print(page)
    SubMarketShare(data, form, page, digit, directory)
    
  } else if (type == "SubMarketGrowth") {
    print(page)
    SubMarketGrowth(data, form, page, digit, directory)
    
  } else if (type == "Performance") {
    print(page)
    Performance(data, form, page, digit, directory)
    
  } else if (type == "Trend") {
    print(page)
    Trend(data, form, page, digit, directory)
    
  } else if (type == "ShareTrend") {
    print(page)
    ShareTrend(data, form, page, digit, directory)
    
  } else if (type == "Ranking") {
    print(page)
    Ranking(data, form, page, digit, directory)
    
  } else if (type == "RegionPerformance") {
    print(page)
    RegionPerformance(data, form, page, digit, directory)
    
  } else if (type == "RegionTrend") {
    print(page)
    RegionTrend(data, form, page, digit, directory)
    
  } else if (type == "CompetitorPerformance") {
    print(page)
    CompetitorPerformance(data, form, page, digit, directory)
    
  } else if (type == "GrowthTrend") {
    print(page)
    GrowthTrend(data, form, page, digit, directory)
    
  } else if (type == "GrowthTrendMAT") {
    print(page)
    GrowthTrendMAT(data, form, page, digit, directory)
    
  } 
}


map(unique(form.table$Page),
    GenerateFile,
    data = data,
    form.table = form.table,
    directory = "05_Internal_Review/test1")




