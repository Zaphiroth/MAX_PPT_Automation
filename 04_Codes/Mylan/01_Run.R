# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Run
# programmer:   Zhe Liu
# Date:         2020-05-12
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


##---- Readin ----
raw.data <- data.frame(read_excel("02_Inputs/Raw data-CMAX.xlsx"), 
                       check.names = FALSE)

form.table <- data.frame(read_excel("03_Outputs/Form_table-CMAX.xlsx"), 
                         check.names = FALSE)


##---- Period ----
if(grepl('Q',raw.data$Date[9],fixed=TRUE)==TRUE) {
  raw.data <- raw.data %>% mutate(Year=str_sub(Date,1,4),Month=as.numeric(str_sub(Date,nchar(raw.data$Date[1])))) %>% 
    mutate(Month=str_pad(as.character(Month*3),2,pad='0')) %>% mutate(Date=paste0(Year,Month)) %>% select(-Year,-Month)
  cycle <- 4
} else {
  cycle <- 12
}

latest.date <- sort(unique(raw.data$Date), decreasing = TRUE)


mat.date <- data.frame(Date = sort(unique(raw.data$Date))) %>% 
  mutate(ymd = ymd(stri_paste(Date, "01")),
         MAT = case_when(
           ymd <= max(ymd) & ymd > max(ymd)-months(x = 12) ~ 
             stri_paste(stri_sub(latest.date[1], 3, 4), "M", stri_sub(latest.date[1], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 12) & ymd > max(ymd)-months(x = 24) ~
             stri_paste(stri_sub(latest.date[1+cycle], 3, 4), "M", stri_sub(latest.date[1+cycle], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 24) & ymd > max(ymd)-months(x = 36) ~
             stri_paste(stri_sub(latest.date[1+2*cycle], 3, 4), "M", stri_sub(latest.date[1+2*cycle], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 36) & ymd > max(ymd)-months(x = 48) ~
             stri_paste(stri_sub(latest.date[1+3*cycle], 3, 4), "M", stri_sub(latest.date[1+3*cycle], 5, 6), " MAT"),
           ymd <= max(ymd)-months(x = 48) & ymd > max(ymd)-months(x = 60) ~
             stri_paste(stri_sub(latest.date[1+4*cycle], 3, 4), "M", stri_sub(latest.date[1+4*cycle], 5, 6), " MAT"),
           TRUE ~ NA_character_
         )) %>% 
  inner_join(raw.data, by = "Date") %>% 
  distinct(Date, MAT) %>% 
  add_count(MAT, name = "n") %>% 
  filter(n == cycle) %>% 
  select(Date, MAT)


ytd.date <- data.frame(Date = sort(unique(raw.data$Date))) %>% 
  mutate(ymd = ymd(stri_paste(Date, "01")),
         YTD = case_when(
           ymd <= max(ymd) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1], 1, 4) ~ 
             stri_paste(stri_sub(latest.date[1], 3, 4), "M", stri_sub(latest.date[1], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 12) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1+cycle], 1, 4) ~
             stri_paste(stri_sub(latest.date[1+cycle], 3, 4), "M", stri_sub(latest.date[1+cycle], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 24) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1+2*cycle], 1, 4) ~
             stri_paste(stri_sub(latest.date[1+2*cycle], 3, 4), "M", stri_sub(latest.date[1+2*cycle], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 36) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1+3*cycle], 1, 4) ~
             stri_paste(stri_sub(latest.date[1+3*cycle], 3, 4), "M", stri_sub(latest.date[1+3*cycle], 5, 6), " YTD"),
           ymd <= max(ymd)-months(x = 48) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1+4*cycle], 1, 4) ~
             stri_paste(stri_sub(latest.date[1+4*cycle], 3, 4), "M", stri_sub(latest.date[1+4*cycle], 5, 6), " YTD"),
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
if (cylce==4) {
  source("04_Codes/Maylan/02_MarketSizeQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/03_SubMarketShareQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/04_SubMarketGrowthQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/05_PerformanceQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/06_TrendQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/07_ShareTrendQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/08_RankingQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/09_RegionPerformanceQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/10_RegionTrendQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/11_CompetitorPerformanceQ.R", encoding = "UTF-8")
  source("04_Codes/Maylan/12_GrowthTrendQ.R", encoding = "UTF-8") 

} else {
  source("04_Codes/Maylan/02_MarketSize.R", encoding = "UTF-8")
  source("04_Codes/Maylan/03_SubMarketShare.R", encoding = "UTF-8")
  source("04_Codes/Maylan/04_SubMarketGrowth.R", encoding = "UTF-8")
  source("04_Codes/Maylan/05_Performance.R", encoding = "UTF-8")
  source("04_Codes/Maylan/06_Trend.R", encoding = "UTF-8")
  source("04_Codes/Maylan/07_ShareTrend.R", encoding = "UTF-8")
  source("04_Codes/Maylan/08_Ranking.R", encoding = "UTF-8")
  source("04_Codes/Maylan/09_RegionPerformance.R", encoding = "UTF-8")
  source("04_Codes/Maylan/10_RegionTrend.R", encoding = "UTF-8")
  source("04_Codes/Maylan/11_CompetitorPerformance.R", encoding = "UTF-8")
  source("04_Codes/Maylan/12_GrowthTrend.R", encoding = "UTF-8")
} 

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
    
  } 
  
}


map(unique(form.table$Page),
    GenerateFile,
    data = data,
    form.table = form.table,
    directory = "05_Internal_Review/test1")




