# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-05-22
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code P23/24 RegionTrend

RegionTrend <- function(data,
                        form,
                        page,
                        digit,
                        directory) {
  
  dataQ <- data %>% mutate(MATY=str_sub(data$MAT,1,2), YTDY=str_sub(data$YTD,1,2), QT=as.numeric(str_sub(data$Date,5,6))/3) %>%
    mutate(MAT = paste0(MATY,'Q4 MAT'), YTD=paste0(YTDY,'Q4 YTD'), MTH=paste0(MATY,'Q',QT)) %>%
    select(-MATY, -YTDY, -QT)
  
  if(unique(form$Period)=='MTH') {
    grouper <- 'MTH' } else {
      grouper <- 'MAT'  
    }
  
  if (unique(form$Period)=='MTH') {  
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1))) %>%
    summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
    ungroup() %>% 
    mutate (Value= Value/digit) %>%
    filter(period %in% tail(unique(period),12)) %>%
    mutate(month=str_sub(period,-1,-1),year=str_sub(period,1,2)) %>%
    group_by(period) %>%
    mutate (total=sum(Value, na.rm = TRUE)) %>%
    ungroup() %>%
    mutate(Share = Value/total) %>%
    arrange(month,region) %>%
    group_by (month,region) %>%
    mutate (EI=(Share / lag(Share))*100) %>% 
    ungroup() %>%
    na.omit() %>%
    select(region,period,Share,EI) %>%
    gather(category,value,Share:EI) %>%
    mutate(Name= paste(region,category,sep="_")) %>%
    spread(period,value) %>%
    arrange(desc(category),desc(Name)) %>%
    select(Name, everything(),-category,-region) %>%
    rename(' ' = Name) 
  
  source("04_Codes/Maylan/14_DisplayFunctionQ.R", encoding = "UTF-8")
  table.file <- DisplayFunction(table.file=table.file, type='MTH')
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
  
  } else {
      
      rollparam <- 4
      table.file <- dataQ %>%
        group_by(PeriodMAT = !!sym(grouper),
                 PeriodMTH = MTH,
                 summary = !!sym(unique(form$Summary1))) %>%
        summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
        ungroup() %>%
        filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),min(16,length(unique(PeriodMTH))))) %>%
        arrange(summary,PeriodMAT, PeriodMTH) %>%
        group_by(summary) %>%
        mutate(RollValue=c(rep(NA,(rollparam-1)),rollsum(value,rollparam))) %>%
        ungroup() %>% na.omit() %>%
        mutate(period = paste(PeriodMTH,unique(form$Period),sep=' ')) %>%
        select(period,summary, RollValue,-value,-PeriodMAT,-PeriodMTH) %>%
        group_by(period) %>% 
        mutate(Share = RollValue / sum(RollValue, na.rm = TRUE)) %>% 
        ungroup() %>% 
        select(-RollValue) %>%
        mutate(month=str_sub(period,4,4),year=str_sub(period,1,2)) %>%
        arrange(month,summary) %>%
        group_by (month,summary) %>%
        mutate (EI=(Share / lag(Share))*100) %>% 
        ungroup() %>%
        na.omit() %>%
        select(summary,period,Share,EI) %>%
        gather(category,value,Share:EI) %>%
        mutate(Name= paste(summary,category,sep="_")) %>%
        spread(period,value) %>% 
        arrange(desc(category),desc(Name)) %>%
        select(Name, everything(),-category,-summary)  %>%
        rename(' ' = Name) 
      
      if(ncol(table.file)>9) {
        table.file <- table.file[,-(ncol(table.file)-8)]
      }
      
      if (nrow(table.file) != 0) {
        source("04_Codes/Maylan/14_DisplayFunctionQ.R", encoding = "UTF-8")
        table.file <- DisplayFunction(table.file=table.file,type=unique(form$Period))
        write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
      } else {
        print ('Warning: Insufficient Data to Calculate Rolling MAT Growth Rates!')
      }
    }
}

