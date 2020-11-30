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

  dateformat <- data.frame(Date=sort(unique(data$Date)),
                           Quarter=rep(NA,length(unique(data$Date))),
                           Year=rep(NA,length(unique(data$Date))))
  dateformat$Quarter <- as.numeric(substr(dateformat$Date[nrow(dateformat)],5,6))/3
  timer <- nrow(dateformat)
  refYr <- as.numeric(substr(dateformat$Date[nrow(dateformat)],3,4))
  iter <- nrow(dateformat)/4
  for (i in 1:iter){
    for (t in 1:4) {
      dateformat$Year[timer+1-t] <- refYr
    }
    timer <- timer - 4
    refYr <- refYr -1
  }

  dataQ <- data %>%
    left_join(dateformat, by = "Date") %>%
    mutate(MATY=str_sub(data$MAT,1,2),
           QT=as.numeric(str_sub(data$Date,5,6))/3) %>%
    mutate(MAT = paste0(Year,'Q', Quarter,' MAT'),
           YTD=paste0(Year,'Q', Quarter,' YTD'),
           MTH=paste0(MATY,'Q',QT)) %>%
    select(-Quarter, -Year,-QT,-MATY)



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
    filter(!is.na(period),
           period %in% tail(unique(period),12)) %>%
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

  table.file <- DisplayFunction(table.file=table.file, type='MTH')
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))

  } else {

      rollparam <- 4
      table.file <- dataQ %>%
        group_by(PeriodMAT = !!sym(grouper),
                 PeriodMTH = MTH,
                 summary = !!sym(unique(form$Summary1))) %>%
        summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                    digit) %>%
        ungroup() %>%
        filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                   min(16,length(unique(PeriodMTH))))) %>%
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
        table.file <- DisplayFunction(table.file=table.file,
                                      type=unique(form$Period))
        write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
      } else {
        print ('Warning: Insufficient Data to Calculate
               Rolling MAT Growth Rates!')
      }
    }
}

