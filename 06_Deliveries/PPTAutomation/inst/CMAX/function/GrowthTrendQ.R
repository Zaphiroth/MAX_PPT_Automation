# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code :  Growth Trend (MAT/RQ)

GrowthTrend <- function(data,
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



  table.file <- dataQ %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in%
                              unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others"))

  if (unique(form$Period)=='MTH') {
    table.file <- table.file %>%
      group_by(period = !!sym(unique(form$Period)),
               summary) %>%
      summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
      ungroup() %>%
      mutate (Value= Value/digit) %>%
      filter(period %in% tail(sort(unique(period)),
                              min(12,length(unique(period))))) %>%
      spread(period,Value) %>%
      adorn_totals("row", na.rm = TRUE, name = "Total") %>%
      gather(period,Value,-summary) %>%
      mutate(month=str_sub(period,-1,-1),year=str_sub(period,1,2)) %>%
      arrange(month,summary) %>%
      group_by (month,summary) %>%
      mutate (Growth=(Value / lag(Value))-1)%>%
      ungroup() %>%
      na.omit() %>%
      select(summary,period,Growth) %>%
      spread(period,Growth) %>%
      mutate(sequence = ifelse(summary == 'Others' ,
                               2, ifelse(summary=='Total',3,1))) %>%
      arrange(sequence,summary) %>%
      select(-sequence) %>%
      rename(' ' = summary) %>%
      mutate(` ` = factor(` `, levels = form$Display)) %>%
      filter(!is.na(` `)) %>%
      arrange(` `)

    table.file <- DisplayFunction(table.file=table.file,type='MTH')
    write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))

  } else {

    rollparam <- 4
    table.file <- table.file %>%
      group_by(PeriodMAT = !!sym(unique(form$Period)),
               PeriodMTH = MTH,
               summary) %>%
      summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
      ungroup() %>%
      mutate (Value= Value/digit) %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(16,length(unique(PeriodMTH))))) %>%
      arrange(summary,PeriodMAT, PeriodMTH) %>%
      group_by(summary) %>%
      mutate(RollValue=c(rep(NA,(rollparam-1)),rollsum(Value,rollparam))) %>%
      ungroup() %>% na.omit() %>%
      mutate(period = paste(PeriodMTH,'MAT',sep=' ')) %>%
      select(period,summary, RollValue,-Value,-PeriodMAT,-PeriodMTH) %>%
      spread(period,RollValue) %>%
      adorn_totals("row", na.rm = TRUE, name = "Total") %>%
      gather(period,Value,-summary) %>%
      mutate(month=str_sub(period,4,4),year=str_sub(period,1,2)) %>%
      arrange(month,summary) %>%
      group_by (month,summary) %>%
      mutate (Growth=(Value / lag(Value))-1)%>%
      ungroup() %>%
      na.omit() %>%
      select(summary,period,Growth) %>%
      spread(period,Growth) %>%
      mutate(sequence = ifelse(summary == 'Others' ,
                               2, ifelse(summary=='Total',3,1))) %>%
      arrange(sequence,summary) %>%
      select(-sequence) %>%
      rename(' ' = summary)

    if (nrow(table.file) != 0) {
      if(ncol(table.file)>9) {
        table.file <- table.file[,-(ncol(table.file)-8)]
      }
      table.file <- DisplayFunction(table.file=table.file,
                                    type=unique(form$Period))
      write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
    } else {
      print ('Warning: Insufficient Data to Calculate
             Rolling MAT Growth Rates!')
    }
  }
}

