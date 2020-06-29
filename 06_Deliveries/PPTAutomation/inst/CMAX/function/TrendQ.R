# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

Trend <- function(data,
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

  table.file <- dataQ %>%
    filter(!!sym(grouper) %in%
             tail(sort(unique(unlist(dataQ[, grouper]))),
                  min(8,length(unique(unlist(dataQ[, grouper])))))) %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in%
                              unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others"))

  if (unique(form$Period)=='MTH') {
    table.file <- table.file %>%
      group_by(period = !!sym(unique(form$Period)), summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      arrange(period) %>%
      setDT() %>%
      dcast(summary ~ period, value.var = "value") %>%
      right_join(distinct(form, Display), by = c("summary" = "Display")) %>%
      rename(` ` = summary)

  } else  {

    rollparam <- 4
    table.file <- table.file %>%
      group_by(PeriodMAT = !!sym(grouper),
               PeriodMTH = MTH,
               summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(12,length(unique(PeriodMTH))))) %>%
      arrange(summary,PeriodMAT, PeriodMTH) %>%
      group_by(summary) %>%
      mutate(RollValue=c(rep(NA,(rollparam-1)),rollsum(value,rollparam))) %>%
      ungroup() %>% na.omit() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(8,length(unique(PeriodMTH))))) %>%
      mutate(period = paste(PeriodMTH,unique(form$Period),sep=' ')) %>%
      select(period,summary, RollValue,-value,-PeriodMAT,-PeriodMTH) %>%
      spread(period,RollValue) %>%
      adorn_totals("row", na.rm = TRUE, name = "Total") %>%
      mutate(sequence = ifelse(summary == 'Others' ,
                               2, ifelse(summary=='Total',3,1))) %>%
      arrange(sequence,summary) %>%
      select(-sequence) %>%
      rename(' ' = summary)

    table.file <- DisplayFunction(table.file=table.file,
                                  type=unique(form$Period))
  }

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}
