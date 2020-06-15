# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# Date:         2020-05-26
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code :  Growth Trend (MAT/RQ)

GrowthTrend <- function(data,
                        form,
                        page,
                        digit,
                        directory) {

  table.file <- data %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others"))

  if (unique(form$Period)=='MTH') {
    table.file <- table.file %>%
      group_by(period = !!sym(unique(form$Period)),
               summary) %>%
      summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
      ungroup() %>%
      mutate (Value= Value/digit) %>%
      filter(period %in% tail(sort(unique(period)),min(36,length(unique(period))))) %>%
      spread(period,Value) %>%
      adorn_totals("row", na.rm = TRUE, name = "Total") %>%
      gather(period,Value,-summary) %>%
      mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
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
      right_join(distinct(form, Display), by = c(" " = "Display"))

    table.file <- DisplayFunction(table.file=table.file,type='MTH')
    write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))

  } else {

    if (unique(form$Period)=='MAT') {
      rollparam <- 12
    } else {
      rollparam <- 3
    }

    table.file <- table.file %>%
      group_by(PeriodMAT = !!sym(unique(form$Period)),
               PeriodMTH = MTH,
               summary) %>%
      summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
      ungroup() %>%
      mutate (Value= Value/digit) %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),min(48,length(unique(PeriodMTH))))) %>%
      arrange(summary,PeriodMAT, PeriodMTH) %>%
      group_by(summary) %>%
      mutate(RollValue=c(rep(NA,(rollparam-1)),rollsum(Value,rollparam))) %>%
      ungroup() %>% na.omit() %>%
      mutate(period = paste(PeriodMTH,'MAT',sep=' ')) %>%
      select(period,summary, RollValue,-Value,-PeriodMAT,-PeriodMTH) %>%
      spread(period,RollValue) %>%
      adorn_totals("row", na.rm = TRUE, name = "Total") %>%
      gather(period,Value,-summary) %>%
      mutate(month=str_sub(period,4,5),year=str_sub(period,1,2)) %>%
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
      if(ncol(table.file)>25) {
        table.file <- table.file[,-(ncol(table.file)-24)]
      }
      table.file <- DisplayFunction(table.file=table.file,type=unique(form$Period))
      write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
    } else {
      print ('Warning: Insufficient Data to Calculate Rolling MAT Growth Rates!')
    }
  }
}

