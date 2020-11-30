# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

MarketSize <- function(data,
                       form,
                       page,
                       digit,
                       directory) {

  ## Display Function
  if (unique(form$Period)=='MTH') {
    last.date <- last(sort(unique(data$Date)))
    last.year <- as.numeric(stri_sub(last.date, 1, 4))
    last.mth <- as.numeric(stri_sub(last.date, 5, 6))
    if(last.mth == 12) { tvar <-4 } else {tvar <-5 }
    display.year <- stri_sub((last.year-tvar):last.year, 3, 4)

    first.mth <- last.mth +1
    if(first.mth > 12) {first.mth <- first.mth -12}
    fdmth <- c(first.mth:12)
    ldmth <- c(1:last.mth)
    rowtimer <- length(fdmth) + length(ldmth) + (length(display.year)-2)*12
    if(last.mth == 12) {
      displayyr <- sort(c(rep(display.year,12)))} else {
        displayyr <- c(rep(display.year[1],length(fdmth)),rep(display.year[2],12),
                       rep(display.year[3],12),rep(display.year[4],12),rep(display.year[5],12),
                       rep(display.year[6],length(ldmth)))
      }
    displaymth <- as.character(c(fdmth,rep(1:12,length(display.year)-2),ldmth))
    displaydf <- data.frame(year=displayyr,mth=displaymth) %>%
      mutate(period=paste0(year,"M", str_pad(mth,2,pad='0')))
    display <- data.frame(period=displaydf$period)

  } else {
    last.date <- last(sort(unique(data$Date)))
    last.year <- as.numeric(stri_sub(last.date, 1, 4))
    display.year <- stri_sub((last.year-4):last.year, 3, 4)
    display.date <- paste0(display.year, "M", stri_sub(last.date, 5, 6), " ", unique(form$Period))
    display <- data.frame(period = display.date)
  }


  ## Main Table
  table.file <- data %>%
    group_by(period = !!sym(unique(form$Period))) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    filter(!is.na(period))

  if (unique(form$Period)=='MTH') {

    table.file <- table.file %>%
      mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
      arrange(month) %>%
      group_by(month) %>%
      mutate(growth = value / lag(value) - 1) %>%
      ungroup() %>%
      arrange(year,month) %>%
      mutate(period = factor(period, levels = display$period)) %>%
      filter(!is.na(period)) %>%
      arrange(period) %>%
      select(Period = period,
             !!sym(paste0("Value(", unique(form$Digit), ")")) := value,
             `Growth%` = growth)

  } else {

    table.file <- table.file %>%
      arrange(period) %>%
      mutate(growth = value / lag(value) - 1) %>%
      mutate(period = factor(period, levels = display$period)) %>%
      filter(!is.na(period)) %>%
      arrange(period) %>%
      select(Period = period,
             !!sym(paste0("Value(", unique(form$Digit), ")")) := value,
             `Growth%` = growth)
  }

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}
