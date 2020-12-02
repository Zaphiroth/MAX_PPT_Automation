# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #



MarketSize <- function(data,
                       form,
                       page,
                       digit,
                       directory) {

  ## Data Function
  dateformat <- data.frame(Date=sort(unique(data$Date)),Quarter=rep(NA,length(unique(data$Date))),
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

  dataQ <- data %>% left_join(dateformat, by = "Date") %>%
    mutate(MAT = paste0(Year,'Q', Quarter,' MAT'), YTD=paste0(Year,'Q', Quarter,' YTD')) %>%
    select(-Quarter, -Year)


  ## Display Function
  if (unique(form$Period)=='MTH') {
    cy <- 4
    last.date <- last(sort(unique(dataQ$Date)))
    last.year <- as.numeric(stri_sub(last.date, 1, 4))
    last.mth <- as.numeric(stri_sub(last.date, 5, 6))
    if(last.mth == 12) { tvar <-4 } else {tvar <-5 }
    display.year <- stri_sub((last.year-tvar):last.year, 3, 4)

    first.mth <- last.mth + (12/cy)
    if(first.mth > 12) {first.mth <- first.mth -12}
    fdmth <- seq(first.mth,12,by=(12/cy))
    ldmth <- seq(3,last.mth,by=(12/cy))
    rowtimer <- length(fdmth) + length(ldmth) + (length(display.year)-2)*cy
    if(last.mth == 12) {
      displayyr <- sort(c(rep(display.year,cy)))} else {
        displayyr <- c(rep(display.year[1],length(fdmth)),rep(display.year[2],cy),
                       rep(display.year[3],cy),rep(display.year[4],cy),rep(display.year[5],cy),
                       rep(display.year[6],length(ldmth)))
      }
    displaymth <- as.character(c(fdmth,rep(seq(3,12,by=(12/cy)),length(display.year)-2),ldmth))
    displaydf <- data.frame(year=displayyr,mth=displaymth) %>%
      mutate(mth=as.numeric(mth)/3) %>%
      mutate(period=paste0(year,"Q",mth))
    display <- data.frame(period=displaydf$period)

  } else {
    last.date <- last(sort(unique(dataQ$Date)))
    last.year <- as.numeric(stri_sub(last.date, 1, 4))
    display.year <- stri_sub((last.year-4):last.year, 3, 4)
    display.mth <- as.numeric(stri_sub(last.date, 5, 6))/3
    display.date <- paste0(display.year, "Q", display.mth , " ", unique(form$Period))
    display <- data.frame(period = display.date)
  }


  ## Main Table
  table.file <- dataQ %>%
    group_by(period = !!sym(unique(form$Period))) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    filter(is.na(period))

  if (unique(form$Period)=='MTH') {

    table.file <- table.file %>%
      mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
      arrange(month) %>%
      group_by(month) %>%
      mutate(growth = value / lag(value) - 1) %>%
      ungroup() %>%
      arrange(year,month) %>%
      mutate(period=paste0(year,'Q',as.numeric(month)/3)) %>%
      right_join(display, by = "period") %>%
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
      right_join(display, by = "period") %>%
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
