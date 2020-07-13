# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketShare <- function(data,
                           form,
                           page,
                           digit,
                           directory) {

  ### Data Function
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


  ### Display Function
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

    display <- data.frame(period = seq(max(as.numeric(dataQ$Date)), as.numeric(max(dataQ$Date))-400, -100)) %>%
      arrange(period) %>%
      mutate(period = paste0(stri_sub(period, 3, 4), "Q", as.numeric(stri_sub(period, 5, 6))/3 , " ", unique(form$Period)))
  }

  display <- display %>%
    mutate (period_index = paste0(period, " Share")) %>%
    setDT() %>%
    melt(measure.vars = c("period", "period_index"), value.name = "period_index") %>%
    select(period_index) %>%
    setDF() %>%
    merge(distinct(form, sub_market = Display))


  ### Main Table
  table.file <- dataQ %>%
    group_by(period = !!sym(unique(form$Period)),
             sub_market = !!sym(unique(form$Summary1))) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    filter(!is.na(period)) %>%
    group_by(period) %>%
    mutate(share = value / sum(value, na.rm = TRUE)) %>%
    ungroup() %>%
    arrange(period) %>%
    setDT() %>%
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE,
         variable.name = "index", value.name = "value") %>%
    mutate(index = ifelse(index == "value", "", "Share"))

  if (unique(form$Period)=='MTH') {
    table.file <- table.file %>%
      mutate(month=str_sub(period,4,5),year=str_sub(period,1,2)) %>%
      mutate(period = paste0(year,'Q', as.numeric(month)/3)) %>%
      select(-year,-month) }

  table.file <- table.file %>%
    unite(period_index, period, index, sep = " ") %>%
    mutate(period_index = trimws(period_index)) %>%
    right_join(display, by = c("period_index", "sub_market")) %>%
    setDT() %>%
    dcast(sub_market ~ period_index, value.var = "value")


  if (unique(form$Period)=='MTH') {
    NameList <- grep(pattern='.[0-9]$',x=colnames(table.file),value=TRUE)
    colnum <- which(colnames(table.file) %in% NameList)

    table.file <- table.file %>%
      select(sub_market,
             colnum,
             ends_with("Share")) %>%
      right_join(distinct(form, Display), by = c("sub_market" = "Display"))

  } else {

    table.file <- table.file %>%
      select(sub_market,
             ends_with(unique(form$Period)),
             ends_with("Share")) %>%
      right_join(distinct(form, Display), by = c("sub_market" = "Display"))
  }



  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}

