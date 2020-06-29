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

  dataQ <- data %>% left_join(dateformat, by = "Date") %>%
          mutate(MAT = paste0(Year,'Q', Quarter,' MAT'),
                 YTD=paste0(Year,'Q', Quarter,' YTD')) %>%
          select(-Quarter, -Year)

  table.file <- dataQ %>%
    group_by(period = !!sym(unique(form$Period))) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                digit) %>%
    ungroup() %>%
    arrange(period) %>%
    mutate(growth = value / lag(value) - 1) %>%
    select(Period = period,
           !!sym(paste0("Value(", unique(form$Digit), ")")) := value,
           `Growth%` = growth)

  if(nrow(table.file) < 5) {
    timer <- 5 - nrow(table.file)
    refperiod <- as.character(table.file$Period[1])
    refother <- substr(refperiod,3,nchar(refperiod))
    refyear <- as.numeric(substr(refperiod,1,2))
    newdf <- data.frame(matrix(nrow=timer,ncol=ncol(table.file)))
    colnames(newdf) <- colnames(table.file)
    for (i in 1: timer) {
      refyear <- refyear -1
      newdf$Period[timer+1-i] <- paste0(refyear,refother)
    }
    table.file <- rbind(newdf,table.file)
  }

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}
