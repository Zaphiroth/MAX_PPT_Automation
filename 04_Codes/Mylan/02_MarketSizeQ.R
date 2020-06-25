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
  
  dataQ <- data %>% mutate(MATY=str_sub(data$MAT,1,2), YTDY=str_sub(data$YTD,1,2)) %>%
          mutate(MAT = paste0(MATY,'Q4 MAT'), YTD=paste0(YTDY,'Q4 YTD')) %>%
          select(-MATY, -YTDY)
  
  table.file <- dataQ %>% 
    group_by(period = !!sym(unique(form$Period))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
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
