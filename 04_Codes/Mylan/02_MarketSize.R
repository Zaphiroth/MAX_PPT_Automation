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
  
  last.date <- last(sort(unique(data$Date)))
  last.year <- as.numeric(stri_sub(last.date, 3, 4))
  display.year <- (last.year-4):last.year
  display.date <- paste0(display.year, "M", stri_sub(last.date, 5, 6), " ", unique(form$Period))
  display <- data.frame(period = display.date)
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(display, by = "period") %>% 
    select(Period = period,
           !!sym(paste0("Sales(", unique(form$Digit), ")")) := value,
           `Growth%` = growth)
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}



