# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


MarketSize <- function(data,
                       form,
                       page,
                       digit,
                       directory) {
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(distinct(form, Display), by = c("period" = "Display")) %>% 
    select(Period = period,
           !!sym(paste0("Value(", unique(form$Digit), ")")) := value,
           `Growth%` = growth)
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}



