# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


MarketSize <- function(data,
                       form,
                       page,
                       directory) {
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(distinct(form, Display), by = c("period" = "Display")) %>% 
    select(Period = period,
           !!sym(paste0(unique(form$Index), "(Mn)")) := value,
           `Growth%` = growth)
  
  table.file
}



