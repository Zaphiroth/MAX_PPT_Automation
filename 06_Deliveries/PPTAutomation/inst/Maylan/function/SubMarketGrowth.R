# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketGrowth <- function(data,
                            form,
                            page,
                            digit,
                            directory) {
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             sub_market = !!sym(unique(form$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "value") %>% 
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    melt(id.vars = "sub_market", variable.factor = FALSE, 
         variable.name = "period", value.name = "value") %>% 
    group_by(sub_market) %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() %>% 
    filter(!is.na(growth)) %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "growth") %>% 
    right_join(distinct(form, Display), by = c("sub_market" = "Display")) %>% 
    rename("Growth%" = sub_market)
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}









