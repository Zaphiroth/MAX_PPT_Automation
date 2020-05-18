# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketShare <- function(data,
                           form,
                           page,
                           directory) {
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             sub_market = !!sym(unique(form$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE, 
         variable.name = "index", value.name = "value") %>% 
    mutate(index = ifelse(index == "value", unique(form$Index), "Share")) %>% 
    unite(period_index, period, index, sep = " ") %>% 
    setDT() %>% 
    dcast(sub_market ~ period_index, value.var = "value") %>% 
    select(sub_market,
           ends_with(unique(form$Index)),
           ends_with("Share")) %>% 
    right_join(distinct(form, Display), by = c("sub_market" = "Display"))
  
  table.file
}



