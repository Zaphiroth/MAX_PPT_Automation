# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


RegionPerfFunc <- function(data,
                           page,
                           form.table,
                           directory) {
  ##---- P21_1 ----
  form21.1 <- form.table %>% 
    filter(Page == "P21_1")
  
  table21.1.1 <- data %>% 
    group_by(period = !!sym(form21.1), product) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form15.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(product ~ period, value.var = "RMB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form15.1$Content)[1]) %>% 
    melt(id.vars = "product", variable.factor = FALSE, 
         variable.name = "period", value.name = "RMB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
}



