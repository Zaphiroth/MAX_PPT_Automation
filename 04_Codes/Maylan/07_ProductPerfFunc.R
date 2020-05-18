# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-15
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


ProductPerfFunc <- function(data,
                            page,
                            form.table,
                            directory) {
  ##---- P15_1 ----
  form15.1 <- form.table %>% 
    filter(Page == "P15_1")
  
  table15.1.1 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form15.1$Summary1)) %in% unique(form15.1$Content), 
                            !!sym(unique(form15.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = MAT, product) %>% 
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
  
  table15.1.2 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form15.1$Summary1)) %in% unique(form15.1$Content), 
                            !!sym(unique(form15.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = YTD, product) %>% 
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
  
  table15.1.3 <- data %>% 
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH), 4, 5)) %>% 
    mutate(product = ifelse(!!sym(unique(form15.1$Summary1)) %in% unique(form15.1$Content), 
                            !!sym(unique(form15.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = MTH, product) %>% 
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
  
  table15.1.4 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form15.1$Summary2)) == "MNC", "TTL MNC", 
                            ifelse(!!sym(unique(form15.1$Summary2)) == "LOCAL", "TTL Local", 
                                   NA_character_))) %>% 
    group_by(period = MAT, product) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form15.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table15.1.5 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form15.1$Summary2)) == "MNC", "TTL MNC", 
                            ifelse(!!sym(unique(form15.1$Summary2)) == "LOCAL", "TTL Local", 
                                   NA_character_))) %>% 
    group_by(period = YTD, product) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form15.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table15.1.6 <- data %>% 
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH), 4, 5)) %>% 
    mutate(product = ifelse(!!sym(unique(form15.1$Summary2)) == "MNC", "TTL MNC", 
                            ifelse(!!sym(unique(form15.1$Summary2)) == "LOCAL", "TTL Local", 
                                   NA_character_))) %>% 
    group_by(period = MTH, product) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form15.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table15.1 <- bind_rows(table15.1.1, table15.1.2, table15.1.3, 
                         table15.1.4, table15.1.5, table15.1.6) %>% 
    mutate(product = factor(product, levels = unique(form15.1$Content)),
           period = factor(period, levels = c(unique(table15.1.1$period), 
                                              unique(table15.1.2$period), 
                                              unique(table15.1.3$period)))) %>% 
    rename(Value = product) %>% 
    tabular(Value * Heading() ~ 
              period * (`RMB(Mn)` + `Growth%` + `Share%` + `ΔShare%` + EI) * Heading() * identity, 
            data = .)
  
  ##---- P16_1 ----
  form16.1 <- form.table %>% 
    filter(Page == "P16_1")
  
  table16.1.1 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form16.1$Summary1)) %in% unique(form16.1$Content), 
                            !!sym(unique(form16.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = MAT, product) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form16.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(product ~ period, value.var = "TAB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form16.1$Content)[1]) %>% 
    melt(id.vars = "product", variable.factor = FALSE, 
         variable.name = "period", value.name = "TAB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table16.1.2 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form16.1$Summary1)) %in% unique(form16.1$Content), 
                            !!sym(unique(form16.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = YTD, product) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form16.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(product ~ period, value.var = "TAB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form16.1$Content)[1]) %>% 
    melt(id.vars = "product", variable.factor = FALSE, 
         variable.name = "period", value.name = "TAB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table16.1.3 <- data %>% 
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH), 4, 5)) %>% 
    mutate(product = ifelse(!!sym(unique(form16.1$Summary1)) %in% unique(form16.1$Content), 
                            !!sym(unique(form16.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = MTH, product) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form16.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(product ~ period, value.var = "TAB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form16.1$Content)[1]) %>% 
    melt(id.vars = "product", variable.factor = FALSE, 
         variable.name = "period", value.name = "TAB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table16.1.4 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form16.1$Summary2)) == "MNC", "TTL MNC", 
                            ifelse(!!sym(unique(form16.1$Summary2)) == "LOCAL", "TTL Local", 
                                   NA_character_))) %>% 
    group_by(period = MAT, product) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form16.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table16.1.5 <- data %>% 
    mutate(product = ifelse(!!sym(unique(form16.1$Summary2)) == "MNC", "TTL MNC", 
                            ifelse(!!sym(unique(form16.1$Summary2)) == "LOCAL", "TTL Local", 
                                   NA_character_))) %>% 
    group_by(period = YTD, product) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form16.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table16.1.6 <- data %>% 
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH), 4, 5)) %>% 
    mutate(product = ifelse(!!sym(unique(form16.1$Summary2)) == "MNC", "TTL MNC", 
                            ifelse(!!sym(unique(form16.1$Summary2)) == "LOCAL", "TTL Local", 
                                   NA_character_))) %>% 
    group_by(period = MTH, product) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form16.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table16.1 <- bind_rows(table16.1.1, table16.1.2, table16.1.3, 
                         table16.1.4, table16.1.5, table16.1.6) %>% 
    mutate(product = factor(product, levels = unique(form16.1$Content)),
           period = factor(period, levels = c(unique(table16.1.1$period), 
                                              unique(table16.1.2$period), 
                                              unique(table16.1.3$period)))) %>% 
    rename(Volume = product) %>% 
    tabular(Volume * Heading() ~ 
              period * (`TAB(Mn)` + `Growth%` + `Share%` + `ΔShare%` + EI) * Heading() * identity, 
            data = .)
  
  
  list(table15.1,
       table16.1)
}



