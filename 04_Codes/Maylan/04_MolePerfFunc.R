# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-14
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


MolePerfFunc <- function(data,
                         page,
                         form.table,
                         directory) {
  ##---- P8_1 ----
  form8.1 <- form.table %>% 
    filter(Page == "P8_1")
  
  table8.1.1 <- data %>% 
    mutate(molecule = ifelse(!!sym(unique(form8.1$Summary1)) %in% unique(form8.1$Content), 
                             !!sym(unique(form8.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = MAT, molecule) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form8.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(molecule ~ period, value.var = "RMB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form8.1$Content)[1]) %>% 
    melt(id.vars = "molecule", variable.factor = FALSE, 
         variable.name = "period", value.name = "RMB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(molecule) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table8.1.2 <- data %>% 
    mutate(molecule = ifelse(!!sym(unique(form8.1$Summary1)) %in% unique(form8.1$Content), 
                             !!sym(unique(form8.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = YTD, molecule) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form8.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(molecule ~ period, value.var = "RMB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form8.1$Content)[1]) %>% 
    melt(id.vars = "molecule", variable.factor = FALSE, 
         variable.name = "period", value.name = "RMB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(molecule) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table8.1.3 <- data %>% 
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH), 4, 5)) %>% 
    mutate(molecule = ifelse(!!sym(unique(form8.1$Summary1)) %in% unique(form8.1$Content), 
                             !!sym(unique(form8.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = MTH, molecule) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form8.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(molecule ~ period, value.var = "RMB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form8.1$Content)[1]) %>% 
    melt(id.vars = "molecule", variable.factor = FALSE, 
         variable.name = "period", value.name = "RMB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(molecule) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table8.1 <- bind_rows(table8.1.1, table8.1.2, table8.1.3) %>% 
    mutate(molecule = factor(molecule, levels = unique(form8.1$Content)),
           period = factor(period, levels = c(unique(table8.1.1$period), 
                                              unique(table8.1.2$period), 
                                              unique(table8.1.3$period)))) %>% 
    rename(Value = molecule) %>% 
    tabular(Value * Heading() ~ 
              period * (`RMB(Mn)` + `Growth%` + `Share%` + `ΔShare%` + EI) * Heading() * identity, 
            data = .)
  
  ##---- P9_1 ----
  form9.1 <- form.table %>% 
    filter(Page == "P9_1")
  
  table9.1.1 <- data %>% 
    mutate(molecule = ifelse(!!sym(unique(form9.1$Summary1)) %in% unique(form9.1$Content), 
                             !!sym(unique(form9.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = MAT, molecule) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form9.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(molecule ~ period, value.var = "TAB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form9.1$Content)[1]) %>% 
    melt(id.vars = "molecule", variable.factor = FALSE, 
         variable.name = "period", value.name = "TAB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(molecule) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table9.1.2 <- data %>% 
    mutate(molecule = ifelse(!!sym(unique(form9.1$Summary1)) %in% unique(form9.1$Content), 
                             !!sym(unique(form9.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = YTD, molecule) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form9.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(molecule ~ period, value.var = "TAB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form9.1$Content)[1]) %>% 
    melt(id.vars = "molecule", variable.factor = FALSE, 
         variable.name = "period", value.name = "TAB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(molecule) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table9.1.3 <- data %>% 
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH), 4, 5)) %>% 
    mutate(molecule = ifelse(!!sym(unique(form9.1$Summary1)) %in% unique(form9.1$Content), 
                             !!sym(unique(form9.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = MTH, molecule) %>% 
    summarise(`TAB(Mn)` = sum(!!sym(unique(form9.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(molecule ~ period, value.var = "TAB(Mn)") %>% 
    adorn_totals("row", na.rm = TRUE, name = unique(form9.1$Content)[1]) %>% 
    melt(id.vars = "molecule", variable.factor = FALSE, 
         variable.name = "period", value.name = "TAB(Mn)") %>% 
    group_by(period) %>% 
    mutate(`Share%` = `TAB(Mn)` / sum(`TAB(Mn)`, na.rm = TRUE) * 2) %>% 
    ungroup() %>% 
    group_by(molecule) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `TAB(Mn)` / lag(`TAB(Mn)`) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table9.1 <- bind_rows(table9.1.1, table9.1.2, table9.1.3) %>% 
    mutate(molecule = factor(molecule, levels = unique(form9.1$Content)),
           period = factor(period, levels = c(unique(table9.1.1$period), 
                                              unique(table9.1.2$period), 
                                              unique(table9.1.3$period)))) %>% 
    rename(Volume = molecule) %>% 
    tabular(Volume * Heading() ~ 
              period * (`TAB(Mn)` + `Growth%` + `Share%` + `ΔShare%` + EI) * Heading() * identity, 
            data = .)
  
  
  list(table8.1,
       table9.1)
}


