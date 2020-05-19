# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


RegionPerformance1 <- function(data,
                              form,
                              page,
                              directory) {
  
  table1 <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             mnf_type = !!sym(unique(form$Summary2))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period, region) %>% 
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(mnf_type == "MNC") %>% 
    group_by(period) %>% 
    mutate(`Con%` = complete_value / sum(complete_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(region) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = complete_value / lag(complete_value) - 1) %>% 
    ungroup() %>% 
    filter(period == max(period)) %>% 
    mutate(`MNC%` = sub_value / complete_value,
           type = paste0("Market ", period),
           region = factor(region, levels = unique(form$Display))) %>% 
    select(region,
           complete_value,
           `Con%`,
           `Growth%`,
           `MNC%`,
           type)
  
  table2 <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary3))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period, region) %>% 
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(product %in% unique(form$Internal)) %>% 
    group_by(period, product) %>% 
    mutate(`Con%` = sub_value / sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(region, product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = sub_value / lag(sub_value) - 1,
           `Share%` = sub_value / complete_value,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`)) %>% 
    ungroup() %>% 
    filter(period == max(period)) %>% 
    mutate(type = paste0("Internal Product Performance ", period),
           region = factor(region, levels = unique(form$Display))) %>% 
    select(region,
           sub_value,
           `Con%`,
           `Growth%`,
           `Share%`,
           `ΔShare%`,
           EI,
           type)
  
  table.file <- tabular(Heading(unique(form$Summary1), character.only = TRUE) * table1$region ~ 
                          Heading(unique(table1$type), character.only = TRUE) * identity * 
                          (Heading(unique(form$Index), character.only = TRUE) * table1$complete_value + 
                             Heading("Con%") * table1$`Con%` + 
                             Heading("Growth%") * table1$`Growth%` + 
                             Heading("MNC%") * table1$`MNC%`) + 
                          Heading(unique(table2$type), character.only = TRUE) * identity * 
                          (Heading(unique(form$Index)) * table2$sub_value + Heading("Con%") * table2$`Con%` + 
                             Heading("Growth%") * table2$`Growth%` + Heading("Share%") * table2$`Share%` + 
                             Heading("ΔShare%") * table2$`ΔShare%` + Heading("EI") * table2$EI))
  
  table.file
}



