# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-13
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketFunc <- function(data,
                          page,
                          form.table,
                          directory) {
  ##---- P4_1 ----
  form4.1 <- form.table %>% 
    filter(Page == "P4_1")
  
  table4.1 <- data %>% 
    group_by(period = !!sym(unique(form4.1$Period)),
             sub_market = !!sym(unique(form4.1$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form4.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE, 
         variable.name = "index", value.name = "value") %>% 
    mutate(index = ifelse(index == "value", unique(form4.1$Index)[1], unique(form4.1$Index)[2])) %>% 
    unite(period_index, period, index, sep = " ") %>% 
    setDT() %>% 
    dcast(sub_market ~ period_index, value.var = "value") %>% 
    select(sub_market,
           ends_with(unique(form4.1$Index)[1]),
           ends_with(unique(form4.1$Index)[2]))
  
  ##---- P4_2 ----
  form4.2 <- form.table %>% 
    filter(Page == "P4_2")
  
  table4.2 <- data %>% 
    group_by(period = !!sym(unique(form4.2$Period)),
             sub_market = !!sym(unique(form4.2$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form4.2$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "value") %>% 
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    melt(id.vars = "sub_market", variable.factor = FALSE, 
         variable.name = "period", value.name = "value") %>% 
    group_by(`Value Growth%` = sub_market) %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() %>% 
    filter(!is.na(growth)) %>% 
    setDT() %>% 
    dcast(`Value Growth%` ~ period, value.var = "growth") %>% 
    right_join(distinct(form4.2, Content), by = c("Value Growth%" = "Content"))
  
  ##---- P5_1 ----
  form5.1 <- form.table %>% 
    filter(Page == "P5_1")
  
  table5.1 <- data %>% 
    group_by(period = !!sym(unique(form5.1$Period)),
             sub_market = !!sym(unique(form5.1$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form5.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE, 
         variable.name = "index", value.name = "value") %>% 
    mutate(index = ifelse(index == "value", unique(form5.1$Index)[1], unique(form5.1$Index)[2])) %>% 
    unite(period_index, period, index, sep = " ") %>% 
    setDT() %>% 
    dcast(sub_market ~ period_index, value.var = "value") %>% 
    select(sub_market,
           ends_with(unique(form5.1$Index)[1]),
           ends_with(unique(form5.1$Index)[2]))
  
  ##---- P5_2 ----
  form5.2 <- form.table %>% 
    filter(Page == "P5_2")
  
  table5.2 <- data %>% 
    group_by(period = !!sym(unique(form5.2$Period)),
             sub_market = !!sym(unique(form5.2$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form5.2$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "value") %>% 
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    melt(id.vars = "sub_market", variable.factor = FALSE, 
         variable.name = "period", value.name = "value") %>% 
    group_by(`Value Growth%` = sub_market) %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() %>% 
    filter(!is.na(growth)) %>% 
    setDT() %>% 
    dcast(`Value Growth%` ~ period, value.var = "growth") %>% 
    right_join(distinct(form5.2, Content), by = c("Value Growth%" = "Content"))
  
  ##---- P6_1 ----
  form6.1 <- form.table %>% 
    filter(Page == "P6_1")
  
  table6.1 <- data %>% 
    group_by(period = !!sym(unique(form6.1$Period)),
             sub_market = !!sym(unique(form6.1$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form6.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE, 
         variable.name = "index", value.name = "value") %>% 
    mutate(index = ifelse(index == "value", unique(form6.1$Index)[1], unique(form6.1$Index)[2])) %>% 
    unite(period_index, period, index, sep = " ") %>% 
    setDT() %>% 
    dcast(sub_market ~ period_index, value.var = "value") %>% 
    select(sub_market,
           ends_with(unique(form6.1$Index)[1]),
           ends_with(unique(form6.1$Index)[2]))
  
  ##---- P6_2 ----
  form6.2 <- form.table %>% 
    filter(Page == "P6_2")
  
  table6.2 <- data %>% 
    group_by(period = !!sym(unique(form6.2$Period)),
             sub_market = !!sym(unique(form6.2$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form6.2$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "value") %>% 
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    melt(id.vars = "sub_market", variable.factor = FALSE, 
         variable.name = "period", value.name = "value") %>% 
    group_by(`Value Growth%` = sub_market) %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() %>% 
    filter(!is.na(growth)) %>% 
    setDT() %>% 
    dcast(`Value Growth%` ~ period, value.var = "growth") %>% 
    right_join(distinct(form6.2, Content), by = c("Value Growth%" = "Content"))
  
  ##---- P7_1 ----
  form7.1 <- form.table %>% 
    filter(Page == "P7_1")
  
  table7.1 <- data %>% 
    group_by(period = !!sym(unique(form7.1$Period)),
             sub_market = !!sym(unique(form7.1$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form7.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE, 
         variable.name = "index", value.name = "value") %>% 
    mutate(index = ifelse(index == "value", unique(form7.1$Index)[1], unique(form7.1$Index)[2])) %>% 
    unite(period_index, period, index, sep = " ") %>% 
    setDT() %>% 
    dcast(sub_market ~ period_index, value.var = "value") %>% 
    select(sub_market,
           ends_with(unique(form7.1$Index)[1]),
           ends_with(unique(form7.1$Index)[2]))
  
  ##---- P7_2 ----
  form7.2 <- form.table %>% 
    filter(Page == "P7_2")
  
  table7.2 <- data %>% 
    group_by(period = !!sym(unique(form7.2$Period)),
             sub_market = !!sym(unique(form7.2$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form7.2$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "value") %>% 
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    melt(id.vars = "sub_market", variable.factor = FALSE, 
         variable.name = "period", value.name = "value") %>% 
    group_by(`Value Growth%` = sub_market) %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() %>% 
    filter(!is.na(growth)) %>% 
    setDT() %>% 
    dcast(`Value Growth%` ~ period, value.var = "growth") %>% 
    right_join(distinct(form7.2, Content), by = c("Value Growth%" = "Content"))
  
  
  list(table4.1,
       table4.2,
       table5.1,
       table5.2,
       table6.1,
       table6.2,
       table7.1,
       table7.2)
}



