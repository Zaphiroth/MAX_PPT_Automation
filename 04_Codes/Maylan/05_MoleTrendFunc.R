# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-14
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


MoleTrendFunc <- function(data,
                          page,
                          form.table,
                          directory) {
  ##---- P10_1 ----
  form10.1 <- form.table %>% 
    filter(Page == "P10_1")
  
  table10.1 <- data %>% 
    filter(!!sym(unique(form10.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form10.1$Period)]))), 24)) %>% 
    mutate(Molecule = ifelse(!!sym(unique(form10.1$Summary1)) %in% unique(form10.1$Content), 
                             !!sym(unique(form10.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = !!sym(unique(form10.1$Period)), Molecule) %>% 
    summarise(value = sum(!!sym(unique(form10.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    dcast(Molecule ~ period, value.var = "value") %>% 
    right_join(distinct(form10.1, Content), by = c("Molecule" = "Content"))
  
  ##---- P11_1 ----
  form11.1 <- form.table %>% 
    filter(Page == "P11_1")
  
  table11.1 <- data %>% 
    filter(!!sym(unique(form11.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form11.1$Period)]))), 24)) %>% 
    mutate(Molecule = ifelse(!!sym(unique(form11.1$Summary1)) %in% unique(form11.1$Content), 
                             !!sym(unique(form11.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = !!sym(unique(form11.1$Period)), Molecule) %>% 
    summarise(value = sum(!!sym(unique(form11.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    dcast(Molecule ~ period, value.var = "value") %>% 
    right_join(distinct(form11.1, Content), by = c("Molecule" = "Content"))
  
  ##---- P12_1 ----
  form12.1 <- form.table %>% 
    filter(Page == "P12_1")
  
  table12.1 <- data %>% 
    filter(!!sym(unique(form12.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form12.1$Period)]))), 24)) %>% 
    mutate(Molecule = ifelse(!!sym(unique(form12.1$Summary1)) %in% unique(form12.1$Content), 
                             !!sym(unique(form12.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = !!sym(unique(form12.1$Period)), Molecule) %>% 
    summarise(value = sum(!!sym(unique(form12.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(value_share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Molecule ~ period, value.var = "value_share") %>% 
    right_join(distinct(form12.1, Content), by = c("Molecule" = "Content"))
  
  ##---- P13_1 ----
  form13.1 <- form.table %>% 
    filter(Page == "P13_1")
  
  table13.1 <- data %>% 
    filter(!!sym(unique(form13.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form13.1$Period)]))), 24)) %>% 
    mutate(Molecule = ifelse(!!sym(unique(form13.1$Summary1)) %in% unique(form13.1$Content), 
                             !!sym(unique(form13.1$Summary1)), 
                             "Others")) %>% 
    group_by(period = !!sym(unique(form13.1$Period)), Molecule) %>% 
    summarise(value = sum(!!sym(unique(form13.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(value_share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Molecule ~ period, value.var = "value_share") %>% 
    right_join(distinct(form13.1, Content), by = c("Molecule" = "Content"))
  
  
  list(table10.1,
       table11.1,
       table12.1,
       table13.1)
}


