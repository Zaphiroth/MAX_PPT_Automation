# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


ProductTrendFunc <- function(data,
                             page,
                             form.table,
                             directory) {
  ##---- P17_1 ----
  form17.1 <- form.table %>% 
    filter(Page == "P17_1")
  
  table17.1 <- data %>% 
    filter(!!sym(unique(form17.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form17.1$Period)]))), 24)) %>% 
    mutate(Product = ifelse(!!sym(unique(form17.1$Summary1)) %in% unique(form17.1$Content), 
                            !!sym(unique(form17.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = !!sym(unique(form17.1$Period)), Product) %>% 
    summarise(value = sum(!!sym(unique(form17.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    dcast(Product ~ period, value.var = "value") %>% 
    right_join(distinct(form17.1, Content), by = c("Product" = "Content"))
  
  ##---- P18_1 ----
  form18.1 <- form.table %>% 
    filter(Page == "P18_1")
  
  table18.1 <- data %>% 
    filter(!!sym(unique(form18.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form18.1$Period)]))), 24)) %>% 
    mutate(Product = ifelse(!!sym(unique(form18.1$Summary1)) %in% unique(form18.1$Content), 
                            !!sym(unique(form18.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = !!sym(unique(form18.1$Period)), Product) %>% 
    summarise(value = sum(!!sym(unique(form18.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    dcast(Product ~ period, value.var = "value") %>% 
    right_join(distinct(form18.1, Content), by = c("Product" = "Content"))
  
  ##---- P19_1 ----
  form19.1 <- form.table %>% 
    filter(Page == "P19_1")
  
  table19.1 <- data %>% 
    filter(!!sym(unique(form19.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form19.1$Period)]))), 24)) %>% 
    mutate(Product = ifelse(!!sym(unique(form19.1$Summary1)) %in% unique(form19.1$Content), 
                            !!sym(unique(form19.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = !!sym(unique(form19.1$Period)), Product) %>% 
    summarise(value = sum(!!sym(unique(form19.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(value_share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product ~ period, value.var = "value_share") %>% 
    right_join(distinct(form19.1, Content), by = c("Product" = "Content"))
  
  ##---- P20_1 ----
  form20.1 <- form.table %>% 
    filter(Page == "P20_1")
  
  table20.1 <- data %>% 
    filter(!!sym(unique(form20.1$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form20.1$Period)]))), 24)) %>% 
    mutate(Product = ifelse(!!sym(unique(form20.1$Summary1)) %in% unique(form20.1$Content), 
                            !!sym(unique(form20.1$Summary1)), 
                            "Others")) %>% 
    group_by(period = !!sym(unique(form20.1$Period)), Product) %>% 
    summarise(value = sum(!!sym(unique(form20.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(value_share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product ~ period, value.var = "value_share") %>% 
    right_join(distinct(form20.1, Content), by = c("Product" = "Content"))
  
  
  list(table17.1,
       table18.1,
       table19.1,
       table20.1)
}





