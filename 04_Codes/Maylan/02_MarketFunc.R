# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-12
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


MarketFunc <- function(data,
                       page,
                       form.table,
                       directory) {
  ##---- P2_1 ----
  form2.1 <- form.table %>% 
    filter(Page == "P2_1")
  
  table2.1 <- data %>% 
    group_by(period = !!sym(unique(form2.1$Period))) %>% 
    summarise(value = sum(!!sym(unique(form2.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(distinct(form2.1, Content), by = c("period" = "Content")) %>% 
    select(Period = period,
           !!sym(unique(form2.1$Index)[1]) := value,
           !!sym(unique(form2.1$Index)[2]) := growth)
  
  ##---- P2_2 ----
  form2.2 <- form.table %>% 
    filter(Page == "P2_2")
  
  table2.2 <- data %>% 
    group_by(period = !!sym(unique(form2.2$Period))) %>% 
    summarise(value = sum(!!sym(unique(form2.2$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(distinct(form2.2, Content), by = c("period" = "Content")) %>% 
    select(Period = period,
           !!sym(unique(form2.2$Index)[1]) := value,
           !!sym(unique(form2.2$Index)[2]) := growth)
  
  ##---- P3_1 ----
  form3.1 <- form.table %>% 
    filter(Page == "P3_1")
  
  table3.1 <- data %>% 
    group_by(period = !!sym(unique(form3.1$Period))) %>% 
    summarise(value = sum(!!sym(unique(form3.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(distinct(form3.1, Content), by = c("period" = "Content")) %>% 
    select(Period = period,
           !!sym(unique(form3.1$Index)[1]) := value,
           !!sym(unique(form3.1$Index)[2]) := growth)
  
  ##---- P3_2 ----
  form3.2 <- form.table %>% 
    filter(Page == "P3_2")
  
  table3.2 <- data %>% 
    group_by(period = !!sym(unique(form3.2$Period))) %>% 
    summarise(value = sum(!!sym(unique(form3.2$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    right_join(distinct(form3.2, Content), by = c("period" = "Content")) %>% 
    select(Period = period,
           !!sym(unique(form3.2$Index)[1]) := value,
           !!sym(unique(form3.2$Index)[2]) := growth)
  
  
  list(table2.1,
       table2.2,
       table3.1,
       table3.2)
}



