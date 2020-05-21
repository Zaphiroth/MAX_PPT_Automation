# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-05-20
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code --- Regional Performance 2: P23/24/27/28/34/35

RegionPerformance2 <- function(data,
                               form,
                               page,
                               directory) {
  
  table3 <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary2))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(period %in% tail(unique(period),2) & product %in% unique(form$Internal)) %>% 
    mutate(type =  paste("Internal Product",product,sep=':')) %>% 
    select(type,period,region,sub_value) 

  table4 <- data %>% 
    mutate(type = ifelse(!!sym(unique(form$Summary2)) %in% unique(form$Display), 
                         !!sym(unique(form$Summary2)), 
                         "Others")) %>%
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             type ) %>%
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
    ungroup() %>% 
    filter(period %in% tail(unique(period),2)) %>%
    mutate (type = factor(type, levels = unique(form$Display))) %>%
    select(type,period,region,sub_value)  
  
  table.file <- table3 %>% rbind(table4) %>% 
    group_by(period,region) %>%
    mutate(total = sum(sub_value)) %>%
    ungroup() %>% 
    mutate(Share = (sub_value/total)*100) %>%
    arrange(type,region) %>%
    group_by(type,region) %>%
    mutate (EV=(Share / lag(Share))*100) %>% na.omit() %>%
    ungroup() %>%
    select(type,region,Share,EV) %>%
    gather (category,value,Share:EV) %>%
    spread(region,value) %>%
    arrange(desc(category)) %>%
    mutate (Name = paste(category,type,sep="_")) %>%
    select (Name, everything(),-type,-category) 
    
  table.file
}




