# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-05-26
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code P44/45 Growth Trend


GrowthTrend <- function(data,
                        form,
                        page,
                        digit,
                        directory) {
  
table.file <- data %>% 
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display), 
                            !!sym(unique(form$Summary1)), 
                            "Others")) %>% 
    group_by(period = !!sym(unique(form$Period)),
             summary) %>%
    summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
    ungroup() %>%
    mutate (Value= Value/digit) %>%
    filter(period %in% tail(unique(period),36)) %>%
    spread(period,Value) %>%
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    gather(period,Value,-summary) %>%
    mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
    arrange(month,summary) %>%
    group_by (month,summary) %>%
    mutate (Growth=(Value / lag(Value))-1)%>% 
    ungroup() %>%
    na.omit() %>%
    select(summary,period,Growth) %>%
    spread(period,Growth) %>%
    mutate(sequence = ifelse(summary == 'Others' , 
                            2, ifelse(summary=='Total',3,1))) %>%
    arrange(sequence,summary) %>%
    select(-sequence) %>%
    rename(' ' = summary) 

table.file
write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}
    
    
  

    
    
    