# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# Date:         2020-05-22
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code P23/24 RegionTrend

RegionTrend <- function(data,
                        form,
                        page,
                        digit,
                        directory) {
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1))) %>%
    summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
    ungroup() %>% 
    mutate (Value= Value/digit) %>%
    filter(period %in% tail(unique(period),36)) %>%
    mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
    group_by(period) %>%
    mutate (total=sum(Value, na.rm = TRUE)) %>%
    ungroup() %>%
    mutate(Share = Value/total) %>%
    arrange(month,region) %>%
    group_by (month,region) %>%
    mutate (EI=(Share / lag(Share))*100) %>% 
    ungroup() %>%
    na.omit() %>%
    select(region,period,Share,EI) %>%
    gather(category,value,Share:EI) %>%
    mutate(Name= paste(region,category,sep="_")) %>%
    spread(period,value) %>%
    arrange(desc(category),desc(Name)) %>%
    select(Name, everything(),-category,-region) 
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}

