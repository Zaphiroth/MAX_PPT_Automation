# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


ShareTrend <- function(data,
                       form,
                       page,
                       digit,
                       directory) {
  
  table.file <- data %>% 
    filter(!!sym(unique(form$Period)) %in% 
             tail(sort(unique(unlist(data[, unique(form$Period)]))), 24)) %>% 
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display), 
                            !!sym(unique(form$Summary1)), 
                            "Others")) %>% 
    group_by(period = !!sym(unique(form$Period)), summary) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(value_share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(summary ~ period, value.var = "value_share") %>% 
    right_join(distinct(form, Display), by = c("summary" = "Display")) %>% 
    rename(!!sym(' ') := summary)
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}




