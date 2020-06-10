# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# Date:         2020-05-22
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code --- Competitor Performance : 27/28/34/35/40/41/42/43

CompetitorPerformance <- function(data,
                                  form,
                                  page,
                                  digit,
                                  directory) {
  
  table1 <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary2))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(!is.na(period)) %>% 
    filter(period %in% tail(unique(period),2) & product %in% unique(form$Internal)) %>% 
    mutate(type = product, sequence = 0) %>% 
    select(type,period,region,sub_value,sequence) 
  
  
  table2 <- data %>% 
    filter (!!sym(unique(form$Summary2)) != unique(form$Internal) ) %>%
    mutate(type = ifelse(!!sym(unique(form$Summary2)) %in% unique(form$Display), 
                         !!sym(unique(form$Summary2)), 
                         "Others")) %>%
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             type ) %>%
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
    ungroup() %>% 
    filter(!is.na(period)) %>% 
    filter(period %in% tail(unique(period),2)) %>%
    mutate(sequence=ifelse(type=='Others',2,1))%>%
    select(type,period,region,sub_value,sequence)  
  
  
  table.file <- bind_rows(table1, table2) %>% 
    group_by(period,region) %>%
    mutate(total = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>% 
    mutate(Share = sub_value/total) %>%
    arrange(type,region) %>%
    group_by(type,region) %>%
    mutate (EI=(Share / lag(Share))*100) %>% 
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE)) %>% 
    setDT() %>% 
    melt(id.vars = c("type", "period", "region", "sequence"), 
         measure.vars = c("Share", "EI"), 
         variable.factor = FALSE) %>% 
    unite("type", variable, type, sep = " _ ") %>% 
    mutate(type = factor(type, levels = c(paste0("Share _ ", c(unique(form$Internal), 
                                                               unique(form$Display), 
                                                               "Others")), 
                                          paste0("EI _ ", c(unique(form$Internal), 
                                                            unique(form$Display), 
                                                            "Others"))))) %>% 
    setDT() %>% 
    dcast(region ~ type, value.var = "value") %>% 
    arrange(region) %>% 
    rename(` ` = region)
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}





