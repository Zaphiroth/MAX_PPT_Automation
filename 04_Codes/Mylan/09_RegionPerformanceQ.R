# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


RegionPerformance <- function(data,
                              form,
                              page,
                              digit,
                              directory) {
  
  dateformat <- data.frame(Date=sort(unique(data$Date)),Quarter=rep(NA,length(unique(data$Date))),
                           Year=rep(NA,length(unique(data$Date))))
  dateformat$Quarter <- as.numeric(substr(dateformat$Date[nrow(dateformat)],5,6))/3
  timer <- nrow(dateformat)
  refYr <- as.numeric(substr(dateformat$Date[nrow(dateformat)],3,4))
  iter <- nrow(dateformat)/4
  for (i in 1:iter){
    for (t in 1:4) {
      dateformat$Year[timer+1-t] <- refYr
    }
    timer <- timer - 4
    refYr <- refYr -1
  }
  
  dataQ <- data %>% left_join(dateformat, by = "Date") %>% mutate(MATY=str_sub(data$MAT,1,2),QT=as.numeric(str_sub(data$Date,5,6))/3) %>%
    mutate(MAT = paste0(Year,'Q', Quarter,' MAT'), YTD=paste0(Year,'Q', Quarter,' YTD'),MTH=paste0(MATY,'Q',QT)) %>%
    select(-Quarter, -Year,-QT,-MATY)
  
  
  table1 <- dataQ %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             mnf_type = !!sym(unique(form$Summary2))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period, region) %>% 
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(mnf_type == "MNC",
           region %in% form$Display) %>% 
    group_by(region = "Total", period) %>% 
    summarise(sub_value = sum(sub_value, na.rm = TRUE),
              complete_value = sum(complete_value, na.rm = TRUE),
              national_value = first(national_value)) %>% 
    ungroup() %>% 
    mutate(`Con%` = complete_value / national_value) %>% 
    group_by(region) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = complete_value / lag(complete_value) - 1) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table2 <- dataQ %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             mnf_type = !!sym(unique(form$Summary2))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period, region) %>% 
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(mnf_type == "MNC") %>% 
    mutate(`Con%` = complete_value / national_value) %>% 
    group_by(region) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = complete_value / lag(complete_value) - 1) %>% 
    ungroup() %>% 
    filter(period == max(period),
           region %in% form$Display) %>% 
    arrange(-complete_value) %>% 
    bind_rows(table1) %>% 
    mutate(`MNC%` = sub_value / complete_value,
           type = paste0("Market ", period),
           region = factor(region, levels = region)) %>% 
    select(region,
           complete_value,
           `Con%`,
           `Growth%`,
           `MNC%`,
           type)
  
  table3 <- dataQ %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary3))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    filter(product %in% unique(form$Internal)) %>% 
    group_by(period) %>% 
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period, region) %>% 
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(region %in% form$Display) %>% 
    group_by(region = "Total", period) %>% 
    summarise(sub_value = sum(sub_value, na.rm = TRUE),
              complete_value = sum(complete_value, na.rm = TRUE),
              national_value = first(national_value)) %>% 
    ungroup() %>% 
    mutate(`Con%` = sub_value / national_value) %>% 
    group_by(region) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = sub_value / lag(sub_value) - 1,
           `Share%` = sub_value / complete_value,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>% 
    ungroup() %>% 
    filter(period == max(period))
  
  table4 <- dataQ %>% 
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary3))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    filter(product %in% unique(form$Internal)) %>% 
    group_by(period) %>% 
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period, region) %>% 
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    mutate(`Con%` = sub_value / sum(sub_value, na.rm = TRUE)) %>% 
    group_by(region) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = sub_value / lag(sub_value) - 1,
           `Share%` = sub_value / complete_value,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>% 
    ungroup() %>% 
    filter(period == max(period),
           region %in% form$Display) %>% 
    bind_rows(table3) %>% 
    mutate(region = factor(region, levels = table2$region)) %>% 
    right_join(distinct(table2, region), by = "region") 
    
table4 <- table4 %>%
    mutate(period=unique(na.omit(table4$period))) %>%
    mutate(type = paste0("Internal Product Performance ", period)) %>% 
    select(region,
           sub_value,
           `Con%`,
           `Growth%`,
           `Share%`,
           `ΔShare%`,
           EI,
           type)
  
  table.file <- tabular(Heading(unique(form$Summary1), character.only = TRUE) * table2$region ~ 
                          Heading(unique(table2$type), character.only = TRUE) * identity * 
                          (Heading(paste0("Value(", unique(form$Digit), ")"), character.only = TRUE) * 
                             table2$complete_value + 
                             Heading("Con%") * table2$`Con%` + 
                             Heading("Growth%") * table2$`Growth%` + 
                             Heading("MNC%") * table2$`MNC%`) + 
                          Heading(unique(table4$type), character.only = TRUE) * identity * 
                          (Heading(paste0("Value(", unique(form$Digit), ")"), character.only = TRUE) * 
                             table4$sub_value + 
                             Heading("Con%") * table4$`Con%` + 
                             Heading("Growth%") * table4$`Growth%` + Heading("Share%") * table4$`Share%` + 
                             Heading("ΔShare%") * table4$`ΔShare%` + Heading("EI") * table4$EI))
  
  table.file=as.matrix(table.file)
  table.file
  write.xlsx(table.file,file=paste0(directory,'/',page,'.xlsx'),col.names=FALSE)
}



