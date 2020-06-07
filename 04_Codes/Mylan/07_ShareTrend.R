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
  
  if(unique(form$Period)=='MTH') {
    grouper <- 'MTH' } else {
      grouper <- 'MAT'  
    }
  
  table.file <- data %>% 
    filter(!!sym(grouper) %in% 
             tail(sort(unique(unlist(data[, grouper]))), 
                  min(24,length(unique(unlist(data[, grouper])))))) %>% 
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display), 
                            !!sym(unique(form$Summary1)), 
                            "Others")) 
  
  if (unique(form$Period)=='MTH') {  
    table.file <- table.file %>%
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
   
     } else {
      
      if (unique(form$Period)=='MAT') {
        rollparam <- 12
      } else {
        rollparam <- 3
      }
      
      table.file <- table.file %>%
        group_by(PeriodMAT = !!sym(grouper),
                 PeriodMTH = MTH,
                 summary) %>%
        summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
        ungroup() %>%
        filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),min(36,length(unique(PeriodMTH))))) %>%
        arrange(summary,PeriodMAT, PeriodMTH) %>%
        group_by(summary) %>%
        mutate(RollValue=c(rep(NA,(rollparam-1)),rollsum(value,rollparam))) %>%
        ungroup() %>% na.omit() %>%
        filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),min(24,length(unique(PeriodMTH))))) %>% 
        mutate(period = paste(PeriodMTH,unique(form$Period),sep=' ')) %>%
        select(period,summary, RollValue,-value,-PeriodMAT,-PeriodMTH) %>%
        group_by(period) %>% 
        mutate(value_share = RollValue / sum(RollValue, na.rm = TRUE)) %>% 
        ungroup() %>% 
        select(-RollValue) %>%
        spread(period,value_share) %>% 
        mutate(sequence = ifelse(summary == 'Others' , 
                                 2, ifelse(summary=='Total',3,1))) %>%
        arrange(sequence,summary) %>%
        select(-sequence) %>%
        rename(' ' = summary) 
      
      source("04_Codes/Mylan/14_DisplayFunction.R", encoding = "UTF-8")
      table.file <- DisplayFunction(table.file=table.file, type=unique(form$Period))
    }

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}





