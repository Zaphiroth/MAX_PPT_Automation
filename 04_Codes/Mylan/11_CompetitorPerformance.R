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
    filter(period %in% tail(unique(period),2)) %>%
    mutate (sequence=ifelse(type=='Others',2,1),
            type = factor(type, levels = c(unique(form$Display),'Others')))%>%
    select(type,period,region,sub_value,sequence)  
  
  
  table.file <- table1 %>% rbind(table2) %>% 
    group_by(period,region) %>%
    mutate(total = sum(sub_value)) %>%
    ungroup() %>% 
    mutate(Share = (sub_value/total)*100) %>%
    arrange(type,region) %>%
    group_by(type,region) %>%
    mutate (EI=(Share / lag(Share))*100) %>% na.omit() %>%
    ungroup() %>%
    select(type,region,Share,EI,sequence) %>%
    gather (category,value,Share:EI) %>%
    spread(region,value) %>%
    mutate (Name = paste(category,"_",type)) %>%
    arrange(desc(category),sequence) %>%
    select (Name, everything(),-type,-category,-sequence) %>% 
    gather (region, value,-Name) %>%
    spread(Name,value) 
  
  colnm <- colnames(table.file[,-1])
  table.file <- table.file[,c('region',grep("^Share(?!.*Others)",colnm,value = TRUE, perl = TRUE),
                              grep("^Share.*(?=.*\\bOthers\\b)",colnm,value=TRUE,perl=TRUE),
                              grep("^EI(?!.*Others)",colnm,value = TRUE, perl = TRUE),
                              grep("^EI.*(?=.*\\bOthers\\b)",colnm,value=TRUE,perl=TRUE))] 
  table.file <- table.file %>% rename(" " = region)
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}





