# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


CompetitorPerformance <- function(data,
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
             product = !!sym(unique(form$Summary2))) %>% 
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(period %in% tail(unique(period),2) & product %in% unique(form$Internal)) %>% 
    mutate(type = product, sequence = 0) %>% 
    select(type,period,region,sub_value,sequence) 
  
  
  table2 <- dataQ %>% 
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

