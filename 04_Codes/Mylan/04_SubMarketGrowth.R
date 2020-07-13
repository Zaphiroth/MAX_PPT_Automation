# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketGrowth <- function(data,
                            form,
                            page,
                            digit,
                            directory) {

table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             sub_market = !!sym(unique(form$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(sub_market ~ period, value.var = "value") %>% 
    adorn_totals("row", na.rm = TRUE, name = "Total") %>% 
    melt(id.vars = "sub_market", variable.factor = FALSE, 
       variable.name = "period", value.name = "value")

if (unique(form$Period)=='MTH') {    
  table.file <- table.file  %>% 
    mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
    arrange(sub_market,month) %>% 
    group_by(sub_market,month) %>%
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() 
    
} else {

table.file <- table.file  %>% 
    group_by(sub_market) %>% 
    arrange(period) %>% 
    mutate(growth = value / lag(value) - 1) %>% 
    ungroup() 
}

table.file <- table.file %>% 
     filter(!is.na(growth)) %>% 
     setDT() %>% 
     dcast(sub_market ~ period, value.var = "growth") %>% 
     right_join(distinct(form, Display), by = c("sub_market" = "Display")) %>% 
     rename("Growth%" = sub_market)
  

## Display Fucntion  
if (unique(form$Period)=='MTH') {
    source("04_Codes/Maylan/14_DisplayFunction.R", encoding = "UTF-8")
    table.file <- DisplayFunction(table.file=table.file,type='MTH',dis_period=60)
    colnames(table.file)[1] <- 'Growth%'
    
  } else {  
  
  pdnm <- colnames(table.file[-1])
  if (length(grep("^(?=\\bNA\\b).*",pdnm,value = TRUE, perl = TRUE)) != 0) {
    pdnm <- pdnm[pdnm != grep("^(?=\\bNA\\b).*",pdnm,value = TRUE, perl = TRUE)]
  }
  
  countpd <- length(pdnm)
  if (countpd<5) {
    timer <- 5 - countpd
    nleng <- nchar(pdnm[1])
    ref.table <- data.frame(PD=pdnm) %>% mutate(Year=str_sub(PD,1,2),Month=str_sub(PD,3,nleng))
    newdf <- data.frame(month=rep(unique(ref.table$Month),timer),year=rep(NA,timer))
    yrindex <- as.numeric(ref.table$Year[1])
    for (i in 1:timer) {
      yrindex <- yrindex -1
      newdf$year[timer+1-i] <- yrindex
    }
    newdf <- newdf %>% mutate(pd=paste0(year,month)) 
    ncolnm <- length(unique(newdf$pd))
    adddf <- data.frame(matrix(nrow=nrow(table.file),ncol=ncolnm)) 
    colnames(adddf) <- unique(newdf$pd)
    table.file <- cbind(table.file[1],adddf,table.file[-1])
    
    cdnm <- colnames(table.file[-1])
    if (length(grep("^(?=\\bNA\\b).*",cdnm,value = TRUE, perl = TRUE)) != 0) {
      table.file <- table.file[,-ncol(table.file)]
    }
  }
}
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}









