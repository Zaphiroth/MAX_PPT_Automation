# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketShare <- function(data,
                           form,
                           page,
                           digit,
                           directory) {
  
  table.file <- data %>% 
    group_by(period = !!sym(unique(form$Period)),
             sub_market = !!sym(unique(form$Summary1))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(share = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    arrange(period) %>% 
    setDT() %>% 
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE, 
         variable.name = "index", value.name = "value") %>% 
    mutate(index = ifelse(index == "value", "", "Share")) %>% 
    unite(period_index, period, index, sep = " ") %>% 
    setDT() %>% 
    dcast(sub_market ~ period_index, value.var = "value") %>% 
    select(sub_market,
           ends_with(unique(form$Period)),
           ends_with("Share")) %>% 
    right_join(distinct(form, Display), by = c("sub_market" = "Display"))
  
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
  
  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}



