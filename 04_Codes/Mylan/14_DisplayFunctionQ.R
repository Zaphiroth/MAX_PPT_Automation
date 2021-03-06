# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

DisplayFunction <- function(table.file, type, dis_period=8
                            ) {
  
  ref.table <- data.frame(period=colnames(table.file[,-1])) %>%
    mutate(month=str_sub(period,4,4),year=str_sub(period,1,2))
  timer <- dis_period-nrow(ref.table)
  
  if (timer > 0 ) {
    newdf <- data.frame(month=rep(NA,timer),year=rep(NA,timer))
    for (i in 1:timer) {
      mthindex <- as.numeric(ref.table$month[nrow(ref.table)])+i
      for (d in 1:2){
        if (mthindex > 4) {
          mthindex <- mthindex - 4
        }}
      newdf$month[i] <- mthindex
    }
    
    if (newdf$month[timer]=='4'){
      yrindex <- as.numeric(ref.table$year[1])-1
    } else {
      yrindex <- as.numeric(ref.table$year[1])
    }
    
    for (i in 1:timer) {
      key <- timer + 1 - i
      mthindex <- as.numeric(newdf$month[key])
      if(key == timer){
        diff <- 1
      } else {
        secindex <- as.numeric(newdf$month[key+1])
        diff <- secindex - mthindex
      }
      if (diff != 1) {
        yrindex <- yrindex -1
      }
      newdf$year[key] <- yrindex
    }
    
    if (type == 'MTH' & dis_period==8) {
      newdf <- newdf %>% 
        mutate(period=paste0(year,str_sub(ref.table$period[1],3,3),month))
    } else {
      newdf <- newdf %>% 
        mutate(period=paste0(year,str_sub(ref.table$period[1],3,3),
                             month,str_sub(ref.table$period[1],5,nchar(ref.table$period[1]))))
    }
    
    ncolnm <- c(newdf$period)
    table2 <- data.frame(matrix(nrow=nrow(table.file),ncol=length(ncolnm))) 
    colnames(table2) <- ncolnm
    table.file <- cbind(table.file[,1],table2,table.file[,2:ncol(table.file)])
  }
  
  table.file
}
