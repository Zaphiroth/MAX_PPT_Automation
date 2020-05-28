# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# Date:         2020-05-26
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code P44/45 Growth Trend

GrowthTrend <- function(data,
                        form,
                        page,
                        digit,
                        directory) {

  table.file <- data %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = !!sym(unique(form$Period)),
             summary) %>%
    summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
    ungroup() %>%
    mutate (Value= Value/digit) %>%
    filter(period %in% tail(sort(unique(period)),min(36,length(unique(period))))) %>%
    spread(period,Value) %>%
    adorn_totals("row", na.rm = TRUE, name = "Total") %>%
    gather(period,Value,-summary) %>%
    mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2)) %>%
    arrange(month,summary) %>%
    group_by (month,summary) %>%
    mutate (Growth=(Value / lag(Value))-1)%>%
    ungroup() %>%
    na.omit() %>%
    select(summary,period,Growth) %>%
    spread(period,Growth) %>%
    mutate(sequence = ifelse(summary == 'Others' ,
                             2, ifelse(summary=='Total',3,1))) %>%
    arrange(sequence,summary) %>%
    select(-sequence) %>%
    rename(' ' = summary)

  ref.table <- data.frame(period=colnames(table.file[-1])) %>%
    mutate(month=str_sub(period,-2,-1),year=str_sub(period,1,2))
  timer <- 24-nrow(ref.table)

  if (timer > 0 ) {
    newdf <- data.frame(month=rep(NA,timer),year=rep(NA,timer))
    for (i in 1:timer) {
      mthindex <- as.numeric(ref.table$month[nrow(ref.table)])+i
      for (d in 1:2){
        if (mthindex > 12) {
          mthindex <- mthindex - 12
        }}
      mthindex <- str_pad(mthindex, 2, pad = "0")
      newdf$month[i] <- mthindex
    }

    if (newdf$month[timer]=='12'){
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
    newdf <- newdf %>% mutate(period=paste0(year,str_sub(ref.table$period[1],-3,-3),month))
    ncolnm <- c(newdf$period)
    table2 <- data.frame(matrix(nrow=nrow(table.file),ncol=length(ncolnm)))
    colnames(table2) <- ncolnm
    table.file <- cbind(table.file[,1],table2,table.file[,2:ncol(table.file)])
  }

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}




