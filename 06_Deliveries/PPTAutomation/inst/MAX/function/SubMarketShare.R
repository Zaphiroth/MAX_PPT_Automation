# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


SubMarketShare <- function(data,
                           form,
                           page,
                           digit,
                           directory) {

  if (unique(form$Period)=='MTH') {
    last.date <- last(sort(unique(data$Date)))
    last.year <- as.numeric(stri_sub(last.date, 1, 4))
    last.mth <- as.numeric(stri_sub(last.date, 5, 6))
    if(last.mth == 12) { tvar <-4 } else {tvar <-5 }
    display.year <- stri_sub((last.year-tvar):last.year, 3, 4)

    first.mth <- last.mth +1
    if(first.mth > 12) {first.mth <- first.mth -12}
    fdmth <- c(first.mth:12)
    ldmth <- c(1:last.mth)
    rowtimer <- length(fdmth) + length(ldmth) + (length(display.year)-2)*12
    if(last.mth == 12) {
      displayyr <- sort(c(rep(display.year,12)))} else {
        displayyr <- c(rep(display.year[1],length(fdmth)),rep(display.year[2],12),
                       rep(display.year[3],12),rep(display.year[4],12),rep(display.year[5],12),
                       rep(display.year[6],length(ldmth)))
      }
    displaymth <- as.character(c(fdmth,rep(1:12,length(display.year)-2),ldmth))
    displaydf <- data.frame(year=displayyr,mth=displaymth) %>%
      mutate(period=paste0(year,"M", str_pad(mth,2,pad='0')))
    display <- data.frame(period=displaydf$period)

  } else {
    display <- data.frame(period = seq(max(data$Date), max(data$Date)-400, -100)) %>%
      arrange(period) %>%
      mutate(period = paste0(stri_sub(period, 3, 4), "M", stri_sub(period, 5, 6), " ", unique(form$Period)))
  }

  display <- display %>%
    mutate (period_index = paste0(period, " Share")) %>%
    setDT() %>%
    melt(measure.vars = c("period", "period_index"), value.name = "period_index") %>%
    select(period_index) %>%
    setDF() %>%
    merge(distinct(form, sub_market = Display))


  table.file <- data %>%
    group_by(period = !!sym(unique(form$Period)),
             sub_market = !!sym(unique(form$Summary1))) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    filter(!is.na(period)) %>%
    group_by(period) %>%
    mutate(share = value / sum(value, na.rm = TRUE)) %>%
    ungroup() %>%
    arrange(period) %>%
    setDT() %>%
    melt(id.vars = c("period", "sub_market"), variable.factor = FALSE,
         variable.name = "index", value.name = "value") %>%
    mutate(index = ifelse(index == "value", "", "Share")) %>%
    unite(period_index, period, index, sep = " ") %>%
    mutate(period_index = trimws(period_index)) %>%
    mutate(period_index = factor(period_index, levels = unique(display$period_index)),
           sub_market = factor(sub_market, levels = unique(display$sub_market))) %>%
    filter(!is.na(period_index), !is.na(sub_market)) %>%
    arrange(period_index, sub_market) %>%
    setDT() %>%
    dcast(sub_market ~ period_index, value.var = "value")

  if (unique(form$Period)=='MTH') {
    NameList <- grep(pattern='.[0-9]$',x=colnames(table.file),value=TRUE)
    colnum <- which(colnames(table.file) %in% NameList)

    table.file <- table.file %>%
      select(sub_market,
             colnum,
             ends_with("Share")) %>%
      mutate(sub_market = factor(sub_market, levels = unique(form$Display))) %>%
      filter(!is.na(sub_market)) %>%
      arrange(sub_market)

  } else {

    table.file <- table.file %>%
      select(sub_market,
             ends_with(unique(form$Period)),
             ends_with("Share")) %>%
      mutate(sub_market = factor(sub_market, levels = unique(form$Display))) %>%
      filter(!is.na(sub_market)) %>%
      arrange(sub_market)
  }

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}




