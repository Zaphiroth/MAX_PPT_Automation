# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

Performance <- function(data,
                        form,
                        page,
                        digit,
                        directory) {

  dateformat <- data.frame(Date=sort(unique(data$Date)),
                           Quarter=rep(NA,length(unique(data$Date))),
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

  dataQ <- data %>%
    left_join(dateformat, by = "Date") %>%
    mutate(MATY=str_sub(data$MAT,1,2),
           QT=as.numeric(str_sub(data$Date,5,6))/3) %>%
    mutate(MAT = paste0(Year,'Q', Quarter,' MAT'),
           YTD=paste0(Year,'Q', Quarter,' YTD'),
           MTH=paste0(MATY,'Q',QT)) %>%
    select(-Quarter, -Year,-QT,-MATY)

  table1 <- dataQ %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in%
                              unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = MAT, summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                digit) %>%
    ungroup() %>%
    filter(!is.na(period)) %>%
    setDT() %>%
    dcast(summary ~ period, value.var = "value") %>%
    adorn_totals("row", na.rm = TRUE, name = "TTL Mkt") %>%
    melt(id.vars = "summary", variable.factor = FALSE,
         variable.name = "period", value.name = "value") %>%
    group_by(period) %>%
    mutate(`Share%` = value / sum(value, na.rm = TRUE) * 2) %>%
    ungroup() %>%
    group_by(summary) %>%
    arrange(period) %>%
    mutate(`Growth%` = value / lag(value) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE)) %>%
    mutate(period_name = "MAT")

  table2 <- dataQ %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in%
                              unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = YTD, summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                digit) %>%
    ungroup() %>%
    filter(!is.na(period)) %>%
    setDT() %>%
    dcast(summary ~ period, value.var = "value") %>%
    adorn_totals("row", na.rm = TRUE, name = "TTL Mkt") %>%
    melt(id.vars = "summary", variable.factor = FALSE,
         variable.name = "period", value.name = "value") %>%
    group_by(period) %>%
    mutate(`Share%` = value / sum(value, na.rm = TRUE) * 2) %>%
    ungroup() %>%
    group_by(summary) %>%
    arrange(period) %>%
    mutate(`Growth%` = value / lag(value) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE)) %>%
    mutate(period_name = "YTD")

  table3 <- dataQ %>%
    filter(stri_sub(MTH, 4, 4) %in% stri_sub(max(MTH, na.rm = TRUE), 4, 4)) %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in%
                              unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = MTH, summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                digit) %>%
    ungroup() %>%
    filter(!is.na(period)) %>%
    setDT() %>%
    dcast(summary ~ period, value.var = "value") %>%
    adorn_totals("row", na.rm = TRUE, name = "TTL Mkt") %>%
    melt(id.vars = "summary", variable.factor = FALSE,
         variable.name = "period", value.name = "value") %>%
    group_by(period) %>%
    mutate(`Share%` = value / sum(value, na.rm = TRUE) * 2) %>%
    ungroup() %>%
    group_by(summary) %>%
    arrange(period) %>%
    mutate(`Growth%` = value / lag(value) - 1,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE)) %>%
    mutate(period_name = "MTH")

  if (!is.na(unique(form$Summary2))) {
    table4 <- dataQ %>%
      mutate(summary = ifelse(!!sym(unique(form$Summary2)) == "MNC",
                              "TTL MNC",
                              ifelse(!!sym(unique(form$Summary2)) == "LOCAL",
                                     "TTL Local",
                                     NA_character_))) %>%
      group_by(period = MAT, summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      filter(!is.na(period)) %>%
      group_by(period) %>%
      mutate(`Share%` = value / sum(value, na.rm = TRUE)) %>%
      ungroup() %>%
      group_by(summary) %>%
      arrange(period) %>%
      mutate(`Growth%` = value / lag(value) - 1,
             `ΔShare%` = `Share%` - lag(`Share%`),
             EI = `Share%` / lag(`Share%`) * 100) %>%
      ungroup() %>%
      filter(period == max(period, na.rm = TRUE)) %>%
      mutate(period_name = "MAT")

    table5 <- dataQ %>%
      mutate(summary = ifelse(!!sym(unique(form$Summary2)) == "MNC",
                              "TTL MNC",
                              ifelse(!!sym(unique(form$Summary2)) == "LOCAL",
                                     "TTL Local",
                                     NA_character_))) %>%
      group_by(period = YTD, summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      filter(!is.na(period)) %>%
      group_by(period) %>%
      mutate(`Share%` = value / sum(value, na.rm = TRUE)) %>%
      ungroup() %>%
      group_by(summary) %>%
      arrange(period) %>%
      mutate(`Growth%` = value / lag(value) - 1,
             `ΔShare%` = `Share%` - lag(`Share%`),
             EI = `Share%` / lag(`Share%`) * 100) %>%
      ungroup() %>%
      filter(period == max(period, na.rm = TRUE)) %>%
      mutate(period_name = "YTD")

    table6 <- dataQ %>%
      filter(stri_sub(MTH, 4, 4) %in%
               stri_sub(max(MTH, na.rm = TRUE), 4, 4)) %>%
      mutate(summary = ifelse(!!sym(unique(form$Summary2)) == "MNC",
                              "TTL MNC",
                              ifelse(!!sym(unique(form$Summary2)) == "LOCAL",
                                     "TTL Local",
                                     NA_character_))) %>%
      group_by(period = MTH, summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      filter(!is.na(period)) %>%
      group_by(period) %>%
      mutate(`Share%` = value / sum(value, na.rm = TRUE)) %>%
      ungroup() %>%
      group_by(summary) %>%
      arrange(period) %>%
      mutate(`Growth%` = value / lag(value) - 1,
             `ΔShare%` = `Share%` - lag(`Share%`),
             EI = `Share%` / lag(`Share%`) * 100) %>%
      ungroup() %>%
      filter(period == max(period, na.rm = TRUE)) %>%
      mutate(period_name = "MTH")

  } else {
    table4 <- data.frame()
    table5 <- data.frame()
    table6 <- data.frame()
  }

  table.file <- bind_rows(table1, table2, table3,
                          table4, table5, table6) %>%
    filter(period_name %in%
             unlist(stri_split_fixed(unique(form$Period), "+"))) %>%
    mutate(` ` = factor(summary, levels = unique(form$Display)),
           period = factor(period, levels = c(unique(table1$period),
                                              unique(table2$period),
                                              unique(table3$period))),
           !!sym(paste0("Sales(", unique(form$Digit), ")")) := value) %>%
    tabular(` ` * Heading() ~
              period * (sym(paste0("Sales(", unique(form$Digit), ")")) +
                          `Growth%` + `Share%` + `ΔShare%` + EI) *
              Heading() * identity,
            data = .)

  table.file=as.matrix(table.file)
  table.file
  write.xlsx(table.file,
             file=paste0(directory,'/',page,'.xlsx'),
             col.names=FALSE)
}
