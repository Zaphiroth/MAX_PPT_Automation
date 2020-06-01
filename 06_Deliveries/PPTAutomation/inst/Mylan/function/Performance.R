# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


Performance <- function(data,
                        form,
                        page,
                        digit,
                        directory) {

  table1 <- data %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = MAT, summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
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

  table2 <- data %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = YTD, summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
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

  table3 <- data %>%
    filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH, na.rm = TRUE), 4, 5)) %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = MTH, summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
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
    table4 <- data %>%
      mutate(summary = ifelse(!!sym(unique(form$Summary2)) == "MNC", "TTL MNC",
                              ifelse(!!sym(unique(form$Summary2)) == "LOCAL", "TTL Local",
                                     NA_character_))) %>%
      group_by(period = MAT, summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
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

    table5 <- data %>%
      mutate(summary = ifelse(!!sym(unique(form$Summary2)) == "MNC", "TTL MNC",
                              ifelse(!!sym(unique(form$Summary2)) == "LOCAL", "TTL Local",
                                     NA_character_))) %>%
      group_by(period = YTD, summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
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

    table6 <- data %>%
      filter(stri_sub(MTH, 4, 5) %in% stri_sub(max(MTH, na.rm = TRUE), 4, 5)) %>%
      mutate(summary = ifelse(!!sym(unique(form$Summary2)) == "MNC", "TTL MNC",
                              ifelse(!!sym(unique(form$Summary2)) == "LOCAL", "TTL Local",
                                     NA_character_))) %>%
      group_by(period = MTH, summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
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
    filter(period_name %in% unlist(stri_split_fixed(unique(form$Period), "+"))) %>%
    mutate(` ` = factor(summary, levels = unique(form$Display)),
           period = factor(period, levels = c(unique(table1$period),
                                              unique(table2$period),
                                              unique(table3$period))),
           !!sym(paste0("Value(", unique(form$Digit), ")")) := value) %>%
    tabular(` ` * Heading() ~
              period * (sym(paste0("Value(", unique(form$Digit), ")")) +
                          `Growth%` + `Share%` + `ΔShare%` + EI) * Heading() * identity,
            data = .)

  table.file=as.matrix(table.file)
  table.file
  write.xlsx(table.file,file=paste0(directory,'/',page,'.xlsx'),col.names=FALSE)
}


