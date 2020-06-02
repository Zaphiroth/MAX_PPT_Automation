# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


RegionPerformance <- function(data,
                              form,
                              page,
                              digit,
                              directory) {

  table1 <- data %>%
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             mnf_type = !!sym(unique(form$Summary2))) %>%
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    group_by(period) %>%
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>%
    group_by(period, region) %>%
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>%
    filter(mnf_type == "MNC",
           region %in% form$Display) %>%
    group_by(region = "Total", period) %>%
    summarise(sub_value = sum(sub_value, na.rm = TRUE),
              complete_value = sum(complete_value, na.rm = TRUE),
              national_value = first(national_value)) %>%
    ungroup() %>%
    mutate(`Con%` = complete_value / national_value) %>%
    group_by(region) %>%
    arrange(period) %>%
    mutate(`Growth%` = complete_value / lag(complete_value) - 1) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE))

  table2 <- data %>%
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             mnf_type = !!sym(unique(form$Summary2))) %>%
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    setDT() %>%
    dcast(period + region ~ mnf_type, value.var = "sub_value", fill = 0) %>%
    mutate(complete_value = MNC + LOCAL) %>%
    group_by(period) %>%
    mutate(national_value = sum(complete_value, na.rm = TRUE)) %>%
    ungroup() %>%
    mutate(`Con%` = complete_value / national_value) %>%
    group_by(region) %>%
    arrange(period) %>%
    mutate(`Growth%` = complete_value / lag(complete_value) - 1) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE),
           region %in% form$Display) %>%
    arrange(-complete_value) %>%
    full_join(distinct(form, Display), by = c("region" = "Display")) %>%
    bind_rows(table1) %>%
    mutate(`MNC%` = MNC / complete_value,
           type = paste0("Market ", period),
           region = factor(region, levels = region)) %>%
    select(region,
           complete_value,
           `Con%`,
           `Growth%`,
           `MNC%`,
           type)

  table3 <- data %>%
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary3))) %>%
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    group_by(period, region) %>%
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>%
    filter(product %in% unique(form$Internal)) %>%
    group_by(period) %>%
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>%
    filter(region %in% form$Display) %>%
    group_by(region = "Total", period) %>%
    summarise(sub_value = sum(sub_value, na.rm = TRUE),
              complete_value = sum(complete_value, na.rm = TRUE),
              national_value = first(national_value)) %>%
    ungroup() %>%
    mutate(`Con%` = sub_value / national_value) %>%
    group_by(region) %>%
    arrange(period) %>%
    mutate(`Growth%` = sub_value / lag(sub_value) - 1,
           `Share%` = sub_value / complete_value,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE))

  table4 <- data %>%
    group_by(period = !!sym(unique(form$Period)),
             region = !!sym(unique(form$Summary1)),
             product = !!sym(unique(form$Summary3))) %>%
    summarise(sub_value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    group_by(period, region) %>%
    mutate(complete_value = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>%
    filter(product %in% unique(form$Internal)) %>%
    group_by(period) %>%
    mutate(national_value = sum(sub_value, na.rm = TRUE)) %>%
    ungroup() %>%
    mutate(`Con%` = sub_value / national_value) %>%
    group_by(region) %>%
    arrange(period) %>%
    mutate(`Growth%` = sub_value / lag(sub_value) - 1,
           `Share%` = sub_value / complete_value,
           `ΔShare%` = `Share%` - lag(`Share%`),
           EI = `Share%` / lag(`Share%`) * 100) %>%
    ungroup() %>%
    filter(period == max(period, na.rm = TRUE),
           region %in% form$Display) %>%
    bind_rows(table3) %>%
    mutate(region = factor(region, levels = table2$region)) %>%
    right_join(distinct(table2, region), by = "region") %>%
    mutate(type = paste0("Internal Product Performance ", first(na.omit(period)))) %>%
    select(region,
           sub_value,
           `Con%`,
           `Growth%`,
           `Share%`,
           `ΔShare%`,
           EI,
           type)

  table.file <- tabular(Heading(unique(form$Summary1), character.only = TRUE) * table2$region ~
                          Heading(unique(table2$type), character.only = TRUE) * identity *
                          (Heading(paste0("Value(", unique(form$Digit), ")"), character.only = TRUE) *
                             table2$complete_value +
                             Heading("Con%") * table2$`Con%` +
                             Heading("Growth%") * table2$`Growth%` +
                             Heading("MNC%") * table2$`MNC%`) +
                          Heading(unique(table4$type), character.only = TRUE) * identity *
                          (Heading(paste0("Value(", unique(form$Digit), ")"), character.only = TRUE) *
                             table4$sub_value +
                             Heading("Con%") * table4$`Con%` +
                             Heading("Growth%") * table4$`Growth%` + Heading("Share%") * table4$`Share%` +
                             Heading("ΔShare%") * table4$`ΔShare%` + Heading("EI") * table4$EI))

  table.file=as.matrix(table.file)
  table.file
  write.xlsx(table.file,file=paste0(directory,'/',page,'.xlsx'),col.names=FALSE)
}



