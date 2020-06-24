# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


MarketSize <- function(data,
                       form,
                       page,
                       digit,
                       directory) {

  dataQ <- data %>%
    mutate(MATY = str_sub(data$MAT, 1, 2),
           YTDY = str_sub(data$YTD, 1, 2)) %>%
    mutate(MAT = paste0(MATY, 'Q4 MAT'),
           YTD = paste0(YTDY, 'Q4 YTD')) %>%
    select(-MATY, -YTDY)

  table.file <- dataQ %>%
    group_by(period = !!sym(unique(form$Period))) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    arrange(period) %>%
    mutate(growth = value / lag(value) - 1) %>%
    right_join(distinct(form, Display), by = c("period" = "Display")) %>%
    select(Period = period,
           !!sym(paste0("Value(", unique(form$Digit), ")")) := value,
           `Growth%` = growth)

  table.file
  write.xlsx(table.file, paste0(directory, '/', page, '.xlsx'))
}
