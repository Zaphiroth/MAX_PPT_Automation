# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# Date:         2020-06-18
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


ShareTrend <- function(data,
                       form,
                       page,
                       digit,
                       directory) {

  dataQ <- data %>%
    mutate(MATY = str_sub(data$MAT, 1, 2),
           YTDY = str_sub(data$YTD, 1, 2),
           QT = as.numeric(str_sub(data$Date, 5, 6)) / 3) %>%
    mutate(MAT = paste0(MATY, 'Q4 MAT'),
           YTD = paste0(YTDY, 'Q4 YTD'),
           MTH = paste0(MATY, 'Q', QT)) %>%
    select(-MATY, -YTDY, -QT)

  if(unique(form$Period) == 'MTH') {
    grouper <- 'MTH'
  } else {
    grouper <- 'MAT'
  }

  table.file <- dataQ %>%
    filter(!!sym(grouper) %in%
             tail(sort(unique(unlist(dataQ[, grouper]))),
                  min(8, length(unique(unlist(dataQ[, grouper])))))) %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others"))

  if (unique(form$Period) == 'MTH') {
    table.file <- table.file %>%
      group_by(period = !!sym(unique(form$Period)), summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
      ungroup() %>%
      group_by(period) %>%
      mutate(value_share = value / sum(value, na.rm = TRUE)) %>%
      ungroup() %>%
      setDT() %>%
      dcast(summary ~ period, value.var = "value_share") %>%
      right_join(distinct(form, Display), by = c("summary" = "Display")) %>%
      rename(!!sym(' ') := summary)

  } else {

    if (unique(form$Period) == 'MAT') {
      rollparam <- 12
    } else {
      rollparam <- 3
    }

    table.file <- table.file %>%
      group_by(PeriodMAT = !!sym(grouper),
               PeriodMTH = MTH,
               summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
      ungroup() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(36, length(unique(PeriodMTH))))) %>%
      arrange(summary, PeriodMAT, PeriodMTH) %>%
      group_by(summary) %>%
      mutate(RollValue = c(rep(NA, (rollparam - 1)), rollsum(value, rollparam))) %>%
      ungroup() %>%
      na.omit() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(24, length(unique(PeriodMTH))))) %>%
      mutate(period = paste(PeriodMTH, unique(form$Period), sep=' ')) %>%
      select(period, summary, RollValue, -value, -PeriodMAT, -PeriodMTH) %>%
      group_by(period) %>%
      mutate(value_share = RollValue / sum(RollValue, na.rm = TRUE)) %>%
      ungroup() %>%
      select(-RollValue) %>%
      spread(period, value_share) %>%
      mutate(sequence = ifelse(summary == 'Others', 2,
                               ifelse(summary=='Total', 3,
                                      1))) %>%
      arrange(sequence, summary) %>%
      select(-sequence) %>%
      rename(' ' = summary)

    table.file <- DisplayFunction(table.file = table.file, type = unique(form$Period))
  }

  table.file
  write.xlsx(table.file, paste0(directory, '/', page, '.xlsx'))
}
