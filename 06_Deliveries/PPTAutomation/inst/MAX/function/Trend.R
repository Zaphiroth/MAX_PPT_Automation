# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-14
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

Trend <- function(data,
                  form,
                  page,
                  digit,
                  directory) {

  if(unique(form$Period) == 'MTH') {
    grouper <- 'MTH'
  } else {
    grouper <- 'MAT'
  }

  table.file <- data %>%
    filter(!!sym(grouper) %in%
             tail(sort(unique(unlist(data[, grouper]))),
                  min(24,length(unique(unlist(data[, grouper])))))) %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in%
                              unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others"))

  if (unique(form$Period) == 'MTH') {
    table.file <- table.file %>%
      group_by(period = !!sym(unique(form$Period)), summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      arrange(period) %>%
      setDT() %>%
      dcast(summary ~ period, value.var = "value") %>%
      right_join(distinct(form, Display), by = c("summary" = "Display")) %>%
      rename(` ` = summary)

  } else  {

    if (unique(form$Period) == 'MAT') {
      rollparam <- 12
    } else {
      rollparam <- 3
    }

    table.file <- table.file %>%
      group_by(PeriodMAT = !!sym(grouper),
               PeriodMTH = MTH,
               summary) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(36, length(unique(PeriodMTH))))) %>%
      arrange(summary, PeriodMAT, PeriodMTH) %>%
      group_by(summary) %>%
      mutate(RollValue = c(rep(NA,(rollparam - 1)),
                           rollsum(value, rollparam))) %>%
      ungroup() %>%
      na.omit() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(24, length(unique(PeriodMTH))))) %>%
      mutate(period = paste(PeriodMTH, unique(form$Period), sep=' ')) %>%
      select(period, summary, RollValue, -value, -PeriodMAT, -PeriodMTH) %>%
      spread(period, RollValue) %>%
      adorn_totals("row", na.rm = TRUE, name = "Total") %>%
      mutate(sequence = ifelse(summary == 'Others', 2 ,
                               ifelse(summary == 'Total', 3,
                                      1))) %>%
      arrange(sequence, summary) %>%
      select(-sequence) %>%
      rename(' ' = summary)

    table.file <- DisplayFunction(table.file = table.file,
                                  type = unique(form$Period))
  }

  table.file
  write.xlsx(table.file, paste0(directory, '/', page, '.xlsx'))
}



