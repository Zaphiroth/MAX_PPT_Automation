# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Function
# Date:         2020-05-22
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
## Test Code P23/24 RegionTrend

RegionTrend <- function(data,
                        form,
                        page,
                        digit,
                        directory) {

  if(unique(form$Period)=='MTH') {
    grouper <- 'MTH' } else {
      grouper <- 'MAT'
    }

  if (unique(form$Period)=='MTH') {

    table.file <- data %>%
      group_by(period = !!sym(unique(form$Period)),
               region = !!sym(unique(form$Summary1))) %>%
      summarise(Value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE)) %>%
      ungroup() %>%
      mutate (Value = Value/digit) %>%
      filter(period %in% tail(unique(period), 36)) %>%
      mutate(month = str_sub(period, -2, -1),
             year = str_sub(period, 1, 2)) %>%
      group_by(period) %>%
      mutate (total = sum(Value, na.rm = TRUE)) %>%
      ungroup() %>%
      mutate(Share = Value / total) %>%
      arrange(month, region) %>%
      group_by(month, region) %>%
      mutate(EI = (Share / lag(Share)) * 100) %>%
      ungroup() %>%
      na.omit() %>%
      select(region, period, Share, EI) %>%
      gather(category, value, Share:EI) %>%
      mutate(Name = paste(region, category, sep="_")) %>%
      spread(period, value) %>%
      arrange(desc(category), desc(Name)) %>%
      select(Name, everything(), -category, -region) %>%
      rename(' ' = Name)

    table.file <- DisplayFunction(table.file = table.file, type='MTH')
    write.xlsx(table.file, paste0(directory, '/', page, '.xlsx'))

  } else {

    if (unique(form$Period) == 'MAT') {
      rollparam <- 12
    } else {
      rollparam <- 3
    }

    table.file <- data %>%
      group_by(PeriodMAT = !!sym(grouper),
               PeriodMTH = MTH,
               summary = !!sym(unique(form$Summary1))) %>%
      summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) /
                  digit) %>%
      ungroup() %>%
      filter(PeriodMTH %in% tail(sort(unique(PeriodMTH)),
                                 min(48, length(unique(PeriodMTH))))) %>%
      arrange(summary, PeriodMAT, PeriodMTH) %>%
      group_by(summary) %>%
      mutate(RollValue = c(rep(NA, (rollparam - 1)),
                           rollsum(value, rollparam))) %>%
      ungroup() %>%
      na.omit() %>%
      mutate(period = paste(PeriodMTH, unique(form$Period), sep=' ')) %>%
      select(period, summary, RollValue, -value, -PeriodMAT, -PeriodMTH) %>%
      group_by(period) %>%
      mutate(Share = RollValue / sum(RollValue, na.rm = TRUE)) %>%
      ungroup() %>%
      select(-RollValue) %>%
      mutate(month = str_sub(period, 4, 5),
             year = str_sub(period, 1, 2)) %>%
      arrange(month, summary) %>%
      group_by(month, summary) %>%
      mutate(EI=(Share / lag(Share)) * 100) %>%
      ungroup() %>%
      na.omit() %>%
      select(summary, period, Share, EI) %>%
      gather(category, value, Share:EI) %>%
      mutate(Name = paste(summary, category, sep="_")) %>%
      spread(period, value) %>%
      arrange(desc(category), desc(Name)) %>%
      select(Name, everything(), -category, -summary)  %>%
      rename(' ' = Name)

    if(ncol(table.file)>25) {
      table.file <- table.file[, -(ncol(table.file) - 24)]
    }

    if (nrow(table.file) != 0) {
      table.file <- DisplayFunction(table.file = table.file,
                                    type = unique(form$Period))
      write.xlsx(table.file, paste0(directory, '/', page, '.xlsx'))
    } else {
      print ('Warning: Insufficient Data to Calculate
             Rolling MAT Growth Rates!')
    }
  }
}

