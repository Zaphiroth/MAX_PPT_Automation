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

  display <- data.frame(period = seq(max(data$Date), max(data$Date)-400, -100)) %>%
    arrange(period) %>%
    mutate(period = paste0(stri_sub(period, 3, 4), "M", stri_sub(period, 5, 6), " ", unique(form$Period)),
           period_index = paste0(period, " Share")) %>%
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
    right_join(display, by = c("period_index", "sub_market")) %>%
    setDT() %>%
    dcast(sub_market ~ period_index, value.var = "value") %>%
    select(sub_market,
           ends_with(unique(form$Period)),
           ends_with("Share")) %>%
    right_join(distinct(form, Display), by = c("sub_market" = "Display"))

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}



