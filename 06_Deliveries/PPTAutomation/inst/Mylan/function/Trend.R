# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-14
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


Trend <- function(data,
                  form,
                  page,
                  digit,
                  directory) {

  table.file <- data %>%
    filter(!!sym(unique(form$Period)) %in%
             tail(sort(unique(unlist(data[, unique(form$Period)]))), 24)) %>%
    mutate(summary = ifelse(!!sym(unique(form$Summary1)) %in% unique(form$Display),
                            !!sym(unique(form$Summary1)),
                            "Others")) %>%
    group_by(period = !!sym(unique(form$Period)), summary) %>%
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>%
    ungroup() %>%
    arrange(period) %>%
    setDT() %>%
    dcast(summary ~ period, value.var = "value") %>%
    right_join(distinct(form, Display), by = c("summary" = "Display")) %>%
    rename(` ` = summary)

  table.file
  write.xlsx(table.file,paste0(directory,'/',page,'.xlsx'))
}


