# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-15
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


ProductRankFunc <- function(data,
                            page,
                            form.table,
                            directory) {
  ##---- P14_1 ----
  form14.1 <- form.table %>% 
    filter(Page == "P14_1")
  
  table14.1.1 <- data %>% 
    filter(MAT %in% head(sort(unique(data$MAT), decreasing = TRUE), 2),
           !!sym(unique(form14.1$Summary1)) %in% unique(form14.1$Content)) %>% 
    group_by(period = MAT,
             Product = !!sym(unique(form14.1$Summary1)),
             Manufactor = !!sym(unique(form14.1$Summary2)),
             `MNC/Local` = !!sym(unique(form14.1$Summary3))) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form14.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(Ranking = rank(-`RMB(Mn)`)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period, value.var = "Ranking") %>% 
    arrange(!!sym(sort(unique(data$MAT), decreasing = TRUE)[1]))
  
  table14.1.2 <- data %>% 
    filter(YTD %in% head(sort(unique(data$YTD), decreasing = TRUE), 2),
           !!sym(unique(form14.1$Summary1)) %in% unique(form14.1$Content)) %>% 
    group_by(period = YTD,
             Product = !!sym(unique(form14.1$Summary1)),
             Manufactor = !!sym(unique(form14.1$Summary2)),
             `MNC/Local` = !!sym(unique(form14.1$Summary3))) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form14.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(Ranking = rank(-`RMB(Mn)`)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period, value.var = "Ranking")
  
  table14.1.3 <- data %>% 
    filter(MTH %in% head(sort(unique(data$MTH), decreasing = TRUE), 2),
           !!sym(unique(form14.1$Summary1)) %in% unique(form14.1$Content)) %>% 
    group_by(period = MTH,
             Product = !!sym(unique(form14.1$Summary1)),
             Manufactor = !!sym(unique(form14.1$Summary2)),
             `MNC/Local` = !!sym(unique(form14.1$Summary3))) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form14.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(Ranking = rank(-`RMB(Mn)`)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period, value.var = "Ranking")
  
  table14.1.4 <- data %>% 
    filter(MAT %in% head(sort(unique(data$MAT), decreasing = TRUE), 2)) %>% 
    group_by(period = MAT,
             Product = !!sym(unique(form14.1$Summary1)),
             Manufactor = !!sym(unique(form14.1$Summary2)),
             `MNC/Local` = !!sym(unique(form14.1$Summary3))) %>% 
    summarise(`RMB(Mn)` = sum(!!sym(unique(form14.1$Calculation)), na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = `RMB(Mn)` / sum(`RMB(Mn)`, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(Product %in% unique(form14.1$Content)) %>% 
    setDT() %>% 
    melt(id.vars = c("Product", "Manufactor", "MNC/Local", "period"), 
         variable.factor = FALSE, variable.name = "index", value.name = "value") %>% 
    unite(period_index, period, index) %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period_index, value.var = "value") %>% 
    adorn_totals("row", fill = NA_character_, na.rm = TRUE, name = "Sum") %>% 
    melt(id.vars = c("Product", "Manufactor", "MNC/Local"), variable.factor = FALSE, 
         variable.name = "period_index", value.name = "value") %>% 
    separate(period_index, c("period", "index"), sep = "_") %>% 
    dcast(Product + Manufactor + `MNC/Local` + period ~ index, value.var = "value") %>% 
    group_by(Product) %>% 
    arrange(period) %>% 
    mutate(`Growth%` = `RMB(Mn)` / lag(`RMB(Mn)`) - 1) %>% 
    ungroup() %>% 
    filter(period == max(period)) %>% 
    select(-period) %>% 
    mutate(index_type = sort(unique(data$MAT), decreasing = TRUE)[1])
  
  table14.1.5 <- table14.1.4 %>% 
    select(Product, Manufactor, `MNC/Local`) %>% 
    mutate(index_type = "Product Info")
  
  table14.1 <- table14.1.1 %>% 
    left_join(table14.1.2, by = c("Product", "Manufactor", "MNC/Local")) %>% 
    left_join(table14.1.3, by = c("Product", "Manufactor", "MNC/Local")) %>% 
    full_join(table14.1.4, by = c("Product", "Manufactor", "MNC/Local")) %>% 
    mutate(` ` = factor(Product, levels = Product)) %>% 
    tabular(` ` ~ 
              Heading("Ranking") * identity * 
              (sym(sort(unique(data$MAT), decreasing = TRUE)[1]) + 
                 sym(sort(unique(data$MAT), decreasing = TRUE)[2]) + 
                 sym(sort(unique(data$YTD), decreasing = TRUE)[1]) + 
                 sym(sort(unique(data$YTD), decreasing = TRUE)[2])) + 
              Heading("Product Info") * identity * 
              (Product + Manufactor + `MNC/Local`) + 
              Heading(sort(unique(data$MAT), decreasing = TRUE)[1], character.only = TRUE) * identity * 
              (`RMB(Mn)` + `Growth%` + `Share%`),
            data = .)
  
  
  list(table14.1)
}



