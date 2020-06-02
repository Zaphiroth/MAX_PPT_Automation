# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Mylan PPT Function
# programmer:   Zhe Liu
# Date:         2020-05-15
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


Ranking <- function(data,
                    form,
                    page,
                    digit,
                    directory) {
  
  table1 <- data %>% 
    filter(MAT %in% head(sort(unique(data$MAT), decreasing = TRUE), 2),
           !!sym(unique(form$Summary1)) %in% unique(form$Display)) %>% 
    group_by(period = MAT,
             Product = !!sym(unique(form$Summary1)),
             Manufactor = !!sym(unique(form$Summary2)),
             `MNC/Local` = !!sym(unique(form$Summary3))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(Ranking = rank(-value)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period, value.var = "Ranking") %>% 
    arrange(!!sym(sort(unique(data$MAT), decreasing = TRUE)[1]))
  
  table2 <- data %>% 
    filter(YTD %in% head(sort(unique(data$YTD), decreasing = TRUE), 2),
           !!sym(unique(form$Summary1)) %in% unique(form$Display)) %>% 
    group_by(period = YTD,
             Product = !!sym(unique(form$Summary1)),
             Manufactor = !!sym(unique(form$Summary2)),
             `MNC/Local` = !!sym(unique(form$Summary3))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(Ranking = rank(-value)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period, value.var = "Ranking")
  
  table3 <- data %>% 
    filter(MTH %in% sort(unique(data$MTH), decreasing = TRUE)[c(1, 13)],
           !!sym(unique(form$Summary1)) %in% unique(form$Display)) %>% 
    group_by(period = MTH,
             Product = !!sym(unique(form$Summary1)),
             Manufactor = !!sym(unique(form$Summary2)),
             `MNC/Local` = !!sym(unique(form$Summary3))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(Ranking = rank(-value)) %>% 
    ungroup() %>% 
    setDT() %>% 
    dcast(Product + Manufactor + `MNC/Local` ~ period, value.var = "Ranking")
  
  table4 <- data %>% 
    filter(MAT %in% head(sort(unique(data$MAT), decreasing = TRUE), 2)) %>% 
    group_by(period = MAT,
             Product = !!sym(unique(form$Summary1)),
             Manufactor = !!sym(unique(form$Summary2)),
             `MNC/Local` = !!sym(unique(form$Summary3))) %>% 
    summarise(value = sum(!!sym(unique(form$Calculation)), na.rm = TRUE) / digit) %>% 
    ungroup() %>% 
    group_by(period) %>% 
    mutate(`Share%` = value / sum(value, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(Product %in% unique(form$Display)) %>% 
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
    mutate(`Growth%` = value / lag(value) - 1) %>% 
    ungroup() %>% 
    filter(period == max(period, na.rm = TRUE)) %>% 
    select(-period) %>% 
    mutate(index_type = sort(unique(data$MAT), decreasing = TRUE)[1])
  
  table5 <- table4 %>% 
    select(Product, Manufactor, `MNC/Local`) %>% 
    mutate(index_type = "Product Info")
  
  table.file <- table1 %>% 
    left_join(table2, by = c("Product", "Manufactor", "MNC/Local")) %>% 
    left_join(table3, by = c("Product", "Manufactor", "MNC/Local")) %>% 
    full_join(table4, by = c("Product", "Manufactor", "MNC/Local")) %>% 
    mutate(` ` = factor(Product, levels = Product),
           !!sym(paste0("Value(", unique(form$Digit), ")")) := value) %>% 
    tabular(` ` ~ 
              Heading("Ranking") * identity * 
              (sym(sort(unique(data$MAT), decreasing = TRUE)[1]) + 
                 sym(sort(unique(data$MAT), decreasing = TRUE)[2]) + 
                 sym(sort(unique(data$YTD), decreasing = TRUE)[1]) + 
                 sym(sort(unique(data$YTD), decreasing = TRUE)[2]) + 
                 sym(sort(unique(data$MTH), decreasing = TRUE)[1]) + 
                 sym(sort(unique(data$MTH), decreasing = TRUE)[13])) + 
              Heading("Product Info") * identity * 
              (Product + Manufactor + `MNC/Local`) + 
              Heading(sort(unique(data$MAT), decreasing = TRUE)[1], character.only = TRUE) * identity * 
              (sym(paste0("Value(", unique(form$Digit), ")")) + `Growth%` + `Share%`),
            data = .)
  
  table.file=as.matrix(table.file)
  table.file
  write.xlsx(table.file,file=paste0(directory,'/',page,'.xlsx'),col.names=FALSE)
}



