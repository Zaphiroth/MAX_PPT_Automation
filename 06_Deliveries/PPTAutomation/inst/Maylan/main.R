# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Maylan PPT Run
# programmer:   Zhe Liu
# Date:         2020-05-12
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


##---- Generation function ----
GenerateFile <- function(page,
                         data,
                         form.table,
                         func.dir,
                         directory) {
  ##---- Identify table type ----
  form <- form.table %>%
    filter(Page == page)

  type <- form %>%
    distinct(Type) %>%
    unlist()

  if (!is.na(unique(form$Region))) {
    region.select <- stri_split_fixed(unique(form$Region), ":", simplify = TRUE) %>%
      as.character()

    data <- data %>%
      filter(!!sym(region.select[1]) == region.select[2])
  }

  if (is.na(unique(form$Digit))) {
    digit <- 1

  } else if (unique(form$Digit) == "K") {
    digit <- 1000

  } else if (unique(form$Digit) == "Mn") {
    digit <- 1000000

  } else {
    digit <- 1
  }

  ##---- Calculate function ----
  source(paste0(func.dir, "/MarketSize.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/SubMarketShare.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/SubMarketGrowth.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/Performance.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/Trend.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/ShareTrend.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/Ranking.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/RegionPerformance.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/RegionTrend.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/CompetitorPerformance.R"), encoding = "UTF-8")

  if (type == "MarketSize") {
    MarketSize(data, form, page, digit, directory)

  } else if (type == "SubMarketShare") {
    SubMarketShare(data, form, page, digit, directory)

  } else if (type == "SubMarketGrowth") {
    SubMarketGrowth(data, form, page, digit, directory)

  } else if (type == "Performance") {
    Performance(data, form, page, digit, directory)

  } else if (type == "Trend") {
    Trend(data, form, page, digit, directory)

  } else if (type == "ShareTrend") {
    ShareTrend(data, form, page, digit, directory)

  } else if (type == "Ranking") {
    Ranking(data, form, page, digit, directory)

  } else if (type == "RegionPerformance") {
    RegionPerformance(data, form, page, digit, directory)

  } else if (type == "RegionTrend") {
    RegionTrend(data, form, page, digit, directory)

  } else if (type == "CompetitorPerformance") {
    CompetitorPerformance(data, form, page, digit, directory)

  }
}







