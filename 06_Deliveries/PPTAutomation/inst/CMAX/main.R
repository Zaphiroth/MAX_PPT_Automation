# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      PPT Run
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

  if (!is.na(unique(form$Restriction))) {
    restriction <- stri_split_fixed(unique(form$Restriction), ",", simplify = TRUE) %>%
      as.character()

    for (i in restriction) {
      rst <- stri_split_fixed(i, ":", simplify = TRUE) %>%
        as.character()

      data <- data %>%
        filter(!!sym(rst[1]) == rst[2])
    }
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
  source(paste0(func.dir, "/MarketSizeQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/SubMarketShareQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/SubMarketGrowthQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/PerformanceQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/TrendQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/ShareTrendQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/RankingQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/RegionPerformanceQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/RegionTrendQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/CompetitorPerformanceQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/GrowthTrendQ.R"), encoding = "UTF-8")
  source(paste0(func.dir, "/DisplayFunctionQ.R"), encoding = "UTF-8")

  if (type == "MarketSize") {
    print(page)
    MarketSize(data, form, page, digit, directory)

  } else if (type == "SubMarketShare") {
    print(page)
    SubMarketShare(data, form, page, digit, directory)

  } else if (type == "SubMarketGrowth") {
    print(page)
    SubMarketGrowth(data, form, page, digit, directory)

  } else if (type == "Performance") {
    print(page)
    Performance(data, form, page, digit, directory)

  } else if (type == "Trend") {
    print(page)
    Trend(data, form, page, digit, directory)

  } else if (type == "ShareTrend") {
    print(page)
    ShareTrend(data, form, page, digit, directory)

  } else if (type == "Ranking") {
    print(page)
    Ranking(data, form, page, digit, directory)

  } else if (type == "RegionPerformance") {
    print(page)
    RegionPerformance(data, form, page, digit, directory)

  } else if (type == "RegionTrend") {
    print(page)
    RegionTrend(data, form, page, digit, directory)

  } else if (type == "CompetitorPerformance") {
    print(page)
    CompetitorPerformance(data, form, page, digit, directory)

  } else if (type == "GrowthTrend") {
    print(page)
    GrowthTrend(data, form, page, digit, directory)

  } else {
    stop("Type Error!")
  }
}







