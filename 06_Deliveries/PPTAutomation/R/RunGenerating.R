#' Setup the prerequest env
#'
#' @author Zhe Liu
#' @description Setup the prerequest env
#' @example prerequest()
#'
prerequest <- function() {
  if (!requireNamespace("pacman")) install.packages("pacman")
  pacman::p_load("zip",
                 "openxlsx",
                 "readxl",
                 "writexl",
                 "RcppRoll",
                 "plyr",
                 "stringi",
                 "feather",
                 "RODBC",
                 "MASS",
                 "car",
                 "data.table",
                 "plotly",
                 "tidyverse",
                 "lubridate",
                 "janitor",
                 "digest",
                 "tables")
}


#' Initiate the shiny app
#'
#' @export
#' @author Zhe Liu
#' @description Initiate the shiny app to do the generating
#' @example RunGenerating()
#'
RunGenerating <- function(inst,
                          data.dir,
                          form.dir,
                          directory) {
  prerequest()

  mainDir <- system.file(inst,
                         package = "PPTAutomation")

  funcDir <- system.file(inst,
                         "function",
                         package = "PPTAutomation")

  if (funcDir == "") {
    stop("Could not find example directory. Try re-installing `PPTAutomation`.",
         call. = FALSE)
  }

  ##---- Readin ----
  ReadWorkbooks <- function(file) {
    filenames <- list.files(path = file,
                            pattern = "*.xlsx", full.names = TRUE)

    file.list <- map(filenames, function(filename) {
      print(filename)
      file.data <- read_excel(filename) %>%
        distinct()
    })

    file.list
  }

  raw.data <- ReadWorkbooks(directory) %>%
    bind_rows()

  form.table <- data.frame(read_excel(form.dir))

  if (is.null(raw.data$Date)) {
    stop("Data is not available for the tool.")
  }

  ##---- Set period ----
  latest.date <- sort(unique(raw.data$Date), decreasing = TRUE)

  mat.date <- data.frame(Date = sort(unique(raw.data$Date))) %>%
    mutate(ymd = ymd(stri_paste(Date, "01")),
           MAT = case_when(
             ymd <= max(ymd) & ymd > max(ymd)-months(x = 12) ~
               stri_paste(stri_sub(latest.date[1], 3, 4), "M", stri_sub(latest.date[1], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 12) & ymd > max(ymd)-months(x = 24) ~
               stri_paste(stri_sub(latest.date[13], 3, 4), "M", stri_sub(latest.date[13], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 24) & ymd > max(ymd)-months(x = 36) ~
               stri_paste(stri_sub(latest.date[25], 3, 4), "M", stri_sub(latest.date[25], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 36) & ymd > max(ymd)-months(x = 48) ~
               stri_paste(stri_sub(latest.date[37], 3, 4), "M", stri_sub(latest.date[37], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 48) & ymd > max(ymd)-months(x = 60) ~
               stri_paste(stri_sub(latest.date[49], 3, 4), "M", stri_sub(latest.date[49], 5, 6), " MAT"),
             TRUE ~ NA_character_
           )) %>%
    select(Date, MAT)

  ytd.date <- data.frame(Date = sort(unique(raw.data$Date))) %>%
    mutate(ymd = ymd(stri_paste(Date, "01")),
           YTD = case_when(
             ymd <= max(ymd) & stri_sub(Date, 1, 4) == stri_sub(latest.date[1], 1, 4) ~
               stri_paste(stri_sub(latest.date[1], 3, 4), "M", stri_sub(latest.date[1], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 12) & stri_sub(Date, 1, 4) == stri_sub(latest.date[13], 1, 4) ~
               stri_paste(stri_sub(latest.date[13], 3, 4), "M", stri_sub(latest.date[13], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 24) & stri_sub(Date, 1, 4) == stri_sub(latest.date[25], 1, 4) ~
               stri_paste(stri_sub(latest.date[25], 3, 4), "M", stri_sub(latest.date[25], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 36) & stri_sub(Date, 1, 4) == stri_sub(latest.date[37], 1, 4) ~
               stri_paste(stri_sub(latest.date[37], 3, 4), "M", stri_sub(latest.date[37], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 48) & stri_sub(Date, 1, 4) == stri_sub(latest.date[49], 1, 4) ~
               stri_paste(stri_sub(latest.date[49], 3, 4), "M", stri_sub(latest.date[49], 5, 6), " YTD"),
             TRUE ~ NA_character_
           )) %>%
    select(Date, YTD)

  mth.date <- data.frame(Date = sort(unique(raw.data$Date))) %>%
    mutate(MTH = stri_paste(stri_sub(Date, 3, 4), "M", stri_sub(Date, 5, 6))) %>%
    select(Date, MTH)

  data <- raw.data %>%
    left_join(mat.date, by = "Date") %>%
    left_join(ytd.date, by = "Date") %>%
    left_join(mth.date, by = "Date") %>%
    filter(!is.na(MAT), !is.na(YTD))

  ##---- Generate files ----
  source(paste0(mainDir, "/main.R"), encoding = "UTF-8")

  map(unique(form.table$Page),
      GenerateFile,
      data = data,
      form.table = form.table,
      func.dir = funcDir,
      directory = directory)
}
