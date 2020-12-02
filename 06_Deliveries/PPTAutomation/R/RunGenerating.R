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
                 "tables",
                 "zoo")
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

  options(java.parameters = "-Xmx2048m",
          scipen = 200,
          stringsAsFactors = FALSE)

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
  filenames <- list.files(path = data.dir,
                          pattern = "*.xlsx", full.names = TRUE)

  file.list <- map(filenames, function(filename) {
    print(filename)
    file.data <- data.frame(read_excel(filename), check.names = FALSE)
  })

  raw.data <- bind_rows(file.list)

  form.table <- data.frame(read_excel(form.dir), check.names = FALSE)

  if (is.null(raw.data$Date)) {
    stop("Data is not available for the tool.")
  }

  ##---- Set period ----
  if(grepl('Q',raw.data$Date[9],fixed=TRUE)==TRUE) {
    raw.data <- raw.data %>%
      mutate(Year=str_sub(Date,1,4),
             Month=as.numeric(str_sub(Date,nchar(raw.data$Date[1])))) %>%
      mutate(Month=str_pad(as.character(Month*3),2,pad='0')) %>%
      mutate(Date=paste0(Year,Month)) %>%
      select(-Year,-Month)

    cycle <- 4
  } else {
    cycle <- 12
  }

  latest.date <- sort(unique(raw.data$Date), decreasing = TRUE)

  mat.date <- data.frame(Date = sort(unique(raw.data$Date))) %>%
    mutate(ymd = ymd(stri_paste(Date, "01")),
           MAT = case_when(
             ymd <= max(ymd) & ymd > max(ymd)-months(x = 12) ~
               stri_paste(stri_sub(latest.date[1], 3, 4), "M",
                          stri_sub(latest.date[1], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 12) & ymd > max(ymd)-months(x = 24) ~
               stri_paste(stri_sub(latest.date[1+cycle], 3, 4), "M",
                          stri_sub(latest.date[1+cycle], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 24) & ymd > max(ymd)-months(x = 36) ~
               stri_paste(stri_sub(latest.date[1+2*cycle], 3, 4), "M",
                          stri_sub(latest.date[1+2*cycle], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 36) & ymd > max(ymd)-months(x = 48) ~
               stri_paste(stri_sub(latest.date[1+3*cycle], 3, 4), "M",
                          stri_sub(latest.date[1+3*cycle], 5, 6), " MAT"),
             ymd <= max(ymd)-months(x = 48) & ymd > max(ymd)-months(x = 60) ~
               stri_paste(stri_sub(latest.date[1+4*cycle], 3, 4), "M",
                          stri_sub(latest.date[1+4*cycle], 5, 6), " MAT"),
             TRUE ~ NA_character_
           )) %>%
    inner_join(raw.data, by = "Date") %>%
    distinct(Date, MAT) %>%
    add_count(MAT, name = "n") %>%
    filter(n == cycle) %>%
    select(Date, MAT)


  ytd.date <- data.frame(Date = sort(unique(raw.data$Date))) %>%
    mutate(ymd = ymd(stri_paste(Date, "01")),
           YTD = case_when(
             ymd <= max(ymd) &
               stri_sub(Date, 1, 4) == stri_sub(latest.date[1], 1, 4) ~
               stri_paste(stri_sub(latest.date[1], 3, 4), "M",
                          stri_sub(latest.date[1], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 12) &
               stri_sub(Date, 1, 4) == stri_sub(latest.date[1+cycle], 1, 4) ~
               stri_paste(stri_sub(latest.date[1+cycle], 3, 4), "M",
                          stri_sub(latest.date[1+cycle], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 24) &
               stri_sub(Date, 1, 4) == stri_sub(latest.date[1+2*cycle], 1, 4) ~
               stri_paste(stri_sub(latest.date[1+2*cycle], 3, 4), "M",
                          stri_sub(latest.date[1+2*cycle], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 36) &
               stri_sub(Date, 1, 4) == stri_sub(latest.date[1+3*cycle], 1, 4) ~
               stri_paste(stri_sub(latest.date[1+3*cycle], 3, 4), "M",
                          stri_sub(latest.date[1+3*cycle], 5, 6), " YTD"),
             ymd <= max(ymd)-months(x = 48) &
               stri_sub(Date, 1, 4) == stri_sub(latest.date[1+4*cycle], 1, 4) ~
               stri_paste(stri_sub(latest.date[1+4*cycle], 3, 4), "M",
                          stri_sub(latest.date[1+4*cycle], 5, 6), " YTD"),
             TRUE ~ NA_character_
           )) %>%
    select(Date, YTD)

  mth.date <- data.frame(Date = sort(unique(raw.data$Date))) %>%
    mutate(MTH = stri_paste(stri_sub(Date, 3, 4), "M",
                            stri_sub(Date, 5, 6))) %>%
    select(Date, MTH)

  data <- raw.data %>%
    distinct() %>%
    left_join(mat.date, by = "Date") %>%
    left_join(ytd.date, by = "Date") %>%
    left_join(mth.date, by = "Date")

  ##---- Generate files ----
  source(paste0(mainDir, "/main.R"), encoding = "UTF-8")

  map(unique(form.table$Page),
      GenerateFile,
      data = data,
      form.table = form.table,
      func.dir = funcDir,
      directory = directory)

  print("Success!")
}
