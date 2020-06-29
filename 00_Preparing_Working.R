# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ProjectName:  MAX PPT Automation
# Purpose:      Generate PPT tabulars using MAX data
# programmer:   Zhe Liu
# Date:         2020-05-07
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


options(java.parameters = "-Xmx2048m",
        scipen = 200,
        stringsAsFactors = FALSE)

##---- loading the required packages ----
suppressPackageStartupMessages({
  require(zip)
  require(openxlsx)
  require(readxl)
  require(writexl)
  require(RcppRoll)
  require(plyr)
  require(stringi)
  require(feather)
  require(RODBC)
  require(MASS)
  require(car)
  require(data.table)
  require(plotly)
  require(tidyverse)
  require(lubridate)
  require(janitor)
  require(RecordLinkage)
  require(digest)
  require(tables)
  require(zoo)
})


##---- setup the directories ----
system("mkdir 01_Background 02_Inputs 03_Outputs 04_Codes 05_Internal_Review 06_Deliveries")

RunGenerating(inst = "MAX", 
              data.dir = "E:/LZ/MAX_PPT_Automation/02_Inputs/MAX", 
              form.dir = "E:/LZ/MAX_PPT_Automation/03_Outputs/测试Form_table.xlsx", 
              directory = "E:/LZ/MAX_PPT_Automation/05_Internal_Review/test1")

RunGenerating(inst = "CMAX", 
              data.dir = "E:/LZ/MAX_PPT_Automation/02_Inputs/CMAX", 
              form.dir = "E:/LZ/MAX_PPT_Automation/03_Outputs/Form_table-CMAX.xlsx", 
              directory = "E:/LZ/MAX_PPT_Automation/05_Internal_Review/test2")
