# Lapse Report generator
# Jadon Thomson 
# 2017-02-03

#
# Initialize  
#

rm(list = ls())
gc()
Path <- getwd()
source(paste0(Path,"/R/Installer.R"))


# Set end date (generally leave 5 months gap)
DateEnd <- as.Date("2016/09/30")

# Begin Date, if blank will take from the beginning of the book
DateStart <- as.Date("2014-01-01")     ####### Start date must always be the first of the month!!! #######




# 
# Set rules
#



Agent = "ALL"          # Type Agents name if you would like to analyse only one agent.

#
#
#
#
#
#

#
# Import Data (put files into a folder with the current date under data folder)
#

source(paste0(Path,"/R/Import_Data.R"))






#
# Clean Data
#

source(paste0(Path,"/R/Clean_Data.R"))






#
# Create Report
#

source(paste0(Path,"/R/Report.R"))
