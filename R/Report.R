# Create Lapse Report

# Create new All_Data matrix depending on the criteria specified in the home page
All_Data <- All_Data_Original

source(paste(Path, "/R_Code/OpenDB.R", sep = ""))

Active_Agents_Data <- dbReadTable(my_new_db, "agents")

if (toupper(Agent) != "ALL") {All_Data <- All_Data[All_Data$SALESAGENT == Agent,]}




## Create report matrix.

no_col      <- month(DateEnd) - month(DateStart) + (year(DateEnd) - year(DateStart))*12
no_row      <- elapsed_months(DateEnd, as.Date("2010-01-01"))
report.df   <- data.frame(duration = 0:no_row)

row_names <- c("NTU", 1:no_row)

DATE <- DateStart
col = 2
row = 1

while (col <= no_col){
 
  # create column with heading as the month and year
  report.df[,as.character(paste(year(DATE), month.abb[month(DATE)]))] <- NA
  

  # for row in 1:no_row
  row = 1
  while (row <= no_row + 1){
    # lapsed = number of policies Commencement Date < DATE and new month(end date) - month(commencement date) = row

    if (row == 1){
      # Expsure for NTU is all policies which commenced before DATE
      exposure <- sum(All_Data$COMMENCEMENTDATE <= DATE & !is.na(All_Data$NEW_END_DATE))
      # The number of policies which NTUed are all policies which commenced before DATE whose Status is NTU
      lapsed <- sum(All_Data$COMMENCEMENTDATE <= DATE &
                                All_Data$STATUS == "NTU" &
                                !is.na(All_Data$NEW_END_DATE))
      report.df[row, col] = lapsed/exposure
    } else {
      # For non-NTU, the exposure is the number of policies which commenced before DATE and who ended after duration
      # Since the number of months between policies which ended on the last day of the month is always one less than the duration,
      #       we need to add one to the difference between the start and end date. We must then exclude all NTUs because NTUs would now have a duration of 1.
      exposure <- sum(All_Data$COMMENCEMENTDATE <= DATE & 
                                  elapsed_months(All_Data$NEW_END_DATE, All_Data$COMMENCEMENTDATE) + 1 >= report.df[row,1] &
                                  elapsed_months(DATE, All_Data$COMMENCEMENTDATE) >= report.df[row,1] &
                                  !is.na(All_Data$NEW_END_DATE) &
                                  All_Data$STATUS != "NTU")
      lapsed <- sum(All_Data$COMMENCEMENTDATE <= DATE &
                                elapsed_months(All_Data$NEW_END_DATE, All_Data$COMMENCEMENTDATE) + 1 == report.df[row,1] &
                                elapsed_months(DATE, All_Data$COMMENCEMENTDATE) >= report.df[row,1] &
                                !is.na(All_Data$NEW_END_DATE) &
                                All_Data$STATUS != "NTU")
      report.df[row, col] = lapsed/exposure
    }
    
    # report.df[row, col] = lapsed/exposure
    row  <-  row + 1
  }  
  
  # create a DATE variable for the column
  if (month(DATE) == 12) {
    DATE <- as.Date(paste(year(DATE)+1, "01", "01", sep="-"))
  }else{
    DATE <- as.Date(paste(year(DATE), month(DATE)+1, "01", sep="-"))
  } 
  
col <- col+1  
}


report.df[1,1] <- "NTU"
report.df <- subset(report.df, select = -duration)
rownames(report.df) <- row_names

rm(col, row, DATE, no_col, no_row, row_names)

report.df[is.na(report.df)] <- ""
if (All_Affinities == "YES" & length(Affinities) > 1) {Affinities <- c("not", Affinities)}

dir.create(file.path(paste0(Path, "/Results"),Sys.Date()), showWarnings = F)
write.csv(report.df, paste0(Path, "/Results/",Sys.Date(),"/", paste0(Affinities, collapse = " "),Groups, Voluntary," Groups ",Agent, " Agents.csv"))
