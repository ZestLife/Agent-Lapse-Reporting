# Create Lapse Report

# Create new All_Data matrix depending on the criteria specified in the home page
All_Data <- All_Data_Original

Agents <- paste(Active_Agents_Data$AgentName[Active_Agents_Data$Active == "YES"], Active_Agents_Data$AgentSurname[Active_Agents_Data$Active == "YES"])
if (toupper(Agent) != "ALL") {Agents = Agent}

All_Data <- All_Data[All_Data$AGENTNAME %in% Agents,]



## Create report matrix.

no_col      <- month(DateEnd) - month(DateStart) + (year(DateEnd) - year(DateStart))*12
no_row      <- ifelse(toupper(Agent) == "All", 4, length(Agents)) * 4 + 4
rows        <- c("NTU","0 - 2", "2 - 6", "6 -  ", "2 months", "6 months", ">6 months")

work_names <- c("Whole Book","Whole Book","Whole Book","Whole Book")
row_names_1 <- c("Whole Book", "", "", "")
work_names_2 <- c("NTU","0 - 2", "2 - 6", "6 -  ")
row_names_2 <- c("NTU", "2 months", "6 months", ">6 months")
for (i in 1:length(Agents)){
  work_names[4*i + 1] = Agents[i]
  row_names_1[4*i + 1] = Agents[i]
  work_names_2[4*i + 1] = "NTU"
  row_names_2[4*i + 1] = "NTU"
  for (j in 1:3){
    work_names[4*i + 1 + j] = Agents[i]
    row_names_1[4*i + 1 + j] = " "
    work_names_2[4*i + 1 + j] = rows[j+1]
    row_names_2[4*i + 1 + j] = rows[j+4]
  }
}

report.df = data.frame(Agent = work_names, Lapse = work_names_2)

DATE <- DateStart
col = 3
row = 1

while (col <= no_col + 1){
 
  # create column with heading as the month and year
  report.df[,as.character(paste(year(DATE), month.abb[month(DATE)]))] <- NA
  
  # Whole Book
  row = 1

  
  # For each Agent
  # for row in 1:no_row
  
  while (row <= no_row){
    
    Current_Agent <- as.character(report.df[row,1])
    # lapsed = number of policies with commencement date = above date and month end date - commencement date < (row months)
    if (report.df[row, 2] == "NTU"){
      # Expsure for NTU is all policies which commenced before DATE
      exposure <- ifelse(report.df[row, 1] == "Whole Book",
                         sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE),
                         sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE & All_Data$AGENTNAME == Current_Agent))
      # The number of policies which NTUed are all policies which commenced on DATE whose Status is NTU
      lapsed <- ifelse(report.df[row, 1] == "Whole Book",
                       sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE & All_Data$STATUS == "NTU"),
                       sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE & All_Data$STATUS == "NTU" & All_Data$AGENTNAME == Current_Agent))
      report.df[row, col] = ifelse(exposure == 0, "-", lapsed/exposure)
      
    } else {
      # For non-NTU, the exposure is the number of policies which commenced on DATE and who ended after the beginning of the period of interest (who were'nt NTU)
      # Since the number of months between policies which ended on the last day of the month is always one less than the duration,
      #       we need to add one to the difference between the start and end date. We must then exclude all NTUs because NTUs would now have a duration of 1.
      exposure <- ifelse(report.df[row, 1] == "Whole Book",
                         sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE & All_Data$STATUS != "NTU"),
                         sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE & All_Data$STATUS != "NTU" & All_Data$AGENTNAME == Current_Agent))
      lapsed <- ifelse(report.df[row, 1] == "Whole Book",
                       sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE &
                             elapsed_months(All_Data$NEW_END_DATE, All_Data$COMMENCEMENTDATEOFPOLICY) + 1 >= substr(report.df[row,2],1,1) &
                             elapsed_months(All_Data$NEW_END_DATE, All_Data$COMMENCEMENTDATEOFPOLICY) + 1 <= ifelse(substr(report.df[row,2],5,5) == " ", 
                                                                                                             12000, substr(report.df[row,2],5,5)) &
                             All_Data$STATUS != "NTU" & All_Data$STATUS != "ACT"),
                       sum(All_Data$COMMENCEMENTDATEOFPOLICY == DATE &
                             elapsed_months(All_Data$NEW_END_DATE, All_Data$COMMENCEMENTDATEOFPOLICY) + 1 >= substr(report.df[row,2],1,1) &
                             elapsed_months(All_Data$NEW_END_DATE, All_Data$COMMENCEMENTDATEOFPOLICY) + 1 <= ifelse(substr(report.df[row,2],5,5) == " ", 
                                                                                                             12000, substr(report.df[row,2],5,5)) &
                             All_Data$STATUS != "NTU" & All_Data$STATUS != "ACT" & All_Data$AGENTNAME == Current_Agent))
      report.df[row, col] = ifelse(exposure == 0, "-", lapsed/exposure)
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


report.df$Agent = row_names_1
report.df$Lapse = row_names_2


rm(col, row, DATE, row_names_1, row_names_2, work_names, work_names_2)

report.df[is.na(report.df)] <- ""

# if (All_Affinities == "YES" & length(Affinities) > 1) {Affinities <- c("not", Affinities)}

dir.create(file.path(paste0(Path, "/Results"),Sys.Date()), showWarnings = F)
write.csv(report.df, paste0(Path, "/Results/",Sys.Date(),"/", "Agent Lapse Report.csv"))



# 
# 
# 
# 
# # Put data into excel template:
# # Open workbook
# wb <- loadWorkbook(paste(Path, "/Results/Template.xlsx", sep = ""))
# # select sheet
# report_sheet = readWorksheet(wb, sheet = getSheets(wb)[1])
# # for all columns and rows in DATAFRAME, replace cells in worksheet
# for (row in 1:nrow(report.df)){
#   for (col in 1:ncol(report.df)){
#     report_sheet[row + 1, col + 1] <- report.df[row, col]
#   }
# }
# # set column heading
# for (col in 1:ncol(report.df)){
#   report_sheet[1, col + 1] <- colnames(report.df)[col]
# }
# 
# if (dir.exists(paste(Path, "/Results/", Sys.Date(), sep = ""))){} else {
#   dir.create(paste(Path, "/Results/", Sys.Date(), sep = ""))
# }
# 
# # Save workbook into new location.
# writeWorksheet(wb,report_sheet,sheet=getSheets(wb)[1],startRow=2,header=F)
# saveWorkbook(wb, paste(Path, "/Results/", Sys.Date(),"/Report.xlsx", sep = ""))
# 
# 
# 


