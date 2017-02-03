# Importing application form data 

# get file names
Data_Folder         = paste(Path, "/Data", sep="")
Data_Folders        = dir(path = Data_Folder)
Latest_Data_Folder  = paste(Data_Folder, max(Data_Folders), sep="/")
file_list           = list.files(paste0(Latest_Data_Folder, "/CSV"))

# import each file
All_Data <- data.frame()

for(file in 1:length(file_list)){
  
  file_location <- paste0(Latest_Data_Folder, "/CSV/", file_list[file])
  
  temp.df <- read.csv(file_location)   # import excel into a temporary data frame
  temp.df <- temp.df[-1,]                                             # the second row of the groups spread sheet has weird stuff in it (column sub-headings)
  
  
  colnames(temp.df) <- gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(colnames(temp.df))))
  temp.df <- temp.df[temp.df$POLICYNUMBER != "",]                                 # Removes summary info at the bottom of spreadsheets
  
  if (nrow(All_Data) == 0) {
    All_Data <- temp.df
  }else{
    
    common_col <- intersect(colnames(temp.df), colnames(All_Data))
    All_Data <- rbind(subset(All_Data, select = common_col),
                      subset(temp.df, select = common_col))              # otherwise append all data to temp
    }                                                           
}


# Calculate Age
yy    <- substr(All_Data$POLICYHOLDERIDNUMBER, 1,2)
mm    <- substr(All_Data$POLICYHOLDERIDNUMBER, 3,4)
dd    <- substr(All_Data$POLICYHOLDERIDNUMBER, 5,6)
bdate <- as.Date(paste0("19",yy,"-",mm,"-",dd))
All_Data$BDATE <- bdate

rm(yy,mm,dd,bdate,Data_Folder,Data_Folders,common_col,file,file_list,file_location,Latest_Data_Folder,temp.df)

