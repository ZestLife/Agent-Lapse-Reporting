# Clean data for lapse report


# Use ID to find gender
gggg    <- as.numeric(substr(All_Data$POLICYHOLDERIDNUMBER, 7,10))
gggg[nchar(as.character(All_Data$POLICYHOLDERIDNUMBER)) != 13] <- NA
gggg[gggg < 5000] <- "F"
gggg[gggg >= 5000] <- "M"

# Calculate start Age
All_Data$COMMENCEMENTDATEOFPOLICY <- DateConv(All_Data$COMMENCEMENTDATEOFPOLICY)
All_Data$STARTAGE <- year(All_Data$COMMENCEMENTDATEOFPOLICY) - year(All_Data$BDATE)
All_Data$STARTAGE[All_Data$STARTAGE < 18] <- NA

# Remove unnecessary columns
Columns <- c("UNIQUEID",
             "AFFINITYGROUP",
             "PRODUCTCODE",
             "POLICYNUMBER",
             "POLICYHOLDERIDNUMBER",
             "COMMENCEMENTDATEOFPOLICY",
             "STATUS",
             "STATUSEFFECTIVEENDDATE",
             "AGENTNAME",
             "GENDER",
             "ACTIVEPREMIUM",
             "TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH")

All_Data <- subset(All_Data, select = Columns)

# Clean Premiums paid
All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH <- gsub(" ","",All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH)
All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH <- as.numeric(All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH)
All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH[is.na(All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH)] = 0

# Clean status
All_Data$STATUS[All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH == 0] <- "NTU"
All_Data$STATUS[All_Data$TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH != 0 & All_Data$STATUS == "NTU"] <- "LAP"


# Clean effective end date
All_Data$STATUSEFFECTIVEENDDATE <- DateConv(All_Data$STATUSEFFECTIVEENDDATE)


# Clean affinity group code
All_Data$AFFINITYGROUP <- as.character(All_Data$AFFINITYGROUP)
string <- All_Data$AFFINITYGROUP[str_sub(All_Data$AFFINITYGROUP, start = -1) == " "]
string <- substr(string, 1, nchar(string) - 1)
All_Data$AFFINITYGROUP[str_sub(All_Data$AFFINITYGROUP, start = -1) == " "] <- string

# New End Date
All_Data$NEW_END_DATE <- All_Data$STATUSEFFECTIVEENDDATE
All_Data$NEW_END_DATE[All_Data$STATUS == "ACT"] <- as.Date(today(tzone = ""))
All_Data$NEW_END_DATE[All_Data$STATUS == "NTU"] <- All_Data$COMMENCEMENTDATE[All_Data$STATUS == "NTU"] 
date_errors <- All_Data$NEW_END_DATE[All_Data$STATUS == "LAP" & day(All_Data$NEW_END_DATE) == 1]
date_errors[month(date_errors) == 12] <- date_errors[month(date_errors) == 12] + 30 
date_errors[month(date_errors) != 12] <- as.Date(paste0(year(date_errors[month(date_errors) != 12]),"-", 
                                                        month(date_errors[month(date_errors) != 12]) + 1,"-", 
                                                        day(date_errors[month(date_errors) != 12]))) - 1
All_Data$NEW_END_DATE[All_Data$STATUS == "LAP" & day(All_Data$NEW_END_DATE) == 1] <- date_errors

# 
# # Premium data
# All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2010 <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2010)) 
# All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2010[is.na(All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2010)] = 0
# All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2011 <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2011)) 
# All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2011[is.na(All_Data$OLDMONTHLYPREMIUMAMOUNTTOENDJAN2011)] = 0
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2012 <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMTOENDJAN2012))
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2012[is.na(All_Data$OLDMONTHLYPREMIUMTOENDJAN2012)] = 0
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2013 <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMTOENDJAN2013))
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2013[is.na(All_Data$OLDMONTHLYPREMIUMTOENDJAN2013)] = 0
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2014 <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMTOENDJAN2014))
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2014[is.na(All_Data$OLDMONTHLYPREMIUMTOENDJAN2014)] = 0
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2015 <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMTOENDJAN2015))
# All_Data$OLDMONTHLYPREMIUMTOENDJAN2015[is.na(All_Data$OLDMONTHLYPREMIUMTOENDJAN2015)] = 0
# All_Data$OLDMONTHLYPREMIUMTOENDDEC15GAPJAN16MPW <- as.numeric(gsub(" ","",All_Data$OLDMONTHLYPREMIUMTOENDDEC15GAPJAN16MPW))
# All_Data$OLDMONTHLYPREMIUMTOENDDEC15GAPJAN16MPW[is.na(All_Data$OLDMONTHLYPREMIUMTOENDDEC15GAPJAN16MPW)] = 0

# Remove rows of NAs
All_Data <- All_Data[!is.na(All_Data$POLICYNUMBER),]

rm(Columns, gggg)


# Create an original copy of the All_Data

All_Data_Original <- All_Data
