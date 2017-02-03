# Installer

require(openxlsx)
require(excel.link)
require(lubridate)
require(stringr)



elapsed_months <- function(end_date, start_date) {
  ed <- as.POSIXlt(end_date)
  sd <- as.POSIXlt(start_date)
  12 * (ed$year - sd$year) + (ed$mon - sd$mon)
}



`%not in%` <- function (x, table) is.na(match(x, table, nomatch=NA_integer_))



DateConv <- function(Cont){
  
  #############################################################################################
  
  Cont <- as.character(toupper(Cont))
  
  numCont <- floor(as.numeric(Cont))
  numCont[is.na(numCont)] <- Cont[is.na(numCont)]
  
  Cont <- numCont
  
  Cont <- gsub(" ", "",   Cont)
  Cont <- gsub("/", "-",  Cont)
  Cont <- gsub("\\", "-", Cont, fixed = TRUE)
  Cont <- gsub(".", "-",  Cont, fixed = TRUE)
  
  #############################################################################################
  
  ContDf <- strsplit(Cont, "-")
  
  n <- max(sapply(ContDf, length))
  l <- lapply(ContDf, function(X) c(X, rep(NA, n - length(X))))
  
  ContDf <- data.frame(t(do.call(cbind, l)), stringsAsFactors = FALSE)
  colnames(ContDf) <- paste("Date", seq(1:ncol(ContDf)), sep = "")
  
  ContDf$Date1[is.na(ContDf$Date1)] <- as.Date(0, origin = "1899-12-30")
  
  Months <- c("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
  
  for (i in 1:ncol(ContDf)) {
    OnNames <- which(ContDf[, i] %in% Months)
    ContDf[, i][OnNames] <- which(Months %in% ContDf[, i][OnNames])
  }
  
  ContDf$OUT <- rep(as.Date(0, origin = "1899-12-30"), length(Cont))
  
  if (ncol(ContDf) == 2) {
    
    #############################################################################################
    # Date is of the numeric format -> 42095 
    # NB - five digits, 
    #       anything more than won't be a valid date for the next 150 years ( the year 2173)
    #       anything less than 5 digits is before 1928
    #       therefore a five digit value is a safe number to indicate a numeric year format
    
    onepartonly <- rep(1, nrow(ContDf))
    
    if (sum(onepartonly) > 0) {
      onepartonlyfinal                   <- NA
      onepartonlyfinal[onepartonly == 1] <- ContDf$Date1[onepartonly == 1]
      
      onepartonlyfinal[nchar(onepartonlyfinal) != 5] <- ""
      
      onepartonlyfinal <-  as.Date(as.numeric(onepartonlyfinal), origin = "1899-12-30")
      
      ContDf$OUT[!is.na(onepartonlyfinal)] <- as.Date(onepartonlyfinal[!is.na(onepartonlyfinal)]) 
    }
    
  } else {
    #############################################################################################
    # Year is left -> yyyy-dd-mm or yyyy-mm-dd
    
    yyPrt                    <- as.numeric(ContDf$Date1)
    yyPrt[nchar(yyPrt) != 4] <- ""
    yyPrt[is.na(yyPrt)]      <- ""
    
    if (sum(yyPrt != "") > 0){
      mmPrt              <- as.numeric(ContDf$Date2)
      mmPrt[yyPrt == ""] <- NA
      ddPrt              <- as.numeric(substr(ContDf$Date3, 1, 2))
      ddPrt[yyPrt == ""] <- NA
      
      yyPrt[ddPrt == 0 | mmPrt == 0] <- 1899
      mmPrt[ddPrt == 0 | mmPrt == 0] <- 12
      ddPrt[ddPrt == 0 | mmPrt == 0] <- 30
      
      YrDat                <- paste(yyPrt, mmPrt, ddPrt, sep = "-")
      YrDat                <- gsub("NA", "", YrDat)
      YrDat[YrDat == "--"] <- NA
      
      mmPrt_ch                   <- ifelse(mmPrt > 12, 1, 0)
      mmPrt_ch[is.na(mmPrt_ch)]  <- FALSE
      
      YrDat[mmPrt_ch == 1] <- paste(yyPrt, ddPrt, mmPrt, sep = "-")[mmPrt_ch == 1]
      
      ContDf$OUT[!(is.na(YrDat))] <- as.Date(YrDat[!(is.na(YrDat))])
    }
    
    #############################################################################################
    # Year is right -> dd-mm-yyyy or mm-dd-yyyy
    
    yyPrt                    <- as.numeric(ContDf$Date3)
    yyPrt[nchar(yyPrt) != 4] <- ""
    yyPrt[is.na(yyPrt)]      <- ""
    
    if (sum(yyPrt != "") > 0){
      ddPrt              <- as.numeric(ContDf$Date1)
      ddPrt[yyPrt == ""] <- NA
      mmPrt              <- as.numeric(ContDf$Date2)
      mmPrt[yyPrt == ""] <- NA
      
      yyPrt[ddPrt == 0 | mmPrt == 0] <- 1899
      mmPrt[ddPrt == 0 | mmPrt == 0] <- 12
      ddPrt[ddPrt == 0 | mmPrt == 0] <- 30
      
      YrDat                <- paste(yyPrt, mmPrt, ddPrt, sep = "-")
      YrDat                <- gsub("NA", "", YrDat)
      YrDat[YrDat == "--"] <- NA
      
      mmPrt_ch                   <- ifelse(mmPrt > 12, 1, 0)
      mmPrt_ch[is.na(mmPrt_ch)]  <- FALSE
      
      YrDat[mmPrt_ch == 1] <- paste(yyPrt, ddPrt, mmPrt, sep = "-")[mmPrt_ch == 1]
      
      ContDf$OUT[!is.na(YrDat)] <- as.Date(YrDat[!is.na(YrDat)])
    }
    
    #############################################################################################
    # Date is of the numeric format -> 42095 
    # NB - five digits, 
    #       anything more than won't be a valid date for the next 150 years ( the year 2173)
    #       anything less than 5 digits is before 1928
    #       therefore a five digit value is a safe number to indicate a numeric year format
    
    onepartonly <- rowSums(is.na(ContDf))
    onepartonly[onepartonly != 0] <- onepartonly[onepartonly != 0] - 1 # Remove 1 for the added "OUT" column, now the entries with a 1 is numonly
    
    if (sum(onepartonly) > 0) {
      onepartonlyfinal                   <- NA
      onepartonlyfinal[onepartonly == 1] <- ContDf$Date1[onepartonly == 1]
      
      onepartonlyfinal[nchar(onepartonlyfinal) != 5] <- ""
      
      onepartonlyfinal <-  as.Date(as.numeric(onepartonlyfinal), origin = "1899-12-30")
      
      ContDf$OUT[!is.na(onepartonlyfinal)] <- as.Date(onepartonlyfinal[!is.na(onepartonlyfinal)]) 
    }
    
  }
  
  #############################################################################################
  # Date format could also be 20050109 or 01092005, but this could also include things like:
  # 20050109 -> 050109 or 50109
  # 01092005 -> 010905 or 10905
  # these will barely makse sense even if you try and interpret them manually
  
  ContDf$OUT[ContDf$OUT <= as.Date("1900/01/01")] <- NA
  
  return(ContDf$OUT)
  
}