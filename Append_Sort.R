## Import, append and sort lapse data
## Jadon
## Jan 2016

##########
# Setup #
#########


###### Make sure the working directory is right. #######
Path <- substr(getwd(),1, nchar(getwd()) - 17)
source(paste(Path, "/Lapse_Prediction/Initialize.R", sep = ""))
#uncommon_cols <- c()




#############
# Load Data #
#############


####################### Convert excel to CSV if they have been updated ############################

# Start
lap_File_List      <-  list.files(paste(Path, "/Data", sep = ""))
lap_File_List      <-  lap_File_List[lap_File_List != "CSV"]
num_lap_file       <-  length(lap_File_List)     #  Number of files in folder - 2 due to folders

for(lapfile in 1:num_lap_file){
  
  # Check to see if CSV should be created
  file_name <- lap_File_List[lapfile]
  CSV_file_name <- paste0(substr(file_name, 1, nchar(file_name) - 5), ".csv") # remove everything after the . rather than just the last 5 characters
  
  if (file.exists(paste(Path, "/Data/CSV/", CSV_file_name, sep=""))) {
    
    fileCSVDate <- file.mtime(paste(Path, "/Data/CSV/", CSV_file_name, sep = ""))
    fileXLSDate <- file.mtime(paste(Path, "/Data/", file_name, sep = ""))
    
    if (fileXLSDate > fileCSVDate){
      excelToCsv(paste(Path, "/Data/CSV/", file_name, sep = ""))
    }
  } else {
    excelToCsv(paste(Path, "/Data/",   file_name, sep = ""))
  }
  
  CSV_lap_File_List  <-  list.files(paste(Path, "/Data/CSV/", sep = ""))
  
  lap_Data <- fread(paste(Path, "/Data/CSV/",CSV_file_name, sep = ""),
                   colClasses  =  "character",
                   header      =  TRUE)
  lap_Data <- as.data.frame(lap_Data)
  
  colnames(lap_Data)  <-  gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(colnames(lap_Data))))
  
  lap_Data <- lap_Data[lap_Data$POLICYNUMBER != "",]

  if(lapfile == 1) {
    All_lap_Data <- lap_Data

  } else{
    
    common_cols <- intersect(colnames(All_lap_Data), colnames(lap_Data)) # Combine only the common columns (in case of missmatches)
    #uncommon_cols <- c(uncommon_cols, setdiff(colnames(All_lap_Data), colnames(lap_Data)))
    #uncommon_cols <- c(uncommon_cols, setdiff(colnames(lap_Data), colnames(All_lap_Data)))
    
    All_lap_Data <- rbind(
      subset(All_lap_Data,  select = common_cols), 
      subset(lap_Data,      select = common_cols)
    )
    
  }
  
  print(lap_File_List[lapfile])

} 

fileXLSDate                            <- file.mtime(paste(Path, "/Data/", lap_File_List[lapfile], sep = ""))
All_lap_Data$COMMENCEMENTDATEOFPOLICY  <- DateConv(All_lap_Data$COMMENCEMENTDATEOFPOLICY)
All_lap_Data$STATUSEFFECTIVEENDDATE    <- DateConv(All_lap_Data$STATUSEFFECTIVEENDDATE)
All_lap_Data$POSTALADDRESS1            <- gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(All_lap_Data$POSTALADDRESS1)))
All_lap_Data$POSTALADDRESS2            <- gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(All_lap_Data$POSTALADDRESS2)))
All_lap_Data$POSTALADDRESS3            <- gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(All_lap_Data$POSTALADDRESS3)))
All_lap_Data$DURATION                  <- as.numeric(((as.Date(substr(fileXLSDate,1,10)) - All_lap_Data$COMMENCEMENTDATEOFPOLICY)/365.25)*12)
All_lap_Data$DURATION[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)] <- as.numeric(((All_lap_Data$STATUSEFFECTIVEENDDATE[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)] - All_lap_Data$COMMENCEMENTDATEOFPOLICY[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)])/365.25)*12)


# Remove data from workspace (to save memory and time)
rm(lap_Data, lapfile, num_lap_file, lap_File_List, common_cols, file_name, CSV_file_name, CSV_lap_File_List, fileCSVDate, fileXLSDate) 

# Sort out status column
# if there is end date = lapse
All_lap_Data$STATUS[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)] <- "LAP"    # could make this 1 and act = 0
# if no end date = active             ########## do these changes at the beginning of the code.
All_lap_Data$STATUS[is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)] <- "ACT"


###################### Calculate the year-to-year increase in QUOTED premiums

All_lap_Data$Increase        <-  1
All_lap_Data$CURRENTPREMIUM  <-  All_lap_Data$YEAR1QUOTEDTOTALPREMIUM
All_lap_Data$LASTPREMIUM     <-  All_lap_Data$YEAR1QUOTEDTOTALPREMIUM
year_len                     <-  ceiling(max(All_lap_Data$DURATION)/12)



for (i in 2:year_len){
    # if duration is less than or equal to i but greater than i - 1
    All_lap_Data$CURRENTPREMIUM[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1)] <-  All_lap_Data[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1),paste0("YEAR", i, "QUOTEDTOTALPREMIUM")]
    # TOOK data$year i premium as currentpremium
    All_lap_Data$LASTPREMIUM[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1)]    <-  All_lap_Data[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1),paste0("YEAR", i - 1, "QUOTEDTOTALPREMIUM")] 
    # TOOK data$year i - 1 premium as lastpremium
}

All_lap_Data$INCREASE <- as.numeric(All_lap_Data$CURRENTPREMIUM)/as.numeric(All_lap_Data$LASTREMIUM)
All_lap_Data$INCREASE[All_lap_Data$INCREASE == 1]    <- NA
All_lap_Data$INCREASE[All_lap_Data$STATUS == "NTU"]  <- NA

rm(year_len)

########### Clean postal codes
# we want province and postal codes
Provinces <- c("WESTERNCAPE",
               "LIMPOPO",
               "MPUMALANGA",
               "GAUTENG",
               "EASTERNCAPE",
               "KWAZULUNATAL",
               "NORTHERNCAPE",
               "FREESTATE",
               "NORTHWEST",
               "WESTERNPROVINCE",
               "CAPETOWN",
               "PRETORIA",
               "EASTLONDON",
               "SOWETO",
               "RONDEBOSCH",
               "DURBAN",
               "JOHANNESBURG",
               "PORTELIZABETH",
               "PIETERMARITZBURG",
               "BENONI")

# look at all of the three postal columns and see if you can find one of the 9 provinces
All_lap_Data$PROVINCE <- ""
All_lap_Data$PROVINCE[All_lap_Data$POSTALADDRESS3 %in% Provinces]  <- All_lap_Data$POSTALADDRESS3[All_lap_Data$POSTALADDRESS3 %in% Provinces]
All_lap_Data$PROVINCE[All_lap_Data$POSTALADDRESS2 %in% Provinces]  <- All_lap_Data$POSTALADDRESS2[All_lap_Data$POSTALADDRESS2 %in% Provinces]
All_lap_Data$PROVINCE[All_lap_Data$POSTALADDRESS1 %in% Provinces]  <- All_lap_Data$POSTALADDRESS1[All_lap_Data$POSTALADDRESS1 %in% Provinces]
rm(Provinces)

# Replace cities with provinces:
All_lap_Data$PROVINCE[All_lap_Data$PROVINCE %in% c("CAPETOWN",
                                                   "RONDEBOSCH",
                                                   "WESTERNPROVINCE")]  <- "WESTERNCAPE"

All_lap_Data$PROVINCE[All_lap_Data$PROVINCE %in% c("PRETORIA",
                                                   "SOWETO",
                                                   "JOHANNESBURG",
                                                   "BENONI")]           <- "GAUTENG"

All_lap_Data$PROVINCE[All_lap_Data$PROVINCE %in% c("PIETERMARITZBURG",
                                                   "DURBAN")]           <- "KWAZULUNATAL"

All_lap_Data$PROVINCE[All_lap_Data$PROVINCE %in% c("PORTELIZABETH",
                                                   "EASTLONDON")]       <- "EASTERNCAPE"


# strip payment day into just numbers
######################### All_lap_Data$ ############## which payment date????????// lkjlk;j ########J
  
All_lap_Data$VOICELOGGED <- DateConv(All_lap_Data$VOICELOGGED)

# from voice log extract day and month.
All_lap_Data$VOICELOGGEDDAY <- format(All_lap_Data$VOICELOGGED, format = "%d")
All_lap_Data$VOICELOGGEDMONTH <- format(All_lap_Data$VOICELOGGED, format = "%m")

# Clean smoking columns
All_lap_Data$SMOKERNONSMOKER <- gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(All_lap_Data$SMOKERNONSMOKER)))
All_lap_Data$SMOKERNONSMOKER[All_lap_Data$SMOKERNONSMOKER == "YES"]   <- "SMOKER"
All_lap_Data$SMOKERNONSMOKER[All_lap_Data$SMOKERNONSMOKER == "NO"]    <- "NON-SMOKER"

# Clean More than 30 cigs column. anything that is a "no" becomes a "no".
colnames(All_lap_Data)[which(names(All_lap_Data) == "DOYOUSMOKEMORETHAN30CIGARETTESPERDAY")] <- "MORETHAN30"
All_lap_Data$MORETHAN30 <- gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(All_lap_Data$MORETHAN30)))
All_lap_Data$MORETHAN30[All_lap_Data$MORETHAN30 %in% c("LESSTHAN30ADAY",
                                                       "LESSTHAN30PERDAY",
                                                       "NON-SMOKER",
                                                       "NOTAPPLICABLE")]       <-  "NO"

All_lap_Data$MORETHAN30[All_lap_Data$MORETHAN30 %in% c("MORETHAN30ADAY",
                                                       "YES",
                                                       "MORETHAN30PERDAY")]    <-  "YES"

# checking if height is non-numeric or if it is less than 100cm. if the height is less than 100 then we add a 100.
All_lap_Data$HEIGHTINCM <- as.numeric(All_lap_Data$HEIGHTINCM)
All_lap_Data$HEIGHTINCM[!as.numeric(All_lap_Data$HEIGHTINCM)] <- mean(All_lap_Data$HEIGHTINCM[!is.na(All_lap_Data$HEIGHTINCM)])
All_lap_Data$HEIGHTINCM[All_lap_Data$HEIGHTINCM < 100 & !is.na(All_lap_Data$HEIGHTINCM)] <- All_lap_Data$HEIGHTINCM[All_lap_Data$HEIGHTINCM < 100 & !is.na(All_lap_Data$HEIGHTINCM)] + 100 

# do the PPB indicator (1 if PPB, 0 if not)
All_lap_Data$PPB <- gsub(" ", "", All_lap_Data$PPBTOCELLCAPTIVE)
All_lap_Data$PPB[All_lap_Data$PPB != ""]  <- 1
All_lap_Data$PPB[All_lap_Data$PPB == ""]  <- 0



# choose the columns you want to keep

Names <- c("AFFINITYGROUP",
           "CURRENTPREMIUM",
           "LASTPREMIUM",
           "INCREASE",
           "EMPLOYEEBATCHNUMBER",
           "TITLEOFPOLICYHOLDER",
           "POLICYHOLDERSURNAME",
           "PROVINCE",
           "POSTALCODE",
           "EMAILADDRESS",
           "PHONEW",
           "PREMIUMPAYERBANK",
           "PREMIUMPAYERBRANCHCODE",
           "ACCOUNTTYPECODE",
           "PREMIUMPAYERDEBITORDERDAY",
           "COMMENCEMENTDATEOFPOLICY",
           "STATUS",
           "DURATION",
           "STATUSEFFECTIVEENDDATE",
           "AGENTNAME",
           "VOICELOGGEDDAY",
           "VOICELOGGEDMONTH",
           "IDNUMBEROFLIFEINSURED",
           "DOBOFLIFEINSURED",
           "GENDER", 
           "LEVELOFINCOME",
           "EDUCATION",
           "SMOKERNONSMOKER",
           "MORETHAN30",
           "HEIGHTINCM",
           "WEIGHT130KGSORWEIGHTOLDPOLICIES",
           "CREDITPROTECTIONDEATHSUMASSURED",
           "CREDITPROTECTIONDISABILITYSUMASSURED",
           "CREDITPROTECTIONTEMPORARYDISABILITYMONTHLYBENEFIT",
           "TEMPORARYDISABILITYPERIOD",
           "RETRENCHMENTBENEFIT",
           "RETRENCHMENTPERIOD",
           "RAFBENEFIT",
           "RAFNIGHTSINHOSPITAL",
           "RAFTYPEOFCOVER",
           "CRITICALILLNESSCOVER",
           "ROADCOVERINCLUDED",
           "CREDITPROVIDER1",
           "CREDITPROVIDER2", 
           "CREDITPROVIDER3",
           "CREDITPROVIDER4",
           "CREDITPROVIDER5",
           "CREDITPROVIDER6",
           "CREDITPROVIDER7",
           "RELATIONSHIPOFBENEFICIARY1",
           "RELATIONSHIPOFBENEFICIARY2",
           "RELATIONSHIPOFBENEFICIARY3",
           "RELATIONSHIPOFBENEFICIARY4",
           "RELATIONSHIPOFBENEFICIARY5",
           "RELATIONSHIPOFBENEFICIARY6",
           "RELATIONSHIPOFBENEFICIARY7",
           "VOICELOGGEDENDORSEMENT",
           "TOTALAMOUNTPAIDSINCEINCEPTIONTOCURRENTMONTH",
           "PPB")

# Remove the unwanted data
All_lap_Data <- subset(All_lap_Data, select = Names)
rm(Names)


# If weight is outside the interquartile range, bring it to IQR boundry.
# first find the interquartile range
weight <- gsub(" ", "", toupper(All_lap_Data$WEIGHT130KGSORWEIGHTOLDPOLICIES))
weight[weight == "NOTANSWERED" | weight == ""]        <-  NA
weight[weight == "YES"]                               <-  130
mean                                                  <-  mean(as.numeric(weight), na.rm = T)
IQR                                                   <-  IQR(weight, na.rm = T)
weight[weight == "NO"]                                <-  mean
weight                                                <-  as.numeric(weight)
weight[weight > mean + 2*IQR]                         <-  mean + IQR
weight[weight < mean - 2*IQR]                         <-  mean - IQR
All_lap_Data$WEIGHT130KGSORWEIGHTOLDPOLICIES          <-  weight
rm(mean, IQR, weight)

# We need the number of beneficiaries and the relationship to the policyholder.
# Cleaning all relationship columns
temp     <-  select(All_lap_Data, contains("RELATIONSHIPOFBENEFICIARY"))
temp     <-  toupper(gsub(" ", "", as.matrix(temp)))
All_lap_Data[colnames(All_lap_Data) %in% colnames(temp)] <- temp 

# finding the names of all the columns with "relationshipofbeneficiary" to count the number of beneficiaries
All_lap_Data$NUMBEROFBENEFICIARIES <- 0
temp[temp != ""] <- 1
temp[temp == ""] <- 0
temp <- apply(temp, 2, as.numeric)
All_lap_Data$NUMBEROFBENEFICIARIES <- rowSums(temp)

# Number of credit providers and cleaning provider names.
temp     <-  select(All_lap_Data, contains("CREDITPROVIDER"))
temp     <-  toupper(gsub(" ", "", as.matrix(temp)))
All_lap_Data[colnames(All_lap_Data) %in% colnames(temp)] <- temp

# count the number of credit providers
All_lap_Data$NOCREDITPROVIDERS          <-  0
temp[temp != ""]                        <-  1
temp[temp == ""]                        <-  0
temp                                    <-  apply(temp, 2, as.numeric)
All_lap_Data$NOCREDITPROVIDERS          <-  rowSums(temp)
rm(temp)


# sort out premium debit day column (take away th and rd....)
All_lap_Data$PREMIUMPAYERDEBITORDERDAY  <-  gsub("last day", 31, All_lap_Data$PREMIUMPAYERDEBITORDERDAY)
All_lap_Data$PREMIUMPAYERDEBITORDERDAY  <-  gsub("st", "", All_lap_Data$PREMIUMPAYERDEBITORDERDAY)
All_lap_Data$PREMIUMPAYERDEBITORDERDAY  <-  gsub("nd", "", All_lap_Data$PREMIUMPAYERDEBITORDERDAY)
All_lap_Data$PREMIUMPAYERDEBITORDERDAY  <-  gsub("th", "", All_lap_Data$PREMIUMPAYERDEBITORDERDAY)

                
