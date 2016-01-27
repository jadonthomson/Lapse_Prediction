# Import, append and sort data

#####################################################################################################
# Setup #
#########

Path <- getwd()
source(paste(Path, "/Lapse_Prediction/Initialize.R", sep = ""))
uncommon_cols <- c()

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
    uncommon_cols <- c(uncommon_cols, setdiff(colnames(All_lap_Data), colnames(lap_Data)))
    uncommon_cols <- c(uncommon_cols, setdiff(colnames(lap_Data), colnames(All_lap_Data)))
    
    All_lap_Data <- rbind(
      subset(All_lap_Data,  select = common_cols), 
      subset(lap_Data,      select = common_cols)
    )
    
  }
  
  print(lap_File_List[lapfile])

} 

fileXLSDate <- file.mtime(paste(Path, "/Data/", lap_File_List[lapfile], sep = ""))
All_lap_Data$COMMENCEMENTDATEOFPOLICY <- DateConv(All_lap_Data$COMMENCEMENTDATEOFPOLICY)
All_lap_Data$STATUSEFFECTIVEENDDATE <- DateConv(All_lap_Data$STATUSEFFECTIVEENDDATE)
All_lap_Data$DURATION <- as.numeric(((as.Date(substr(fileXLSDate,1,10)) - All_lap_Data$COMMENCEMENTDATEOFPOLICY)/365.25)*12)
All_lap_Data$DURATION[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)] <- as.numeric(((All_lap_Data$STATUSEFFECTIVEENDDATE[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)] - All_lap_Data$COMMENCEMENTDATEOFPOLICY[!is.na(All_lap_Data$STATUSEFFECTIVEENDDATE)])/365.25)*12)


# Remove data from workspace (to save memory and time)
rm(lap_Data, lapfile, num_lap_file, lap_File_List, common_cols,file_name,CSV_file_name) 




###################### Calculate the year-to-year increase in QUOTED premiums

All_lap_Data$Increase       <-  1
All_lap_Data$CURRENTPREMIUM <-  All_lap_Data$YEAR1QUOTEDTOTALPREMIUM
All_lap_Data$LASTREMIUM     <-  All_lap_Data$YEAR1QUOTEDTOTALPREMIUM
year_len <- ceiling(max(All_lap_Data$DURATION)/12)



for (i in 2:year_len){
     # if duration is less than or equal to i but greater than i - 1
    All_lap_Data$CURRENTPREMIUM[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1)] <-  All_lap_Data[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1),paste0("YEAR", i, "QUOTEDTOTALPREMIUM")]
              # TOOK data$year i premium as currentpremium
    All_lap_Data$LASTREMIUM[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1)]     <-  All_lap_Data[All_lap_Data$DURATION <= 12*i & All_lap_Data$DURATION >= 12*(i - 1),paste0("YEAR", i - 1, "QUOTEDTOTALPREMIUM")] 
              # TOOK data$year i - 1 premium as lastpremium
}

All_lap_Data$INCREASE <- as.numeric(All_lap_Data$CURRENTPREMIUM)/as.numeric(All_lap_Data$LASTREMIUM)
All_lap_Data$INCREASE[All_lap_Data$INCREASE == 1] <- NA
All_lap_Data$INCREASE[All_lap_Data$STATUS == "NTU"] <- NA

############ Remove columns

All_lap_Data <- subset(All_lap_Data, select = C("POLICYNUMBER",
                                                "POLICYHOLDERFIRSTNAME",
                                                "POLICYHOLDERSURNAME",
                                                "ACTIVEQUOTEDPREMIUM",
                                                "INITIALLIFEINSURED",
                                                "FIRSTNAMELIFEINSURED",
                                                "SURNAMELIFEINSURED",
                                                "DOBOFLIFEINSURED",
                                                "IDNUMBEROFLIFEINSURED",
                                                "REINSURER",
                                                "COMMISSIONEXCLVAT",
                                                "ZLADMINFEEINCLVAT",
                                                "ZLBINDERFEEINCLVAT",
                                                "YEAR1QUOTEDDEATHPREMIUM",
                                                "YEAR1QUOTEDDISABILITYPREMIUM",
                                                "YEAR1QUOTEDTEMPORARYDISABILITYPREMIUM",
                                                "YEAR1QUOTEDCRITICALILLNESSPREMIUM",
                                                "QUOTEDRETRENCHMENTPREMIUMY1Y10",
                                                "QUOTEDRAFPREMIUMY1Y10", 
                                                "QUOTEDROADCOVERPREMIUMY1Y10", 
                                                "QUOTEDFUTUREVAPPREMIUMY1Y10"))








########### Clean postal codes
# we want province and postal codes
# look at all of the three postal columns and see if you can find one of the 9 provinces
# leave postal codes.

# strip payment day into just numbers

# from voice log extract day and month.
                     
                      