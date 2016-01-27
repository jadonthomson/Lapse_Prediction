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
  if (file.exists(paste(Path, "/Data/CSV/",paste0(substr(lap_File_List[lapfile], 1, 4), ".csv"), sep = ""))) {
    
    fileCSVDate <- file.mtime(paste(Path, "/Data/CSV/", paste0(substr(lap_File_List[lapfile], 1, 4), ".csv"), sep = ""))
    fileXLSDate <- file.mtime(paste(Path, "/Data/", lap_File_List[lapfile], sep = ""))
    
    if (fileXLSDate > fileCSVDate){
      excelToCsv(paste(Path, "/Data/CSV/", lap_File_List[lapfile], sep = ""))
    }
  } else {
    excelToCsv(paste(Path, "/Data/",   lap_File_List[lapfile], sep = ""))
  }
  
  CSV_lap_File_List  <-  list.files(paste(Path, "/Data/CSV/", sep = ""))
  
  lap_Data <- fread(paste(Path, "/Data/CSV/",paste0(substr(lap_File_List[lapfile], 1, nchar(lap_File_List[lapfile]) - 5), ".csv"), sep = ""),
                   colClasses  =  "character",
                   header      =  TRUE)
  lap_Data <- as.data.frame(lap_Data)
  lap_Data <- lap_Data[lap_Data$`Policy Number` != "",]
  
  colnames(lap_Data)  <-  gsub(" ","",gsub("[^[:alnum:] ]", "", toupper(colnames(lap_Data))))
  
  if(lapfile == 1) {
    All_lap_Data <- lap_Data
  } else{
    
    common_cols <- intersect(colnames(All_lap_Data), colnames(lap_Data)) # Combine only the common columns (in case of missmatches)
    uncommon_cols <- c(uncommon_cols, setdiff(colnames(All_lap_Data), colnames(lap_Data)))
    
    All_lap_Data <- rbind(
      subset(All_lap_Data,  select = common_cols), 
      subset(lap_Data,      select = common_cols)
    )
    
  }
  
  print(lap_File_List[lapfile])

All_lap_Data$COMMENCEMENTDATEOFPOLICY <- DateConv(All_lap_Data$COMMENCEMENTDATEOFPOLICY)
All_lap_Data$STATUSEFFECTIVEENDDATE <- DateConv(All_lap_Data$STATUSEFFECTIVEENDDATE)

} 

# Remove data from workspace (to save memory and time)
rm(lap_Data, lapfile, num_lap_file, lap_File_List, common_cols) 




###################### Calculate the year-to-year increase in QUOTED premiums
for (i in 2:10){
  All_lap_Data$Increase <- 0
  ### NTU -> NA
  All_lap_Data[All_lap_Data$STATUS == "NTU", "Increase"] <- NA
  ### duration < 12   -> NA
  All_lap_Data$Increase[(All_lap_Data$STATUSEFFECTIVEENDDATE - All_lap_Data$COMMENCEMENTDATEOFPOLICY)/365.25 <= 1] <- NA
  ### increase = (year 2)/(year 1)
  All_lap_Data$Increase[!is.na(All_lap_Data$Increase)] <- (All_lap_Data[!is.na(All_lap_Data$Increase),paste0("YEAR", as.character(i), "QUOTEDTOTALPREMIUM")])/(All_lap_Data[!is.na(All_lap_Data$Increase),paste0("YEAR", as.character(i - 1), "QUOTEDTOTALPREMIUM")])
  colnames(All_lap_Data)[which(names(All_lap_Data) == "Increase")] <- paste0("Increase",i)
}

                        
                       
                        



