# Import, append and sort data

#####################################################################################################
# Setup #
#########

Path <- getwd()
source(paste(Path, "/Lapse_Prediction/Initialize.R", sep = ""))

#############
# Load Data #
#############






####################### Convert excel to CSV if they have been updated ############################

# Start
lap_File_List      <-  list.files(paste(Path, "/Data", sep = ""))
num_lap_file       <-  length(lap_File_List) - 1    #  Number of files in folder - 2 due to folders

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
  
  #if(lapfile == 1) {
  #  All_lap_Data <- lap_Data
  #} else{
  #  
  #  common_cols <- intersect(colnames(All_lap_Data), colnames(lap_Data)) # Combine only the common columns (in case of missmatches)
  #  
  #  All_lap_Data <- rbind(
  #    subset(All_lap_Data,  select = common_cols), 
  #    subset(lap_Data,      select = common_cols)
  #  )
    
  }
  
  print(lap_File_List[lapfile])
  
}

# Remove data from workspace (to save memory and time)
rm(cl_Data, clfile, num_cl_file, cl_File_List, common_cols, fileCSVDate, fileXLSDate) 






####################### Import each of the three CSVs

                        ############ Make sure no formulas are copied
                        
                        ############ Headings might be a problem

                        ############ Sub the column headings using the gsub functions

  

####################### Append CSVs by column headers #########################




###################### Calculate the year-to-year increase in QUOTED premiums
                        ### NTU -> NA
                        ### duration < 12   -> NA
                        ### increase = (year 2)/(year 1)



# Load Correct Claims File Here
Dat <- read.xlsx(paste(Path,"/Data/Lapse_Data.xlsx", sep = ""), 
                 sheet = "Lapse_Data", 
                 detectDates = TRUE, 
                 colNames = TRUE)
