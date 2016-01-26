# Import, append and sort data

#####################################################################################################
# Setup #
#########

Path <- getwd()
source(paste(Path, "/Initialize.R", sep = ""))

#############
# Load Data #
#############






####################### Convert excel to CSV if they have been updated ############################







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
