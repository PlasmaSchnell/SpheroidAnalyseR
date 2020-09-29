################### R CODE INFORMATION   ##################################
#                                                                         #
# filename :        Spheroid_calc_rev2.R                                  #
# Revision :        2                                                     #      
# Revision Date :   17-Nov-2019                                           #
# Author :          Joe Wilkinson                                         #
#                                                                         #
# Project :         Spheroid Analysis                                     #
#                   Rhiannon Barrow,  Dr Lucy Stead                       #
#                   Glioma Genomics                                       #
#                   Leeds Institute of Medical Research                   #
#                   St James's University Hospital, Leeds  LS9 7TF        #
#                                                                         #
###########################################################################

#suppress console messages
sink()
sink("NUL")
# for Mac OS or Linux, please replace above command with: sink("/dev/null")

library(tidyverse)
library(ggthemes)
library(gridExtra)
library(readxl)
library(openxlsx)
library(plotrix)
library(writexl)
library(ggpubr)

# define file separator
slash = "//"

# read configuration file 


suppressMessages(Sph_config_data <- read_xlsx("Sph_Config_File_Rev2.xlsx", sheet = 1, col_names = TRUE,  .name_repair = "universal")) 


rawfilename <- Sph_config_data$Processing.parameters[1]
input_path <- Sph_config_data$Processing.parameters[2]

input_file <- paste(input_path,slash,rawfilename,".xlsx",sep="")


#extract processing parameters : directories, filenammes, limits etc

RobZ_LoLim <- as.numeric(Sph_config_data$Processing.parameters[11])
RobZ_UpLim <- as.numeric(Sph_config_data$Processing.parameters[12])

# extract backup CSV filename for merging processed data files later.

tempfilename <- Sph_config_data$Processing.parameters[7] 
temp_output_path <- Sph_config_data$Processing.parameters[8]

merge_output_file<- paste(temp_output_path,slash,tempfilename, ".csv",sep="")

reportfilename <- Sph_config_data$Processing.parameters[4] 
report_output_path <- Sph_config_data$Processing.parameters[5]

report_output_file <- paste(report_output_path,slash,reportfilename, ".xlsx",sep="")

# console screen message

for (j in 1:12)
{
   message("\n") 
}
message("importing files...... ","\n")


# check if report file already exists, if it does, keep original file and save new report file as a duplicate  ie filename(2).xlsx  
# Also provide a message at the end of the code execution to this effect

TF_checkfile <- "FALSE"
TF_checkCSVfile <- "FALSE"

if (file.exists(report_output_file))
 { 
   
   TF_checkfile <- 'TRUE'
   reportfilename2 <- paste(reportfilename, "(2)", sep="")
   report_output_file <- paste(report_output_path,slash,reportfilename2, ".xlsx",sep="") 
   
   Msg_file_duplicated_1 <- paste("Report file already exists - saving a duplicate : ", reportfilename2, ".xlsx",sep="") 
   Msg_file_duplicated_2 <-'Enter a new report filename in the config file and process again, or rename the duplicate file.'
}
 
# check if Export file already exists, if it does, keep original file and save new Export file as a duplicate  ie filename(2).xlsx  
# Also provide a message at the end of the code execution to this effect

TF_checkCSVfile <- 'FALSE'

if (file.exists(merge_output_file))
{ 
   TF_checkCSVfile <- 'TRUE'
   tempfilename2 <- paste(tempfilename, "(2)", sep="")
   merge_output_file<- paste(temp_output_path,slash,tempfilename2, ".csv",sep="")
  
   
   Msg_CSVfile_duplicated_1 <- paste("Output data CSV file already exists - saving a duplicate : ", tempfilename2  , ".CSV",sep="") 
   Msg_CSVfile_duplicated_2 <-'Enter a new Output data CSV filename in the config file and process again, or rename the duplicate file.'
}


platesetupname <- Sph_config_data$Processing.parameters[10]
TF_apply_thresholds <- Sph_config_data$Processing.parameters[14]
TF_outlier_override <- Sph_config_data$Processing.parameters[15]
TF_copytomergedir <- Sph_config_data$Processing.parameters[16]


##############################################################


############################################################
##  Read Plate Setup data from Spheroid configuration XLSX file
##############################################################
 

suppressMessages(Plate_Setup <- read_xlsx("Sph_Config_File_Rev2.xlsx", sheet = platesetupname, .name_repair = "universal"))
 
# subset treatment index dataset

P_setup <-   select(Plate_Setup, plate_row, p_col_1, p_col_2, p_col_3, p_col_4, p_col_5, p_col_6, p_col_7, p_col_8, p_col_9, p_col_10, p_col_11, p_col_12)
P_treatment <- select(Plate_Setup, Treatment_Index, Treatment.Label, Time_Date, Cell_line, Passage_No, Radiation_dosage, Drug_1, Conc_1, Drug_2, Conc_2, Drug_3, Conc_3)
 
# select relevant rows

P_setup <- P_setup[1:8,]
P_treatment <- P_treatment[1:32,]
 
# merge setup and treatment dataframes into a lookup table with a treatment index for each cell 

lookup <-  data.frame("Row"= 1:96, "Col" = 1:96,"T_I" =0)

#count number of treatments in the pllate, and check that they match the number of rows in the datase
Ccount <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

Ccount[1] <- length(which(P_setup$p_col_1 >= 1))
Ccount[2] <- length(which(P_setup$p_col_2 >= 1))
Ccount[3] <- length(which(P_setup$p_col_3 >= 1))
Ccount[4] <- length(which(P_setup$p_col_4 >= 1))
Ccount[5] <- length(which(P_setup$p_col_5 >= 1))
Ccount[6] <- length(which(P_setup$p_col_6 >= 1))
Ccount[7] <- length(which(P_setup$p_col_7 >= 1))
Ccount[8] <- length(which(P_setup$p_col_8 >= 1))
Ccount[9] <- length(which(P_setup$p_col_9 >= 1))
Ccount[10] <- length(which(P_setup$p_col_10 >= 1))
Ccount[11] <- length(which(P_setup$p_col_11 >= 1))
Ccount[12] <- length(which(P_setup$p_col_12 >= 1))

checksum <- sum(Ccount)





colnames(P_treatment)[1] <- "T_I"

# insert A:H in row fields, and assign relevant Treatment Index (T_I) 

j=1
i=1
k=1

 for (j in 1:8)
   
  
     {
    for (i in 1:12) 
    
       {
       
       k = ((j-1)*12)+i
       
       lookup$Row[k]<- P_setup$plate_row[j]
       lookup$Col[k] <- i
      if (i==1) { lookup$T_I[k] = P_setup$p_col_1[j]} else {
          if (i==2) { lookup$T_I[k] = P_setup$p_col_2[j]} else{                                          
             if (i==3) { lookup$T_I[k] = P_setup$p_col_3[j]} else {
                if (i==4) { lookup$T_I[k] = P_setup$p_col_4[j]} else {
                   if (i==5) { lookup$T_I[k] = P_setup$p_col_5[j]} else {
                      if (i==6) { lookup$T_I[k] = P_setup$p_col_6[j]} else {
                         if (i==7) { lookup$T_I[k] = P_setup$p_col_7[j]}else {
                         if (i==8) { lookup$T_I[k] = P_setup$p_col_8[j]} else {
                            if (i==9) { lookup$T_I[k] = P_setup$p_col_9[j]} else{                                          
                               if (i==10) { lookup$T_I[k] = P_setup$p_col_10[j]} else {
                                  if (i==11) { lookup$T_I[k] = P_setup$p_col_11[j]} else {
                                     if (i==12) { lookup$T_I[k] = P_setup$p_col_12[j]}}}}}}}}}}}}
          
                      
}
      
 } 


## merge lookup and P_treatment dataframes by T_I.

newdata <- merge(lookup,P_treatment, by="T_I") 

 
 
  
###################################################################  
# Import raw spheroid dataset (read excel file using read_excel function ).  see link below for more info.
# https://blog.rstudio.com/2015/04/15/readxl-0-1-0/  


#   Spheroid_data <- suppressMessages(read.xlsx(input_file, sheet = "JobView",cols=c(1:5, 7:12, 15) .name_repair = "universal"))

#Spheroid_data <- read.xlsx(input_file, sheet = "JobView",cols=c(1:5, 7:12, 15:20, 30:38), detectDates=TRUE)
#Spheroid_data <- read_excel(input_file, "JobView",cols=c(1:5, 7:12, 15:20, 30:38)
Spheroid_data <- suppressMessages(read_excel(input_file, "JobView",col_names=TRUE, .name_repair = "universal"))

datalength <- nrow(Spheroid_data)
checksumOK <- "TRUE"

if (checksum != datalength)  
{checksumOK <- "FALSE"} 

if (checksumOK == "FALSE")
{
   message("#### Warning : plate setup checksum mismatch. ####","\n")
   message("Number of rows of data in raw file (",datalength  ,") does not match the plate setup checksum (",checksum,") on ",platesetupname,"\n")
   message("Output file corruption may occur - please recheck your plate setup","\n")
   message("Processing halted","\n")
   message("\n")
   sink()
   stop()
}

if (checksumOK == "TRUE")

   {
   message("** Plate setup checksum is valid for raw dataset ","\n")
   message("Number of rows of data in raw file (",datalength  ,") matches the checksum (",checksum,") on ",platesetupname,"\n")
}


# extract job information 

Job_Info_data <- select(Spheroid_data, Project.Folder, Project.ID,	Project.Name,	Jobdef.ID, Jobdef.Name, Jobrun.Finish.Time, Jobrun.Folder, Jobrun.ID, Jobrun.Name, Jobrun.UUID)
#Job_Info_data <- Job_Info_data[1:1,]









# extract relevant columns and create new dataset Spheroid_data_1

Spheroid_data_1 <- select(Spheroid_data, Well.Name, Spheroid_Area.TD.Area, Spheroid_Area.TD.Perimeter.Mean, Spheroid_Area.TD.Circularity.Mean, Spheroid_Area.TD.Count, Spheroid_Area.TD.EqDiameter.Mean, Spheroid_Area.TD.VolumeEqSphere.Mean)


# rename columns 
colnames(Spheroid_data_1) <- c("Well.Name", "Area", "Perimeter", "Circularity","Count", "Diameter", "Volume"  )

# replace any zero values with null value to calculate median values properly

Spheroid_data_1$Area <- ifelse(Spheroid_data_1$Area==0, NA, Spheroid_data_1$Area) 
Spheroid_data_1$Perimeter <- ifelse(Spheroid_data_1$Perimeter==0, NA, Spheroid_data_1$Perimeter)
Spheroid_data_1$Circularity <- ifelse(Spheroid_data_1$Circularity==0, NA, Spheroid_data_1$Circularity)
Spheroid_data_1$Diameter <- ifelse(Spheroid_data_1$Diameter==0, NA, Spheroid_data_1$Diameter)
Spheroid_data_1$Volume <- ifelse(Spheroid_data_1$Volume==0, NA, Spheroid_data_1$Volume)

Spheroid_data_1$Area<- as.numeric(Spheroid_data_1$Area)

## extract min/max thresholds from config file
# Apply Thresholds from pre-screen tab


if (TF_apply_thresholds == "TRUE") 

   {
prescreen <- suppressMessages(read_xlsx("Sph_Config_File_Rev2.xlsx", sheet = "Pre-Screen", .name_repair = "universal"))

TH_Area_min <- prescreen$Area[2]
TH_Area_max <- prescreen$Area[1]
TH_Diameter_min <- prescreen$Diameter[2]
TH_Diameter_max <- prescreen$Diameter[1]
TH_Circularity_min <- prescreen$Circularity[2]
TH_Circularity_max <- prescreen$Circularity[1]
TH_Perimeter_min <- prescreen$Perimeter[2]
TH_Perimeter_max <- prescreen$Perimeter[1]
TH_Volume_min <- prescreen$Volume[2]
TH_Volume_max <- prescreen$Volume[1]

   Spheroid_data_1$Area <- ifelse(Spheroid_data_1$Area < TH_Area_max & Spheroid_data_1$Area > TH_Area_min,  Spheroid_data_1$Area, NA)   
   Spheroid_data_1$Diameter <- ifelse(Spheroid_data_1$Diameter<TH_Diameter_max & Spheroid_data_1$Diameter>TH_Diameter_min, Spheroid_data_1$Diameter, NA)
   Spheroid_data_1$Circularity <- ifelse(Spheroid_data_1$Circularity<TH_Circularity_max & Spheroid_data_1$Circularity>TH_Circularity_min, Spheroid_data_1$Circularity, NA)
   Spheroid_data_1$Volume <- ifelse(Spheroid_data_1$Volume<TH_Volume_max & Spheroid_data_1$Volume>TH_Volume_min, Spheroid_data_1$Volume, NA)
   Spheroid_data_1$Perimeter <- ifelse(Spheroid_data_1$Perimeter<TH_Perimeter_max & Spheroid_data_1$Perimeter>TH_Perimeter_min, Spheroid_data_1$Perimeter, NA)

 }   
  









#Area_A <- select(Spheroid_data_1, Well.Name, Spheroid_Area.TD.Area)

# split well name into row(character) and column(number)

well_1ch <- substr(Spheroid_data_1$Well.Name, start = 1, stop = 1)
well_2no <- as.numeric(substr(Spheroid_data_1$Well.Name, start = 2, stop = 3))


#add 2 cols to dataset for well column / row 

Spheroid_data_1$Well.row <- well_1ch
Spheroid_data_1$Well.col <- well_2no

Spheroid_data_1$Row <- well_1ch
Spheroid_data_1$Col <- well_2no

Spheroid_data_1a <- Spheroid_data_1[with(Spheroid_data_1, order(Well.row, Well.col)),]



Sph_Treat <- merge(Spheroid_data_1a,newdata, by=c("Row", "Col"))

Sph_Treat <- Sph_Treat[with(Sph_Treat, order(Well.row, Well.col)),]



message("Processing data...... ","\n")


#################################################################

######### data_summary - 
######### helper function used to calculate the mean and the standard dev, for the variable of interest, in each group :

data_summary <- function(data, varname, groupnames){
   require(plyr)
   summary_func <- function(x, col){
      c(mean = mean(x[[col]], na.rm=TRUE),
        sd = sd(x[[col]], na.rm=TRUE),
        se = std.error(x[[col]], na.rm=TRUE),
      med = median(x[[col]], na.rm=TRUE))
   }
   data_sum<-ddply(data, groupnames, .fun=summary_func,
                   varname)
   data_sum <- rename(data_sum, c("mean" = varname))
   return(data_sum)
   
}
#####  end of function



#####  Remove Area outliers ###################

Sph_Treat_datsum <- data_summary(Sph_Treat, varname="Area", 
                     groupnames=c("T_I"))

Sph_Treat_datsum <- Sph_Treat_datsum[ -c(2:4) ]
Sph_Treat_datsum <- rename(Sph_Treat_datsum, c("med" = "Area_Median"))

# merge Medians into main dataset ##

Sph_Treat_Robz_A <- merge(Sph_Treat,Sph_Treat_datsum, by=c("T_I"))


#get deviation from median (MAD is defined as the median of the absolute deviations from the datas median

Sph_Treat_Robz_A$Area_Devn <-  Sph_Treat_Robz_A$Area - Sph_Treat_Robz_A$Area_Median

#get abs deviation
Sph_Treat_Robz_A$Area_Abs_Devn <-  abs(Sph_Treat_Robz_A$Area - Sph_Treat_Robz_A$Area_Median)

#remove treatment columns, temporarily, add back in later

Sph_Treat_Robz_A <- Sph_Treat_Robz_A[,-c(11:23)]

Sph_Treat_Robz_A$Robz_LoLim <- RobZ_LoLim
Sph_Treat_Robz_A$Robz_HiLim <- RobZ_UpLim 

###calculate Mean Abs Deviations for TI groups

Sph_Treat_datsum2 <- data_summary(Sph_Treat_Robz_A, varname="Area_Abs_Devn", groupnames=c("T_I"))

Sph_Treat_datsum2 <- rename(Sph_Treat_datsum2, c("med" = "Area_MAD"))

Sph_Treat_datsum2 <- Sph_Treat_datsum2[, -c(2:4) ]

### merge MADs into main dataset

Sph_Treat_Robz_A <- merge(Sph_Treat_Robz_A,Sph_Treat_datsum2, by=c("T_I"))



#calulate Area Robust Z scores

Sph_Treat_Robz_A$Area_RobZ <- Sph_Treat_Robz_A$Area_Devn / Sph_Treat_Robz_A$Area_MAD
Sph_Treat_Robz_A$Area_status <-(ifelse(Sph_Treat_Robz_A$Area_RobZ >= RobZ_UpLim | Sph_Treat_Robz_A$Area_RobZ <= RobZ_LoLim , '1', '0'))


########################  end of Area outlier removal ######################


#####  Remove Diameter outliers, new dataset Sph_Treat_Robz_AD (ie AD = Area+Diameter) ###################

Sph_Treat_datsum <- data_summary(Sph_Treat, varname="Diameter", 
                                 groupnames=c("T_I"))

Sph_Treat_datsum <- Sph_Treat_datsum[ -c(2:4) ]
Sph_Treat_datsum <- rename(Sph_Treat_datsum, c("med" = "Diameter_Median"))

# merge Medians into main dataset ##

Sph_Treat_Robz_AD <- merge(Sph_Treat_Robz_A,Sph_Treat_datsum, by=c("T_I"))


#get deviation from median (MAD is defined as the median of the absolute deviations from the data's median

Sph_Treat_Robz_AD$Diameter_Devn <-  Sph_Treat_Robz_AD$Diameter - Sph_Treat_Robz_AD$Diameter_Median

#get abs deviation
Sph_Treat_Robz_AD$Diameter_Abs_Devn <-  abs(Sph_Treat_Robz_AD$Diameter - Sph_Treat_Robz_AD$Diameter_Median)


###calculate Mean Abs Deviations for TI groups

Sph_Treat_datsum2 <- data_summary(Sph_Treat_Robz_AD, varname="Diameter_Abs_Devn", groupnames=c("T_I"))

Sph_Treat_datsum2 <- rename(Sph_Treat_datsum2, c("med" = "Diameter_MAD"))

Sph_Treat_datsum2 <- Sph_Treat_datsum2[, -c(2:4) ]

###merge MAD's into main dataset

Sph_Treat_Robz_AD <- merge(Sph_Treat_Robz_AD,Sph_Treat_datsum2, by=c("T_I"))



#calulate Diameter Robust Z scores

Sph_Treat_Robz_AD$Diameter_RobZ <- Sph_Treat_Robz_AD$Diameter_Devn / Sph_Treat_Robz_AD$Diameter_MAD
Sph_Treat_Robz_AD$Diameter_status <-(ifelse(Sph_Treat_Robz_AD$Diameter_RobZ >= RobZ_UpLim | Sph_Treat_Robz_AD$Diameter_RobZ <= RobZ_LoLim , '1', '0'))


########################  end of Diameter outlier removal ######################


#####  Remove Volume outliers, new dataset Sph_Treat_Robz_ADV (ie ADV = Area+Diameter+Volume) ###################

Sph_Treat_datsum <- data_summary(Sph_Treat, varname="Volume", 
                                 groupnames=c("T_I"))

Sph_Treat_datsum <- Sph_Treat_datsum[ -c(2:4) ]
Sph_Treat_datsum <- rename(Sph_Treat_datsum, c("med" = "Volume_Median"))

# merge Medians into main dataset ##

Sph_Treat_Robz_ADV <- merge(Sph_Treat_Robz_AD,Sph_Treat_datsum, by=c("T_I"))


#get deviation from median (MAD is defined as the median of the absolute deviations from the data's median

Sph_Treat_Robz_ADV$Volume_Devn <-  Sph_Treat_Robz_ADV$Volume - Sph_Treat_Robz_ADV$Volume_Median

#get abs deviation
Sph_Treat_Robz_ADV$Volume_Abs_Devn <-  abs(Sph_Treat_Robz_ADV$Volume - Sph_Treat_Robz_ADV$Volume_Median)


###calculate Mean Abs Deviations for TI groups

Sph_Treat_datsum2 <- data_summary(Sph_Treat_Robz_ADV, varname="Volume_Abs_Devn", groupnames=c("T_I"))

Sph_Treat_datsum2 <- rename(Sph_Treat_datsum2, c("med" = "Volume_MAD"))

Sph_Treat_datsum2 <- Sph_Treat_datsum2[, -c(2:4) ]

###merge MAD's into main dataset

Sph_Treat_Robz_ADV <- merge(Sph_Treat_Robz_ADV,Sph_Treat_datsum2, by=c("T_I"))



#calulate Volume Robust Z scores

Sph_Treat_Robz_ADV$Volume_RobZ <- Sph_Treat_Robz_ADV$Volume_Devn / Sph_Treat_Robz_ADV$Volume_MAD
Sph_Treat_Robz_ADV$Volume_status <-(ifelse(Sph_Treat_Robz_ADV$Volume_RobZ >= RobZ_UpLim | Sph_Treat_Robz_ADV$Volume_RobZ <= RobZ_LoLim , '1', '0'))


########################  end of Volume outlier removal ######################



#####  Remove Perimeter outliers, new dataset Sph_Treat_Robz_ADVP (ie ADVP = Area+Diameter+Volume+ Perimeter) ###################

Sph_Treat_datsum <- data_summary(Sph_Treat, varname="Perimeter", 
                                 groupnames=c("T_I"))

Sph_Treat_datsum <- Sph_Treat_datsum[ -c(2:4) ]
Sph_Treat_datsum <- rename(Sph_Treat_datsum, c("med" = "Perimeter_Median"))

# merge Medians into main dataset ##

Sph_Treat_Robz_ADVP <- merge(Sph_Treat_Robz_ADV,Sph_Treat_datsum, by=c("T_I"))


#get deviation from median (MAD is defined as the median of the absolute deviations from the data's median

Sph_Treat_Robz_ADVP$Perimeter_Devn <-  Sph_Treat_Robz_ADVP$Perimeter - Sph_Treat_Robz_ADVP$Perimeter_Median

#get abs deviation
Sph_Treat_Robz_ADVP$Perimeter_Abs_Devn <-  abs(Sph_Treat_Robz_ADVP$Perimeter - Sph_Treat_Robz_ADVP$Perimeter_Median)


###calculate Mean Abs Deviations for TI groups

Sph_Treat_datsum2 <- data_summary(Sph_Treat_Robz_ADVP, varname="Perimeter_Abs_Devn", groupnames=c("T_I"))

Sph_Treat_datsum2 <- rename(Sph_Treat_datsum2, c("med" = "Perimeter_MAD"))

Sph_Treat_datsum2 <- Sph_Treat_datsum2[, -c(2:4) ]

###merge MAD's into main dataset

Sph_Treat_Robz_ADVP <- merge(Sph_Treat_Robz_ADVP,Sph_Treat_datsum2, by=c("T_I"))



#calulate Perimeter Robust Z scores

Sph_Treat_Robz_ADVP$Perimeter_RobZ <- Sph_Treat_Robz_ADVP$Perimeter_Devn / Sph_Treat_Robz_ADVP$Perimeter_MAD
Sph_Treat_Robz_ADVP$Perimeter_status <-(ifelse(Sph_Treat_Robz_ADVP$Perimeter_RobZ >= RobZ_UpLim | Sph_Treat_Robz_ADVP$Perimeter_RobZ <= RobZ_LoLim , '1', '0'))


########################  end of Perimeter outlier removal ######################


#####  Remove Circularity outliers, new dataset Sph_Treat_Robz_ADVP (ie ADVPC = Area+Diameter+Volume+ Perimeter + Circularity) ###################

Sph_Treat_datsum <- data_summary(Sph_Treat, varname="Circularity", 
                                 groupnames=c("T_I"))

Sph_Treat_datsum <- Sph_Treat_datsum[ -c(2:4) ]
Sph_Treat_datsum <- rename(Sph_Treat_datsum, c("med" = "Circularity_Median"))

# merge Medians into main dataset ##

Sph_Treat_Robz_ADVPC <- merge(Sph_Treat_Robz_ADVP,Sph_Treat_datsum, by=c("T_I"))


#get deviation from median (MAD is defined as the median of the absolute deviations from the data's median

Sph_Treat_Robz_ADVPC$Circularity_Devn <-  Sph_Treat_Robz_ADVPC$Circularity - Sph_Treat_Robz_ADVPC$Circularity_Median

#get abs deviation
Sph_Treat_Robz_ADVPC$Circularity_Abs_Devn <-  abs(Sph_Treat_Robz_ADVPC$Circularity - Sph_Treat_Robz_ADVPC$Circularity_Median)


###calculate Mean Abs Deviations for TI groups

Sph_Treat_datsum2 <- data_summary(Sph_Treat_Robz_ADVPC, varname="Circularity_Abs_Devn", groupnames=c("T_I"))

Sph_Treat_datsum2 <- rename(Sph_Treat_datsum2, c("med" = "Circularity_MAD"))

Sph_Treat_datsum2 <- Sph_Treat_datsum2[, -c(2:4) ]

###merge MAD's into main dataset

Sph_Treat_Robz_ADVPC <- merge(Sph_Treat_Robz_ADVPC,Sph_Treat_datsum2, by=c("T_I"))



#calculate Circularity Robust Z scores

Sph_Treat_Robz_ADVPC$Circularity_RobZ <- Sph_Treat_Robz_ADVPC$Circularity_Devn / Sph_Treat_Robz_ADVPC$Circularity_MAD
Sph_Treat_Robz_ADVPC$Circularity_status <-(ifelse(Sph_Treat_Robz_ADVPC$Circularity_RobZ >= RobZ_UpLim | Sph_Treat_Robz_ADVPC$Circularity_RobZ <= RobZ_LoLim , '1', '0'))

# reposition count, Robz_LoLim, Robz_HiLim columns in dataset

#AAtest0_Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC
#AAtest43_Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[,c(1:4, 15, 16,8, 5:7, 9:14,17:43 )]

Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[,c(1:4, 14, 15,8, 5:7, 9:13,16:42 )]

###################################################################################
#message("Generating plots...... ","\n")

########################  Start of Manual Outlier Override ######################

#TF_outlier_override == FALSE

#######  Apply override flags from config file to RobZ processed dataset 

ADCVP_Override_data <- suppressMessages(read_xlsx("Sph_Config_File_Rev2.xlsx", sheet = "Manual Override", range = cell_cols("B:G"), n_max=96))
  
A_OD <- ADCVP_Override_data
colnames(A_OD)<- c("Well.Name", "A_Flag","D_Flag","C_Flag","V_Flag","P_Flag")

# merge outlier override flags into main dataframe, rename as ADVPC_OVR.

ADVPC_OVR <- merge(Sph_Treat_Robz_ADVPC, A_OD, by= "Well.Name")

ADVPC_OVR <- ADVPC_OVR[with(ADVPC_OVR, order(Row, Col)),]
 

############ start manual utlier overrides

if (TF_outlier_override == 'TRUE')

{

#######  Area overrides

AS<- as.numeric(ADVPC_OVR$Area_status)
AS2<- as.numeric(ADVPC_OVR$Area_status)
AFlag <-  ADVPC_OVR$A_Flag
AS <- if_else(AFlag == "x", abs(AS2 - 1),  AS2 )
ADVPC_OVR$Area_status <- as.character(AS)
ADVPC_OVR <- ADVPC_OVR[with(ADVPC_OVR, order(Row, Col)),]


#################

#######  Diameter overrides

DS<- as.numeric(ADVPC_OVR$Diameter_status)
DS2<- as.numeric(ADVPC_OVR$Diameter_status)
DFlag <-  ADVPC_OVR$D_Flag
DS <- if_else(DFlag == "x", abs(DS2 - 1),  DS2 )
ADVPC_OVR$Diameter_status <- as.character(DS)
ADVPC_OVR <- ADVPC_OVR[with(ADVPC_OVR, order(Row, Col)),]



#################


#######  Circularity overrides

CS<- as.numeric(ADVPC_OVR$Circularity_status)
CS2<- as.numeric(ADVPC_OVR$Circularity_status)
CFlag <-  ADVPC_OVR$C_Flag
CS <- if_else(CFlag == "x", abs(CS2 - 1),  CS2 )
ADVPC_OVR$Circularity_status <- as.character(CS)
ADVPC_OVR <- ADVPC_OVR[with(ADVPC_OVR, order(Row, Col)),]

#################




#######  Volume overrides

VS<- as.numeric(ADVPC_OVR$Volume_status)
VS2<- as.numeric(ADVPC_OVR$Volume_status)
VFlag <-  ADVPC_OVR$V_Flag
VS <- if_else(VFlag == "x", abs(VS2 - 1),  VS2 )
ADVPC_OVR$Volume_status <- as.character(VS)
ADVPC_OVR <- ADVPC_OVR[with(ADVPC_OVR, order(Row, Col)),]

#################



#######  Perimeter overrides

PS<- as.numeric(ADVPC_OVR$Perimeter_status)
PS2<- as.numeric(ADVPC_OVR$Perimeter_status)
PFlag <-  ADVPC_OVR$P_Flag
PS <- if_else(PFlag == "x", abs(PS2 - 1),  PS2 )
ADVPC_OVR$Perimeter_status <- as.character(PS)

ADVPC_OVR <- ADVPC_OVR[with(ADVPC_OVR, order(Row, Col)),]


###############################################################################
###### All Override flags processed, rewrite override edited data to main dataset



Sph_Treat_Robz_ADVPC <- ADVPC_OVR[,c(1:42)]
#Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]

}


Sph_Treat_Robz_ADVPC<- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]



### create Area, Diameter etc columns with oUtliers removed eg (Area_OR), both via RobZ and Manual Override




Sph_Treat_Robz_ADVPC$Area_OR <- ifelse(Sph_Treat_Robz_ADVPC$Area_status=='0', Sph_Treat_Robz_ADVPC$Area, NA)
Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]

Sph_Treat_Robz_ADVPC$Diameter_OR <- ifelse(Sph_Treat_Robz_ADVPC$Diameter_status=='0', Sph_Treat_Robz_ADVPC$Diameter, NA)
Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]

Sph_Treat_Robz_ADVPC$Circularity_OR <- ifelse(Sph_Treat_Robz_ADVPC$Circularity_status=='0', Sph_Treat_Robz_ADVPC$Circularity, NA)
Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]

Sph_Treat_Robz_ADVPC$Volume_OR <- ifelse(Sph_Treat_Robz_ADVPC$Volume_status=='0', Sph_Treat_Robz_ADVPC$Volume, NA)
Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]

Sph_Treat_Robz_ADVPC$Perimeter_OR <- ifelse(Sph_Treat_Robz_ADVPC$Perimeter_status=='0', Sph_Treat_Robz_ADVPC$Perimeter, NA)
Sph_Treat_Robz_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(Row, Col)),]





###################################################################################


######################################################################
###########  START REPORT / PLOTTING 
######################################################################

Sph_Treat_Robz_ADVPC$T_I <- as.numeric(Sph_Treat_Robz_ADVPC$T_I)
Sph_Treat_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(T_I)),]


#subset data for outlier override table()... dfA = dataframe Area, dfD = Diameter etc...


### dfA####

dfA <- select(Sph_Treat_ADVPC, Well.Name ,T_I, Area, Area_status )
dfA$Area_status <- as.numeric(dfA$Area_status)


### dfD####

dfD <- select(Sph_Treat_ADVPC,Well.Name, T_I, Diameter,Diameter_status)
dfD$Diameter_status <- as.numeric(dfD$Diameter_status)


#################################

### dfC####

dfC <- select(Sph_Treat_ADVPC,Well.Name, T_I, Circularity,Circularity_status)

dfC$Circularity_status <- as.numeric(dfC$Circularity_status)


############################

### dfV####

dfV <- select(Sph_Treat_ADVPC,Well.Name, T_I, Volume,Volume_status)
dfV$Volume_status <- as.numeric(dfV$Volume_status)


############################

### dfP####

dfP <- select(Sph_Treat_ADVPC,Well.Name, T_I, Perimeter,Perimeter_status)
dfP$Perimeter_status <- as.numeric(dfP$Perimeter_status)

############################



############################################
## Generate plots and Write into spreadsheet

AreaPointcolour <- "black"
PerimeterPointcolour <- "black"
CircularityPointcolour <- "black"
CircularityPointcolour <- "black"
VolumePointcolour <- "black"
OutlierPointcolour <- "red"
Pointsize <- 1.5
OutlierPointSize  <- 1.5

Arealabel = expression("Area" ~ (mu~m^{2} ))
Diameterlabel = expression("Diameter" ~ (mu~m ))
Volumelabel = expression("Volume" ~ (mu~m^{3} ))
Perimeterlabel = expression("Perimeter" ~ (mu~m ))



cols <- c("1" = "red", "0" = "black")
#cols <- c("Outlier" = "red", "Normal" = "black")


######### Area PLots

p_Area_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Area, colour = Area_status), size = Pointsize,  show.legend=FALSE ) + 
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
      facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
   ylab(Arealabel)+xlab('Column') +  labs(title= " Area data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
   scale_colour_manual(values = cols) 

p_Area2 <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Area, colour = Area_status), size = Pointsize,  show.legend=FALSE ) + 
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   facet_wrap(~Sph_Treat_ADVPC$Row, nrow=1) +
   ylab(Arealabel)+xlab('Column') +  labs(title= " Area data by Row with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10)) +
   scale_colour_manual(values = cols) 

# dotplot : needs all _outlier removed datasets ASCVP, eg Area_OR 
p_Area_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Area_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                  geom = "crossbar", width = 0.25)  + 
   ylab(Arealabel) +xlab('Treatment Index') +  labs(title= "Area  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))


###### write plot png files to Excel sheet

wb <- createWorkbook()


addWorksheet(wb, "Area plots", gridLines = FALSE) 
addWorksheet(wb, "Dia.plots", gridLines = FALSE)
addWorksheet(wb, "Circ.plots", gridLines = FALSE)
addWorksheet(wb, "Vol.plots", gridLines = FALSE)
addWorksheet(wb, "Perim.plots", gridLines = FALSE)
addWorksheet(wb, "Area data", gridLines = FALSE)
addWorksheet(wb, "Dia.data", gridLines = FALSE)
addWorksheet(wb, "Circ.data", gridLines = FALSE)
addWorksheet(wb, "Vol.data", gridLines = FALSE)
addWorksheet(wb, "Perim.data", gridLines = FALSE)
addWorksheet(wb, "Main dataset", gridLines = FALSE)
addWorksheet(wb, "Export dataset", gridLines = FALSE)
addWorksheet(wb, "Summary", gridLines = FALSE)
addWorksheet(wb, "All plots", gridLines = FALSE)



print(p_Area_new)
#insertPlot(wb, 1, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
insertPlot(wb, 1, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")

suppressMessages(print(p_Area_dotplot_new))

insertPlot(wb, 1, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")


#################################################################################
#########   Plots
#message("Progress status :  Area processing completed" ,"\n")
message("Area processing completed.......")

Sph_Treat_ADVPC <- Sph_Treat_ADVPC[with(Sph_Treat_ADVPC, order(T_I)),]


########################### Diameter PLots

p_Diameter_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Diameter, colour = Diameter_status), size = Pointsize,  show.legend=FALSE ) + 
   # geom_point(data = Sph_Data_Out, aes(x= Col, y=Diameter_OO) , size = OutlierPointSize, colour = OutlierPointcolour ) +ylab(' Diameter  (um^2)')   xlab('Column') +
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   
   facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
   #   ylab(Arealabel) +xlab('Column') +  labs(title= " Area data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))
   ylab(Diameterlabel)+xlab('Column') +  labs(title= " Diameter data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
   scale_colour_manual(values = cols) 

p_Diameter2 <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Diameter, colour = Diameter_status), size = Pointsize,  show.legend=FALSE ) + 
   # geom_point(data = Sph_Data_Out, aes(x= Col, y=Diameter_OO) , size = OutlierPointSize, colour = OutlierPointcolour ) +ylab(' Diameter  (um^2)')   xlab('Column') +
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   facet_wrap(~Sph_Treat_ADVPC$Row, nrow=1) +
   #   ylab(Diameterlabel) +xlab('Column') +  labs(title= " Diameter data with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))
   ylab(Diameterlabel)+xlab('Column') +  labs(title= " Diameter data by Row with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10)) +
   scale_colour_manual(values = cols) 

p_Diameter_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Diameter_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                       geom = "crossbar", width = 0.25)  + 
   ylab(Diameterlabel) +xlab('Treatment Index') +  labs(title= "Diameter  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))

###### write plot png files to Excel sheet

#print(p_Diameter2) 
#insertPlot(wb, 2, xy = c("A", 1),width = 9, height = 3.5, fileType = "png", units = "in")

print(p_Diameter_new)
insertPlot(wb, 2, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")

suppressMessages(print(p_Diameter_dotplot_new))
insertPlot(wb, 2, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")

########################################End of diameter plots ###################
#message("Progress status :  Diameter processing completed" ,"\n")
message("Diameter processing completed.......")

########################### Circularity PLots

p_Circularity_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Circularity, colour = Circularity_status), size = Pointsize,  show.legend=FALSE ) + 
    scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   
   facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
   ylab('Circularity')+xlab('Column') +  labs(title= " Circularity data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
   scale_colour_manual(values = cols) 

p_Circularity2 <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Circularity, colour = Circularity_status), size = Pointsize,  show.legend=FALSE ) + 
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   facet_wrap(~Sph_Treat_ADVPC$Row, nrow=1) +
   ylab('Circularity')+xlab('Column') +  labs(title= " Circularity data by Row with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10)) +
   scale_colour_manual(values = cols) 

p_Circularity_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Circularity_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                               geom = "crossbar", width = 0.25)  + 
   ylab('Circularity') +xlab('Treatment Index') +  labs(title= "Circularity  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))

###### write plot png files to Excel sheet

#print(p_Circularity2) 
#insertPlot(wb, 3, xy = c("A", 1),width = 9, height = 3.5, fileType = "png", units = "in")

print(p_Circularity_new)
insertPlot(wb, 3, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")

suppressMessages(print(p_Circularity_dotplot_new))
insertPlot(wb, 3, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")

########################################End of Circularity plots ###################
#message("Progress status :  Circularity processing completed","\n")
message("Circularity processing completed.......")


########################### Volume PLots

p_Volume_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Volume, colour = Volume_status), size = Pointsize,  show.legend=FALSE ) + 
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   
   facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
   ylab(Volumelabel)+xlab('Column') +  labs(title= " Volume data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
   scale_colour_manual(values = cols) 

p_Volume2 <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Volume, colour = Volume_status), size = Pointsize,  show.legend=FALSE ) + 
  scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   facet_wrap(~Sph_Treat_ADVPC$Row, nrow=1) +
   ylab(Volumelabel)+xlab('Column') +  labs(title= " Volume data by Row with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10)) +
   scale_colour_manual(values = cols) 

p_Volume_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Volume_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                               geom = "crossbar", width = 0.25)  + 
   ylab(Volumelabel) +xlab('Treatment Index') +  labs(title= "Volume  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))

###### write plot png files to Excel sheet

#print(p_Volume2) 
#insertPlot(wb, 4, xy = c("A", 1),width = 9, height = 3.5, fileType = "png", units = "in")

print(p_Volume_new)
insertPlot(wb, 4, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")

suppressMessages(print(p_Volume_dotplot_new))
insertPlot(wb, 4, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")

########################################End of Volume plots ###################
#message("Progress status :  Volume processing completed","\n")
message("Volume processing completed.......")

########################### Perimeter PLots

p_Perimeter_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Perimeter, colour = Perimeter_status), size = Pointsize,  show.legend=FALSE ) + 
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+

   facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
   ylab(Perimeterlabel)+xlab('Column') +  labs(title= " Perimeter data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
   scale_colour_manual(values = cols) 

p_Perimeter2 <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Perimeter, colour = Perimeter_status), size = Pointsize,  show.legend=FALSE ) + 
   scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
   facet_wrap(~Sph_Treat_ADVPC$Row, nrow=1) +
   ylab(Perimeterlabel)+xlab('Column') +  labs(title= " Perimeter data by Row with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10)) +
   scale_colour_manual(values = cols) 

#Perbinwidth = diff(range(Sph_Treat_ADVPC$Perimeter_OR,na.rm=TRUE))/10

p_Perimeter_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Perimeter_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray")) + 
   stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,geom = "crossbar", width = 0.25)  + 
   ylab(Perimeterlabel) +xlab('Treatment Index') +  labs(title= "Perimeter  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))

#p_Perimeter_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Perimeter_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=Perbinwidth, stackdir = "center", fill= "darkgray")) + 
#   stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,geom = "crossbar", width = 0.25)  + 
#   ylab(Perimeterlabel) +xlab('Treatment Index') +  labs(title= "Perimeter  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))




###### write plot png files to Excel sheet

#print(p_Perimeter2) 
#insertPlot(wb, 5, xy = c("A", 1),width = 9, height = 3.5, fileType = "png", units = "in")

print(p_Perimeter_new)
insertPlot(wb, 5, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")

suppressMessages(print(p_Perimeter_dotplot_new))
insertPlot(wb, 5, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")

########################################End of Perimeter plots ###################
#message("Progress status :  Perimeter processing completed","\n")
message("Perimeter processing completed.......","\n")

message("Generating report file.......","\n")





#########################
#### Start writing ADCVP datasets into spreadsheets.


#  create formatting styles
datastyleC <- createStyle(fontSize = 10, fontColour = rgb(0,0,0),halign = "center", valign = "center")
datastyleR <- createStyle(fontSize = 10, fontColour = rgb(0,0,0), halign = "right", valign = "center")
datastyleL <- createStyle(fontSize = 10, fontColour = rgb(0,0,0), halign = "left", valign = "center")


styleCentre <- createStyle(halign = "center", valign = "center")
#redstyle <- createStyle(fontSize = 10, fontColour = rgb(1,0,0),textDecoration = c("bold"),halign = "center", valign = "center")
#redstyle <- createStyle(fontSize = 10, fontColour = rgb(1,0,0),bgFill = "yellow", textDecoration = c("bold"),halign = "center", valign = "center")
redstyle <- createStyle(fontSize = 10, fontColour = "black",bgFill = "lightpink",halign = "center", valign = "center")
styleborderR <- createStyle(border = "Right")
styleborderTBLR <- createStyle(border = "TopBottomLeftRight", borderColour = "black", halign = "center")
styleArea <- createStyle(border = "Right", halign = "center", valign = "center", numFmt = "#,##0.00")
styleCirc <- createStyle(border = "Right", halign = "center", valign = "center", numFmt = "#,##0.0000")
datestyle <- createStyle(numFmt = "hh:mm dd-mmm-yy" )
dateTBLRstyle <- createStyle(numFmt = "hh:mm dd-mmm-yy" , halign = "right", border = "TopBottomLeftRight" )
dpstyle <- createStyle(halign = "right", valign = "center", numFmt = "0.0000")
greyshadeTBLR <- createStyle(border = "TopBottomLeftRight", borderColour = "black", halign = "center", fgFill= "gray85")
styleBold <- createStyle(fontSize = 10, fontColour = "black",textDecoration = c("bold"))
styleItalic <- createStyle(fontSize = 10, fontColour = "black",textDecoration = c("italic"))
styleBoldlarge <- createStyle(fontSize = 14, fontColour = "black",textDecoration = c("bold"))

###################################
#insert Area data into tab

### write dfA dataset to table in 2nd tab.
A_Rows = nrow(dfA)+1

writeDataTable(wb, 6, dfA, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

addStyle(wb, sheet = 6, datastyleC, rows = 1 , cols = c(2:4), gridExpand = TRUE)
addStyle(wb, sheet = 6, datastyleC, rows = 1:A_Rows , cols = c(2:6), gridExpand = TRUE)
#addStyle(wb, sheet = 6, datastyleC, rows = 1:A_Rows , cols = 4, gridExpand = TRUE)
#header row centred

addStyle(wb, sheet = 6, styleCentre, rows = 1 , cols = c(2:4), gridExpand = TRUE)

setColWidths(wb, sheet = 6, cols = c(1:6), widths = c(4,14, 8, 16, 14, 14 ))
setColWidths(wb, sheet = 6, cols = c(1,5), hidden = rep(TRUE, length(5)))
addStyle(wb, sheet = 6, styleArea, rows = 2:A_Rows, cols = 4, gridExpand = T, stack = T)

# format red for outliers
conditionalFormatting(wb, sheet = 6, cols= 2:5, rows=2:A_Rows, rule = "$E2==1", style = redstyle)
freezePane(wb, sheet = 6 ,  firstActiveRow = 2,  firstActiveCol = 1)


############################################

#insert Diameter data into tab

#addStyle(wb, sheet = 1, datastyle, rows = 1:100 , cols = 16:22, gridExpand = TRUE)

writeDataTable(wb, 7, dfD, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

#freezePane(wb, sheet = 4 ,  firstActiveRow = 2,  firstActiveCol = 1)

addStyle(wb, sheet = 7, datastyleC, rows = 1:A_Rows , cols = c(2:3, 5:6), gridExpand = TRUE)
addStyle(wb, sheet = 7, datastyleR, rows = 1:A_Rows , cols = 4, gridExpand = TRUE)
addStyle(wb, sheet = 7, styleCentre, rows = 1 , cols = 4, gridExpand = TRUE)

setColWidths(wb, sheet = 7, cols = c(1:6), widths = c(4,14, 8, 16, 14, 14 ))
setColWidths(wb, sheet = 7, cols = c(1,5), hidden = rep(TRUE, length(5)))
addStyle(wb, sheet = 7, styleArea, rows = 2:A_Rows, cols = 4, gridExpand = T, stack = T)


conditionalFormatting(wb, sheet = 7, cols= 2:5, rows=2:A_Rows, rule = "$E2==1", style = redstyle)
freezePane(wb, sheet = 4 ,  firstActiveRow = 2,  firstActiveCol = 1)

##############################################

#insert Circularity data into tab

writeDataTable(wb, 8, dfC, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

#freezePane(wb, sheet = 6 ,  firstActiveRow = 2,  firstActiveCol = 1)

addStyle(wb, sheet = 8, datastyleC, rows = 1:100 , cols = c(2:3, 5:6), gridExpand = TRUE)
addStyle(wb, sheet = 8, datastyleR, rows = 1:100 , cols = 4, gridExpand = TRUE)
addStyle(wb, sheet = 8, styleCentre, rows = 1 , cols = 4, gridExpand = TRUE)


#setColWidths(wb, sheet = 8, cols = c(1,5), hidden = rep(TRUE, length(5)))
setColWidths(wb, sheet = 8, cols = c(1:6), widths = c(4,14, 8, 16, 14, 14 ))

setColWidths(wb, sheet = 8, cols = c(1,5), hidden = rep(TRUE, length(5)))
addStyle(wb, sheet = 8, styleCirc, rows = 2:A_Rows, cols = 4, gridExpand = T, stack = T)

conditionalFormatting(wb, sheet = 8, cols= 2:5, rows=2:97, rule = "$E2==1", style = redstyle)
freezePane(wb, sheet = 8,  firstActiveRow = 2,  firstActiveCol = 1)

##############################################

#insert Volume data into tab

writeDataTable(wb, 9, dfV, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

#freezePane(wb, sheet = 8 ,  firstActiveRow = 2,  firstActiveCol = 1)

addStyle(wb, sheet = 9, datastyleC, rows = 1:100 , cols = c(2:3, 5:6), gridExpand = TRUE)
addStyle(wb, sheet = 9, datastyleR, rows = 1:100 , cols = 4, gridExpand = TRUE)
addStyle(wb, sheet = 9, styleCentre, rows = 1 , cols = 4, gridExpand = TRUE)


setColWidths(wb, sheet = 9, cols = c(1:6), widths = c(4,14, 8, 16, 14, 14 ))
setColWidths(wb, sheet = 9, cols = c(1,5), hidden = rep(TRUE, length(5)))
addStyle(wb, sheet = 9, styleArea, rows = 2:A_Rows, cols = 4, gridExpand = T, stack = T)



conditionalFormatting(wb, sheet = 9, cols= 2:5, rows=2:97, rule = "$E2==1", style = redstyle)
freezePane(wb, sheet = 9 ,  firstActiveRow = 2,  firstActiveCol = 1)

##############################################

#insert Perimeter data into tab

writeDataTable(wb, 10, dfP, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)



addStyle(wb, sheet = 10, datastyleC, rows = 1:100 , cols = c(2:3, 5:6), gridExpand = TRUE)
addStyle(wb, sheet = 10, datastyleR, rows = 1:100 , cols = 4, gridExpand = TRUE)
addStyle(wb, sheet = 10, styleCentre, rows = 1 , cols = 4, gridExpand = TRUE)


setColWidths(wb, sheet = 10, cols = c(1:6), widths = c(4,14, 8, 16, 14, 14 ))
setColWidths(wb, sheet = 10, cols = c(1,5), hidden = rep(TRUE, length(5)))
addStyle(wb, sheet = 10 , styleArea, rows = 2:A_Rows, cols = 4, gridExpand = T, stack = T)


conditionalFormatting(wb, sheet = 10, cols= 2:5, rows=2:97, rule = "$E2==1", style = redstyle)
freezePane(wb, sheet = 10 ,  firstActiveRow = 2,  firstActiveCol = 1)

##############################################
#insert Final Processed Dataset data into tab

#writeData(wb, 11, Sph_Treat_Robz_ADVPC, colNames=TRUE)
#writeDataTable(wb, 10, dfP, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

writeDataTable(wb, 11, Sph_Treat_Robz_ADVPC, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)
#setColWidths(wb, 11, cols, widths = "auto", hidden = rep(FALSE, length(cols)))

addStyle(wb, sheet = 11, datastyleC, rows = 2:100 , cols = c(1:7, 14, 15, 18, 24, 30, 36, 42), gridExpand = TRUE)
addStyle(wb, sheet = 11, datastyleC, rows = 1 , cols = c(1:47), gridExpand = TRUE)

setColWidths(wb, sheet = 11, cols = c(9:47), widths = 16 )
setColWidths(wb, sheet = 11, cols = c(4:8), widths = 10 )
setColWidths(wb, sheet = 11, cols = c(1:3), widths = 8 )

freezePane(wb, sheet = 11,  firstActiveRow = 2)
#  Save outlier analysis report , 


#######################################
#saveWorkbook(wb, report_output_file, overwrite = TRUE)

#######################################

###################################################################################

## helper function for Means and SE's, by T_I

data_summary_TI <- function(data, varname, groupnames){
   require(plyr)
   summary_func <- function(x, col){
      c(mean = mean(x[[col]], na.rm=TRUE),
        se = std.error(x[[col]], na.rm=TRUE))
      
   }
   data_sum<-ddply(data, groupnames, .fun=summary_func,
                   varname)
   data_sum <- rename(data_sum, c("mean" = varname))
   return(data_sum)
   
}







# Write .CSV file :   Treatments + ADVPC data , all outliers removed, once checkbox enabled in config file.

#if (TF_copytomergedir == 'TRUE')  
#   {

   
   # remove all unnecessary columns for calcs, add treatment data and write to Temp CSV file for merging later

#Sph_ADVPC <- Sph_Treat_ADVPC[,c(1:4, 53:57)]

Sph_ADVPC <- Sph_Treat_ADVPC[,c(1:4, 43:47)]

# merge with treatment data 

Sph_ADVPC  <- merge(Sph_ADVPC, P_treatment, by= "T_I")
Sph_ADVPC <- Sph_ADVPC[with(Sph_ADVPC, order(Row, Col)),]

Job_Info_data <- select(Spheroid_data, Jobrun.Finish.Time)
Sph_ADVPC$Job.Date <- as.Date(Job_Info_data$Jobrun.Finish.Time, origin="1900-01-01")

#work out time diff in days ETST.

Sph_ADVPC$ETST.d <-  difftime(Sph_ADVPC$Job.Date, Sph_ADVPC$Time_Date, units= "days")
Sph_ADVPC$ETST.h <-  difftime(Sph_ADVPC$Job.Date, Sph_ADVPC$Time_Date, units= "hours")


Sph_ADVPC$Filename <- rawfilename
Sph_ADVPCa <-  Sph_ADVPC[,c(1:2,5:22)]

## calculate Means and SE's for A D C V P items 

aArea_means <- data_summary_TI(Sph_ADVPCa, varname="Area_OR", 
                                 groupnames=c("T_I"))

aDiameter_means <- data_summary_TI(Sph_ADVPCa, varname="Diameter_OR", 
                              groupnames=c("T_I"))

aCircularity_means <- data_summary_TI(Sph_ADVPCa, varname="Circularity_OR", 
                              groupnames=c("T_I"))


aVolume_means <- data_summary_TI(Sph_ADVPCa, varname="Volume_OR", 
                              groupnames=c("T_I"))


aPerimeter_means <- data_summary_TI(Sph_ADVPCa, varname="Perimeter_OR", 
                              groupnames=c("T_I"))


#merge all mean+SE datasets sequentially

a_allmeans <- merge(aArea_means, aDiameter_means, by= "T_I")
a_allmeans1 <- merge(a_allmeans, aCircularity_means, by= "T_I")
a_allmeans2 <- merge(a_allmeans1, aVolume_means, by= "T_I")
a_allmeans3 <- merge(a_allmeans2, aPerimeter_means, by= "T_I")
colnames(a_allmeans3)<- c("T_I", "Area_Mean","Area_SE","Diameter_Mean","Diameter_SE","Circularity_Mean","Circularity_SE","Volume_Mean","Volume_SE", "Perimeter_Mean","Perimeter_SE" )

# subset TI and treatment items from main dataset

Treatments <-  Sph_ADVPC[,c(1,8:20)]

#merge final means dataset with Treatment data items


ADVPC_means <- merge(a_allmeans3, P_treatment, by= "T_I") 

job.date = Job_Info_data$Jobrun.Finish.Time[1]
ADVPC_means$Job.Date <- job.date



# calculate time diff in days and hours ETST.

ADVPC_means$ETST.d <-  as.numeric(difftime(ADVPC_means$Job.Date, ADVPC_means$Time_Date, units= "days"))
ADVPC_means$ETST.h <-  as.numeric(difftime(ADVPC_means$Job.Date, ADVPC_means$Time_Date, units= "hours"))
ADVPC_means$Filename <- rawfilename


#(write TI group means only) 


#ADVPC_means <- ADVPC_means[with(ADVPC_means, order(T_I)),]
aatest_ADVPC_means <- ADVPC_means

names(ADVPC_means)[13] <- "Date_Treated"
names(ADVPC_means)[23] <- "Date_Scanned"
ADVPC_means <-  ADVPC_means[,c(1,12, 2:11,14:22, 13, 23:26)]

ADVPC_means <- ADVPC_means[with(ADVPC_means, order(T_I)),]

if (TF_copytomergedir == 'TRUE')  
{

write.csv(ADVPC_means, file = merge_output_file, col.names = TRUE, row.names = FALSE, append = FALSE)

}

###########end of write CSV Export file#########################

### start write Mean/SE  CSV data to report at tab 12
##############################################

#writeData(wb, 12, Sph_Treat_Robz_ADVPC, colNames=TRUE)
#writeDataTable(wb, 10, dfP, startRow = 1, startCol = 2, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

writeDataTable(wb, 12, ADVPC_means, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

addStyle(wb, sheet = 12, datastyleC, rows = 1 , cols = c(1:26), gridExpand = TRUE)
addStyle(wb, sheet = 12, datastyleC, rows = 2:33 , cols = c(13:21), gridExpand = TRUE)
addStyle(wb, sheet = 12, datastyleR, rows = 2:33 , cols = c(3:12,24:26), gridExpand = TRUE)

addStyle(wb, sheet = 12, datestyle, rows = 2:33 , cols = c(22,23), gridExpand = TRUE)
addStyle(wb, sheet = 12, dpstyle, rows = 2:33 , cols = c( 3:12,24:25), gridExpand = TRUE)

addStyle(wb, sheet = 12, datastyleC, rows = 2:33 , cols = 1, gridExpand = TRUE)
addStyle(wb, sheet = 12, datastyleR, rows = 2:33 , cols = c(2,26), gridExpand = TRUE)


setColWidths(wb, sheet = 12, cols = c(3:12, 24:25), widths = 16 )
setColWidths(wb, sheet = 12, cols = c(13:21, 24, 25), widths = 10 )
setColWidths(wb, sheet = 12, cols = c(22, 23), widths = 20 )

setColWidths(wb, sheet = 12, cols = 26, widths = 30 )
setColWidths(wb, sheet = 12, cols = 2, widths = 20)


freezePane(wb, sheet = 12,  firstActiveRow = 2)

############################################


# read plate setup from config file, copy into tab 13 of report 

platelayout <- suppressMessages(read.xlsx("Sph_Config_File_Rev2.xlsx", sheet = platesetupname, rows = c(1:9), cols = c(2:14), 
                    colNames = TRUE))
colnames(platelayout) <- c(" ","1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")


writeData(wb, 13, platelayout, startRow = 12, startCol = 1, rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)

setColWidths(wb, sheet = 13, cols = 1:13, widths = 5)
#setColWidths(wb, sheet = 13, cols = 1, widths = 9)
addStyle(wb, sheet = 13, styleborderTBLR, rows= 13:20 , cols= 2:13, gridExpand = TRUE)
#addStyle(wb, sheet = 13, datastyleC, rows = 12, cols = 2:13, gridExpand = TRUE)
addStyle(wb, sheet = 13, greyshadeTBLR, rows= 12 , cols= 2:13, gridExpand = TRUE)
addStyle(wb, sheet = 13, greyshadeTBLR, rows= 13:20 , cols= 1, gridExpand = TRUE)




#copy treatment Index definitions into tab 13 of report

#P_treatment2 <- subset(P_treatment, Cell_line != "", select= T_I:Conc_3)


colnames(P_treatment)[1] <- "T_I"
colnames(P_treatment)[5] <- "Passage"
colnames(P_treatment)[6] <- "RA.dose"

writeData(wb, 13, P_treatment, startRow = 12, startCol = 15, rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)

setColWidths(wb, sheet = 13, cols = 16, widths = 20)
setColWidths(wb, sheet = 13, cols = 17, widths = 16)
setColWidths(wb, sheet = 13, cols = c(15,19:26), widths = 8)
setColWidths(wb, sheet = 13, cols = 18, widths = 12)
setColWidths(wb, sheet = 13, cols = 14, widths = 8)

addStyle(wb, sheet = 13, datastyleC, rows= 12:44 , cols= 15:26, gridExpand = TRUE)
addStyle(wb, sheet = 13, styleborderTBLR, rows= 13:44 , cols= 15:26, gridExpand = TRUE)
addStyle(wb, sheet = 13, dateTBLRstyle, rows= 13:45 , cols= 17, gridExpand = TRUE)
addStyle(wb, sheet = 13, greyshadeTBLR, rows= 12 , cols= 15:26, gridExpand = TRUE)
addStyle(wb, sheet = 13, greyshadeTBLR, rows= 13:44 , cols= 15, gridExpand = TRUE)


### create Summary tab, detailing all treatment parameters used, outlier overrides, threshold limits etc 


c1 <- "Plate setup"
writeData(wb, 13, c1, startCol = 1, startRow = 11)

c2 <- "Treatment parameters"
writeData(wb, 13, c2, startCol = 15, startRow = 11)

addStyle(wb, sheet = 13, styleBold, rows= 11 , cols= 1, gridExpand = TRUE)
addStyle(wb, sheet = 13, styleBold, rows= 11 , cols= 15, gridExpand = TRUE)



#t1 <- "      Summary of processing parameters used"
t1 <- "      Summary of parameters used for processing"

p0 <- "       Raw data filename (.xlsx)"
p1 <- "       Plate setup tab" 
p2 <- "       Robust z low limit"
p3 <- "       Robust z high limit"
p4 <- "       Apply Thresholds"
p5 <- "       Apply outlier overrides"
psep <- ":"

#par0 <- paste(rawfilename,".xlsx",sep="")
par0 <- rawfilename
par1 <- platesetupname
par2 <- RobZ_LoLim 
par3 <- RobZ_UpLim 
par4 <- TF_apply_thresholds
par5 <- TF_outlier_override

writeData(wb, 13, t1, startCol = 1, startRow = 1)
writeData(wb, 13, p0, startCol = 4, startRow = 3)
writeData(wb, 13, p1, startCol = 4, startRow = 4)
writeData(wb, 13, p2, startCol = 4, startRow = 5)
writeData(wb, 13, p3, startCol = 4, startRow = 6)
writeData(wb, 13, p4, startCol = 4, startRow = 8)
writeData(wb, 13, p5, startCol = 4, startRow = 9)

writeData(wb, 13, psep, startCol = 5, startRow = 3)
writeData(wb, 13, psep, startCol = 5, startRow = 4)
writeData(wb, 13, psep, startCol = 5, startRow = 5)
writeData(wb, 13, psep, startCol = 5, startRow = 6)
writeData(wb, 13, psep, startCol = 5, startRow = 8)
writeData(wb, 13, psep, startCol = 5, startRow = 9)

          
#writeData(wb, 13, p0, startCol = 1, startRow = 3)
#writeData(wb, 13, p1, startCol = 1, startRow = 4)
#writeData(wb, 13, p2, startCol = 1, startRow = 5)
#writeData(wb, 13, p3, startCol = 1, startRow = 6)
#writeData(wb, 13, p4, startCol = 1, startRow = 8)
#writeData(wb, 13, p5, startCol = 1, startRow = 9)

addStyle(wb, sheet = 13, styleBoldlarge, rows= 1 , cols= 1, gridExpand = TRUE)
addStyle(wb, sheet = 13, styleCentre, rows= c(3:9) , cols= 5, gridExpand = TRUE)
addStyle(wb, sheet = 13, datastyleL, rows= c(3:9) , cols= 6, gridExpand = TRUE)
addStyle(wb, sheet = 13, datastyleR, rows= c(3:9) , cols= 4, gridExpand = TRUE)

writeData(wb, 13, par0, startCol = 6, startRow = 3)
writeData(wb, 13, par1, startCol = 6, startRow = 4)
writeData(wb, 13, par2, startCol = 6, startRow = 5)
writeData(wb, 13, par3, startCol = 6, startRow = 6)
writeData(wb, 13, par4, startCol = 6, startRow = 8)
writeData(wb, 13, par5, startCol = 6, startRow = 9)



#copy outlier overrides into summary worksheet
if (TF_outlier_override == 'TRUE')

{

   c3 <- "Manual outlier override"
   writeData(wb, 13, c3, startCol = 2, startRow = 22)
   addStyle(wb, sheet = 13, styleBold, rows= 22 , cols= 2, gridExpand = TRUE)
   

colnames(ADCVP_Override_data) <- c("Well", "Area", "Dia.", "Circ.","Vol.", "Perim." )
     
writeData(wb, 13, ADCVP_Override_data, startRow = 23, startCol = 2,  rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)
addStyle(wb, sheet = 13, styleborderTBLR, rows= 23:119 , cols= 2:7, gridExpand = TRUE)
addStyle(wb, sheet = 13, greyshadeTBLR, rows= 23 , cols= 2:7, gridExpand = TRUE)
addStyle(wb, sheet = 13, greyshadeTBLR, rows= 23:119 , cols= 2, gridExpand = TRUE)
 
}

if (TF_apply_thresholds == 'TRUE')
{
   Thresholdlayout <- read.xlsx("Sph_Config_File_Rev2.xlsx", sheet = "Pre-Screen", rows = c(1:3), cols = c(2:7), 
                            colNames = TRUE)  
   Thresholdlayout_s <- select(Thresholdlayout, Area, Diameter, Volume, Perimeter, Circularity)
   writeData(wb, 13, Thresholdlayout_s, startRow = 3, startCol = 16,  rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)  
   
   th1 <- "High limit" 
   th2 <- "Low limit"
   TH_T <-"Pre-screen thresholds"
   com_T1 <- "Comments" 
   com_T2 <- "….................comment 1 …............................................................................................................." 
   com_T3 <- "….................comment 2 …............................................................................................................."
    
   writeData(wb, 13, th1, startCol = 15, startRow = 4)
   writeData(wb, 13, th2, startCol = 15, startRow = 5)
   writeData(wb, 13, TH_T, startCol = 16, startRow = 2)
   writeData(wb, 13, com_T1, startCol = 16, startRow = 7)
   writeData(wb, 13, com_T2, startCol = 16, startRow = 8)
   writeData(wb, 13, com_T3, startCol = 16, startRow = 9)
   
   
   addStyle(wb, sheet = 13, styleBold, rows= c(2,7) , cols= 16, gridExpand = TRUE)
   addStyle(wb, sheet = 13, styleItalic, rows= c(8:9) , cols= 16, gridExpand = TRUE)
   
   addStyle(wb, sheet = 13, styleborderTBLR, rows= 4:5 , cols= 16:20, gridExpand = TRUE)
   addStyle(wb, sheet = 13, greyshadeTBLR, rows= 3 , cols= 16:20, gridExpand = TRUE)
  
   
}


saveWorkbook(wb, report_output_file, overwrite = TRUE)



if (TF_checkfile == "TRUE")
{
   message("\n")
#   message("#### Report file overwrite notification ####","\n")
   message("#### Report file overwrite notification ####")
 message(Msg_file_duplicated_1)
 #  message("\n")
#message("####  CORRECTIVE ACTION ####", "\n")
   message(Msg_file_duplicated_2)
   message("\n")
  }

if (TF_checkCSVfile == "TRUE")
{
 # message("\n")
 # message("#### Output CSV file overwrite notification ####","\n")
   message("#### Output CSV file overwrite notification ####")
   message(Msg_CSVfile_duplicated_1)
   #message("\n")
   #message("####  CORRECTIVE ACTION ####", "\n")
   message(Msg_CSVfile_duplicated_2)
   message("\n")
}


message("All processing completed successfully.", "\n")
message("\n") 

message("Please disregard the following message - (refers to column name changes while importing xlsx file.) ")
sink()


