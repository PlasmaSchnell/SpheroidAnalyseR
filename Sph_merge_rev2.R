################### R CODE INFORMATION   ##################################
#                                                                         #
# filename :        Sph_merge.R                                           #
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


library(tidyverse)
library(openxlsx)

sink()
sink("NUL")

for (j in 1:12)
{
  message("\n") 
}

message("Merging files......")

for (j in 1:3)
{
  message("\n") 
}

slash = "//"
TF_checkMergefile <- "FALSE"

# read config file : filenames and directory info

suppressMessages(merge_params <- read_excel("Sph_Config_File_Rev2.xlsx", sheet = "Merge files", .name_repair = "universal"))
mergepath <- merge_params$Input_field[2]

merge_output_path <- merge_params$Input_field[4]

merge_file <- merge_params$Input_field[6]

merge_output_file_xlsx <- paste(merge_output_path,slash, merge_file,'.xlsx',sep="")

message("\n")

# Error handling - check if overwriting existing merged xlsx file

if (file.exists(merge_output_file_xlsx))
{ 
  
  TF_checkMergefile <- 'TRUE'
  merge_file2 <- paste(merge_file, "(2)", sep="")
  merge_output_file_xlsx <- paste(merge_output_path,slash, merge_file2,'.xlsx',sep="")
  
  Msg_mergefile_duplicated_1 <- paste("Merged output data file already exists - saving a duplicate : ", merge_file2, ".xlsx",sep="") 
  Msg_mergefile_duplicated_2 <-'Enter a new merged output data filename in the config file and process again, or rename the duplicate file.'
}

# merge all files, using multmerge function 
# refer to : https://stackoverflow.com/questions/30242065/trying-to-merge-multiple-csv-files-in-r


multmerge = function(mypath){
  filenames=list.files(path=mypath, full.names=TRUE)
  datalist = lapply(filenames, function(x){read.csv(file=x,header=T)})
  Reduce(function(x,y) {merge(x,y,all = TRUE)}, datalist)
}

full_data = multmerge(mergepath)

## sort merged dataset full_data and write merged data to csv file.

full_data_A  <- full_data[with(full_data , order(ETST.d, T_I)),]
full_data_B <- full_data_A

full_data_B$Date_Treated <- as.POSIXct(full_data_A$Date_Treated, format = "%Y-%m-%d %H:%M:%S")
full_data_B$Date_Scanned <- as.POSIXct(full_data_A$Date_Scanned, format = "%Y-%m-%d %H:%M:%S")


#option to save CSV, rem out for now..
#write.csv(full_data_A,merge_output_file )

#CSV merge file written
###################################
## start writing merge data to formatted xlsx file

#  create formatting styles
datastyleC <- createStyle(fontSize = 10, fontColour = rgb(0,0,0),halign = "center", valign = "center")
datastyleR <- createStyle(fontSize = 10, fontColour = rgb(0,0,0), halign = "right", valign = "center")
datastyleL <- createStyle(fontSize = 10, fontColour = rgb(0,0,0), halign = "left", valign = "center")

styleCentre <- createStyle(halign = "center", valign = "center")

datestyle <- createStyle(numFmt = "hh:mm  dd-mmm-yyyyy")
dpstyle <- createStyle(halign = "right", valign = "center", numFmt = "0.0000")



wb_mergedata <- createWorkbook()
addWorksheet(wb_mergedata , "Merged data", gridLines = FALSE)

writeDataTable(wb_mergedata, 1, full_data_B, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = FALSE)
maxrows <-nrow(full_data_A)+1


addStyle(wb_mergedata, sheet = 1, datestyle, rows = 2:maxrows , cols = c(22,23), gridExpand = TRUE)

addStyle(wb_mergedata, sheet = 1, dpstyle, rows = 2:maxrows , cols = c( 3:12,24:25), gridExpand = TRUE)

addStyle(wb_mergedata, sheet = 1, datastyleC, rows = 2:maxrows , cols = c(1,13:21, 24,25), gridExpand = TRUE)
addStyle(wb_mergedata, sheet = 1, datastyleC, rows = 1 , cols = c(1:26), gridExpand = TRUE)

addStyle(wb_mergedata, sheet = 1, datastyleR, rows = 2:maxrows , cols = c(3:12), gridExpand = TRUE)


setColWidths(wb_mergedata, sheet = 1, cols = c(3:12, 24:25), widths = 16 )
setColWidths(wb_mergedata, sheet = 1, cols = c(13:21), widths = 10 )
setColWidths(wb_mergedata, sheet = 1, cols = c(22, 23), widths = 24 )

setColWidths(wb_mergedata, sheet = 1, cols = 26, widths = 30 )
setColWidths(wb_mergedata, sheet = 1, cols = 2, widths = 20)


freezePane(wb_mergedata, sheet = 1,  firstActiveRow = 2)

#  Save outlier analysis report , 

saveWorkbook(wb_mergedata, merge_output_file_xlsx, overwrite = TRUE)

message("Merge files completed....", "\n") 

# check for overwriting existing file, message about duplicate filename if TRUE

if ( TF_checkMergefile == "TRUE")
{
  message("\n")
  # message("#### Merged Output file overwrite notification ####","\n")
  message("#### Merged Output CSV file overwrite notification ####", "\n")
  message(Msg_mergefile_duplicated_1)
  #message("\n")
  message(Msg_mergefile_duplicated_2)
  message("\n")
}


message("Files merged successfully to : ",merge_output_file_xlsx, "\n")

sink()

######################## End

