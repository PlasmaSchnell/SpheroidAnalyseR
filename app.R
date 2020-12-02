#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(tidyverse)
library(ggthemes)
library(gridExtra)
library(readxl)
library(openxlsx)
library(plotrix)
library(writexl)
library(ggpubr)

#######Predefine functions ######


#####  helper function for Means and SE's
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


##### helper function for Means and SE's, by T_I

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

##### function calculating outliers
cal_z_score = function(df_sph_treat, df_prev, varname, RobZ_LoLim, RobZ_UpLim){
  med_name = paste0(varname,"_Median")
  Devn_name = paste0(varname,"_Devn")
  Abs_Devn_name = paste0(varname,"_Abs_Devn")
  
  MAD_name = paste0(varname,"_MAD")
  RobZ_name = paste0(varname,"_RobZ")
  status_name= paste0(varname,"_status")
  
  Sph_Treat_datsum <- data_summary(df_sph_treat, varname=varname, 
                                   groupnames=c("T_I"))
  
  Sph_Treat_datsum <- Sph_Treat_datsum[ -c(2:4) ]
  Sph_Treat_datsum <- rename(Sph_Treat_datsum, c("med" = med_name))
  
  
  Sph_Treat_Robz <- merge(df_prev,Sph_Treat_datsum, by=c("T_I"))
  
  Sph_Treat_Robz[,Devn_name] <-  Sph_Treat_Robz[,varname] - Sph_Treat_Robz[,med_name]
  
  #get abs deviation
  Sph_Treat_Robz[,Abs_Devn_name] <-  abs(Sph_Treat_Robz[,varname]  - Sph_Treat_Robz[,med_name])
  
  Sph_Treat_Robz$Robz_LoLim <- RobZ_LoLim
  Sph_Treat_Robz$Robz_HiLim <- RobZ_UpLim 
  
  
  Sph_Treat_datsum2 <- data_summary(Sph_Treat_Robz, varname=Abs_Devn_name, groupnames=c("T_I"))
  
  Sph_Treat_datsum2 <- rename(Sph_Treat_datsum2, c("med" = MAD_name))
  
  Sph_Treat_datsum2 <- Sph_Treat_datsum2[, -c(2:4) ]
  
  ### merge MADs into main dataset
  Sph_Treat_Robz <- merge(Sph_Treat_Robz,Sph_Treat_datsum2, by=c("T_I"))
  
  #calulate Area Robust Z scores
  
  Sph_Treat_Robz[,RobZ_name] <- Sph_Treat_Robz[,Devn_name]/ Sph_Treat_Robz[,MAD_name]
  Sph_Treat_Robz[,status_name] <-(ifelse(Sph_Treat_Robz[,RobZ_name] >= RobZ_UpLim | 
                                           Sph_Treat_Robz[,RobZ_name]<= RobZ_LoLim , '1', '0'))
  return(Sph_Treat_Robz)
  
}
#####  end of function

##### function for drawing outliers
draw_outlier_plot = function(df, value){
  df$is_Outlier = is.na(df[,value])
  levels(df$is_Outlier) = c(TRUE, FALSE)
  colours = c("TRUE" = "red", "FALSE" = "white")
  
  ggplot(df, aes(x = Row, y = Col, fill = is_Outlier,label=Well.Name)) +
    geom_tile() + 
    scale_y_continuous(breaks=1:12) + 
    scale_fill_manual(values=colours, drop=FALSE) + geom_text()
}


draw_z_score_outlier_plot = function(df, value){
  value= paste0(value,"_status")
  
  # levels(df[,value]) = c("1", "0", "NA")
  # colours = c("1" = "red", "0" = "white", "NA"='grey')
  
  df[,value] = factor(df[,value])
  levels(df[,value]) = c(1,0,as.integer(NA))
  colours = c("1" = "red", "0" = "white", "NA"='grey')
  # colours = c(1 = "red", 0 = "white", NA='grey')
  
  
  ggplot(df, aes_string(x = "Row", y = "Col", fill = value,label="Well.Name")) +
    geom_tile() + 
    scale_y_continuous(breaks=1:12) + 
    scale_fill_manual(values=colours, drop=FALSE) + geom_text()
}


##### function for generating the report

gen_report = function(Sph_Treat_Robz_ADVPC,
                      df_treat,df_setup,
                      Spheroid_data,
                      p_Area_new,p_Area_dotplot_new,
                      p_Diameter_new ,p_Diameter_dotplot_new,
                      p_Circularity_new,p_Circularity_dotplot_new,
                      p_Volume_new, p_Volume_dotplot_new,
                      p_Perimeter_new, p_Perimeter_dotplot_new,
                      rawfilename,platesetupname,
                      RobZ_LoLim,RobZ_UpLim,
                      TF_apply_thresholds,
                      TF_outlier_override,
                      TH_Area_max=NA,TH_Area_min=NA,
                      TH_Diameter_max=NA,TH_Diameter_min=NA,
                      TH_Volume_max=NA, TH_Volume_min=NA,
                      TH_Perimeter_max = NA, TH_Perimeter_min=NA,
                      TH_Circularity_max = NA, TH_Circularity_min=NA,
                      TF_copytomergedir= FALSE
                      ){

  
  Sph_Treat_Robz_ADVPC$T_I <- as.numeric(Sph_Treat_Robz_ADVPC$T_I)
  Sph_Treat_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(T_I)),]
  
  #subset data for outlier override table()... dfA = dataframe Area, dfD = Diameter etc...
  ### dfA####
  dfA <- select(Sph_Treat_ADVPC, Well.Name ,T_I, Area, Area_status )
  dfA$Area_status <- as.numeric(dfA$Area_status)
  ### dfD####
  dfD <- select(Sph_Treat_ADVPC,Well.Name, T_I, Diameter,Diameter_status)
  dfD$Diameter_status <- as.numeric(dfD$Diameter_status)
  
  ### dfC####
  dfC <- select(Sph_Treat_ADVPC,Well.Name, T_I, Circularity,Circularity_status)
  dfC$Circularity_status <- as.numeric(dfC$Circularity_status)
  ### dfV####
  dfV <- select(Sph_Treat_ADVPC,Well.Name, T_I, Volume,Volume_status)
  dfV$Volume_status <- as.numeric(dfV$Volume_status)
  ### dfP####
  dfP <- select(Sph_Treat_ADVPC,Well.Name, T_I, Perimeter,Perimeter_status)
  dfP$Perimeter_status <- as.numeric(dfP$Perimeter_status)
  
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
  
  ### writing area plots  ###
  suppressMessages(print(p_Area_new))
  insertPlot(wb, 1, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  suppressMessages(print(p_Area_dotplot_new))
  insertPlot(wb, 1, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  ### writing diameter plots  ###
  suppressMessages(print(p_Diameter_new))
  insertPlot(wb, 2, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  suppressMessages(print(p_Diameter_dotplot_new))
  insertPlot(wb, 2, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")

  ###### write Circularity plots
  suppressMessages(print(p_Circularity_new))
  insertPlot(wb, 3, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  suppressMessages(print(p_Circularity_dotplot_new))
  insertPlot(wb, 3, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  ###### write Volume plots
  suppressMessages(print(p_Volume_new))
  insertPlot(wb, 4, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  suppressMessages(print(p_Volume_dotplot_new))
  insertPlot(wb, 4, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  ###### write Perimeter plots
  suppressMessages(print(p_Perimeter_new))
  insertPlot(wb, 5, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  suppressMessages(print(p_Perimeter_dotplot_new))
  insertPlot(wb, 5, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  

  #########################
  #### Start writing ADCVP datasets into spreadsheets.


  #create formatting styles
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
  writeDataTable(wb, 11, Sph_Treat_Robz_ADVPC, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = TRUE)

  addStyle(wb, sheet = 11, datastyleC, rows = 2:100 , cols = c(1:7, 14, 15, 18, 24, 30, 36, 42), gridExpand = TRUE)
  addStyle(wb, sheet = 11, datastyleC, rows = 1 , cols = c(1:47), gridExpand = TRUE)

  setColWidths(wb, sheet = 11, cols = c(9:47), widths = 16 )
  setColWidths(wb, sheet = 11, cols = c(4:8), widths = 10 )
  setColWidths(wb, sheet = 11, cols = c(1:3), widths = 8 )

  freezePane(wb, sheet = 11,  firstActiveRow = 2)
  #  Save outlier analysis report ,
  
  #######################################
  
  #Sph_ADVPC <- Sph_Treat_ADVPC[,c(1:4, 53:57)]
  Sph_ADVPC <- Sph_Treat_ADVPC
  
  dropped_cols=setdiff(colnames(Sph_ADVPC), 
                       intersect(colnames(Sph_ADVPC), colnames(df_treat)))
  dropped_cols=c('T_I',dropped_cols)
  Sph_ADVPC = Sph_ADVPC[, dropped_cols]

  # merge with treatment data

  Sph_ADVPC  <- merge(Sph_ADVPC, df_treat, by= "T_I")
  
  Sph_ADVPC <- Sph_ADVPC[with(Sph_ADVPC, order(Row, Col)),]
  
  Job_Info_data <- select(Spheroid_data, Jobrun.Finish.Time)

  Sph_ADVPC$Job.Date <- as.Date(Job_Info_data$Jobrun.Finish.Time, origin="1900-01-01")
  
  #work out time diff in days ETST.

  Sph_ADVPC$ETST.d <-  difftime(Sph_ADVPC$Job.Date, Sph_ADVPC$Time_Date, units= "days")
  Sph_ADVPC$ETST.h <-  difftime(Sph_ADVPC$Job.Date, Sph_ADVPC$Time_Date, units= "hours")

  

  Sph_ADVPC$Filename <- rawfilename
  Sph_ADVPCa <-  Sph_ADVPC
  
  # calculate Means and SE's for A D C V P items
  
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
  colnames(a_allmeans3)<- c("T_I", "Area_Mean","Area_SE","Diameter_Mean","Diameter_SE","Circularity_Mean",
                            "Circularity_SE","Volume_Mean","Volume_SE", "Perimeter_Mean","Perimeter_SE" )

  #subset TI and treatment items from main dataset


  #merge final means dataset with Treatment data items

  
  ADVPC_means <- merge(a_allmeans3, df_treat, by= "T_I")

  job.date = Job_Info_data$Jobrun.Finish.Time[1]
  ADVPC_means$Job.Date <- job.date

  # calculate time diff in days and hours ETST.

  ADVPC_means$ETST.d <-  as.numeric(difftime(ADVPC_means$Job.Date, ADVPC_means$Time_Date, units= "days"))
  ADVPC_means$ETST.h <-  as.numeric(difftime(ADVPC_means$Job.Date, ADVPC_means$Time_Date, units= "hours"))
  ADVPC_means$Filename <- rawfilename



  ADVPC_means <- ADVPC_means[with(ADVPC_means, order(T_I)),]

  if (TF_copytomergedir == 'TRUE')
  {

    write.csv(ADVPC_means, file = merge_output_file, col.names = TRUE, row.names = FALSE, append = FALSE)

  }
  
  

  ###########end of write CSV Export file#########################

  ### start write Mean/SE  CSV data to report at tab 12
  ##############################################

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

  # ############################################
  # read plate setup from config file, copy into tab 13 of report
  # use df_setup which read from <layout file>.csv

  writeData(wb, 13, df_setup, startRow = 12, startCol = 1, rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)

  setColWidths(wb, sheet = 13, cols = 1:13, widths = 5)
  #setColWidths(wb, sheet = 13, cols = 1, widths = 9)
  addStyle(wb, sheet = 13, styleborderTBLR, rows= 13:20 , cols= 2:13, gridExpand = TRUE)
  #addStyle(wb, sheet = 13, datastyleC, rows = 12, cols = 2:13, gridExpand = TRUE)
  addStyle(wb, sheet = 13, greyshadeTBLR, rows= 12 , cols= 2:13, gridExpand = TRUE)
  addStyle(wb, sheet = 13, greyshadeTBLR, rows= 13:20 , cols= 1, gridExpand = TRUE)

  # #copy treatment Index definitions into tab 13 of report
  #
  writeData(wb, 13, df_treat, startRow = 12, startCol = 15, rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)

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
  #
  #
  #
  t1 <- "      Summary of parameters used for processing"

  p0 <- "       Raw data filename (.xlsx)"
  p1 <- "       Plate setup tab"
  p2 <- "       Robust z low limit"
  p3 <- "       Robust z high limit"
  p4 <- "       Apply Thresholds"
  p5 <- "       Apply outlier overrides"
  psep <- ":"

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
  if (TF_outlier_override == TRUE)
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

  if (TF_apply_thresholds == TRUE)
  {
    # Thresholdlayout <- read.xlsx("Sph_Config_File_Rev2.xlsx", sheet = "Pre-Screen", rows = c(1:3), cols = c(2:7),
    #                              colNames = TRUE)
    # Thresholdlayout_s <- select(Thresholdlayout, Area, Diameter, Volume, Perimeter, Circularity)
    Thresholdlayout_s <- data.frame(Area=c(TH_Area_max, TH_Area_min), Diameter=c(TH_Diameter_max,TH_Diameter_min),
                                    Volume=c(TH_Volume_max, TH_Volume_min), Perimeter=c(TH_Perimeter_max,TH_Perimeter_min),
                                    Circularity= c(TH_Circularity_max,TH_Circularity_min))

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
  return(wb)
}


manual_outlier_selections = outer(LETTERS[1:8], 1:12, FUN = "paste")
dim(manual_outlier_selections) =NULL
# Define UI for application that draws a histogram
ui <- navbarPage("SpheroidAnalyseR",

      # Application title
    tabPanel("Data Input",

    # Sidebar with a slider input for number of bins 
    sidebarLayout(
        sidebarPanel(
        fileInput(inputId = "raw_data","Choose Raw Data",accept=".xlsx",buttonLabel = "Browse"
        ),
        fileInput(inputId = "layout","Choose Plate Layout",accept=".csv",buttonLabel="Browse"
        ),
        fileInput(inputId = "treat_data","Choose Treatment Definitions",accept=".csv",buttonLabel="Browse"
        )
        ),
        # Show a plot of the generated distribution
        mainPanel(
           h2("Preview the data"),           
           tableOutput("data_preview"), 
           h2("Show the treatment groups on the plate"),
           selectInput("select_treatment",label="Choose the value to view the layout",
                       list()),
           plotOutput("show_layout")
        )
    )
    ),
    tabPanel("Outliers",
             sidebarLayout(
                 sidebarPanel(
                     textInput("z_low","Robust z-score low limit",value = -1.96),
                     textInput("z_high","Robust z-score high limit",value = 1.96),
                     checkboxInput("pre_screen","Apply Pre-screen thresholds?",value = TRUE),
                     checkboxInput("override","Apply Manual overrides?",value = FALSE),
                     conditionalPanel(condition ="input.pre_screen==1",
                        #- Values above zero
                        numericInput("area_threshold_low","Area lower limit",value=1,min=0),
                        numericInput("area_threshold_high","Area higher limit",value=1630000,min=0),
                        numericInput("diam_threshold_low","Diameter lower limit",value=100,min=0),
                        numericInput("diam_threshold_high","Diameter higher limit",value=1440,min=0),
                        numericInput("vol_threshold_low","Volume lower limit",value=1000,min=0),
                        numericInput("vol_threshold_high","Volume higher limit",value=1440,min=0),
                        numericInput("perim_threshold_low","Perimeter lower limit",value=100,min=0),
                        numericInput("perim_threshold_high","Perimeter higher limit",value=6276,min=0),
                        numericInput("circ_threshold_low","Circularity lower limit",value=0.01,min=0),
                       numericInput("circ_threshold_high","Circularity higher limit",value=1,min=0)
                     )
                     ,
                     #
                     conditionalPanel(condition="input.override==1",
                                       selectInput("area_manual_outliers",label="Area Outliers to be removed",
                                                   manual_outlier_selections,
                                                   multiple=TRUE,selectize=TRUE),
                                      
                                      selectInput("diameter_manual_outliers",label="Diameter Outliers to be removed",
                                                  manual_outlier_selections,
                                                  multiple=TRUE,selectize=TRUE),
                                      selectInput("volume_manual_outliers",label="Volume Outliers to be removed",
                                                  manual_outlier_selections,
                                                  multiple=TRUE,selectize=TRUE),
                                      selectInput("perimeter_manual_outliers",label="Perimeter Outliers to be removed",
                                                  manual_outlier_selections,
                                                  multiple=TRUE,selectize=TRUE),
                                      selectInput("circularity_manual_outliers",label="Circularity Outliers to be removed",
                                                  manual_outlier_selections,
                                                  multiple=TRUE,selectize=TRUE)
                     )
                     ,
                     actionButton("outlier_btn", "Remove outliers"),
                     downloadButton("downloadData_btn", "Download"),
                     textOutput("textStatus")
                 ),
                 # Show a plot of the generated distribution
                 mainPanel(
                    
                     h2("Plate layout after pre-sreen outlier removal (if applied)"),
                  
                     selectInput("select_outlier_values",label="Choose the value to view outliers",
                                 list("Area", "Diameter", "Volume", "Perimeter" , "Circularity")),
                     
                     plotOutput("outlierPlot"),
                     
                     h2("Plate layout after robust Z-score outlier removal"),
                     
                     selectInput("select_z_scores",label="Choose the value to view result plots",
                                 list("Area", "Diameter", "Volume", "Perimeter" , "Circularity")),
                     
                     plotOutput("resultPlot"),
                     plotOutput("periPlot")
                     
                 )
             ) 
    )
)


# Define server logic required to draw a histogram
server <- function(input, output,session) {

    df_output <- NULL
    df_outliers <- NULL
    global_wb <-NULL
     
    df_plot_treat <-NULL
    df_plot_layout <-NULL
    
    output$data_preview <- renderTable(
        {
        data <- readRawData()
        if(!is.null(data))
            head(data, n=3)
        }

    )
    

    readRawData <- reactive({
        
        file <- input$raw_data

        
        if(!is.null(file)){
            ext <- tools::file_ext(file$datapath)
            req(file)
            validate(need(ext == "xlsx","Please upload an xlsx file"))
            data <- readxl::read_xlsx(file$datapath)
        }

    })
    
    readLayout <- reactive({
        layout_file <- input$layout
        
        if(!is.null(layout_file)){
            ext <- tools::file_ext(layout_file$datapath)
            req(layout_file)
            validate(need(ext == "csv","Please upload a layout in csv format"))
            layout <- readr::read_csv(layout_file$datapath)  %>% 
            #     #layout <- readr::read_csv("layout_example.csv") %>% 
                dplyr::rename("Row"=X1) %>%
                tidyr::pivot_longer(-Row,names_to="Col",
                                    values_to="Index") %>%
                mutate(Index = as.factor(Index),Col=as.numeric(Col)) %>%
                filter(!is.na(Row))
            
            df_plot_layout <<- layout
            layout
        }
    }
    )
    
    readTreatment <- reactive({
        treat_file <- input$treat_data
        ## TO DO: Need to check that 1st column is named Index
        if(!is.null(treat_file)){
            ext <- tools::file_ext(treat_file$datapath)
            req(treat_file)
            validate(need(ext == "csv","Please upload a layout in csv format"))
            treatments <- readr::read_csv(treat_file$datapath) %>% 
                mutate_all(as.factor)
            # update the value for treatment
            selections = colnames(treatments)
            updateSelectInput(session, "select_treatment",
                              choices = selections,
                              selected = head(selections, 1))
            
            df_plot_treat<<-treatments
            treatments
            
        }
    }
    )
    
    draw_layout<-function(df_layout){
      df_layout$Well.Name = paste0(df_layout$Row,df_layout$Col)
      ggplot(df_layout, aes(x = Row, y = Col, label=Well.Name)) +
        geom_tile() +
        geom_text() +
        scale_y_continuous(breaks=1:12)
    }
    draw_layout_with_treat <-function(df_layout, df_treat, value){
      layout <- left_join(df_layout,df_treat)
      layout$Well.Name = paste0(layout$Row,layout$Col)
      
      ggplot(layout, aes_string(x = "Row", y = "Col",fill=value,label="Well.Name")) +
        geom_tile() +
        scale_y_continuous(breaks=1:12)+ geom_text()
      
    }
    output$show_layout <- renderPlot(
      {
        if(!is.null(df_plot_layout)){
          if(!is.null(df_plot_treat)){
            draw_layout_with_treat(df_plot_layout, df_plot_treat, input$select_treatment)
          }
        else{
          draw_layout(df_plot_layout)
        }

        }
      }
    )
    
    
    observeEvent(input$layout, {
      
      layout_file <- input$layout
      
      if(!is.null(layout_file)){
        ext <- tools::file_ext(layout_file$datapath)
        req(layout_file)
        validate(need(ext == "csv","Please upload a layout in csv format"))
        layout <- readr::read_csv(layout_file$datapath)  %>% 
          #     #layout <- readr::read_csv("layout_example.csv") %>% 
          dplyr::rename("Row"=X1) %>%
          tidyr::pivot_longer(-Row,names_to="Col",
                              values_to="Index") %>%
          mutate(Index = as.factor(Index),Col=as.numeric(Col)) %>%
          filter(!is.na(Row))
        
        df_plot_layout <<- layout
        layout
      }
      
      output$show_layout <- renderPlot(
        {
          if(!is.null(df_plot_layout)){
            if(!is.null(df_plot_treat)){
              draw_layout_with_treat(df_plot_layout, df_plot_treat, input$select_treatment)
            }
            else{
              draw_layout(df_plot_layout)
            }
            
          }
        }
      )
      
    })
    
    observeEvent(input$treat_data, {
      treat_file <- input$treat_data
      ## TO DO: Need to check that 1st column is named Index
      if(!is.null(treat_file)){
        ext <- tools::file_ext(treat_file$datapath)
        req(treat_file)
        validate(need(ext == "csv","Please upload a layout in csv format"))
        treatments <- readr::read_csv(treat_file$datapath) %>% 
          mutate_all(as.factor)
        # update the value for treatment
        selections = colnames(treatments)
        updateSelectInput(session, "select_treatment",
                          choices = selections,
                          selected = head(selections, 1))
        
        df_plot_treat<<-treatments
        treatments
        
      }
      
      output$show_layout <- renderPlot(
        {
          if(!is.null(df_plot_layout)){
            if(!is.null(df_plot_treat)){
              draw_layout_with_treat(df_plot_layout, df_plot_treat, input$select_treatment)
            }
            else{
              draw_layout(df_plot_layout)
            }
            
          }
        }
      )
      
    })
    
   
    output_report <- eventReactive(input$outlier_btn, {

      
      validate(
        need(is.numeric(input$area_threshold_low) & input$area_threshold_low>0, "Please input a positive area threshold (numeric)"),
        need(is.numeric(input$area_threshold_high) & input$area_threshold_high>0, "Please input a positive area threshold (numeric)"),
        
        need(is.numeric(input$diam_threshold_low) & input$diam_threshold_low>0, "Please input a positive diameter threshold (numeric)"),
        need(is.numeric(input$diam_threshold_high) & input$diam_threshold_high>0, "Please input a positive diameter threshold (numeric)"),
        
        need(is.numeric(input$vol_threshold_low) & input$vol_threshold_low>0, "Please input a positive volume threshold (numeric)"),
        need(is.numeric(input$vol_threshold_high) & input$vol_threshold_high>0, "Please input a positive diameter threshold (numeric)"),
        
        need(is.numeric(input$perim_threshold_low) & input$perim_threshold_low>0, "Please input a positive Perimeter threshold (numeric)"),
        need(is.numeric(input$perim_threshold_high) & input$perim_threshold_high>0, "Please input a positive Perimeter threshold (numeric)"),
        
        need(is.numeric(input$circ_threshold_low) & input$circ_threshold_low>0, "Please input a positive Circularity threshold (numeric)"),
        need(is.numeric(input$circ_threshold_high) & input$circ_threshold_high>0, "Please input a positive Circularity threshold (numeric)"),
        
        need(input$raw_data, "please upload the raw data"),
        need(input$layout, "please upload the layout file"),
        need(input$treat_data, "please upload the treatment file")
      )
      

      
      #- after validating, use the thresholding to remove outliers
      #- do the thresholding and generate the wells

      #- read from the input
      rawfilename = input$raw_data$datapath
      platesetupname<- input$layout$datapath
      treat_file <- input$treat_data
      # 
      # 
      df_setup =read.csv(platesetupname)
      df_treat = read.csv(treat_file$datapath)
      #- import raw data
      Spheroid_data <- suppressMessages(read_excel(rawfilename, "JobView",col_names=TRUE, .name_repair = "universal"))
      
      
      # setup_file = "layout_example.csv"
      # treatment_file = "treatment_example.csv"
      # 
      # input_file = "GBM58 - for facs - day 3.xlsx"
      # 
      # df_setup =read.csv(setup_file)
      # df_treat = read.csv(treatment_file)
      # #- import raw data
      # Spheroid_data <- suppressMessages(read_excel(input_file, "JobView",col_names=TRUE, .name_repair = "universal"))
      
      
      
      RobZ_LoLim <- input$z_low
      RobZ_UpLim <- input$z_high
      
      TF_apply_thresholds <- input$pre_screen
      TF_outlier_override <- input$override
      TF_copytomergedir <- FALSE
      

      merge_output_file = "temp_merge_file.xlsx"
      report_output_file = "report_file.xlsx"
      

      colnames(df_treat)[1] <- "T_I"
      
      
      Ccount <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      for (j in 1:length(Ccount)){
        Ccount[j] <- length(which(df_setup[,1+j] >= 1))
      }
      checksum <- sum(Ccount)
      
      
      df_lookup <-  data.frame("Row"= 1:96, "Col" = 1:96,"T_I" =0)
      
      for (i_row in 1:8)
      {
        for (i_col in 1:12) 
        {
          idx = ((i_row-1)*12)+i_col
          
          
          df_lookup$Row[idx]<- df_setup[i_row,1]
          df_lookup$Col[idx] <- i_col
          
          df_lookup$T_I[idx] =df_setup[i_row , i_col+1]
          
        }
        
      }
      
      df_merge <- merge(df_lookup,df_treat, by="T_I") 
      

      datalength <- nrow(Spheroid_data)
      checksumOK <- "TRUE"
      
      if (checksum != datalength)  
      {checksumOK <- "FALSE"} 
      
      #- do the check
      if (checksumOK == "FALSE")
      {
        message("#### Warning : plate setup checksum mismatch. ####","\n")
        message("Number of rows of data in raw file (",datalength  ,") does not match the plate setup checksum (",checksum,") \n")
        message("Output file corruption may occur - please recheck your plate setup","\n")
        message("Processing halted","\n")
        message("\n")
        sink()
        stop()
      }
      #- do the check
      if (checksumOK == "TRUE")
        
      {
        message("** Plate setup checksum is valid for raw dataset ","\n")
        message("Number of rows of data in raw file (",datalength  ,") matches the checksum (",checksum,") \n")
      }
      
      
      
      # extract job information 
      Job_Info_data <- select(Spheroid_data, Project.Folder, Project.ID,	Project.Name,	Jobdef.ID, Jobdef.Name, Jobrun.Finish.Time, 
                              Jobrun.Folder, Jobrun.ID, Jobrun.Name, Jobrun.UUID)
      #Job_Info_data <- Job_Info_data[1:1,]
      
      
      # extract relevant columns and create new dataset Spheroid_data_1
      
      Spheroid_data_1 <- select(Spheroid_data, Well.Name, Spheroid_Area.TD.Area, Spheroid_Area.TD.Perimeter.Mean,
                                Spheroid_Area.TD.Circularity.Mean, Spheroid_Area.TD.Count, Spheroid_Area.TD.EqDiameter.Mean, Spheroid_Area.TD.VolumeEqSphere.Mean)
      
      
      colnames(Spheroid_data_1) <- c("Well.Name", "Area", "Perimeter", "Circularity","Count", "Diameter", "Volume"  )
      
      
      # replace any zero values with null value to calculate median values properly
      
      Spheroid_data_1$Area <- ifelse(Spheroid_data_1$Area==0, NA, Spheroid_data_1$Area) 
      Spheroid_data_1$Perimeter <- ifelse(Spheroid_data_1$Perimeter==0, NA, Spheroid_data_1$Perimeter)
      Spheroid_data_1$Circularity <- ifelse(Spheroid_data_1$Circularity==0, NA, Spheroid_data_1$Circularity)
      Spheroid_data_1$Diameter <- ifelse(Spheroid_data_1$Diameter==0, NA, Spheroid_data_1$Diameter)
      Spheroid_data_1$Volume <- ifelse(Spheroid_data_1$Volume==0, NA, Spheroid_data_1$Volume)
      
      Spheroid_data_1$Area<- as.numeric(Spheroid_data_1$Area)
      
      
      if (TF_apply_thresholds == TRUE) {
        TH_Area_min <- input$area_threshold_low
        TH_Area_max <- input$area_threshold_high
        
        TH_Diameter_min <- input$diam_threshold_low
        TH_Diameter_max <- input$diam_threshold_high
        
        TH_Volume_min <- input$vol_threshold_low
        TH_Volume_max <- input$vol_threshold_high
        
        TH_Perimeter_min <- input$perim_threshold_low
        TH_Perimeter_max <- input$perim_threshold_high
        
        TH_Circularity_min <- input$circ_threshold_low
        TH_Circularity_max <- input$circ_threshold_high
        
        Spheroid_data_1$Area <- ifelse(Spheroid_data_1$Area < TH_Area_max & Spheroid_data_1$Area > TH_Area_min,  Spheroid_data_1$Area, NA)   
        Spheroid_data_1$Diameter <- ifelse(Spheroid_data_1$Diameter<TH_Diameter_max & Spheroid_data_1$Diameter>TH_Diameter_min, Spheroid_data_1$Diameter, NA)
        Spheroid_data_1$Circularity <- ifelse(Spheroid_data_1$Circularity<TH_Circularity_max & Spheroid_data_1$Circularity>TH_Circularity_min, Spheroid_data_1$Circularity, NA)
        Spheroid_data_1$Volume <- ifelse(Spheroid_data_1$Volume<TH_Volume_max & Spheroid_data_1$Volume>TH_Volume_min, Spheroid_data_1$Volume, NA)
        Spheroid_data_1$Perimeter <- ifelse(Spheroid_data_1$Perimeter<TH_Perimeter_max & Spheroid_data_1$Perimeter>TH_Perimeter_min, Spheroid_data_1$Perimeter, NA)
        
      }   
      
      # split well name into row(character) and column(number)
      
      well_1ch <- substr(Spheroid_data_1$Well.Name, start = 1, stop = 1)
      well_2no <- as.numeric(substr(Spheroid_data_1$Well.Name, start = 2, stop = 3))
      
      
      #add 2 cols to dataset for well column / row 
      
      Spheroid_data_1$Well.row <- well_1ch
      Spheroid_data_1$Well.col <- well_2no
      
      Spheroid_data_1$Row <- well_1ch
      Spheroid_data_1$Col <- well_2no
      
      #- sorted by well Charcter and number
      Spheroid_data_1a <- Spheroid_data_1[with(Spheroid_data_1, order(Well.row, Well.col)),]
      df_sph_treat <- merge(Spheroid_data_1a,df_merge, by=c("Row", "Col"))
      df_tph_treat <- df_sph_treat[with(df_sph_treat, order(Well.row, Well.col)),]
      

      ######### data_summary - and outlier removing
      ######### helper function used to calculate the mean and the standard dev, for the variable of interest, in each group :
      
      df_a = cal_z_score(df_sph_treat, df_sph_treat, "Area", RobZ_UpLim = RobZ_UpLim,RobZ_LoLim = RobZ_LoLim)
      df_d = cal_z_score(df_sph_treat,df_a, "Diameter", RobZ_UpLim = RobZ_UpLim,RobZ_LoLim = RobZ_LoLim)
      df_v = cal_z_score(df_sph_treat,df_d, "Volume", RobZ_UpLim = RobZ_UpLim,RobZ_LoLim = RobZ_LoLim)
      df_p = cal_z_score(df_sph_treat,df_v, "Perimeter", RobZ_UpLim = RobZ_UpLim,RobZ_LoLim = RobZ_LoLim)
      
      Sph_Treat_Robz_ADVPC =cal_z_score(df_sph_treat,df_p,"Circularity" , RobZ_UpLim = RobZ_UpLim,RobZ_LoLim = RobZ_LoLim)
      
      
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
      
      ## assign to df_output
      df_output <<- Sph_Treat_Robz_ADVPC
      
      ######################################################################
      ###########  START REPORT / PLOTTING 
      ######################################################################
      
      Sph_Treat_Robz_ADVPC$T_I <- as.numeric(Sph_Treat_Robz_ADVPC$T_I)
      Sph_Treat_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(T_I)),]


      ############################################
      ## Generate plots and Write into spreadsheet
      
      #- ploting config
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
      
      
      p_Area_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Area_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                           geom = "crossbar", width = 0.25)  + 
        ylab(Arealabel) +xlab('Treatment Index') +  labs(title= "Area  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))
      
      
      
      ########################### Diameter PLots
      p_Diameter_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Diameter, colour = Diameter_status), size = Pointsize,  show.legend=FALSE ) + 
        # geom_point(data = Sph_Data_Out, aes(x= Col, y=Diameter_OO) , size = OutlierPointSize, colour = OutlierPointcolour ) +ylab(' Diameter  (um^2)')   xlab('Column') +
        scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
        facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
        #   ylab(Arealabel) +xlab('Column') +  labs(title= " Area data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))
        ylab(Diameterlabel)+xlab('Column') +  labs(title= " Diameter data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
        scale_colour_manual(values = cols) 
      
      p_Diameter_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Diameter_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                                   geom = "crossbar", width = 0.25)  + 
        ylab(Diameterlabel) +xlab('Treatment Index') +  labs(title= "Diameter  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))
      
      
      ########################### Circularity PLots
      
      p_Circularity_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Circularity, colour = Circularity_status), size = Pointsize,  show.legend=FALSE ) + 
        scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
        
        facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
        ylab('Circularity')+xlab('Column') +  labs(title= " Circularity data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
        scale_colour_manual(values = cols) 
      
      p_Circularity_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Circularity_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                                         geom = "crossbar", width = 0.25)  + 
        ylab('Circularity') +xlab('Treatment Index') +  labs(title= "Circularity  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))
      
      
      
      
      ########################### Volume PLots
      
      p_Volume_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Volume, colour = Volume_status), size = Pointsize,  show.legend=FALSE ) + 
        scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
        
        facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
        ylab(Volumelabel)+xlab('Column') +  labs(title= " Volume data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
        scale_colour_manual(values = cols) 

      
      p_Volume_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Volume_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                               geom = "crossbar", width = 0.25)  + 
        ylab(Volumelabel) +xlab('Treatment Index') +  labs(title= "Volume  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))
      
      ########################### Perimeter PLots
      
      p_Perimeter_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Perimeter, colour = Perimeter_status), size = Pointsize,  show.legend=FALSE ) + 
        scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
        
        facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
        ylab(Perimeterlabel)+xlab('Column') +  labs(title= " Perimeter data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))+ 
        scale_colour_manual(values = cols) 

      p_Perimeter_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Perimeter_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray")) + 
        stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,geom = "crossbar", width = 0.25)  + 
        ylab(Perimeterlabel) +xlab('Treatment Index') +  labs(title= "Perimeter  : Mean +/- SE  (outliers removed)") + theme_bw() + theme(plot.title = element_text(size = 10))
      
      
     
      
      global_wb <<- gen_report(Sph_Treat_Robz_ADVPC=Sph_Treat_Robz_ADVPC,
                               df_treat = df_treat,df_setup = df_setup,
                               Spheroid_data=Spheroid_data,
                               p_Area_new = p_Area_new, p_Area_dotplot_new = p_Area_dotplot_new,
                               p_Diameter_new =p_Diameter_new, p_Diameter_dotplot_new =p_Diameter_dotplot_new,
                               p_Circularity_new = p_Circularity_new, p_Circularity_dotplot_new = p_Circularity_dotplot_new,
                               p_Volume_new = p_Volume_new, p_Volume_dotplot_new = p_Volume_dotplot_new,
                               p_Perimeter_new= p_Perimeter_new , p_Perimeter_dotplot_new = p_Perimeter_dotplot_new,
                               rawfilename = rawfilename ,platesetupname=platesetupname,
                               RobZ_LoLim = RobZ_LoLim, RobZ_UpLim = RobZ_UpLim,
                               TF_apply_thresholds = TF_apply_thresholds,
                               TF_outlier_override = TF_outlier_override,
                               TH_Area_max=TH_Area_max,TH_Area_min=TH_Area_min,
                               TH_Diameter_max=TH_Diameter_max,TH_Diameter_min=TH_Diameter_min,
                               TH_Volume_max=TH_Volume_max, TH_Volume_min=TH_Volume_min,
                               TH_Perimeter_max = TH_Perimeter_max, TH_Perimeter_min=TH_Perimeter_min,
                               TH_Circularity_max = TH_Circularity_max, TH_Circularity_min=TH_Circularity_min
                               )
      
      message("All processing completed successfully.", "\n")
      message("\n") 
      
      
      
      # p_list = list(p_Area_new,p_Area_dotplot_new, 
      #             p_Diameter_new,p_Diameter_dotplot_new,
      #             p_Circularity_new,p_Circularity_dotplot_new,
      #             p_Volume_new,p_Volume_dotplot_new,
      #             p_Perimeter_new,p_Perimeter_dotplot_new)
      "Report generated, please download the report"
    })
    
    output$textStatus<-  renderText({
      output_report()
    })

      output$outlierPlot <- renderPlot({
        validate(need(df_output,"Please run the outlier removal"))
          message(df_output)
          df_outlier_temp = df_output
          draw_outlier_plot(df_outlier_temp,input$select_outlier_values)
      })
      
      
      output$resultPlot <- renderPlot({
        validate(need(df_output,"Please run the outlier removal"))
        Sph_Treat_ADVPC = df_output
        draw_z_score_outlier_plot(Sph_Treat_ADVPC,input$select_z_scores)
      
        
      })
      
      ### event from clicking the remove outlier button
      ### draw layout 
      observeEvent(input$outlier_btn, {
        output_report()
        
        output$outlierPlot <- renderPlot({
          df_outlier_temp = df_output
          draw_outlier_plot(df_outlier_temp,input$select_outlier_values)

        })

        
        output$resultPlot <- renderPlot({
          message(typeof(df_output))
          Sph_Treat_ADVPC = df_output
          draw_z_score_outlier_plot(Sph_Treat_ADVPC,input$select_z_scores)
        })
        
      })
      
 
    ### Download report
    output$downloadData_btn <- downloadHandler(

      filename = function() {
        paste( "report.xlsx", sep = "")
      },
      content = function(file) {
        # write.csv(df_output, file, row.names = FALSE)
        saveWorkbook(global_wb, file, overwrite = TRUE)
        
      }
    )
       
}

# Run the application 
shinyApp(ui = ui, server = server)
