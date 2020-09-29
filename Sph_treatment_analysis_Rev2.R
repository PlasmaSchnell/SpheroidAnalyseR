################### R CODE INFORMATION   ##################################
#                                                                         #
# filename :        Sph_treatment_analysis_rev2.R                         #
# Revision :        2                                                     #      
# Revision Date :   20-Nov-2019                                           #
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
library(ggthemes)
library(gridExtra)
library(readxl)
library(openxlsx)
library(plotrix)
library(writexl)
library(ggpubr)
library(varhandle)
library(ggpubr)
library(ggsignif)

########## read configuration file "sph_config_file.xlsx" ##############
#sink()
sink("NUL")

slash = "//"

message("\n", "\n")

suppressMessages(reportdata <- read_excel("Sph_Config_File_Rev2.xlsx", sheet = "Treatment analysis", col_names = TRUE, .name_repair = "universal"))

mergefilename <- paste(reportdata$Input[1],".xlsx", sep="")
filepath <- reportdata$Input[3] 
mergefilename2 <- paste(filepath, slash ,mergefilename, sep="")

message("Reading data........", mergefilename2, "\n")


p_colour<-reportdata$plot_options[1]
p_theme <- reportdata$plot_options[2]

#inputs for plot colour scheme and style
if (p_theme == "Clean") { th_c <- theme_classic()} else  { th_c <- theme_bw()} 
if (p_colour == "Greyscale") { scalecol <- scale_colour_grey(start=0.8, end=0.2)} else  { scalecol <- scale_colour_hue()} 
if (p_colour == "Greyscale") { scalefill <- scale_fill_grey(start = 0.8 , end = 0.2)} else  { scalefill <- scale_fill_hue()}

#Title font size  
TF_size <- 10
#Title align (0 = L, 0.5 = C, 1 = R)
Tadj = 0.5

TF_Detailplot <- reportdata$plot_options[6]
TF_Quicklook <- reportdata$plot_options[15]
p_start <-as.numeric(reportdata$plot_options[7])
p_end <- as.numeric(reportdata$plot_options[8])
p_min <- as.numeric(reportdata$plot_options[9])
p_max <- as.numeric(reportdata$plot_options[10])

# check if overwriting Treatment Analysis Report (TAR) file which already exists

reportfilename <- paste(reportdata$Input[5],".xlsx", sep="")
reportfilename2 <- paste(reportdata$Input[5],"(2)", ".xlsx", sep="")
filepath <- reportdata$Input[7] 
TreatmentReportfilename <- paste(filepath, slash ,reportfilename, sep="")

TF_check_TAR_file <- 'FALSE'

if (file.exists(TreatmentReportfilename))
{ 
  
  TF_check_TAR_file <- 'TRUE'

  TreatmentReportfilename <- paste(filepath, slash ,reportfilename2, sep="")
  
  Msg_TARfile_duplicated_1 <- paste("Merged Treatment Analysis Report file already exists - saving a duplicate : ", reportfilename2,sep="") 
  Msg_TARfile_duplicated_2 <-'Enter a new Treatment Analysis Report filename in the config file and run code again, or rename the duplicate file.'
}




######################################################################
###########  START REPORT / PLOTTING  using Merged - Mean + SE tab
######################################################################

#read merged Mean +SE dataset from directory location + filename
full_merged <- read_excel(mergefilename2,  "Merged - Mean + SE",col_names = TRUE, .name_repair = "universal")

#read full merged dataset (all data) from directory location + filename
full_merged_ALL <- read_excel(mergefilename2,  "Merged - All data",col_names = TRUE, .name_repair = "universal")


# create Treatment workbook  (T_wb) 

T_wb <- createWorkbook()

if (TF_Quicklook=='TRUE') {addWorksheet(T_wb, 'Quicklook', gridLines = FALSE)} 
addWorksheet(T_wb, 'Area Bar', gridLines = FALSE) 
addWorksheet(T_wb, 'Area Line', gridLines = FALSE)
addWorksheet(T_wb, 'Area Line EB', gridLines = FALSE)
addWorksheet(T_wb, '%Area Bar', gridLines = FALSE)
addWorksheet(T_wb, 'Detail %Area Bar', gridLines = FALSE)
addWorksheet(T_wb, 'Conc_1 Area Bubble', gridLines = FALSE)
addWorksheet(T_wb, 'Data', gridLines = FALSE)
addWorksheet(T_wb, 'TestPlot', gridLines = FALSE)

######################


############### PLots #####################
#general items (labels, titles, etc) for all plots


Areaplottitle <- 'Area vs Elapsed Time Since Treatment (days)'
Detailtitle <- 'Detail plot % Area (w.r.t. 100% at 0mM conc.) vs E.T.S.T. (d)' 
PCplottitle <-  '% Area (w.r.t. 100% at 0mM conc.) vs E.T.S.T. (d)'
Bubbletitle <- 'Drug concentration vs E.T.S.T.(d), Area as point size'

conc_legendtitle <- 'mM drug1'
Arealegend <- "Mean Area"

Conc_label <- expression("Drug concentration" ~ (mu~M))
xlabel_ETSTd <- "Elapsed Time Since Treatment (days)"
Arealabel = expression("Area" ~ (mu~m^{2} ))
Diameterlabel = expression("Diameter" ~ (mu~m ))
Volumelabel = expression("Volume" ~ (mu~m^{3} ))
Perimeterlabel = expression("Perimeter" ~ (mu~m ))
pc_Arealabel = 'Percentage Area (100% at 0mM conc)'


#prepare data: change ETST.d and drug concentration to factors

#full_merged$ETST.d <- as.factor(full_merged$ETST.d)

full_merged_A <- full_merged

full_merged$ETST.d <- as.numeric(full_merged_A$ETST.d)
full_merged$Conc_1 <- as.factor(full_merged_A$Conc_1)
full_merged$Conc_2 <- as.factor(full_merged_A$Conc_2)
full_merged$Conc_3 <- as.factor(full_merged_A$Conc_3)

#order data by ETST.d

full_merged_B <- full_merged
full_merged  <- full_merged_B[with(full_merged_B , order(ETST.d, T_I)),]
full_merged$ETST.d <- as.factor(full_merged$ETST.d)





#############################
# PLOT 1 : histogram : Area vs ETST.d
message("Generating plot 1......", "\n")

suppressMessages( p_Area1 <-ggplot(data=full_merged, aes(x= ETST.d, y=Area_Mean, fill=Conc_1)) +
  geom_bar(stat="identity", position=position_dodge()) +theme_bw() +  xlab(xlabel_ETSTd) + ylab(Arealabel) + scale_fill_discrete(name = conc_legendtitle) + 
  geom_errorbar(aes(ymin=Area_Mean, ymax=Area_Mean + Area_SE), width= 0.5, size=0.3,  position=position_dodge(.9))+ theme(legend.position="right") +scalefill + th_c )
  
    p_Area1 <-  p_Area1 + ggtitle(Areaplottitle) + theme( plot.title = element_text(hjust = Tadj, size = TF_size))  

    print(p_Area1)
insertPlot(T_wb, sheet="Area Bar" , xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")
if (TF_Quicklook=='TRUE') {insertPlot(T_wb, sheet="Quicklook" , xy = c("A", 1), width = 6, height = 3,  fileType = "png", units = "in")}

#############################
# PLOT 2 : line plot Area vs ETST.d 
message("Generating plot 2......", "\n")


p_Area2 <- ggplot(data=full_merged, aes(x=as.factor(ETST.d), y=Area_Mean, group=Conc_1)) + 
    geom_line(aes( color=Conc_1, linetype= Conc_1)) +  geom_point(aes(color=Conc_1)) +
    xlab(xlabel_ETSTd) + ylab(Arealabel) + scalecol + th_c

   p_Area2 <-  p_Area2 + ggtitle(Areaplottitle) + theme( plot.title = element_text(hjust = Tadj, size = TF_size))

   
   print(p_Area2)
insertPlot(T_wb, sheet="Area Line" , xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")
if (TF_Quicklook=='TRUE') {insertPlot(T_wb, sheet="Quicklook" , xy = c("A", 19), width = 6, height = 3,  fileType = "png", units = "in")}



#############################
# PLOT 2a : line plot Area vs ETST.d (WITH ERROR BARS)
message("Generating plot 2a......", "\n")
pd <- position_dodge(0.05)


p_Area2a <- ggplot(data= full_merged, aes(x=ETST.d, y=Area_Mean, group = Conc_1, colour = Conc_1)) + 
  geom_errorbar(aes(ymin=Area_Mean-Area_SE, ymax=Area_Mean+Area_SE), width= 1, position=pd) +
  geom_line(position=pd, aes(linetype= Conc_1) ) + geom_point(aes(colour=Conc_1)) + xlab(xlabel_ETSTd) + ylab(Arealabel)+
  th_c + scalecol
  


    
 ### 
#  if (p_colour=="Greyscale") {p_Area2a<-p_Area2a+ scale_colour_grey(start = 0, end = .9) }
#if (p_theme == "Clean") {p_Area2a <- p_Area2a + theme_classic()} else {p_Area2a <- p_Area2a + theme_bw()}


p_Area2a <-  p_Area2a + ggtitle(Areaplottitle) + theme( plot.title = element_text(hjust = Tadj, size = TF_size))

 print(p_Area2a)
insertPlot(T_wb, sheet="Area Line EB" , xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")

#if (TF_Quicklook=='TRUE') {insertPlot(T_wb, sheet="Quicklook" , xy = c("A", 19), width = 6, height = 3,  fileType = "png", units = "in")}


#############################
# PLOT 3 : percentage Area histogram : Area vs ETST.d
message("Generating plot 3......", "\n")

full_merged_pc <- full_merged

#add column to full_merged_pc with percentages of Area at 0 dosage = 100%
#get list of variables conc and EDST
# calculate number of groups and rows in each group

l_Conc_1 <- unique(full_merged_pc$Conc_1)
l_ETST.d <-  unique(full_merged_pc$ETST.d)
drug_type <- unique(full_merged_pc$Drug_1)
len_full_merged_pc <- length(full_merged_pc$ETST.d)
l_C <- length(l_Conc_1)
l_E <- length(l_ETST.d)
cycles <- len_full_merged_pc/l_C
fixpc <- 1
full_merged_pc$Area_pc<- 1

legendtitle <- paste("mM", drug_type, sep = " ")

#full_merged_pc$Conc_1 <- as.character(full_merged_pc$Conc_1)

i<-1
j<-1

#FUnction with nested for loop to calculate Area at 100% for 0 mM concentration and Areas in each group relative to 100%. 

full_merged_pc$AreaSE_pc <-0

for (j in 1:l_E) 
{  
  for (i in 1:l_C)
  {
    h <- (j*l_C) +i - l_C
    
    if (full_merged_pc$Conc_1[h] == "0") 
    {fix100pc <- full_merged_pc$Area_Mean[h]
    full_merged_pc$Area_pc[h] <- 100 
    full_merged_pc$AreaSE_pc[h] <- (full_merged_pc$Area_SE[h] /full_merged_pc$Area_Mean[h])* full_merged_pc$Area_pc[h]
    
    }
    
    else {full_merged_pc$Area_pc[h] <- full_merged_pc$Area_Mean[h]/fix100pc*100
    full_merged_pc$AreaSE_pc[h] <- (full_merged_pc$Area_SE[h] /full_merged_pc$Area_Mean[h])* full_merged_pc$Area_pc[h]
    
    }        
    
  }
  
}


suppressMessages( p_Area3 <-ggplot(data=full_merged_pc, aes(x=as.factor(ETST.d), y=Area_pc, fill= Conc_1)) +
  geom_bar(stat="identity", position=position_dodge()) +theme_bw() + xlab(xlabel_ETSTd) + ylab(pc_Arealabel) + scale_fill_discrete(name = "mM TMZ") + 
  coord_cartesian(ylim=c(0,130)) + theme(legend.position="right")+ th_c + scalefill+
    geom_errorbar(aes(ymin=Area_pc, ymax=Area_pc + AreaSE_pc), width= 0.5,size= 0.3, position=position_dodge(.9)))
 
    p_Area3 <- p_Area3 + ggtitle(PCplottitle) + theme( plot.title = element_text(hjust = Tadj, size = TF_size))
    
  
   print(p_Area3)
insertPlot(T_wb, sheet='%Area Bar', xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")
if (TF_Quicklook=='TRUE') {
  
  
  insertPlot(T_wb, sheet="Quicklook" , xy = c("L", 1), width = 6, height = 3,  fileType = "png", units = "in")}
  

##############################
# PLot 4 zoom on percentage plot

if(TF_Detailplot == 'TRUE')  
   {
  message("Generating plot 4 ......", "\n")
  
subFM <- full_merged_pc
subFM$ETST.d <- unfactor(full_merged_pc$ETST.d)


newdata2 <- subset(subFM, ETST.d >= p_start & ETST.d <= p_end, select = c("ETST.d", "Area_pc", "AreaSE_pc", "Conc_1"))

#subsetpc$ETST.d <- as.factor(subsetpc$ETST.d)


suppressMessages(p_Area4 <-ggplot(data=newdata2, aes(x=as.factor(ETST.d), y=Area_pc, fill= Conc_1)) +
  geom_bar(stat="identity", position=position_dodge()) +theme_bw() + xlab(xlabel_ETSTd) + ylab(pc_Arealabel) + scale_fill_discrete(name = "mM TMZ") + 
  geom_errorbar(aes(ymin=Area_pc, ymax=Area_pc + AreaSE_pc), width= 0.3, size= 0.3, position=position_dodge(.9))+
  coord_cartesian(ylim=c(p_min,p_max)) + theme(legend.position="right") +scalefill +th_c )


p_Area4 <- p_Area4 + ggtitle(Detailtitle) + theme( plot.title = element_text(hjust = Tadj, size = TF_size))
  
print(p_Area4)
insertPlot(T_wb, sheet='Detail %Area Bar' , xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")
}

  if (TF_Quicklook=='TRUE' & TF_Detailplot == 'TRUE') {insertPlot(T_wb, sheet="Quicklook" , xy = c("L", 19), width = 6, height = 3,  fileType = "png", units = "in")}


##############################
# PLOT 5 bubble plot

message("Generating plot 5......", "\n")

p_Area5 <- ggplot(full_merged, aes(x=as.factor(ETST.d),y= Conc_1)) + geom_point(shape=21,colour= 'black', fill = 'lightgrey', aes( size = Area_Mean))+ 
  
  xlab(xlabel_ETSTd) + ylab(Conc_label) 

if (p_theme =="Clean") {p_Area5 <- p_Area5 + theme_classic()} else {p_Area5 <- p_Area5 + theme_bw()}

  p_Area5 <- p_Area5 + ggtitle(Bubbletitle) + theme( plot.title = element_text(hjust = Tadj, size = TF_size))

  print(p_Area5)
insertPlot(T_wb, sheet='Conc_1 Area Bubble' , xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")



########################################

#############################
# PLOT 6 Mean comparisons

#full_merged_ALL$ETST.d <- unfactor(full_merged_ALL$ETST.d)
full_merged_ALL$Conc_1 <- as.factor(full_merged_ALL$Conc_1)
full_merged_ALL$Conc_2 <- as.factor(full_merged_ALL$Conc_2)
full_merged_ALL$Conc_3 <- as.factor(full_merged_ALL$Conc_3)

#order data by ETST.d

full_merged_ALL_A <- full_merged_ALL
full_merged_ALL  <- full_merged_ALL_A[with(full_merged_ALL_A , order(ETST.d, T_I)),]


message("Generating plot 6......", "\n")


#subset the data to suit same as detail plot w/ p_start and p_end as limits

f_m_A <-full_merged_ALL

ss_f_m_A <- subset(f_m_A, ETST.d >= p_start & ETST.d <= p_end, select = c(ETST.d, Area_OR, Conc_1))

ss_f_m_A$ETST.d <- as.factor(ss_f_m_A$ETST.d)  


p_Area6 <- ggbarplot(ss_f_m_A, x = "ETST.d", y = "Area_OR",  add = c("mean_se"),fill = "Conc_1", position = position_dodge(0.8)) +  

  stat_compare_means(method = "anova", label.y = 40)+      # Add global p-value
  stat_compare_means(label = "p.signif", method = "t.test",
                     ref.group = "0")   
  
  
  
  
  # stat_compare_means(label = "p.signif", method = "t.test",  ref.group = "30") 
 #stat_compare_means(aes(group = Conc_1)), + comparisons = my_comparisons), label = "p.signif")

  #stat_compare_means(aes(group = Conc_1) +
  #stat_compare_means( comparisons = my_comparisons))+
  



# <- ggbarplot(ss_f_m_A, x = "ETST.d", y = "Area_OR",  add = c("mean_se", "jitter"), position = position_dodge(0.8)) +
#  stat_compare_means(label = "p.signif", method = "t.test",  ref.group = ".all.") 






#print(p_Area6)

#insertPlot(T_wb, sheet='TestPlot' , xy = c("B", 2), width = 12, height = 6,  fileType = "png", units = "in")


########################################










#write full merged dataset to "Data" tab

#  create formatting styles
datastyleC <- createStyle(fontSize = 10, fontColour = rgb(0,0,0),halign = "center", valign = "center")
datastyleR <- createStyle(fontSize = 10, fontColour = rgb(0,0,0), halign = "right", valign = "center")
datastyleL <- createStyle(fontSize = 10, fontColour = rgb(0,0,0), halign = "left", valign = "center")
styleBoldlarge <- createStyle(fontSize = 14, fontColour = "black",textDecoration = c("bold"))
styleCentre <- createStyle(halign = "center", valign = "center")
datestyle <- createStyle(numFmt = "hh:mm  dd-mmm-yyyyy")
dpstyle <- createStyle(halign = "right", valign = "center", numFmt = "0.0000")

writeDataTable(T_wb, sheet = "Data", full_merged, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = FALSE)
maxrows <-nrow(full_merged)+1


addStyle(T_wb, sheet = "Data", datestyle, rows = 2:maxrows , cols = c(22,23), gridExpand = TRUE)

addStyle(T_wb, sheet = "Data", dpstyle, rows = 2:maxrows , cols = c( 3:12,24:25), gridExpand = TRUE)

addStyle(T_wb, sheet = "Data", datastyleC, rows = 2:maxrows , cols = c(1,13:21, 24,25), gridExpand = TRUE)
addStyle(T_wb, sheet = "Data", datastyleC, rows = 1 , cols = c(1:26), gridExpand = TRUE)

addStyle(T_wb, sheet = "Data", datastyleR, rows = 2:maxrows , cols = c(3:12), gridExpand = TRUE)


setColWidths(T_wb, sheet = "Data", cols = c(3:12, 24:25), widths = 16 )
setColWidths(T_wb, sheet = "Data", cols = c(13:21), widths = 10 )
setColWidths(T_wb, sheet = "Data", cols = c(22, 23), widths = 24 )

setColWidths(T_wb, sheet = "Data", cols = 26, widths = 30 )
setColWidths(T_wb, sheet = "Data", cols = 2, widths = 20)


freezePane(T_wb, sheet = "Data",  firstActiveRow = 2)


#  Save outlier analysis report , 

saveWorkbook(T_wb, TreatmentReportfilename, overwrite = TRUE)

message("Report generated successfully : ",reportfilename, "\n")
#message(TreatmentReportfilename, "\n")
message("\n")


# check for overwriting existing file, message about duplicate filename if TRUE

if ( TF_check_TAR_file == "TRUE")
{
  message("\n")
  
  message("#### Treatment Analysis Report xlsx file overwrite notification ####", "\n")
  message(Msg_TARfile_duplicated_1)
  #message("\n")
  message(Msg_TARfile_duplicated_2)
  message("\n")
}


message("Files written successfully : ",reportfilename, "\n")


sink()