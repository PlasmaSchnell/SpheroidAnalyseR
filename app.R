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
library(shiny)
library(shinyjs)
library(DT)
####### Check file formats ######
check_layout = function(df){
  error_message=c("Check if the layout is 8x12 (row names and column names excluded)",
                  "Column names error, please refer to the template",
                  "Row names error, please refer to the template")
  
  dim_check = all(dim(df) == c(8,12))
  
  colnames_check = tryCatch( all(colnames(df) == as.character(1:12)), error = function(e){F})
  rownames_check = tryCatch( all(row.names(df) ==c("A", "B", "C", "D",
                                                   "E", "F", "G", "H")), error = function(e){F})
  type_check = tryCatch( all(sapply(df, typeof) == "integer" |sapply(df, typeof) == "logical"),
                         error = function(e){F})
  
  error_list = c(dim_check, colnames_check, rownames_check,type_check)
  return(error_list)
}


check_treatment = function(df){
  error_message=c("Check if the layout is 8x12 (row names and column names excluded)",
                  "Column names error, please refer to the template",
                  "Row names error, please refer to the template")
  
  dim_check = all(dim(df) == c(6,8))
  colnames_check = colnames_check = tryCatch( all(colnames(df) == c("Index", "Treatment.Label",
                                                                    "Time_Date", "Cell_line",
                                                                    "Passage_No", "Radiation_dosage",
                                                                    "Drug_1", "Conc_1"))
                                              , error = function(e){F})
  
  rownames_check = tryCatch(all(row.names(df) ==as.character(1:6)), error = function(e){F})
  
  type_check = tryCatch(all(sapply(df, typeof) == c("integer","character","character","character","integer","integer","character","integer")),
                        error = function(e){F})
  
  
  error_list = c(dim_check, colnames_check, rownames_check,type_check)
  return(error_list)
}



#######Predefine functions ######

cbind.fill <- function(...) {                                                                                                                                                       
  transpoted <- lapply(list(...),t)                                                                                                                                                 
  transpoted_dataframe <- lapply(transpoted, as.data.frame)                                                                                                                         
  return (data.frame(t(rbind.fill(transpoted_dataframe))))                                                                                                                          
} 
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

update_df_OR_by_status<- function(df){
  ### create Area, Diameter etc columns with oUtliers removed eg (Area_OR), both via RobZ and Manual Override
  df$Area_OR <- ifelse(df$Area_status=='0', df$Area, NA)
  df <- df[with(df, order(Row, Col)),]
  
  df$Diameter_OR <- ifelse(df$Diameter_status=='0', df$Diameter, NA)
  df <- df[with(df, order(Row, Col)),]
  
  df$Circularity_OR <- ifelse(df$Circularity_status=='0', df$Circularity, NA)
  df <- df[with(df, order(Row, Col)),]
  
  df$Volume_OR <- ifelse(df$Volume_status=='0', df$Volume, NA)
  df <- df[with(df, order(Row, Col)),]
  
  df$Perimeter_OR <- ifelse(df$Perimeter_status=='0', df$Perimeter, NA)
  df <- df[with(df, order(Row, Col)),]
  
  return(df)
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

  
  df[is.na(df[,value]), value] <- 2
  df[,value] = factor(df[,value])
  

  levels(df[,value]) = c(0,1,2)
  

  colours = c("1" = "red", "0" = "white", "2"='tan1')
  

  # colours = c(1 = "red", 0 = "white", NA='grey')
  
  
  ggplot(df, aes_string(x = "Row", y = "Col", fill = value,label="Well.Name")) +
    geom_tile() + 
    scale_y_continuous(breaks=1:12) + 
    scale_fill_manual(labels = c("Not an outlier", "Outlier determined by robust Z score", "Outliers determined from pre-set thresholds"),
                      values=colours, drop=FALSE) + geom_text()
}


draw_plot_1 = function(df, value){
  #- ploting config
  AreaPointcolour <- "black"
  PerimeterPointcolour <- "black"
  CircularityPointcolour <- "black"
  CircularityPointcolour <- "black"
  VolumePointcolour <- "black"
  OutlierPointcolour <- "red"
  Pointsize <- 1.5
  OutlierPointSize  <- 1.5
  
  cols <- c("1" = "red", "0" = "black")
  #cols <- c("Outlier" = "red", "Normal" = "black")
  
  ######### Area PLots
  y_value = value
  colour_value = paste0(value,"_status")
  
  ggplot()  + geom_point(data = df, aes_string(x= "Col", y=y_value, colour = colour_value), size = Pointsize,  show.legend=FALSE ) + 
    scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
    facet_wrap(~ df$T_I, nrow=1) +
    ylab(y_value)+xlab('Column') +  labs(title= paste0(y_value, " data by Treatment Index with outliers in red")) +
    theme_bw() + 
    theme(plot.title = element_text(size = 10))+ 
    scale_colour_manual(values = cols) 
}


draw_plot_2 = function(df, value){
  #- ploting config
  AreaPointcolour <- "black"
  PerimeterPointcolour <- "black"
  CircularityPointcolour <- "black"
  CircularityPointcolour <- "black"
  VolumePointcolour <- "black"
  OutlierPointcolour <- "red"
  Pointsize <- 1.5
  OutlierPointSize  <- 1.5
  
  cols <- c("1" = "red", "0" = "black")
  #cols <- c("Outlier" = "red", "Normal" = "black")
  
  ######### Area PLots
  y_value = paste0(value, '_OR')

  
  ggerrorplot(df, x= "T_I", y=y_value, desc_stat ="mean_se", 
              add = "dotplot", error.plot="errorbar", 
              addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + 
    stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                       geom = "crossbar", width = 0.25)  + 
    ylab(value) +
    xlab('Treatment Index') + 
    labs(title= paste0(value,"  : Mean +/- SE  (outliers removed)")) + theme_bw() +
    theme(plot.title = element_text(size = 10))
  
  
}

# help function for turning Sph_ADVPC to ADVPC_means
Sph_ADVPC_2_ADVPC_means <-function(Sph_ADVPC,df_treat,Spheroid_data,rawfilename){
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
  
  Sph_ADVPC$ETST.d <-  difftime(Sph_ADVPC$Job.Date, as.Date(Sph_ADVPC$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "days")
  Sph_ADVPC$ETST.h <-  difftime(Sph_ADVPC$Job.Date, as.Date(Sph_ADVPC$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "hours")
  
  
  
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
  
  ADVPC_means <- merge(a_allmeans3, df_treat, by= "T_I")
  
  job.date = Job_Info_data$Jobrun.Finish.Time[1]
  ADVPC_means$Job.Date <- job.date
  
  # calculate time diff in days and hours ETST.
  
  ADVPC_means$ETST.d <-  as.numeric(difftime(ADVPC_means$Job.Date, as.Date(ADVPC_means$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "days"))
  ADVPC_means$ETST.h <-  as.numeric(difftime(ADVPC_means$Job.Date, as.Date(ADVPC_means$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "hours"))
  ADVPC_means$Filename <- rawfilename
  
  # ADVPC_means$Job.Date <- as.POSIXct(ADVPC_means$Job.Date, format = "%Y-%m-%d %H:%M:%S")

  # ADVPC_means$Time_Date <- as.POSIXct(as.Date(ADVPC_means$Time_Date, "%m/%d/%Y", origin="1900-01-01"), format = "%Y-%m-%d %H:%M:%S")
  
  ADVPC_means <- ADVPC_means[with(ADVPC_means, order(T_I)),]
  return(ADVPC_means)
}

##### function for generating the report

gen_report = function(Sph_Treat_Robz_ADVPC,
                      df_treat,df_setup,
                      Spheroid_data,
                      df_origin,
                      # p_Area_new,p_Area_dotplot_new,
                      # p_Diameter_new ,p_Diameter_dotplot_new,
                      # p_Circularity_new,p_Circularity_dotplot_new,
                      # p_Volume_new, p_Volume_dotplot_new,
                      # p_Perimeter_new, p_Perimeter_dotplot_new,
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

  ######################################################################
  ###########  START REPORT / PLOTTING 
  ######################################################################
  
  Sph_Treat_Robz_ADVPC$T_I <- as.numeric(Sph_Treat_Robz_ADVPC$T_I)
  Sph_Treat_ADVPC <- Sph_Treat_Robz_ADVPC[with(Sph_Treat_Robz_ADVPC, order(T_I)),]
  
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
    ylab(Arealabel)+xlab('Column') +  labs(title= " Area data by Treatment Index with outliers in red") + 
    theme_bw() + theme(plot.title = element_text(size = 10))+ 
    scale_colour_manual(values = cols) 
  
  
  p_Area_dotplot_new <- ggerrorplot(Sph_Treat_ADVPC, x= "T_I", y="Area_OR", desc_stat ="mean_se", add = "dotplot", error.plot="errorbar", addparams= list(color="black", shape=21, binaxis="y", binwidth=2000, stackdir = "center", fill= "darkgray"))  + stat_summary(fun.y = mean, fun.ymin = mean, fun.ymax = mean,
                                                                                                                                                                                                                                                                       geom = "crossbar", width = 0.25)  + 
    ylab(Arealabel) +xlab('Treatment Index') +  labs(title= "Area  : Mean +/- SE  (outliers removed)") +
    theme_bw() + theme(plot.title = element_text(size = 10))
  
  
  
  ########################### Diameter PLots
  p_Diameter_new <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Diameter, colour = Diameter_status), size = Pointsize,  show.legend=FALSE ) + 
    # geom_point(data = Sph_Data_Out, aes(x= Col, y=Diameter_OO) , size = OutlierPointSize, colour = OutlierPointcolour ) +ylab(' Diameter  (um^2)')   xlab('Column') +
    scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
    facet_wrap(~Sph_Treat_ADVPC$T_I, nrow=1) +
    #   ylab(Arealabel) +xlab('Column') +  labs(title= " Area data by Treatment Index with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10))
    ylab(Diameterlabel)+xlab('Column') +  labs(title= " Diameter data by Treatment Index with outliers in red") +
    theme_bw() + theme(plot.title = element_text(size = 10))+ 
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
  png("p11.png",width = 1024, height = 400, res=150)
  print(p_Area_new) 
  dev.off()
  insertImage(wb, 1, "p11.png", width = 9, height = 3.5 , startRow = 1,startCol = 'A')

  
  png("p12.png",width = 1024, height = 400, res=150)
  print(p_Area_dotplot_new)
  dev.off()
  insertImage(wb, 1, "p12.png", width = 9, height = 3.5,startRow = 19,startCol = 'A')
  ### writing diameter plots  ###
  png("p21.png",width = 1024, height = 400, res=150)
  print(p_Diameter_new) 
  dev.off()
  insertImage(wb, 2, "p21.png", width = 9, height = 3.5 , startRow = 1,startCol = 'A')
  
  
  png("p22.png",width = 1024, height = 400, res=150)
  print(p_Diameter_dotplot_new)
  dev.off()
  insertImage(wb, 2, "p22.png", width = 9, height = 3.5,startRow = 19,startCol = 'A')
  
  ###### write Circularity plots
  png("p31.png",width = 1024, height = 400, res=150)
  print(p_Circularity_new) 
  dev.off()
  insertImage(wb, 3, "p31.png", width = 9, height = 3.5 , startRow = 1,startCol = 'A')
  
  
  png("p32.png",width = 1024, height = 400, res=150)
  print(p_Circularity_dotplot_new)
  dev.off()
  insertImage(wb, 3, "p32.png", width = 9, height = 3.5,startRow = 19,startCol = 'A')
  
  ###### write Volume plots
  png("p41.png",width = 1024, height = 400, res=150)
  print(p_Volume_new) 
  dev.off()
  insertImage(wb, 4, "p41.png", width = 9, height = 3.5 , startRow = 1,startCol = 'A')
  
  
  png("p42.png",width = 1024, height = 400, res=150)
  print(p_Volume_dotplot_new)
  dev.off()
  insertImage(wb, 4, "p42.png", width = 9, height = 3.5,startRow = 19,startCol = 'A')
  
  ###### write Perimeter plots
  png("p51.png",width = 1024, height = 400, res=150)
  print(p_Perimeter_new) 
  dev.off()
  insertImage(wb, 5, "p51.png", width = 9, height = 3.5 , startRow = 1,startCol = 'A')
  
  
  png("p52.png",width = 1024, height = 400, res=150)
  print(p_Perimeter_dotplot_new)
  dev.off()
  insertImage(wb, 5, "p52.png", width = 9, height = 3.5,startRow = 19,startCol = 'A')
  
  unlink(c("p11", "p12", "p21", "p22","p31", "p32","p41", "p42","p51", "p52"))
  
  # suppressMessages(print(p_Area_new))
  # insertPlot(wb, 1, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # suppressMessages(print(p_Area_dotplot_new))
  # insertPlot(wb, 1, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  # ### writing diameter plots  ###
  # suppressMessages(print(p_Diameter_new))
  # insertPlot(wb, 2, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # suppressMessages(print(p_Diameter_dotplot_new))
  # insertPlot(wb, 2, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")

  # ###### write Circularity plots
  # suppressMessages(print(p_Circularity_new))
  # insertPlot(wb, 3, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # suppressMessages(print(p_Circularity_dotplot_new))
  # insertPlot(wb, 3, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # ###### write Volume plots
  # suppressMessages(print(p_Volume_new))
  # insertPlot(wb, 4, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # suppressMessages(print(p_Volume_dotplot_new))
  # insertPlot(wb, 4, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # ###### write Perimeter plots
  # suppressMessages(print(p_Perimeter_new))
  # insertPlot(wb, 5, xy = c("A", 1), width = 9, height = 3.5,  fileType = "png", units = "in")
  # 
  # suppressMessages(print(p_Perimeter_dotplot_new))
  # insertPlot(wb, 5, xy = c("A", 19), width = 9, height = 3.5,  fileType = "png", units = "in")
  
  

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
  
  ADVPC_means<-Sph_ADVPC_2_ADVPC_means(Sph_ADVPC,df_treat,Spheroid_data,rawfilename)
  
  print(ADVPC_means$Time_Date)
  print(ADVPC_means$Job.Date)
  ### use ADVPC_means to generate merge file
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

  #
  # addStyle(wb, sheet = 12, datestyle, rows = 2:33 , cols = c(22,23), gridExpand = TRUE)
  addStyle(wb, sheet = 12, datestyle, rows = 2:33 , cols = c(which(colnames(df_excel)=="Job.Date"),
                                                             which(colnames(df_excel)=="Time_Date")), gridExpand = TRUE)
  
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


  ##### Put the manual outlier override ######
  c3 <- "Manual outlier override"
  writeData(wb, 13, c3, startCol = 2, startRow = 22)
  addStyle(wb, sheet = 13, styleBold, rows= 22 , cols= 2, gridExpand = TRUE)
  
  
  if(is.null(df_origin)){
    df_manual_override =data.frame(matrix(nrow = 1, ncol = 5))
    
  }else{
    d_a = Sph_Treat_Robz_ADVPC[(Sph_Treat_Robz_ADVPC$Area_status != df_origin$Area_status) & !is.na(Sph_Treat_Robz_ADVPC$Area_status),
                         'Well.Name']
    d_p = Sph_Treat_Robz_ADVPC[(Sph_Treat_Robz_ADVPC$Perimeter_status != df_origin$Perimeter_status) & !is.na(Sph_Treat_Robz_ADVPC$Perimeter_status),
                                  'Well.Name']
    d_c = Sph_Treat_Robz_ADVPC[(Sph_Treat_Robz_ADVPC$Circularity_status != df_origin$Circularity_status) & !is.na(Sph_Treat_Robz_ADVPC$Circularity_status),
                                  'Well.Name']
    d_d = Sph_Treat_Robz_ADVPC[(Sph_Treat_Robz_ADVPC$Diameter_status != df_origin$Diameter_status) & !is.na(Sph_Treat_Robz_ADVPC$Diameter_status),
                                  'Well.Name']
    d_v = Sph_Treat_Robz_ADVPC[(Sph_Treat_Robz_ADVPC$Volume_status != df_origin$Volume_status) & !is.na(Sph_Treat_Robz_ADVPC$Volume_status),
                                  'Well.Name']

    
    df_manual_override = cbind.fill(d_a,d_p, d_c, d_d, d_v)
    
  }

  colnames(df_manual_override) <- c("Area", "Perim.", "Circ.", "Dia.","Vol." )

  writeData(wb, 13, df_manual_override, startRow = 23, startCol = 2,  rowNames = FALSE, keepNA = FALSE, withFilter=FALSE)
  addStyle(wb, sheet = 13, styleborderTBLR, rows= 23:119 , cols= 2:6, gridExpand = TRUE)
  addStyle(wb, sheet = 13, greyshadeTBLR, rows= 23 , cols= 2:6, gridExpand = TRUE)
 
    
  
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

generate_merge_result = function(df,Spheroid_data, rawfilename,df_treat){

  df$T_I <- as.numeric(df$T_I)
  Sph_Treat_ADVPC <- df[with(df, order(T_I)),]
  
  Sph_ADVPC <- Sph_Treat_ADVPC
  ADVPC_means<-Sph_ADVPC_2_ADVPC_means(Sph_ADVPC,df_treat,Spheroid_data,rawfilename)
  
  # ADVPC_means<-Sph_ADVPC_2_ADVPC_means(Sph_ADVPC,df_treat,Spheroid_data)
  # 
  # dropped_cols=setdiff(colnames(Sph_ADVPC), 
  #                      intersect(colnames(Sph_ADVPC), colnames(df_treat)))
  # dropped_cols=c('T_I',dropped_cols)
  # Sph_ADVPC = Sph_ADVPC[, dropped_cols]
  # 
  # 
  # # merge with treatment data
  # 
  # Sph_ADVPC  <- merge(Sph_ADVPC, df_treat, by= "T_I")
  # 
  # Sph_ADVPC <- Sph_ADVPC[with(Sph_ADVPC, order(Row, Col)),]
  # 
  # Job_Info_data <- select(Spheroid_data, Jobrun.Finish.Time)
  # 
  # Sph_ADVPC$Job.Date <- as.Date(Job_Info_data$Jobrun.Finish.Time, origin="1900-01-01")
  # 
  # #work out time diff in days ETST.
  # 
  # Sph_ADVPC$ETST.d <-  difftime(Sph_ADVPC$Job.Date, as.Date(Sph_ADVPC$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "days")
  # Sph_ADVPC$ETST.h <-  difftime(Sph_ADVPC$Job.Date, as.Date(Sph_ADVPC$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "hours")
  # 
  # 
  # 
  # Sph_ADVPC$Filename <- rawfilename
  # Sph_ADVPCa <-  Sph_ADVPC
  # 
  # # calculate Means and SE's for A D C V P items
  # 
  # aArea_means <- data_summary_TI(Sph_ADVPCa, varname="Area_OR",
  #                                groupnames=c("T_I"))
  # 
  # aDiameter_means <- data_summary_TI(Sph_ADVPCa, varname="Diameter_OR",
  #                                    groupnames=c("T_I"))
  # 
  # aCircularity_means <- data_summary_TI(Sph_ADVPCa, varname="Circularity_OR",
  #                                       groupnames=c("T_I"))
  # 
  # 
  # aVolume_means <- data_summary_TI(Sph_ADVPCa, varname="Volume_OR",
  #                                  groupnames=c("T_I"))
  # 
  # 
  # aPerimeter_means <- data_summary_TI(Sph_ADVPCa, varname="Perimeter_OR",
  #                                     groupnames=c("T_I"))
  # 
  # 
  # #merge all mean+SE datasets sequentially
  # a_allmeans <- merge(aArea_means, aDiameter_means, by= "T_I")
  # a_allmeans1 <- merge(a_allmeans, aCircularity_means, by= "T_I")
  # a_allmeans2 <- merge(a_allmeans1, aVolume_means, by= "T_I")
  # a_allmeans3 <- merge(a_allmeans2, aPerimeter_means, by= "T_I")
  # colnames(a_allmeans3)<- c("T_I", "Area_Mean","Area_SE","Diameter_Mean","Diameter_SE","Circularity_Mean",
  #                           "Circularity_SE","Volume_Mean","Volume_SE", "Perimeter_Mean","Perimeter_SE" )
  # 
  # #subset TI and treatment items from main dataset
  # 
  # #merge final means dataset with Treatment data items
  # ADVPC_means <- merge(a_allmeans3, df_treat, by= "T_I")
  # 
  # job.date = Job_Info_data$Jobrun.Finish.Time[1]
  # ADVPC_means$Job.Date <- job.date
  # 
  # # calculate time diff in days and hours ETST.
  # 
  # ADVPC_means$ETST.d <-  as.numeric(difftime(ADVPC_means$Job.Date, as.Date(ADVPC_means$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "days"))
  # ADVPC_means$ETST.h <-  as.numeric(difftime(ADVPC_means$Job.Date, as.Date(ADVPC_means$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "hours"))
  # ADVPC_means$Filename <- rawfilename
  # 
  # # names(ADVPC_means)[13] <- "Date_Treated"
  # # names(ADVPC_means)[23] <- "Date_Scanned"
  # # ADVPC_means <-  ADVPC_means[,c(1,12, 2:11,14:22, 13, 23:26)]
  # 
  # ADVPC_means <- ADVPC_means[with(ADVPC_means, order(T_I)),]

  return(ADVPC_means)  
}


#################### 
##### Shiny App ####
####################

manual_outlier_selections = outer(LETTERS[1:8], paste0(1:12,'.'), FUN = "paste")
dim(manual_outlier_selections) =NULL

value_selections = list("Area", "Diameter", "Volume", "Perimeter" , "Circularity")

# Define UI for application that draws a histogram
ui <- navbarPage("SpheroidAnalyseR",

      # Application title
    tabPanel("Data Input",

    # Sidebar with a slider input for number of bins 
    sidebarLayout(
        sidebarPanel(
          tags$head(tags$style(".shiny-output-error{color: red;font-weight: bold;}")),
          fileInput(inputId = "raw_data","Choose Raw Data",accept=".xlsx",buttonLabel = "Browse", multiple = TRUE),
          
          fileInput(inputId = "layout","Choose Plate Layout",accept=".csv",buttonLabel="Browse"),
          textOutput("text_layout"),
          
          fileInput(inputId = "treat_data","Choose Treatment Definitions",accept=".csv",buttonLabel="Browse"),
          textOutput("text_treat"),
          downloadButton("downloadRawTemplate_btn", "Download the raw data template"),
          downloadButton("downloadLayoutTemplate_btn", "Download the plate layout template"),
          downloadButton("downloadTreatTemplate_btn", "Download the treatment template")
        ),
        # Show a plot of the generated distribution
        mainPanel(
          DT::dataTableOutput('x1'),
           h2("Preview the data"),    
           selectInput("select_file_preview",label="Choose a raw file to preview",
                       list()),
           
           # tableOutput("data_preview"), 
           DT::dataTableOutput("data_preview"),

           h2("Show the treatment groups on the plate"),
           selectInput("select_treatment",label="Choose the value to view the layout",
                       list()),
           plotOutput("show_layout")
        )
    )
    ),
    tabPanel("Outlier Removal",
                 sidebarPanel(
                   useShinyjs(), 
                   
                    selectInput("select_file",label="Choose a raw file",
                                list()),
                    
                    selectInput("select_view_value",label="Choose a value to plot the outliers",
                                value_selections),
                    
                    downloadButton("downloadData_btn", "Download"),
                    textOutput("textStatus"),
                    helpText("Please run the outlier removal before downloading the report."),
                    
                    actionButton("outlier_btn", "Remove outliers"),
                    helpText("The remove outlier button will plot the data based 
                    on the metric selected above. If you are happy to proceed
                    with default settings which will remove outliers based on
                    the pre-set thresholds in the blue box then click remove outliers.
                    Please use the blue box settings below if
                    you would like to manually adjust settings prior to outlier removal 
                    (this can be done any number of times)"),
                    
                    
                    
                    selectInput("manual_outliers",label="Toggle cells status normal/outliers",
                                manual_outlier_selections,
                                multiple=TRUE,selectize=TRUE),
                    
                    actionButton("manual_outliers_btn", "apply manual adjustment"),
                    helpText("Once you have performed outlier removal,
                    if you are unhappy with any of the automated choices 
                    you have an option here to manually change the status of any given well 
                    and then go to the top and press remove outliers again"),
                    textOutput("manualStatus"),
                

                   fluidRow(
                     column(12,
                            style = "background-color:#c0ecfa;",
                            
                            numericInput("z_score","Robust z-score",value = 1.96, step=0.01, min=0),
                            helpText("Robust-Z-Score looks at the spread of the data and 
                            removes things that are classed as outliers by a Robust-Z-Score 
                            (citation – I will get this to you soon)"),
                            
                            checkboxInput("pre_screen","Apply Pre-screen thresholds?",value = TRUE),
                            helpText("Pre-screen thresholds are to allow you to say that you will only include spheroids that lie within a range.
                            This is because there could be an empty well,
                            or an error with imaging that has produced a value that you don’t want to be included in any analysis."),
                            
                            # checkboxInput("override","Apply Manual overrides?",value = FALSE),
                            
                            
                            conditionalPanel(condition ="input.pre_screen==1",
                                             #- Values above zero
                                             numericInput("area_threshold_low","Area lower limit",value=1,min=0),
                                             numericInput("area_threshold_high","Area higher limit",value=1630000,min=0),
                                             numericInput("diam_threshold_low","Diameter lower limit",value=100,min=0),
                                             numericInput("diam_threshold_high","Diameter higher limit",value=1440,min=0),
                                             numericInput("vol_threshold_low","Volume lower limit",value=1000,min=0),
                                             numericInput("vol_threshold_high","Volume higher limit",value=1567543610,min=0),
                                             numericInput("perim_threshold_low","Perimeter lower limit",value=100,min=0),
                                             numericInput("perim_threshold_high","Perimeter higher limit",value=6276,min=0),
                                             numericInput("circ_threshold_low","Circularity lower limit",value=0.01,min=0),
                                             numericInput("circ_threshold_high","Circularity higher limit",value=1,min=0)
                            )       
                            
                     )
                   )
                   


                 ),
                 # Show a plot of the generated distribution
                 mainPanel(
                    useShinyjs(), 

                    strong("Plate layout after pre-sreen outlier removal (if applied)",id='title_pre'),
                    
                     plotOutput("outlierPlot"),
                     
                    strong("Plate layout after robust Z-score outlier removal"),
                    plotOutput("resultPlot"),
                    
                    strong("Plot 1"),
                    plotOutput("plot1Plot"),
                    
                    strong("Plot 2"),
                    plotOutput("plot2Plot")
                     # selectInput("select_z_scores",label="Choose the value to view result plots",
                     #             list("Area", "Diameter", "Volume", "Perimeter" , "Circularity")),
                     
                     
                     
                     
                 )
             
    ),
    
    tabPanel("Merging",
             sidebarPanel(
               
               h3("Only available when all files have been processed"),
               
               checkboxInput("processed_file_chk","Use previous reports",value = FALSE),
               
               # checkboxInput("override","Apply Manual overrides?",value = FALSE),

               conditionalPanel(condition ="input.processed_file_chk==1",
                                #- Values above zero
                                fileInput(inputId = "processed_data","Choose previous reports",accept=".xlsx",buttonLabel = "Browse", multiple = TRUE),
                                textOutput("text_processed")
               ),
               
               
               textInput("mergeName_text", "Merged file name", value = paste0("merge_file_", Sys.Date(),".xlsx")),
               downloadButton("downloadMerge_btn", "Download the merged file"),
               h2(""),
               textInput("configName_text", "Config file name", value = paste0("config_file_", Sys.Date(),".csv")),
               downloadButton("downloadConfig_btn", "Download the config file")),
             
             mainPanel(
               h2("Raw files"),
               DT::dataTableOutput("merge_file"),
               h2("Previous reports"),
               DT::dataTableOutput("processed_data_table")
               # tableOutput("merge_file")
               
             )
    ),
    
    tabPanel("Help",
               mainPanel(
                 includeHTML("help_file_page.html")
               )
    )

    
    
)


# Define server logic required to draw a histogram
server <- function(input, output,session) {
    
    shinyjs::disable("downloadMerge_btn")
    shinyjs::disable("downloadConfig_btn")
    
    use_previous_report <-FALSE
    
    df_output <- NULL
    df_origin<-NULL

    global_wb <-NULL
     
    df_plot_treat <-NULL
    df_plot_layout <-NULL
    
    
    ###preparing for the  generate the excel
    global_rawfilename <-NULL
    global_platesetupname <-NULL
    
    global_df_setup <-NULL
    global_df_treat <-NULL
    global_Spheroid_data <-NULL
    
    global_RobZ_LoLim <-NULL
    global_RobZ_UpLim <-NULL
    
    global_TF_apply_thresholds <-NULL
    global_TF_outlier_override <-NULL
    global_TF_copytomergedir <-NULL
    

    global_TH_Area_min <-NULL
    global_TH_Area_max <-NULL
    
    global_TH_Diameter_min <-NULL
    global_TH_Diameter_max <-NULL
    
    global_TH_Volume_min <-NULL
    global_TH_Volume_max <-NULL
    
    global_TH_Perimeter_min <-NULL
    global_TH_Perimeter_max <-NULL
    
    global_TH_Circularity_min <-NULL
    global_TH_Circularity_max <-NULL

    
    df_batch_detail <-data.frame()
    df_prev_report_detail<-data.frame()
    
    n_cb_raw<-0
    n_cb_prev<-0
    
    df_output_list <- FALSE
    df_origin_list <- FALSE
    df_spheroid_list <- FALSE
    
    
    # help function for creating checkbox input in table
    shinyInput = function(FUN, len, id, value, ...) {
      if (length(value) == 1) value <- rep(value, len)
      inputs = character(len)
      for (i in seq_len(len)) {
        inputs[i] = as.character(FUN(paste0(id, i), label = NULL, value = value[i]))
      }
      inputs
    }
    
    # obtain the values of inputs
    shinyValue = function(id, len) {
      unlist(lapply(seq_len(len), function(i) {
        value = input[[paste0(id, i)]]
        if (is.null(value)) TRUE else value
      }))
    }
    


    #### input panel ####
    
    ## Read the raw file (first file if multiple files have been uploaded)
    ## event: Upload raw file
    observeEvent(input$raw_data, {
      
      file <- input$raw_data
      
      ## validate the file path is correct
      if(!is.null(file)){
        for (datapath in file$datapath){
          ext <- tools::file_ext(file$datapath)
          req(file)
          validate(need(ext == "xlsx","Please upload an xlsx file"))
        }
        
      ## validate the file is correct  
        
      
      # get the original file names
      ori_filename = input$raw_data$name
      
      # Create selections for raw files in both input and outlier removal panels.
      updateSelectInput(session, "select_file_preview",
                        choices = ori_filename,
                        selected = head(ori_filename, 1))
      
      updateSelectInput(session, "select_file",
                        choices = ori_filename,
                        selected = head(ori_filename, 1))
      
      n_cb_raw <<- length(ori_filename)
      ##initial the variables and UI for merging
      df_batch_detail<<-data.frame(
        
        Use = shinyInput(checkboxInput, n_cb_raw, 'cb_raw_', value = TRUE, width='1px'),                            
        File_name=ori_filename, Processed=rep(c(FALSE), times = length(ori_filename)),
                                   Robust_Z_Low_Limit=NA,
                                   Robust_Z_Upper_Limit= NA,
                                   Threshold_Applied = NA,
                                   Threshold_Area_Min =NA,
                                   Threshold_Area_Max =NA,
                                   Threshold_Diameter_Min =NA,
                                   Threshold_Diameter_Max =NA,
                                   Threshold_Volume_Min =NA,
                                   Threshold_Volume_Max =NA,
                                   Threshold_Perimeter_Min =NA,
                                   Threshold_Perimeter_Max =NA,
                                   Threshold_Circularity_Min =NA,
                                   Threshold_Circularity_Max =NA
                                   )
      
      df_output_list <<- vector("list", length = length(ori_filename))
      df_origin_list <<- vector("list", length = length(ori_filename))
      df_spheroid_list <<- vector("list", length = length(ori_filename))
      
      ## update file list in merge tab
      
      #Display table with checkbox buttons
      
      # output$merge_file <- 
      #   DT::renderDataTable({
      #     DT::datatable(
      #       cbind(cb_raw = shinyInput(checkboxInput, n_cb_raw, 'cb_raw_', value = TRUE, width='1px'),
      #             df_batch_detail),
      #       # df_batch_detail,
      #       options = list(scrollX = TRUE))})
      ##

      # Init the table for merging
      output$merge_file <-
        DT::renderDataTable({
          DT::datatable(
            df_batch_detail,
            # df_batch_detail,
            escape = FALSE, selection = 'none',
            options = list(
              pageLength = 5,
              scrollY = TRUE,
              scrollX = TRUE,
              preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'),
              drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); } ')
            )
          )
        })
      
      observe({
        shinyValue('cb_raw_', n_cb_raw)
        if(all(df_batch_detail$Processed[shinyValue('cb_raw_', n_cb_raw)]) & any(df_batch_detail$Processed)){
          shinyjs::enable("downloadMerge_btn")
          shinyjs::enable("downloadConfig_btn")
        }else{
          shinyjs::disable("downloadMerge_btn")
          shinyjs::disable("downloadConfig_btn")
        }
        
      })
      
      
      }
    })
    

    
    ## update the raw data preview in the input panel
    ## event: After uploading the raw data
    output$data_preview <- DT::renderDataTable(
      {
        file <- input$raw_data
        name_id = which(file$name == input$select_file_preview)
        print(input$select_file_preview)
        print(file$datapath[name_id])
        if(!is.null(file)){
          ext <- tools::file_ext(file$datapath[name_id])
          req(file)
          validate(need(ext == "xlsx","Please upload an xlsx file"))
          data <- readxl::read_xlsx(file$datapath[name_id])
          DT::datatable(
            data,
            options = list(scrollX = TRUE,
                           scrollY = TRUE,
                           pageLength = 5))
        }
      }
    )
    
    ## read layout file
    ## If treament file has been uploaded:
    ##    draw layout vs treament variables plot in layout & treatment preview
    ## else:
    ##    only draw the layout in layout & treatment preview
    ## event: After uploading the layout file
    output$text_layout <- reactive({
      layout_file <- input$layout
      
      if(!is.null(layout_file)){
        ext <- tools::file_ext(layout_file$datapath)
        req(layout_file)
        validate(need(ext == "csv","Please upload a layout in csv format"))
        
        df_check = read.csv(layout_file$datapath,
                            row.names = 1,header=TRUE, check.names = FALSE)
        
        error_list = check_layout(df_check)

        validate(need(error_list[1] == TRUE,"Dimension error\nCheck if the file is 8x12 (row names and column names excluded)"))
        validate(need(error_list[2] == TRUE,"Column names error\nplease refer to the template"))
        validate(need(error_list[3] == TRUE,"Row names error\nplease refer to the template"))
        validate(need(error_list[4] == TRUE,"Value error\nCheck values in the layout file"))
        
        
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
      ""
    }
    )
    
    
    # observeEvent(input$layout, {
    #   layout_file <- input$layout
    #   
    #   if(!is.null(layout_file)){
    #     ext <- tools::file_ext(layout_file$datapath)
    #     req(layout_file)
    #     validate(need(ext == "csv","Please upload a layout in csv format"))
    #     
    #     df_check = read.csv(layout_file$datapath,
    #                    row.names = 1,header=TRUE, check.names = FALSE)
    #     
    #     error_list = check_layout(df_check)
    #     print(error_list)
    #     validate(need(error_list[1] == TRUE,"Check if the layout is 8x12 (row names and column names excluded)"))
    #     validate(need(error_list[2] == TRUE,"Column names error, please refer to the template"))
    #     validate(need(error_list[3] == TRUE,"Row names error, please refer to the template"))
    #     validate(need(error_list[4] == TRUE,"Check values in the layout file"))
    # 
    #     
    #     layout <- readr::read_csv(layout_file$datapath)  %>% 
    #       #     #layout <- readr::read_csv("layout_example.csv") %>% 
    #       dplyr::rename("Row"=X1) %>%
    #       tidyr::pivot_longer(-Row,names_to="Col",
    #                           values_to="Index") %>%
    #       mutate(Index = as.factor(Index),Col=as.numeric(Col)) %>%
    #       filter(!is.na(Row))
    #     
    #     df_plot_layout <<- layout
    #     layout
    #   }
    #   
    #   output$show_layout <- renderPlot(
    #     {
    #       if(!is.null(df_plot_layout)){
    #         if(!is.null(df_plot_treat)){
    #           draw_layout_with_treat(df_plot_layout, df_plot_treat, input$select_treatment)
    #         }
    #         else{
    #           draw_layout(df_plot_layout)
    #         }
    #         
    #       }
    #     }
    #   )
    #   
    # })
    
    ## read treament file
    ## update the treatment variable selections
    ## If layout file has been uploaded:
    ##    draw layout vs treament variables plot in layout & treatment preview
    ## event: After uploading the treatment file
    
    output$text_treat<- reactive({
      treat_file <- input$treat_data
      ## TO DO: Need to check that 1st column is named Index
      if(!is.null(treat_file)){
        ext <- tools::file_ext(treat_file$datapath)
        req(treat_file)
        validate(need(ext == "csv","Please upload a treatment in csv format"))
        
        df_check = read.csv(treat_file$datapath,
                     header=TRUE)
        
        error_list = check_treatment(df_check)
        
        validate(need(error_list[1] == TRUE,"Dimension error\nCheck if the file is 6x8 (row names and column names excluded)"))
        validate(need(error_list[2] == TRUE,"Column names error\nplease refer to the template"))
        validate(need(error_list[3] == TRUE,"Row names error\nplease refer to the template"))
        validate(need(error_list[4] == TRUE,"Value error\nCheck values in the treatment file"))

        
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
      ""
    })
    
    # observeEvent(input$treat_data, {
    #   treat_file <- input$treat_data
    #   ## TO DO: Need to check that 1st column is named Index
    #   if(!is.null(treat_file)){
    #     ext <- tools::file_ext(treat_file$datapath)
    #     req(treat_file)
    #     validate(need(ext == "csv","Please upload a layout in csv format"))
    #     treatments <- readr::read_csv(treat_file$datapath) %>% 
    #       mutate_all(as.factor)
    #     # update the value for treatment
    #     selections = colnames(treatments)
    #     updateSelectInput(session, "select_treatment",
    #                       choices = selections,
    #                       selected = head(selections, 1))
    #     
    #     df_plot_treat<<-treatments
    #     treatments
    #     
    #   }
    #   output$show_layout <- renderPlot(
    #     {
    #       if(!is.null(df_plot_layout)){
    #         if(!is.null(df_plot_treat)){
    #           draw_layout_with_treat(df_plot_layout, df_plot_treat, input$select_treatment)
    #         }
    #         else{
    #           draw_layout(df_plot_layout)
    #         }
    #         
    #       }
    #     }
    #   )
    #   
    # })
    
    ###functions for drawing layout & treatment preview
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
    
    
    #### outlier removal panel ####
  
    ## remove outliers of the selected raw file in the outlier removal panel
    ## event: click the outlier removal button
    
    output_report <- eventReactive(input$outlier_btn, {
      validate(
        need(is.numeric(input$z_score) & input$z_score>0, "Please input a positive Z score (numeric)"),
        
        # need(is.numeric(input$z_low), "Please check if the low limit of the Z score is numeric"),
        # need(is.numeric(input$z_high), "Please check if the high limit of the Z score is numeric"),
        
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
      
      name_id = which(input$raw_data$name == input$select_file)
      rawfilename = input$raw_data$datapath[name_id]  
        
      ori_rawfilename = input$select_file
      
      # rawfilename = input$raw_data$datapath
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
      
      
      RobZ_LoLim <- -(input$z_score)
      RobZ_UpLim <- input$z_score
      
      # RobZ_LoLim <- input$z_low
      # RobZ_UpLim <- input$z_high
      
      TF_apply_thresholds <- input$pre_screen
      TF_outlier_override <- FALSE
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
      
      
      ### Apply manual override ###
      
      
      ### need lists to store dataframe
      
      ## assign to df_output

      df_output_list[[name_id]] <<-Sph_Treat_Robz_ADVPC
      # print(df_origin_list)
      if(is.null(df_origin_list[[name_id]])){
        df_origin_list[[name_id]]<<-Sph_Treat_Robz_ADVPC
      }
      df_spheroid_list[[name_id]]<<-Spheroid_data
      

      df_output <<- Sph_Treat_Robz_ADVPC
      if(is.null(df_origin)){
        df_origin<<- Sph_Treat_Robz_ADVPC
      }
      global_Spheroid_data <<- Spheroid_data
      
      
      
      
      #### preparing for the  generate the excel
      
      ## variables that are same across different data
      # layout setup and treatment
      global_df_setup <<- df_setup
      global_df_treat <<- df_treat
      # name of the setup
      global_platesetupname <<-input$layout$name
      
      
      ## variables that are different across different data
      ## stored in a dataframe 
      # names:
      global_rawfilename <<- ori_rawfilename

      
      df_temp_batch_detail = df_batch_detail

      df_temp_batch_detail$Robust_Z_Low_Limit[name_id] = RobZ_LoLim
      df_temp_batch_detail$Robust_Z_Upper_Limit[name_id] = RobZ_UpLim
      df_temp_batch_detail$Threshold_Applied[name_id] = TF_apply_thresholds
      
      
      global_RobZ_LoLim <<- RobZ_LoLim
      global_RobZ_UpLim <<- RobZ_UpLim
      
      global_TF_apply_thresholds <<- TF_apply_thresholds
      global_TF_outlier_override <<-TF_outlier_override
      global_TF_copytomergedir <<- TF_copytomergedir

      if(TF_apply_thresholds){
        df_temp_batch_detail$Threshold_Area_Min[name_id] = TH_Area_min
        df_temp_batch_detail$Threshold_Area_Max[name_id] = TH_Area_max
        
        df_temp_batch_detail$Threshold_Diameter_Min[name_id] = TH_Diameter_min
        df_temp_batch_detail$Threshold_Diameter_Max[name_id] = TH_Diameter_max
        
        df_temp_batch_detail$Threshold_Perimeter_Min[name_id] = TH_Perimeter_min
        df_temp_batch_detail$Threshold_Perimeter_Max[name_id] = TH_Perimeter_max
        
        df_temp_batch_detail$Threshold_Volume_Min[name_id] = TH_Volume_min
        df_temp_batch_detail$Threshold_Volume_Max[name_id] = TH_Volume_max
        
        df_temp_batch_detail$Threshold_Circularity_Min[name_id] = TH_Circularity_min
        df_temp_batch_detail$Threshold_Circularity_Max[name_id] = TH_Circularity_max
        
        ## original way to storing these values
        global_TH_Area_min <<- TH_Area_min 
        global_TH_Area_max <<- TH_Area_max
        
        global_TH_Diameter_min <<- TH_Diameter_min
        global_TH_Diameter_max <<- TH_Diameter_max
        
        global_TH_Volume_min <<- TH_Volume_min
        global_TH_Volume_max <<- TH_Volume_max
        
        global_TH_Perimeter_min <<- TH_Perimeter_min
        global_TH_Perimeter_max <<- TH_Perimeter_max
        
        global_TH_Circularity_min <<- TH_Circularity_min
        global_TH_Circularity_max <<-TH_Circularity_max
      }
      df_temp_batch_detail$Processed[name_id] = TRUE
      
      df_batch_detail<<-df_temp_batch_detail
      
      if(all(df_batch_detail$Processed[shinyValue('cb_raw_', n_cb_raw)]) & any(df_batch_detail$Processed)){
        shinyjs::enable("downloadMerge_btn")
        shinyjs::enable("downloadConfig_btn")
      }else{
        shinyjs::disable("downloadMerge_btn")
        shinyjs::disable("downloadConfig_btn")
      }
      
      # global_wb <<- gen_report(Sph_Treat_Robz_ADVPC=Sph_Treat_Robz_ADVPC,
      #                          df_treat = df_treat,df_setup = df_setup,
      #                          Spheroid_data=Spheroid_data,
      #                          # p_Area_new = p_Area_new, p_Area_dotplot_new = p_Area_dotplot_new,
      #                          # p_Diameter_new =p_Diameter_new, p_Diameter_dotplot_new =p_Diameter_dotplot_new,
      #                          # p_Circularity_new = p_Circularity_new, p_Circularity_dotplot_new = p_Circularity_dotplot_new,
      #                          # p_Volume_new = p_Volume_new, p_Volume_dotplot_new = p_Volume_dotplot_new,
      #                          # p_Perimeter_new= p_Perimeter_new , p_Perimeter_dotplot_new = p_Perimeter_dotplot_new,
      #                          rawfilename = rawfilename ,platesetupname=platesetupname,
      #                          RobZ_LoLim = RobZ_LoLim, RobZ_UpLim = RobZ_UpLim,
      #                          TF_apply_thresholds = TF_apply_thresholds,
      #                          TF_outlier_override = TF_outlier_override,
      #                          TH_Area_max=TH_Area_max,TH_Area_min=TH_Area_min,
      #                          TH_Diameter_max=TH_Diameter_max,TH_Diameter_min=TH_Diameter_min,
      #                          TH_Volume_max=TH_Volume_max, TH_Volume_min=TH_Volume_min,
      #                          TH_Perimeter_max = TH_Perimeter_max, TH_Perimeter_min=TH_Perimeter_min,
      #                          TH_Circularity_max = TH_Circularity_max, TH_Circularity_min=TH_Circularity_min
      #                          )
      
      message("All processing completed successfully.", "\n")
      message("\n") 
      
      
      
      # p_list = list(p_Area_new,p_Area_dotplot_new, 
      #             p_Diameter_new,p_Diameter_dotplot_new,
      #             p_Circularity_new,p_Circularity_dotplot_new,
      #             p_Volume_new,p_Volume_dotplot_new,
      #             p_Perimeter_new,p_Perimeter_dotplot_new)
      "Report generated, please download the report"
    })
    
    ## renew the text in the outlier removal tab
    ## event: after outlier removal
    output$textStatus<-  renderText({
      output_report()
    })

      # output$outlierPlot <- renderPlot({
      #   message(input$select_view_value)
      #   validate(need(df_output,"Please run the outlier removal"))
      #   message(paste("plot",input$select_view_value))
      #   df_outlier_temp = df_output
      #   draw_outlier_plot(df_outlier_temp,input$select_view_value)
      # })
      # 
      # 
      # output$resultPlot <- renderPlot({
      #   message(input$select_view_value)
      #   validate(need(df_output,"Please run the outlier removal"))
      #   Sph_Treat_ADVPC = df_output
      #   draw_z_score_outlier_plot(Sph_Treat_ADVPC,input$select_view_value)
      # 
      # })
      # 
      # output$plot1Plot <- renderPlot({
      #   message(input$select_view_value)
      #   validate(need(df_output,"Please run the outlier removal"))
      #   draw_plot_1(df_output,input$select_view_value )
      #   
      # })
      # 
      # output$plot2Plot <- renderPlot({
      #   message(input$select_view_value)
      #   validate(need(df_output,"Please run the outlier removal"))
      #   draw_plot_2(df_output,input$select_view_value )
      # })
      
    ## function for updating plots in the outlier removal panel
      update_plots = function(df, value){
        output$outlierPlot <- renderPlot({
          validate(need(df,"Please run the outlier removal"))
          draw_outlier_plot(df,value)
          
        })
        
        output$resultPlot <- renderPlot({
          validate(need(df,"Please run the outlier removal"))
          draw_z_score_outlier_plot(df,value)
        })
        
        output$plot1Plot <- renderPlot({
          validate(need(df,"Please run the outlier removal"))
          draw_plot_1(df,value)
          
        })
        
        output$plot2Plot <- renderPlot({
          validate(need(df,"Please run the outlier removal"))
          draw_plot_2(df,value )
        })
        
      }
      
      ## drawing plot in the outlier removal tab
      ## event: select the viewing value
      
      observeEvent(c(input$select_view_value , input$select_file), {
        print("selected clicked")
        if(typeof(df_output_list)=="list"){
          name_id = which(input$raw_data$name == input$select_file)
          df_temp = df_output_list[[name_id]]
          
          update_plots(df_temp , input$select_view_value)
          
        }
        

      })
 
      ## drawing plot in the outlier removal tab
      ## event: click the outlier removal button
      observeEvent(input$outlier_btn, {
        output_report()
        name_id = which(input$raw_data$name == input$select_file)
        df_temp = df_output_list[[name_id]]
        
        update_plots(df_temp , input$select_view_value)
        
        # output$merge_file <- DT::renderDataTable({
        #   DT::datatable(
        #     df_batch_detail,options = list(scrollX = TRUE))})
        
        output$merge_file <-
          DT::renderDataTable({
            DT::datatable(
              df_batch_detail,
              # df_batch_detail,
              escape = FALSE, selection = 'none',
              options = list(
                pageLength = 5,
                scrollY = TRUE,
                scrollX = TRUE,
                preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'),
                drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); } ')
              )
            )
          })
        
      })
      

      ## update the current selected data using manual override
      ## event: click the apply button
      
      manual_override <- eventReactive(input$manual_outliers_btn,{

        validate(need(df_output,"Please run the outlier removal"))
        validate(need(input$manual_outliers,"Please select wells to override"))
        
        
        name_id = which(input$raw_data$name == input$select_file)
        
        validate(need(df_output_list[[name_id]],"Please run the outlier removal"))
        
        df_temp = df_output_list[[name_id]]
        
        selections=str_split(input$manual_outliers,'\\.')
        value_col=  paste0(input$select_view_value, "_status")
        
        check_non_exist_well=FALSE
        for(i in 1:length(selections)){
          split_result = str_split(selections[i][[1]],' ')[[1]]
          row_id = split_result[1]
          col_id = split_result[2]
          if(!any(df_temp$Row==row_id & df_temp$Col==col_id, na.rm = TRUE)){
            check_non_exist_well=TRUE
            break
          }
        }

        validate(need(!check_non_exist_well,"Please select existed wells"))

        
        for(i in 1:length(selections)){
          split_result = str_split(selections[i][[1]],' ')[[1]]
          row_id = split_result[1]
          col_id = split_result[2]

          previous = df_temp[df_temp$Row==row_id & df_temp$Col==col_id, value_col]
          
          message(paste("value ",row_id, col_id, is.na(previous)))
          message(paste("value ",row_id, col_id, previous))
          if(!is.na(previous)){
            if(previous =='1'){
              df_temp[df_temp$Row==row_id & df_temp$Col==col_id, value_col]='0'
            }
            if(previous =='0'){
              df_temp[df_temp$Row==row_id & df_temp$Col==col_id, value_col]='1'
            }
            
          }

        }
        df_output_list[[name_id]] <<- update_df_OR_by_status(df_temp)
        
        

        "Override successfully"
      }
      )
      
      
      ### event from clicking the remove outlier button
      output$manualStatus<-  renderText({
        manual_override()
      })
      
      observeEvent(input$manual_outliers_btn, {


        manual_override()

        name_id = which(input$raw_data$name == input$select_file)
        
        df_temp = df_output_list[[name_id]]
        
        update_plots(df_temp , input$select_view_value)
        
        # output$merge_file <- DT::renderDataTable({
        #   DT::datatable(
        #     df_batch_detail,options = list(scrollX = TRUE))})
        
        output$merge_file <-
          DT::renderDataTable({
            DT::datatable(
              df_batch_detail,
              # df_batch_detail,
              escape = FALSE, selection = 'none',
              options = list(
                pageLength = 5,
                scrollY = TRUE,
                scrollX = TRUE,
                preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'),
                drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); } ')
              )
            )
          })
        
        
      })
      
      ## hide the pre screen thresholds 
      ## event click the tick box of showing the threshold
      observeEvent(input$pre_screen,{
        message(input$pre_screen)
        if(input$pre_screen==TRUE){
          show("outlierPlot")
          show("select_outlier_values")
          shinyjs::show("title_pre")
        }else{
          shinyjs::hide("title_pre")
          hide("outlierPlot")
          hide("select_outlier_values")
        }
      })

      observeEvent(input$downloadData_btn,{
        message("Preparing")

      })
      
    ### functions of downloading  
      
    ### Download report of the current selected file
    output$downloadData_btn <- downloadHandler(

      filename = function() {
        
        file_name = paste0("report_",input$select_file)
        paste(file_name, sep = "")
      },
      content = function(file) {
        message("Preparing")
        
        ## get the name id
        name_id = which(input$raw_data$name == input$select_file)
        ## get the correct variable
        df_file_detail = df_batch_detail[name_id,]

        global_wb <<- gen_report(Sph_Treat_Robz_ADVPC=df_output_list[[name_id]],
                                 df_treat = global_df_treat,df_setup = global_df_setup,
                                 Spheroid_data=df_spheroid_list[[name_id]],
                                 df_origin = df_origin_list[[name_id]],
                                 rawfilename = df_file_detail$File_name ,platesetupname=global_platesetupname,
                                 RobZ_LoLim = df_file_detail$Robust_Z_Low_Limit, RobZ_UpLim = df_file_detail$Robust_Z_Upper_Limit,
                                 TF_apply_thresholds = df_file_detail$Threshold_Applied,
                                 TF_outlier_override = global_TF_outlier_override,
                                 TH_Area_max=df_file_detail$Threshold_Area_Max,TH_Area_min=df_file_detail$Threshold_Area_Min,
                                 TH_Diameter_max=df_file_detail$Threshold_Diameter_Max,TH_Diameter_min=df_file_detail$Threshold_Diameter_Min,
                                 TH_Volume_max=df_file_detail$Threshold_Volume_Max, TH_Volume_min=df_file_detail$Threshold_Volume_Min,
                                 TH_Perimeter_max = df_file_detail$Threshold_Perimeter_Max, TH_Perimeter_min=df_file_detail$Threshold_Perimeter_Min,
                                 TH_Circularity_max = df_file_detail$Threshold_Circularity_Max, TH_Circularity_min=df_file_detail$Threshold_Circularity_Min)
        
        
        
        # global_wb <<- gen_report(Sph_Treat_Robz_ADVPC=df_output,
        #                          df_treat = global_df_treat,df_setup = global_df_setup,
        #                          Spheroid_data=global_Spheroid_data,
        #                          df_origin = df_origin,
        #                          rawfilename = global_rawfilename ,platesetupname=global_platesetupname,
        #                          RobZ_LoLim = global_RobZ_LoLim, RobZ_UpLim = global_RobZ_UpLim,
        #                          TF_apply_thresholds = global_TF_apply_thresholds,
        #                          TF_outlier_override = global_TF_outlier_override,
        #                          TH_Area_max=global_TH_Area_max,TH_Area_min=global_TH_Area_min,
        #                          TH_Diameter_max=global_TH_Diameter_max,TH_Diameter_min=global_TH_Diameter_min,
        #                          TH_Volume_max=global_TH_Volume_max, TH_Volume_min=global_TH_Volume_min,
        #                          TH_Perimeter_max = global_TH_Perimeter_max, TH_Perimeter_min=global_TH_Perimeter_min,
        #                          TH_Circularity_max = global_TH_Circularity_max, TH_Circularity_min=global_TH_Circularity_min)
        
        saveWorkbook(global_wb, file, overwrite = TRUE)
        
      }
    )
    
    output$downloadRawTemplate_btn <- downloadHandler(
      
      filename = function() {
        paste( "template_raw_data.xlsx", sep = "")
      },
      content = function(file) {
        
        df_temp = read_excel("template_raw_data.xlsx", "JobView",col_names=TRUE, .name_repair = "universal")
        write.xlsx(df_temp,file ,sheetName = "JobView" )
        
      }
    )
    
    
    output$downloadLayoutTemplate_btn <- downloadHandler(
      
      filename = function() {
        paste( "template_layout.csv", sep = "")
      },
      content = function(file) {
        df_temp = read.csv("template_layout.csv", row.names = 1,header=TRUE, check.names = FALSE)
        
        write.csv(df_temp,file,na="")
        # saveWorkbook(global_wb, file, overwrite = TRUE)
        
      }
    )
    
    output$downloadTreatTemplate_btn <- downloadHandler(
      
      filename = function() {
        paste( "template_treatment.csv", sep = "")
      },
      content = function(file) {
        df_temp = read.csv("template_treatment.csv",
                           header=TRUE)
        
        write.csv(df_temp,file,
                  na="",row.names=FALSE)
        # saveWorkbook(global_wb, file, overwrite = TRUE)
        
      }
    )
    
    ## functions in merge part
    
    ## Read in previous report in merge tab
    ## event: upload previous report files
    output$text_processed <- reactive({
      processed_file <- input$processed_data
      ## validate the file path is correct
      if(!is.null(processed_file)){
        for (datapath in processed_file$datapath){
          print(datapath)
          ext <- tools::file_ext(processed_file$datapath)
          req(processed_file)
          validate(need(ext == "xlsx","Please upload an xlsx file"))
        }
        
        use_previous_report<<- TRUE
        for (path in processed_file$name){
          print(path)
        }
        
        
        ## Init the table for previous report
        
        n_cb_prev <<- length(processed_file$name)

        df_prev_report_detail<<-data.frame(
          Use = shinyInput(checkboxInput, n_cb_prev, 'cb_prev_', value = TRUE, width='1px')      ,
          File_name=input$processed_data$name
        )
        
        output$processed_data_table <-DT::renderDataTable({
              DT::datatable(
                df_prev_report_detail,
                escape = FALSE, selection = 'none',
                options = list(
                  scrollX = TRUE,
                  preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'),
                  drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); } ')
                )
              )
        })
        
        ## observe the checkboxes for previous report table
        observe({
          shinyValue('cb_prev_', n_cb_prev)
          if(any(shinyValue('cb_prev_', n_cb_prev)) == FALSE){
            use_previous_report<<- FALSE
          }
          
        })
        
        ""
      }
        
        ## validate the file is correct  
        
        
        # get the original file names
        # ori_filename = input$processed_file$name
        # 
        # print(ori_filename)
      
      
    })

    ## merge processed raw files into the merge file
    ## previous report can also be added in merging if provided
    ## it is possible to select which files to merge
    merge_results = function(){
      result_list = list()
      
      
      raw_file_cb_result = shinyValue('cb_raw_', n_cb_raw)
      print(raw_file_cb_result)
      for(i in 1:length(df_batch_detail$File_name)){
        if (raw_file_cb_result[i]==TRUE){
          result_list = append(result_list,
                              list(generate_merge_result(df_output_list[[i]],df_spheroid_list[[i]], df_batch_detail$File_name[i],global_df_treat)))
        }

      }
      
      prev_file_cb_result = shinyValue('cb_prev_', n_cb_prev)
      if(input$processed_file_chk==TRUE & use_previous_report==TRUE){
        
        
        prev_list = lapply(input$processed_data$datapath[prev_file_cb_result], 
                           function(x){read_excel(x, "Export dataset",col_names=TRUE, .name_repair = "universal")})
        
        prev_list = lapply(prev_list, function(x){
          x$Job.Date=as.Date(x$Job.Date, origin="1900-01-01")
          return(x)}
        )
        
        print(length(result_list))
        result_list = c(result_list,prev_list)
        d1 = prev_list[[1]]
        print(dim(d1))
        print(colnames(d1))
        print(rownames(d1))
        print(d1$Job.Date)
        print(d1$Time_Date)
        
        d1 = result_list[[1]]
        
      }
      
      
      full_data = Reduce(function(x,y) {merge(x,y,all = TRUE)}, result_list)
      
      print(colnames(full_data))
      full_data_A  <- full_data[with(full_data , order(ETST.d, T_I)),]
      full_data_B <- full_data_A
      
      
      # print(full_data_A)
      
      full_data_B$Job.Date <- as.Date(full_data_A$Job.Date, origin="1900-01-01")
      full_data_B$Job.Date <- as.POSIXct(full_data_B$Job.Date, format = "%Y-%m-%d %H:%M:%S")
      # print(full_data_B$Job.Date)
      
      
      # full_data_B$Time_Date <- as.POSIXct(full_data_A$Time_Date, format = "%Y-%m-%d %H:%M:%S")
      # print(full_data_B)
      
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
      
      
      addStyle(wb_mergedata, sheet = 1, dpstyle, rows = 2:maxrows , cols = c( 3:12,24:25), gridExpand = TRUE)
      addStyle(wb_mergedata, sheet = 1, datestyle, rows = 2:maxrows , cols = c(
                                                                 which(colnames(full_data_B)=="Time_Date"),
                                                                 which(colnames(full_data_B)=="Job.Date")), gridExpand = TRUE)
      # addStyle(wb_mergedata, sheet = 1, datestyle, rows = 2:maxrows , cols = c(22,23), gridExpand = TRUE)
      
      
      
      addStyle(wb_mergedata, sheet = 1, datastyleC, rows = 2:maxrows , cols = c(1,13:18,20:21, 24,25), gridExpand = TRUE)
      addStyle(wb_mergedata, sheet = 1, datastyleC, rows = 1 , cols = c(1:26), gridExpand = TRUE)
      
      addStyle(wb_mergedata, sheet = 1, datastyleR, rows = 2:maxrows , cols = c(3:12), gridExpand = TRUE)
      
      
      setColWidths(wb_mergedata, sheet = 1, cols = c(3:12, 24:25), widths = 16 )
      setColWidths(wb_mergedata, sheet = 1, cols = c(13:21), widths = 10 )
      setColWidths(wb_mergedata, sheet = 1, cols = c(22, 23), widths = 24 )
      
      setColWidths(wb_mergedata, sheet = 1, cols = 26, widths = 30 )
      setColWidths(wb_mergedata, sheet = 1, cols = 2, widths = 20)
      
      
      freezePane(wb_mergedata, sheet = 1,  firstActiveRow = 2)
      
      #  Save outlier analysis report , 
      return(wb_mergedata)
      
      
    }
    ## Download the merge file
    ## event after clicking the download merge file button
    output$downloadMerge_btn <- downloadHandler(
      filename = function() {
        paste(input$mergeName_text, sep = "")
      },
      content = function(file) {
        
        
        
        saveWorkbook(merge_results(), file, overwrite = TRUE)
        # write.csv(df_temp,file)
        # saveWorkbook(global_wb, file, overwrite = TRUE)
        
      }
    )
    
    ## Download the config file
    ## event after clicking the download config file button
    output$downloadConfig_btn <- downloadHandler(
      filename = function() {
        paste(input$configName_text, sep = "")
      },
      content = function(file) {
        
        df_prev_report_detail$is_previous_report = TRUE
        df_batch_detail$is_previous_report = FALSE
        
        df_config = rbind.fill(df_batch_detail,  df_prev_report_detail)
        df_config$Use=NULL
        write.csv(df_config, file,row.names = FALSE)
        # write.csv(df_temp,file)
        # saveWorkbook(global_wb, file, overwrite = TRUE)
        
      }
    )
       
}

# Run the application 
shinyApp(ui = ui, server = server)
