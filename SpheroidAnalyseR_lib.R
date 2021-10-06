# File name: SpheroidAnalyseR_lib.R (https://github.com/markdunning/SpheroidAnalyseR)
# Author: Yichen He
# Date: May-2021 
# Logic functions behind SpheroidAnalyseR shiny app. 
# Most of functions are based on Joe Wilkinson's original codes.
#
# Project:          Spheroid Analysis                                     
#                   Rhiannon Barrow, Joe Wilkinson, Mark Dunning, Dr Lucy Stead                       
#                   Glioma Genomics                                       
#                   Leeds Institute of Medical Research                   
#                   St James's University Hospital, Leeds  LS9 7TF 

library(tidyverse)
library(ggthemes)
library(gridExtra)
library(readxl)
library(openxlsx)
library(plotrix)
library(writexl)
library(ggpubr)



####### Check file formats ######

# Check the layout file
# Used in reading the layout file
check_layout = function(df){
  dim_check = all(dim(df) == c(8,12))
  
  colnames_check = tryCatch( all(colnames(df) == as.character(1:12)), error = function(e){F})
  rownames_check = tryCatch( all(row.names(df) ==c("A", "B", "C", "D",
                                                   "E", "F", "G", "H")), error = function(e){F})
  type_check = tryCatch( all(sapply(df, typeof) == "integer" |sapply(df, typeof) == "logical"),
                         error = function(e){F})
  
  error_list = c(dim_check, colnames_check, rownames_check,type_check)
  return(error_list)
}


# Check the treatment file
# Used in reading the treatment file
check_treatment = function(df){
  # dim_check = all(dim(df) == c(6,8))

  dim_check = (dim(df)[1]==6 & dim(df)[2]>=8)
  # colnames_check = colnames_check = tryCatch( all(colnames(df) == c("Index", "Treatment.Label",
  #                                                                   "Time_of_treatment", "Cell_line",
  #                                                                   "Passage_No", "Radiation_dosage",
  #                                                                   "Drug_1", "Conc_1"))
  #                                             , error = function(e){F})
  
  
  colnames_check = tryCatch( all(c("Index", "Treatment.Label",
                                                                    "Time_of_treatment", "Cell_line",
                                                                    "Passage_No", "Radiation_dosage",
                                                                    "Drug_1", "Conc_1") %in% colnames(df))
                                              , error = function(e){F})
  
  
  rownames_check = tryCatch(all(row.names(df) ==as.character(1:6)), error = function(e){F})
  
  type_check = TRUE
  # type_check = tryCatch(all(sapply(df, typeof) == c("integer","character","character","character","integer","integer","character","integer")),
  #                       error = function(e){F})
  
  
  error_list = c(dim_check, colnames_check, rownames_check,type_check)
  return(error_list)
}



# Check the raw file
# Used in reading the treatment file
check_raw = function(df){
  # dim_check = all(dim(df) == c(6,8))
  


  
  
  colnames_check = tryCatch( all(c("Jobrun.Finish.Time", "Well.Name",
                                                    "Spheroid_Area.TD.Area", "Spheroid_Area.TD.Perimeter.Mean",
                                                    "Spheroid_Area.TD.Circularity.Mean", "Spheroid_Area.TD.Count",
                                                    "Spheroid_Area.TD.EqDiameter.Mean",
                                                    "Spheroid_Area.TD.VolumeEqSphere.Mean") %in% colnames(df))
                                              , error = function(e){F})
  

  
  # rownames_check = tryCatch(all(row.names(df) ==as.character(1:6)), error = function(e){F})
  
  type_check = TRUE
  # type_check = tryCatch(all(sapply(df, typeof) == c("integer","character","character","character","integer","integer","character","integer")),
  #                       error = function(e){F})
  
  
  error_list = c(colnames_check,type_check)
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


##### function calculating outliers based on Z score
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
# draw_outlier_plot = function(df, value){
#   df$is_Outlier = is.na(df[,value])
#   levels(df$is_Outlier) = c(TRUE, FALSE)
#   colours = c("TRUE" = "red", "FALSE" = "white")
#   
#   ggplot(df, aes(x = Row, y = Col, fill = is_Outlier,label=Well.Name)) +
#     geom_tile() + 
#     scale_y_continuous(breaks=1:12) + 
#     scale_fill_manual(values=colours, drop=FALSE) + geom_text()
# }


draw_z_score_outlier_plot = function(df, value, TF_apply_thresholds=TRUE){
  
  
  value= paste0(value,"_status")
  ## fill un-used cells
  for(c in 1:12){
    for(r in LETTERS[1:8]){
      well_name = paste0(r,sprintf("%02d", c))
      if ( !(well_name %in% df$Well.Name)){
        
        df_row =df[1,]
        df_row$Col = c
        df_row$Row = r
        df_row$Well.Name = paste0(r,sprintf("%02d", c))
        
        df_row[,value] = '3'
        
        df<- rbind(df, df_row)
      }
    }
  }
  
  
  if(TF_apply_thresholds==TRUE){
    df[is.na(df[,value]), value] <- 2
  }else{
    df[is.na(df[,value]), value] <- 1
  }
  
  
  
  
  
  # df[,value] = factor(df[,value],levels = c("0","1","2"), ordered=TRUE)
  df[,value] = factor(df[,value],levels = c("0","1","2","3"), ordered=TRUE)
  
  
  # colours = c("1" = "red", "0" = "white", "2"='tan1')
  colours = c("1" = "red", "0" = "white", "2"='tan1' , '3' = 'grey')
  # col_labels = c("Not an outlier", "Outlier determined by robust Z score", "Outliers determined from pre-set thresholds" )
  col_labels = c("Not an outlier", "Outlier determined by robust Z score", 
                 "Outliers determined from pre-set thresholds",
                 "Unused cells")
  
  # colours = c(1 = "red", 0 = "white", NA='grey')
  
  df$Col = as.factor(df$Col)
  ggplot(df, aes_string(x = "Col", y = "Row", fill = value,label="Well.Name")) +
    geom_tile() + 
    scale_y_discrete(limits = rev) +scale_x_discrete(position = "top") +
    scale_fill_manual(labels = col_labels,
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
  
  # Sph_ADVPC$ETST.d <-difftime(Sph_ADVPC$Job.Date, as.Date(Sph_ADVPC$Time_Date, "%m/%d/%Y", origin="1900-01-01"), units= "days")
  Sph_ADVPC$Days_since_treatment <-floor(difftime(Sph_ADVPC$Job.Date,
                                                  as.Date(Sph_ADVPC$Time_of_treatment, "%m/%d/%Y", origin="1900-01-01"), units= "days"))
  
  Sph_ADVPC$Hours_since_treatment <-  floor(difftime(Sph_ADVPC$Job.Date,
                                                     as.Date(Sph_ADVPC$Time_of_treatment, "%m/%d/%Y", origin="1900-01-01"), units= "hours"))
  
  
  
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
  
  ADVPC_means$Days_since_treatment <-  floor(as.numeric(difftime(ADVPC_means$Job.Date, as.Date(ADVPC_means$Time_of_treatment, "%m/%d/%Y", origin="1900-01-01"), units= "days")))
  ADVPC_means$Hours_since_treatment <-  floor(as.numeric(difftime(ADVPC_means$Job.Date, as.Date(ADVPC_means$Time_of_treatment, "%m/%d/%Y", origin="1900-01-01"), units= "hours")))
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
  
  # for(temp_file_name in c("p11", "p12", "p21", "p22","p31", "p32","p41", "p42","p51", "p52")){
  #   
  #   file.remove(paste0(temp_file_name,".png"))
  # }
  
  
  
  
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
  
  print(ADVPC_means$Time_of_treatment)
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
  addStyle(wb, sheet = 12, datestyle, rows = 2:33 , cols = c(which(colnames(ADVPC_means)=="Job.Date"),
                                                             which(colnames(ADVPC_means)=="Time_of_treatment")), gridExpand = TRUE)
  
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

## Generate the merged dataframe.
generate_merge_result = function(df,Spheroid_data, rawfilename,df_treat){
  
  df$T_I <- as.numeric(df$T_I)
  Sph_Treat_ADVPC <- df[with(df, order(T_I)),]
  
  Sph_ADVPC <- Sph_Treat_ADVPC
  ADVPC_means<-Sph_ADVPC_2_ADVPC_means(Sph_ADVPC,df_treat,Spheroid_data,rawfilename)
  
  return(ADVPC_means)  
}