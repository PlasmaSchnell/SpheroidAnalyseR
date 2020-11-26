#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

# library(shiny)
# library(ggplot2)
# library(dplyr)


library(tidyverse)
library(ggthemes)
library(gridExtra)
library(readxl)
library(openxlsx)
library(plotrix)
library(writexl)
library(ggpubr)


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
                     ,conditionalPanel(condition="input.override==1",
                                       selectInput("manual_outliers",label="Outliers to be removed",paste(LETTERS[1:8],1:12),multiple=TRUE,selectize=TRUE)
                     )
                     ,actionButton("outlier_btn", "Remove outliers")
                 ),
                 # Show a plot of the generated distribution
                 mainPanel(
                     h2("Plate layout"),
                     plotOutput("outlierPlot"),
                     plotOutput("AreaPlot"),
                     plotOutput("DiaPlot"),
                     plotOutput("CirPlot"),
                     plotOutput("VolPlot"),
                     plotOutput("PeriPlot"),

                     
                     verbatimTextOutput("textOutput")
                     
                 )
             ) 
    )
)


# Define server logic required to draw a histogram
server <- function(input, output) {

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
            treatments
        }

    }
    )
    
    output$show_layout <- renderPlot(
        {
            layout <- readLayout()
            treatments <- readTreatment()

            if(!is.null(layout)){
                if(!is.null(treatments)){
                    layout <- left_join(layout,treatments) %>%
                        tidyr::pivot_longer(Index:last_col(), names_to="Factor",values_to="Label")
                    ggplot(layout, aes(x = Row, y = Col, fill = Label)) + geom_tile() + scale_y_continuous(breaks=1:12) + facet_wrap(~Factor,ncol=1
                                                                                                                                            )
                } else{
            layout <- mutate(layout, Cell = ifelse(!is.na(Index), paste(Row,Col),""))
            ggplot(layout, aes(x = Row, y = Col, fill = Index,label=Cell)) + geom_tile() + geom_text() + scale_y_continuous(breaks=1:12)
                }
            }
        }

    )

    detectOutliers <- function(layout){
        
        rand_outliers <- sample(1:nrow(layout),5)
        rand_outliers <- ifelse(1:nrow(layout) %in% rand_outliers, "X","")
        layout %>% mutate(Outlier = rand_outliers,Outlier = ifelse(is.na(Index),NA,Outlier))
    }
    
    output$outlierPlot <- renderPlot(
    ### a dummy function for now that will just show some randomly selected outliers    
    {
            
        rawData <- readRawData()
        layout <- readLayout()
        
        
        if(!is.null(rawData) & !is.null(layout)){
            
            layout <- detectOutliers(layout) %>% 
                mutate(Cell = ifelse(Outlier=="X",paste(Row,Col),""))
            ## detectOutliers will add an extra column Outlier indicating if cell is an outlier
            ggplot(layout, aes(x = Row, y = Col, fill = Outlier,label=Cell)) + geom_tile() + scale_y_continuous(breaks=1:12) + scale_fill_manual(values=c("white","red","grey")) + geom_text()
            
        
        }
            
        }
    )
    
    # randomVals <- eventReactive(input$outlier_btn, {
    #   
    #   
    #   validate(
    #     need(is.numeric(input$area_threshold_low) & input$area_threshold_low>0, "Please input a positive area threshold (numeric)"),
    #     need(is.numeric(input$area_threshold_high) & input$area_threshold_high>0, "Please input a positive area threshold (numeric)"),
    #     
    #     need(is.numeric(input$diam_threshold_low) & input$diam_threshold_low>0, "Please input a positive diameter threshold (numeric)"),
    #     need(is.numeric(input$diam_threshold_high) & input$diam_threshold_high>0, "Please input a positive diameter threshold (numeric)"),
    #     
    #     need(is.numeric(input$vol_threshold_low) & input$vol_threshold_low>0, "Please input a positive volume threshold (numeric)"),
    #     need(is.numeric(input$vol_threshold_high) & input$vol_threshold_high>0, "Please input a positive diameter threshold (numeric)"),
    #     
    #     need(is.numeric(input$perim_threshold_low) & input$perim_threshold_low>0, "Please input a positive Perimeter threshold (numeric)"),
    #     need(is.numeric(input$perim_threshold_high) & input$perim_threshold_high>0, "Please input a positive Perimeter threshold (numeric)"),
    #     
    #     need(is.numeric(input$circ_threshold_low) & input$circ_threshold_low>0, "Please input a positive Circularity threshold (numeric)"),
    #     need(is.numeric(input$circ_threshold_high) & input$circ_threshold_high>0, "Please input a positive Circularity threshold (numeric)")
    #     
    #   )
    #   
    #   file <- input$raw_data
    #   layout_file <- input$layout
    #   treat_file <- input$treat_data
    #   
    #   if(!is.null(file) && !is.null(layout_file) && !is.null(treat_file)){
    #     ext <- tools::file_ext(file$datapath)
    #     req(file)
    #     validate(need(ext == "xlsx","Please upload an xlsx file"))
    #     Spheroid_data <- suppressMessages(read_excel(file$datapath, "JobView",col_names=TRUE, .name_repair = "universal"))
    #     
    #     
    #     ext <- tools::file_ext(treat_file$datapath)
    #     req(treat_file)
    #     validate(need(ext == "csv","Please upload a layout in csv format"))
    #     df_treat <- readr::read_csv(treat_file$datapath) %>% 
    #       mutate_all(as.factor)
    #     
    #     ext <- tools::file_ext(layout_file$datapath)
    #     req(layout_file)
    #     validate(need(ext == "csv","Please upload a layout in csv format"))
    #     df_setup <- readr::read_csv(layout_file$datapath)
    #     
    #   }
    #   
    #   message(dim(Spheroid_data))
    #   message(dim(df_treat))
    #   return("as")
    # }
    # )
    
    randomVals <- eventReactive(input$outlier_btn, {

      
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
        need(is.numeric(input$circ_threshold_high) & input$circ_threshold_high>0, "Please input a positive Circularity threshold (numeric)")
   
      )
      
      # file <- input$raw_data
      # layout_file <- input$layout
      # treat_file <- input$treat_data
      # 
      # if(!is.null(file) && !is.null(layout_file) && !is.null(treat_file)){
      #   message("s")
      # }
      
      # runif(input$diam_threshold_low)
      
      #- after validating, use the thresholding to remove outliers
      #- do the thresholding and generate the wells
      
      
      file <- input$raw_data
      layout_file <- input$layout
      treat_file <- input$treat_data
      
      #- read from the input
      
      setup_file = "layout_example.csv"
      treatment_file = "treatment_example.csv"
      
      input_file = "GBM58 - for facs - day 3.xlsx"
      
      df_setup =read.csv(setup_file)
      df_treat = read.csv(treatment_file)
      #- import raw data
      Spheroid_data <- suppressMessages(read_excel(input_file, "JobView",col_names=TRUE, .name_repair = "universal"))
      
      
      
      RobZ_LoLim <- input$z_low
      RobZ_UpLim <- input$z_high
      
      TF_apply_thresholds <- input$pre_screen
      TF_outlier_override <- input$override
      TF_copytomergedir <- FALSE
      

      merge_output_file = "temp_merge_file.xlsx"
      report_output_file = "report_file.xlsx"
      
      #-  TODO check duplicate merge and report file
      

      
      # read file from the file browser
      if(!is.null(file) && !is.null(layout_file) && !is.null(treat_file)){
        ext <- tools::file_ext(file$datapath)
        req(file)
        validate(need(ext == "xlsx","Please upload an xlsx file"))
        Spheroid_data <- suppressMessages(read_excel(file$datapath, "JobView",col_names=TRUE, .name_repair = "universal"))

        ext <- tools::file_ext(treat_file$datapath)
        req(treat_file)
        validate(need(ext == "csv","Please upload a layout in csv format"))
        df_treat <- readr::read_csv(treat_file$datapath) %>%
          mutate_all(as.factor)

        ext <- tools::file_ext(layout_file$datapath)
        req(layout_file)
        validate(need(ext == "csv","Please upload a layout in csv format"))
        df_setup <- readr::read_csv(layout_file$datapath)

      }
      
      
      
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
        message("Number of rows of data in raw file (",datalength  ,") does not match the plate setup checksum (",checksum,") on ",platesetupname,"\n")
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
        message("Number of rows of data in raw file (",datalength  ,") matches the checksum (",checksum,") on ",platesetupname,"\n")
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
      
      
      
      
      #### function for outlier removing
      
      cal_z_score = function(df_sph_treat, df_prev, varname){
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
      
      df_a = cal_z_score(df_sph_treat, df_sph_treat, "Area")
      df_d = cal_z_score(df_sph_treat,df_a, "Diameter")
      df_v = cal_z_score(df_sph_treat,df_d, "Volume")
      df_p = cal_z_score(df_sph_treat,df_v, "Perimeter")
      
      Sph_Treat_Robz_ADVPC =cal_z_score(df_sph_treat,df_p,"Circularity" )
      
      
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
      
      p_Perimeter2 <- ggplot()  + geom_point(data = Sph_Treat_ADVPC, aes(x= Col, y=Perimeter, colour = Perimeter_status), size = Pointsize,  show.legend=FALSE ) + 
        scale_x_continuous(limits=c(0,12),breaks = c(0,6,12))+
        facet_wrap(~Sph_Treat_ADVPC$Row, nrow=1) +
        ylab(Perimeterlabel)+xlab('Column') +  labs(title= " Perimeter data by Row with outliers in red") + theme_bw() + theme(plot.title = element_text(size = 10)) +
        scale_colour_manual(values = cols) 
      
      p_list = list(p_Area_new,p_Area_dotplot_new, 
                  p_Diameter_new,p_Diameter_dotplot_new,
                  p_Circularity_new,p_Circularity_dotplot_new,
                  p_Volume_new,p_Volume_dotplot_new,
                  p_Perimeter_new,p_Perimeter2)
      

    })
    
    output$AreaPlot <- renderPlot({
      grid.arrange(grobs=randomVals()[1:2] , ncol=2)
    })  
    
    output$DiaPlot <- renderPlot({
      grid.arrange(grobs=randomVals()[3:4] , ncol=2)
    })  
    output$CirPlot <- renderPlot({
      grid.arrange(grobs=randomVals()[5:6] , ncol=2)
    })  
    output$VolPlot <- renderPlot({
      grid.arrange(grobs=randomVals()[7:8] , ncol=2)
    })  
    output$PeriPlot <- renderPlot({
      grid.arrange(grobs=randomVals()[9:10] , ncol=2)
    })  
    
  
    
    # output$textOutput <- renderPrint({ randomVals() }) 
    
    #- validating thresholds
    # validate_thresholds <- reactive({
    #   validate(
    #     need(is.numeric(input$area_threshold_low) & input$area_threshold_low>0, "Please input a positive area threshold (numeric)"),
    #     need(is.numeric(input$area_threshold_high) & input$area_threshold_high>0, "Please input a positive area threshold (numeric)"),
    #     
    #     need(is.numeric(input$diam_threshold_low) & input$diam_threshold_low>0, "Please input a positive diameter threshold (numeric)"),
    #     need(is.numeric(input$diam_threshold_high) & input$diam_threshold_high>0, "Please input a positive diameter threshold (numeric)"),
    #     
    #     need(is.numeric(input$vol_threshold_low) & input$vol_threshold_low>0, "Please input a positive volume threshold (numeric)"),
    #     need(is.numeric(input$vol_threshold_high) & input$vol_threshold_high>0, "Please input a positive diameter threshold (numeric)"),
    #     
    #     need(is.numeric(input$perim_threshold_low) & input$perim_threshold_low>0, "Please input a positive Perimeter threshold (numeric)"),
    #     need(is.numeric(input$perim_threshold_high) & input$perim_threshold_high>0, "Please input a positive Perimeter threshold (numeric)"),
    #     
    #     need(is.numeric(input$circ_threshold_low) & input$circ_threshold_low>0, "Please input a positive Circularity threshold (numeric)"),
    #     need(is.numeric(input$circ_threshold_high) & input$circ_threshold_high>0, "Please input a positive Circularity threshold (numeric)")
    #   )
    #   
    # })
    # output$textOutput <- renderPrint({ validate_thresholds() }) 
    
}

# Run the application 
shinyApp(ui = ui, server = server)
