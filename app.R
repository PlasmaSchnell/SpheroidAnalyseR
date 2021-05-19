# File name: app.R (https://github.com/markdunning/SpheroidAnalyseR)
# Author: Yichen He
# Date: May-2021
# This is a Shiny web app that :
#   1. Allows users to upload spheroid raw files, plate layout and Treatment files
#   2. Removes outliers on raw files based on thresholds and Robust Z-score. Users can download the outlier analysis report
#   3. Merges outlier analysis reports into the final merged file
                                                                  
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
library(shiny)
library(shinyjs)
library(DT)

source('SpheroidAnalyseR_lib.R')


manual_outlier_selections = outer(LETTERS[1:8], paste0(1:12,'.'), FUN = "paste")
dim(manual_outlier_selections) =NULL

value_selections = list("Area", "Diameter", "Volume", "Perimeter" , "Circularity")

#################### 
##### Shiny App ####
####################

# Define UI for application that draws a histogram
ui <- navbarPage(#"SpheroidAnalyseR",
  
  
  # Title
  title=div(img(src="leeds_logo.png",style = "margin:-30px 10px"),"SpheroidAnalyseR"),
  windowTitle = HTML("SpheroidAnalyseR"),
  
  ## tab panels
  #   Data input
  #   Outlier removal
  #   Merging
  #   Plotting
  #   Help
    tabPanel("Data Input",
    
    # Sidebar with a slider input for number of bins 
    sidebarLayout(
        sidebarPanel(
          # Style
          tags$head(
            tags$style(
              HTML('
              * {
                font-family: Arial, sans-serif !important;
              }

            .navbar {min-height:58px !important;}

            .navbar-nav {
                font-size: 18px;
                text-align: center;
              }

            .navbar-nav > li > a, .navbar-brand {
            min-height:58px 
            }

            .navbar-brand{
                  font-size: 24px;
                  font-weight: bold;
            }')
            )),
          
          
          #Side bar layout
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
                            # style = "background-color:#c0ecfa;",
                            style = "background-color:#F6F1E4",
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

                    # strong("Plate layout after pre-sreen outlier removal (if applied)",id='title_pre'),
                    # 
                    #  plotOutput("outlierPlot"),
                     
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
               
               # h4("Only available when all files have been processed"),
               helpText("Click merge button to generate merged report. Plots can be added to the report. "),
               actionButton("merge_btn", "Merge"),
               textOutput("text_create_merge"),

               
               
               textInput("mergeName_text", "Merged file name", value = paste0("merge_file_", Sys.Date(),".xlsx")),
               
               helpText("Files can only be downloaded if all selected raw data have been processed"),
               downloadButton("downloadMerge_btn", "Download the merged file")
               # strong(""),
               # textInput("configName_text", "Config file name", value = paste0("config_file_", Sys.Date(),".csv")),
               # downloadButton("downloadConfig_btn", "Download the config file")

               ),
             
             mainPanel(
               h4("Raw files"),
               DT::dataTableOutput("merge_file"),
               h4("Previous reports"),
               helpText("Previous report can be uploaded for merging"),
               checkboxInput("processed_file_chk","Use previous reports",value = FALSE),

               # checkboxInput("override","Apply Manual overrides?",value = FALSE),

               conditionalPanel(condition ="input.processed_file_chk==1",
                                #- Values above zero
                                fileInput(inputId = "processed_data","Choose previous reports",accept=".xlsx",buttonLabel = "Browse", multiple = TRUE),
                                textOutput("text_processed")
               ),
               DT::dataTableOutput("processed_data_table")
               # tableOutput("merge_file")
               
             )
    ),
    
    tabPanel("Plotting",
             sidebarPanel(
                h4("Choose plot configuration and add plot"),
                selectInput("sel_plot",label="Choose plot types",
                            c("Bar","Point","Dot","Box")),
                
                selectInput("sel_y",label="Choose value for y axis",
                            c()),
                
                selectInput("sel_x",label="Choose value for x axis",
                            c()),
                
                selectInput("sel_gp_1",label="Choose value for colouring",
                            c()),
                selectInput("sel_gp_2",label="Choose value 2 for colouring",
                            c()),
                textInput("plotName_text", "Name of the plot", value = "Plot"),
                
                checkboxInput("bw_plot_chk","Black & white plot",value = FALSE),
                
                actionButton("plt_m_plt_btn", "Plot"),
                actionButton("add_m_plt_btn", "Add to the report"),
                textOutput("text_merge_plt")
               ),
             
             mainPanel( 
               h4("Plot"),
                plotOutput("mergePlot"))
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
    shinyjs::disable("downloadData_btn")
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

    TF_apply_thresholds <-TRUE
    
    df_batch_detail <-data.frame()
    df_prev_report_detail<-data.frame()
    
    n_cb_raw<-0
    n_cb_prev<-0
    
    df_output_list <- FALSE
    df_origin_list <- FALSE
    df_spheroid_list <- FALSE
    
    layout_file_error = FALSE
    treat_file_error = FALSE
    
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

      }
    })
    

    
    ## update the raw data preview in the input panel
    ## event: After uploading the raw data
    output$data_preview <- DT::renderDataTable(
      {
        updateNumericInput(session, "z_score", value = 1.96)

        
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

        if(all(error_list)==FALSE){
          layout_file_error<<-TRUE
        }else{
          layout_file_error<<-FALSE
        }
        
        validate(need(error_list[1] == TRUE,"Dimension error\nCheck if the file is 8x12 (row names and column names excluded)"),
                 need(error_list[2] == TRUE,"Column names error\nplease refer to the template"),
                 need(error_list[3] == TRUE,"Row names error\nplease refer to the template"),
                 need(error_list[4] == TRUE,"Value error\nCheck values in the layout file"))
        
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
        
        if(all(error_list)==FALSE){
          treat_file_error<<-TRUE
        }else{
          treat_file_error<<-FALSE
        }
        
        validate(need(error_list[1] == TRUE,"Dimension error\nCheck if the file is 6x8 (row names and column names excluded)"),
                 need(error_list[2] == TRUE,"Column names error\nplease refer to the template"),
                 need(error_list[3] == TRUE,"Row names error\nplease refer to the template"),
                 need(error_list[4] == TRUE,"Value error\nCheck values in the treatment file"))

        
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
    
    
    ### Functions for plotting layout & treatment preview
    draw_layout<-function(df_layout){
      
      df_layout$Col = as.factor(df_layout$Col)
      
      df_layout$Well.Name = paste0(df_layout$Row,df_layout$Col)
      ggplot(df_layout, aes(x = Col, y = Row, label=Well.Name)) +
        geom_tile() +
        geom_text()  +scale_y_discrete(limits = rev) +scale_x_discrete(position = "top") 
    }
    draw_layout_with_treat <-function(df_layout, df_treat, value){
      layout <- left_join(df_layout,df_treat)
      
      layout$Col = as.factor(layout$Col)
      layout$Well.Name = paste0(layout$Row,layout$Col)
      
      ggplot(layout, aes_string(x = "Col", y = "Row",fill=value,label="Well.Name")) +
        geom_tile() +
         geom_text() + scale_y_discrete(limits = rev)+scale_x_discrete(position = "top")
      
    }
    
    
    #### outlier removal panel ####
  
    ## remove outliers of the selected raw file in the outlier removal panel
    ## event: click the outlier removal button

    output_report <- eventReactive(input$outlier_btn, {
      print("file errors")
      print(treat_file_error)
      print(layout_file_error)
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
        need(input$treat_data, "please upload the treatment file"),
        
        need(treat_file_error==FALSE, "Treatment file has errors"),
        need(layout_file_error==FALSE, "Layout file has errors")
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
      
      ## Used in testing the outlier removal method
      
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
      
      TF_apply_thresholds <<- input$pre_screen
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
      checksumOK <- TRUE
      
      if (checksum != datalength)  
      {checksumOK <- FALSE} 
      

      validate(
        need(checksumOK, paste0("Number of rows of data in raw file (",
                                datalength,
                                ") does not match the plate setup checksum ("
                                ,checksum,") \nPlease upload "))
      )
      
      # #- do the check
      # if (checksumOK == "FALSE")
      # {
      #   message("#### Warning : plate setup checksum mismatch. ####","\n")
      #   message("Number of rows of data in raw file (",datalength  ,") does not match the plate setup checksum (",checksum,") \n")
      #   message("Output file corruption may occur - please recheck your plate setup","\n")
      #   message("Processing halted","\n")
      #   message("\n")
      #   sink()
      #   stop()
      # }
      # #- do the check
      # if (checksumOK == "TRUE")
      #   
      # {
      #   message("** Plate setup checksum is valid for raw dataset ","\n")
      #   message("Number of rows of data in raw file (",datalength  ,") matches the checksum (",checksum,") \n")
      # }
      
      
      
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
        
        
        pre_threshold_cond<-Spheroid_data_1$Area < TH_Area_max & Spheroid_data_1$Area > TH_Area_min &
          Spheroid_data_1$Diameter<TH_Diameter_max & Spheroid_data_1$Diameter>TH_Diameter_min &
          Spheroid_data_1$Circularity<TH_Circularity_max & Spheroid_data_1$Circularity>TH_Circularity_min &
          Spheroid_data_1$Volume<TH_Volume_max & Spheroid_data_1$Volume>TH_Volume_min &
          Spheroid_data_1$Perimeter<TH_Perimeter_max & Spheroid_data_1$Perimeter>TH_Perimeter_min
        
        # AND version
        Spheroid_data_1$Area <- ifelse(pre_threshold_cond,  Spheroid_data_1$Area, NA)   
        Spheroid_data_1$Diameter <- ifelse(pre_threshold_cond, Spheroid_data_1$Diameter, NA)
        Spheroid_data_1$Circularity <- ifelse(pre_threshold_cond, Spheroid_data_1$Circularity, NA)
        Spheroid_data_1$Volume <- ifelse(pre_threshold_cond, Spheroid_data_1$Volume, NA)
        Spheroid_data_1$Perimeter <- ifelse(pre_threshold_cond, Spheroid_data_1$Perimeter, NA)
        
        # Previous OR version
        
        # Spheroid_data_1$Area <- ifelse(Spheroid_data_1$Area < TH_Area_max & Spheroid_data_1$Area > TH_Area_min,  Spheroid_data_1$Area, NA)   
        # Spheroid_data_1$Diameter <- ifelse(Spheroid_data_1$Diameter<TH_Diameter_max & Spheroid_data_1$Diameter>TH_Diameter_min, Spheroid_data_1$Diameter, NA)
        # Spheroid_data_1$Circularity <- ifelse(Spheroid_data_1$Circularity<TH_Circularity_max & Spheroid_data_1$Circularity>TH_Circularity_min, Spheroid_data_1$Circularity, NA)
        # Spheroid_data_1$Volume <- ifelse(Spheroid_data_1$Volume<TH_Volume_max & Spheroid_data_1$Volume>TH_Volume_min, Spheroid_data_1$Volume, NA)
        # Spheroid_data_1$Perimeter <- ifelse(Spheroid_data_1$Perimeter<TH_Perimeter_max & Spheroid_data_1$Perimeter>TH_Perimeter_min, Spheroid_data_1$Perimeter, NA)
        
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
      shinyjs::enable("downloadData_btn")
      
      df_batch_detail<<-df_temp_batch_detail
      
      # if(all(df_batch_detail$Processed[shinyValue('cb_raw_', n_cb_raw)]) & any(df_batch_detail$Processed)){
      #   shinyjs::enable("downloadMerge_btn")
      #   shinyjs::enable("downloadConfig_btn")
      # }else{
      #   shinyjs::disable("downloadMerge_btn")
      #   shinyjs::disable("downloadConfig_btn")
      # }
      

      
      message("All processing completed successfully.", "\n")
      message("\n") 
      
      
      
      # p_list = list(p_Area_new,p_Area_dotplot_new, 
      #             p_Diameter_new,p_Diameter_dotplot_new,
      #             p_Circularity_new,p_Circularity_dotplot_new,
      #             p_Volume_new,p_Volume_dotplot_new,
      #             p_Perimeter_new,p_Perimeter_dotplot_new)
      
      check_empty_cells_1 = which(Spheroid_data[,c("Spheroid_Area.TD.Area","Spheroid_Area.TD.FillArea.Mean","Spheroid_Area.TD.Perimeter.Mean" ,    
                                      "Spheroid_Area.TD.Circularity.Mean","Spheroid_Area.TD.Count","Spheroid_Area.TD.EqDiameter.Mean" ,   
                                      "Spheroid_Area.TD.VolumeEqSphere.Mean" ,"Spheroid_Area.TD.Roughness.Mean","Spheroid_Area.TD.ShapeFactor.Mean" )] ==0 
                     , arr.ind=TRUE)
      
      
      
      check_empty_cells_2 = which(is.na(Spheroid_data[,c("Spheroid_Area.TD.Area","Spheroid_Area.TD.FillArea.Mean","Spheroid_Area.TD.Perimeter.Mean" ,    
                                             "Spheroid_Area.TD.Circularity.Mean","Spheroid_Area.TD.Count","Spheroid_Area.TD.EqDiameter.Mean" ,   
                                             "Spheroid_Area.TD.VolumeEqSphere.Mean" ,"Spheroid_Area.TD.Roughness.Mean","Spheroid_Area.TD.ShapeFactor.Mean" )]), arr.ind=TRUE)
      
      
      check_empty_cells = rbind(check_empty_cells_1, check_empty_cells_2)
      
      
      empty_rows = unique(check_empty_cells[,1])
      
      if(length(empty_rows)!=0){
        empty_cells = paste0(Spheroid_data$Well.Name[empty_rows],collapse = " ")
        paste0("Report generated, please download the report. These cells may have empty measurements: ",empty_cells, ". Please Check the raw file" )
      }
      else{
        "Report generated, please download the report"
      }

    })
    
    ## renew the text in the outlier removal tab
    ## event: after outlier removal
    output$textStatus<-  renderText({
      output_report()
    })


      
    ## function for updating plots in the outlier removal panel
      update_plots = function(df, value){
        # output$outlierPlot <- renderPlot({
        #   validate(need(df,"Please run the outlier removal"))
        #   draw_outlier_plot(df,value)
        #   
        # })
        
        output$resultPlot <- renderPlot({
          validate(need(df,"Please run the outlier removal"))
          draw_z_score_outlier_plot(df,value, TF_apply_thresholds = TF_apply_thresholds)
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
          
          if(df_batch_detail$Processed[name_id] == TRUE){
            shinyjs::enable("downloadData_btn")
          }else{
            shinyjs::disable("downloadData_btn")
          }
          
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
          # show("outlierPlot")
          show("select_outlier_values")
          # shinyjs::show("title_pre")
        }else{
          # shinyjs::hide("title_pre")
          # hide("outlierPlot")
          hide("select_outlier_values")
        }
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
          }else{
            use_previous_report<<- TRUE
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

      

      full_data_A  <- full_data[with(full_data , order(Days_since_treatment, T_I)),]
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
                                                                 which(colnames(full_data_B)=="Time_of_treatment"),
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
      
      ## add the plot part
      if(length(l_merge_plts)>0){
        addWorksheet(wb_mergedata , "plots", gridLines = FALSE)
        for(i_plt in 1:length(l_merge_plts)){
          
          png(paste0(i_plt,".png"),width = 1024, height = 400, res=150)
          print(length(l_merge_plts)) 
          print(l_merge_plts[[i_plt]]) 
          dev.off()
          insertImage(wb_mergedata, 2, paste0(i_plt,".png"), width = 9, height = 3.5 , startRow = (i_plt-1)*20+1,startCol = 'A')
          
          # file.remove(paste0(i_plt,".png"))
          
        }
      }

      
      
            
      df_prev_report_detail$is_previous_report = TRUE
      df_batch_detail$is_previous_report = FALSE
      
      df_config = rbind.fill(df_batch_detail,  df_prev_report_detail)
      df_config$Use=NULL
      addWorksheet(wb_mergedata , "Merge Config", gridLines = FALSE)
      
      if(length(l_merge_plts)>0){
        writeDataTable(wb_mergedata, 3, df_config, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = FALSE)
      }else{
        writeDataTable(wb_mergedata, 2, df_config, startRow = 1, startCol = 1, tableStyle = "TableStyleLight8",  rowNames = FALSE, keepNA = FALSE)
      }
      
      
      #  Save outlier analysis report , 
      return(wb_mergedata)
      
      
    }
    
    
    df_test=data.frame(a=c(1,2,3,4),b=c(5,6,7,8))
    
    l_merge_plts = list()
    merge_plt=NA
    full_data=NA
    

    
    ## Render the plot on the merge tab
    output$mergePlot <- renderPlot({
      print("plot merge")
      
      plt <- plot_merge_plot()
      merge_plt<<-plt
      plt
      
    })
    
    ## Help funtion on output$mergePlot <- renderPlot 
    plot_merge_plot <- eventReactive(input$plt_m_plt_btn,{
      # Check if there is a merge dataframe
      validate(need(!is.na(full_data),"Please make a plot before adding"))
      

      validate(need(!(input$sel_gp_1=="NULL" & input$sel_gp_2 !="NULL"), "Please select the first colouring variable"))



      full_data$Days_since_treatment = floor(full_data$Days_since_treatment)
      full_data$Hours_since_treatment = floor(full_data$Hours_since_treatment)
      
      # Get variables from selections
      y_var=  paste0(input$sel_y,"_Mean")
      y_se = paste0(input$sel_y,"_SE")
      ymin_var = paste0(y_var,"-",y_se)
      ymax_var = paste0(y_var,"+",y_se)
      
      
      x_var = input$sel_x

      


      gp_full_data=full_data
     
      
      if(input$sel_gp_2 !="NULL"){
        # Check if variables are duplicated
        validate(need(input$sel_x!=input$sel_gp_1 &
                        input$sel_x!=input$sel_gp_2 &
                        input$sel_gp_2!=input$sel_gp_1,"Please select non-duplicate variables."))
        gp_var = paste0('interaction(',input$sel_gp_1,',', input$sel_gp_2, ')')
        
        full_data[input$sel_gp_2] = as.factor(full_data[,input$sel_gp_2])


        gp_full_data = gp_full_data %>% dplyr::group_by_(x_var,input$sel_gp_1,input$sel_gp_2)%>% summarise_(Mean = paste0("mean(", y_var,")"),
                                                                    SE = paste0("mean(", y_se,")"))
        
        gp_var = paste0('factor(' , gp_var , ')')
        
      }else if (input$sel_gp_1 !="NULL"){
        # Check if variables are duplicated
        validate(need(input$sel_x!=input$sel_gp_1 ,"Please select non-duplicate variables."))
        full_data[input$sel_gp_1] = as.factor(full_data[,input$sel_gp_1])
        gp_var = input$sel_gp_1
        
        gp_full_data = gp_full_data %>% dplyr::group_by_(x_var,input$sel_gp_1)%>% summarise_(Mean = paste0("mean(", y_var,")"),
                                                                    SE = paste0("mean(", y_se,")"))
        gp_var = paste0('factor(' , gp_var , ')')
      }else{
        gp_var = "NULL"
        
        gp_full_data = gp_full_data %>% dplyr::group_by_(x_var)%>% summarise_(Mean = paste0("mean(", y_var,")"),
                                                                    SE = paste0("mean(", y_se,")"))
        
      }
      
      colnames(gp_full_data)[colnames(gp_full_data) == 'Mean'] <-y_var
      colnames(gp_full_data)[colnames(gp_full_data) == 'SE'] <-y_se
      
      
      # gp_var = "Conc_1"
      # 
      # full_data[,gp_var] = as.factor(full_data[,gp_var] )
      
      
      #Plot different types of plots using ggplot
      
      if(input$sel_plot=="Point"){

        # gp_full_data[,x_var] = as.inte(gp_full_data[,x_var]) 

        
        
        p=ggplot(gp_full_data, aes_string(x=x_var, y=y_var, group=gp_var, color=gp_var)) + 
          geom_errorbar(aes_string(ymin=ymin_var, ymax=ymax_var), width=.1) +
          geom_line() + geom_point(aes_string(shape=gp_var))


      }
      
      if(input$sel_plot=="Dot"){
        full_data[,x_var] = as.factor(full_data[,x_var])
        
        p=ggplot(full_data, aes_string(x=x_var, y=y_var)) + 
          geom_dotplot(binaxis='y', stackdir='center') +
           stat_summary(fun.y=mean, geom="point", shape=18,
                            size=3, color="red")
      }
      
      if(input$sel_plot=="Bar"){

        gp_full_data[,x_var]=lapply(gp_full_data[,x_var], 
                         as.factor)
        
        # gp_full_data[,x_var] = as.factor(gp_full_data[,x_var]) 
        p=ggplot(gp_full_data, aes_string( x=x_var,y=y_var, group=gp_var, fill=gp_var)) + 
          geom_bar(stat="identity", position=position_dodge())+
          geom_errorbar(aes(ymin=Area_Mean-Area_SE, ymax=Area_Mean+Area_SE),width=.2,
                        position=position_dodge(.9)) 

      }

      if(input$sel_plot=="Box"){
        
        full_data_2 = full_data
        full_data_2[,x_var] = as.factor(full_data_2[,x_var])
        
        p=ggplot(full_data_2, aes_string(x=x_var, y=y_var,fill=gp_var)) + 
          geom_boxplot()
      
      }
      

      if(input$sel_plot=="Box" | input$sel_plot=="Bar"){
        if(input$bw_plot_chk == TRUE){
          p=p+scale_fill_grey()
        }else{
          p=p+scale_fill_colorblind()
        }
      }

      if(input$sel_plot=="Point"){
        if(input$bw_plot_chk == TRUE){
          p=p+scale_color_grey()
        }else{
          p=p+scale_color_colorblind()
        }
      }

      
      
      
      p=p+theme_minimal() + ggtitle(input$plotName_text) +
        theme(axis.title=element_text(size=16),
              axis.text=element_text(size=14),
              legend.title=element_text(size=16),
              legend.text=element_text(size=14)
              )
      
      return(p)
      # ggplot(df_test)+geom_point( aes(x=a,y=b))
    })
    
    
    output$text_merge_plt<-renderText({
      add_merge_plot()
    })
    
    add_merge_plot <- eventReactive(input$add_m_plt_btn,{
      validate(need(!is.na(merge_plt),"Please make a plot before adding"))
      l_merge_plts[[length(l_merge_plts)+1]] <<- merge_plt
      
      "Plot added"
      
    })
    
    

    
    
    output$text_create_merge<-renderText({
      make_merge_data()
    })
    
    ## create the merged data
    ## event after clicking the merge button
    
    make_merge_data <- eventReactive(input$merge_btn,{
      validate(need(length(df_batch_detail)>0 | (input$processed_file_chk==TRUE & use_previous_report==TRUE),
                    "No files have been uploaded"))

      # print(any(shinyValue('cb_raw_', n_cb_raw)) | 
      #         ((input$processed_file_chk==TRUE & use_previous_report==TRUE) | input$processed_file_chk==FALSE))
      
      validate(need(any(shinyValue('cb_raw_', n_cb_raw)) | 
                      ((input$processed_file_chk==TRUE & use_previous_report==TRUE) | input$processed_file_chk==FALSE),
                    "No files have been selected"))
      
      # print((input$processed_file_chk==TRUE & use_previous_report==TRUE) | input$processed_file_chk==FALSE)
    
      validate(need((length(df_batch_detail)==0 | (all(df_batch_detail$Processed[shinyValue('cb_raw_', n_cb_raw)]) & any(df_batch_detail$Processed))),
                    "Please check all selected files have been processed"))
      
      result_list = list()
      
      
      raw_file_cb_result = shinyValue('cb_raw_', n_cb_raw)
      print(raw_file_cb_result)
      if(length(df_batch_detail$File_name)>0){
        for(i in 1:length(df_batch_detail$File_name)){
          if (raw_file_cb_result[i]==TRUE){
            result_list = append(result_list,
                                 list(generate_merge_result(df_output_list[[i]],df_spheroid_list[[i]], df_batch_detail$File_name[i],global_df_treat)))
          }
          
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
        
        d1 = result_list[[1]]
        
      }
      
      
      full_data <<- Reduce(function(x,y) {merge(x,y,all = TRUE)}, result_list)
      
      #Update the selection for plotting
      y_values =c("Area", "Diameter", "Volume", "Perimeter" , "Circularity")

      updateSelectInput(session, "sel_y",
                        choices = y_values,
                        selected = head(y_values, 1))
    
      
      
      y_names = c(paste0(y_values, "_SE"),paste0(y_values, "_Mean"))
      
      ori_col_names = colnames(full_data)
      x_values = ori_col_names[!(ori_col_names %in% y_names)]
      x_values=x_values[x_values!="T_I"]
      
      

      updateSelectInput(session, "sel_x",
                        choices = x_values,
                        selected = head(x_values, 1))  
      
      gp_values = c("NULL",x_values)
      updateSelectInput(session, "sel_gp_1",
                        choices = gp_values,
                        selected = head(gp_values, 1))  

      updateSelectInput(session, "sel_gp_2",
                        choices = gp_values,
                        selected = head(gp_values, 1))  
      
      # enable the download button
      shinyjs::enable("downloadMerge_btn")
      shinyjs::enable("downloadConfig_btn")
      
      
      "Merge result created. Plots can be added and ready for download."
      
    })
    

    
    
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
