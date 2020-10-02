#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)
library(ggplot2)
library(dplyr)

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
                     checkboxInput("pre_screen","Apply Pre-screen thresholds?",value = FALSE),
                     checkboxInput("override","Apply Manual overrides?",value = FALSE),
                     conditionalPanel(condition ="input.pre_screen==1",
                     textInput("area_threshold_low","Area lower limit",value=1),
                     textInput("area_threshold_high","Area higher limit",value=1630000),
                     textInput("diam_threshold_low","Diameter lower limit",value=100),
                     textInput("diam_threshold_high","Diameter higher limit",value=1440),
                     textInput("vol_threshold_low","Volume lower limit",value=1000),
                     textInput("vol_threshold_high","Volume higher limit",value=1440),
                     textInput("perim_threshold_low","Perimeter lower limit",value=100),
                     textInput("perim_threshold_high","Perimeter higher limit",value=6276),
                     textInput("circ_threshold_low","Circularity lower limit",value=0.01),
                     textInput("circ_threshold_high","Circularity higher limit",value=1)
                     )
                     ,conditionalPanel(condition="input.override==1",
                                       selectInput("manual_outliers",label="Outliers to be removed",paste(LETTERS[1:8],1:12),multiple=TRUE,selectize=TRUE)
                     )
                 ),
                 # Show a plot of the generated distribution
                 mainPanel(
                     h2("Plate layout"),
                     plotOutput("outlierPlot")
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
            layout <- readr::read_csv(layout_file$datapath) %>% 
                #layout <- readr::read_csv("layout_example.csv") %>% 
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
    
}

# Run the application 
shinyApp(ui = ui, server = server)
