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
                     conditionalPanel(condition ="input.override==1",
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
                 ),
                 # Show a plot of the generated distribution
                 mainPanel(
                     h2("Plate layout with identified outliers will appear here"),
                     plotOutput("outlierPlot"),
                     h2("Editable table of outliers")
                 )
             ) 
    )
)


# Define server logic required to draw a histogram
server <- function(input, output) {

    output$data_preview <- renderTable(
        {
            file <- input$raw_data
            if(!is.null(file)){
                ext <- tools::file_ext(file$datapath)
                req(file)
                validate(need(ext == "xlsx","Please upload an xlsx file"))
                head(readxl::read_xlsx(file$datapath),n=3)
            }
        }
    )
    
    output$show_layout <- renderPlot(
        {
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
                    mutate(Index = as.character(Index),Col=as.numeric(Col)) %>% 
                    filter(!is.na(Row))
             
            ggplot(layout, aes(x = Row, y = Col, col = Index)) + geom_point(size=10) + scale_y_continuous(breaks=1:12)     
            }
        }
            
    )
}

# Run the application 
shinyApp(ui = ui, server = server)
