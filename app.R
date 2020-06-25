#

library(shiny)
library(DT)
library(tidyverse)
library(rvest)
library(openxlsx)

source("C:/Users/medwrl/Documents/GAMSTOP/GAMSTOP/gamstop1.R")


ui <- fluidPage(
  actionButton("go", "Download from Gambling Commission"),
  checkboxInput("IncludeInactiveDomain", "Include Inactive Domains ?",
               value =TRUE),
  checkboxInput("Long", "Show all domains and activities ?",
                value = FALSE),
  
  dataTableOutput("table")
)

server <- function(input, output) {
  
  dat <- eventReactive(input$go, {
    GAMSTOP_Licence(IncludeInactiveDomain = input$IncludeInactiveDomain,
                    Long = input$Long)
  })

    output$table <- renderDataTable({
    dat()
  }, rownames = FALSE)
}

shinyApp(ui = ui, server = server)

