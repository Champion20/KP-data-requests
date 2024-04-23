# AUTHOR:   R. Pineteh | USAID
# PURPOSE:  Adhoc KP data processing for sharing with SANAC
# REF ID:   b4f36b95 
# LICENSE:  MIT
# DATE:     2024-04-18
# UPDATED: 

# PREPARATORY STEPS -------------------------------------------------------------
  # 1. Create a project for KP data with relevant sub folders using si_setup()-only required the first time you run the script
  # 2. Download the latest Genie from DATIM containing the period for which KP data is requested
  # 3. Unzip the Genie and save in the data folder within your project
  # 4. Change the period (fiscal year and quarter) in the script to that for which KP data is requested

# DEPENDENCIES ------------------------------------------------------------

# Load required packages & libraries
library(tidyverse)
library(gagglr)
library(openxlsx)
library(readr)

 
# GLOBAL VARIABLES --------------------------------------------------------
  
  # SI specific paths/functions  
si_setup()    
load_secrets()

  # msd path- read the most recent MSD from the Genie folder. Prior to this step download the latest Genie and save in the data folder
  path_msd <- return_latest("Data", "SITE.*South Africa")
  
  #grab metadata about MSD
  metadata <- get_metadata() #list of MSD metadata elements
  

# IMPORT ------------------------------------------------------------------
  
  # Load data 
  df_msd <- read_psd(path_msd)
  
  # change the period on the filter based on the fiscal year and quarter for which KP data is requested
  
  df_data <- df_msd %>% reshape_msd() %>%  filter(period=="FY24Q1")
  
  glimpse(df_data)
  
# MUNGE -------------------------------------------------------------------
  
  # Grab all kP standardised disaggregates and relevant columns
  
  kp_data <- df_data %>% 
    filter( standardizeddisaggregate %in% c("KeyPop",
                                            "KeyPop/ARTNoContactReason/HIVStatus", 
                                            "KeyPop/HIVSelfTest", 
                                            "KeyPop/HIVStatus" ,
                                            "KeyPop/Indication/HIVStatus", 
                                            "KeyPop/Result", 
                                            "KeyPop/RTRI/HIVStatus", 
                                            "KeyPop/Status",
                                            "KeyPopAbr"),
    ) %>% 
    select( psnu, community,prime_partner_name , indicator ,standardizeddisaggregate,  otherdisaggregate, otherdisaggregate_sub,funding_agency,mech_code,mech_name,period,value)

  glimpse(kp_data)
  
  #Change to the relevant fiscal year and quarter when summarising
  
  kp_result <-  kp_data %>% 
    group_by(psnu, community,prime_partner_name , indicator , otherdisaggregate_sub) %>% 
    summarise(FY24Q1= sum(value,na.rm = TRUE)) # Change to relevant period
  
  glimpse(kp_result)
  
  # Create individual dataframes for each indicator 
  
  TX_NEW <- kp_result %>% 
    filter(indicator == "TX_NEW")
 
   TX_CURR <- kp_result %>% 
    filter(indicator == "TX_CURR")
   
   PrEP_NEW <- kp_result %>% 
     filter(indicator == "PrEP_NEW")
   
   PrEP_CT <- kp_result %>% 
     filter(indicator == "PrEP_CT")

   
   HTS_TST <- kp_result %>% 
     filter(indicator == "HTS_TST")
   
   HTS_SELF <- kp_result %>% 
     filter(indicator == "HTS_SELF")
   
   HTS_TST_POS <- kp_result %>% 
     filter(indicator == "HTS_TST_POS")
   
   # KP_PREV indicator is reported semi annually, run the codes for KP_PREV when submitting data Q2 and Q4 only
   # KP_PREV <- kp_result %>% 
   #   filter(indicator == "KP_PREV")
 
  
  #Create a workbook with separate worksheets for each indicators
     # Change the name of the output document from "KP results FY24Q1.xlsx" to the relevant fiscal year and quarter
   
  write.xlsx(HTS_TST,"Dataout/KP results FY24Q1.xlsx",  sheetName="HTS_TST",append=TRUE)
  
  wb<-loadWorkbook("Dataout/KP results FY24Q1.xlsx")
  
  addWorksheet(wb,"HTS_SELF")
  writeData(wb,sheet="HTS_SELF",x=HTS_SELF)
  
  addWorksheet(wb,"HTS_TST_POS")
  writeData(wb,sheet="HTS_TST_POS",x=HTS_TST_POS)
  
  addWorksheet(wb,"TX_NEW")
  writeData(wb,sheet="TX_NEW",x=TX_NEW)
  
  addWorksheet(wb,sheetName = "TX_CURR")
  writeData(wb,sheet = "TX_CURR",x=TX_CURR)
  
  
  addWorksheet(wb,sheetName = "PrEP_NEW")
  writeData(wb,sheet = "PrEP_NEW",x=PrEP_NEW)
  
  
  addWorksheet(wb,sheetName = "PrEP_CT")
  writeData(wb,sheet = "PrEP_CT",x=PrEP_CT)
  
  # addWorksheet(wb,sheetName = "KP_PREV")
  # writeData(wb,sheet = "KP_PREV",x=KP_PREV)
  
  saveWorkbook(wb,"Dataout/KP results FY24Q1.xlsx",overwrite = T)
  
