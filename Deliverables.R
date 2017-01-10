if(is.na(match(c("devtools"),installed.packages()[,"Package"]))) install.packages(new.packages) else library(devtools)
suppressMessages(devtools::install_github("marcuskhl/BasicSettings"));library(BasicSettings)

#~~~Notes~~~#
# This script extracts outputs from Demand Side Models in M:/Technology/DATA/Media Management/Demand side/Models/Completed
#~~~Notes~~~#



#~~~Functions Start~~~#
table.cleaning.IRDs <- function(df){ # added .IRDs because other models might look slightly different
  df <- df[-2,]
  colnames(df)[1:7] <- df[1,1:7]
  
  names.length <- fn.piper(length,names)
  # right fill the colnames
  for (i in 8:names.length(df)){
    if(is.na(df[1,i])){
      df[1,i] <- df[1,i-1]
    }
  }
  colnames(df)[8:length(df)] <- paste(df[1,8:length(df)], "xax", df[2,8:length(df)], sep = "")
  df <- cbind.data.frame(df[,1:7],df[,8:names.length(df)][,!is.na(df[2,8:names.length(df)])])
  
  df <- df[-1:-2,]
  df <- df[complete.cases(df[,1:7]),] # remove sub-region separation
  
  df <- df %>% gather(BLAH, value, 8:names.length(df))
  df$Metric <- lapply(strsplit(df$BLAH, "xax"), "[", 1)
  df$Year <- lapply(strsplit(df$BLAH, "xax"), "[", 2)
  df <- column.rm(df, "BLAH")
  df$Year <- f2n(df$Year)
  df$Metric <- unlist(df$Metric) # artifact of gather
  df <- df[!is.na(df$value),] # sometimes 2011 is NA sometimes it is zero, need to decide
  df <- TRAX_country_aggregation(df)
  # df <- df.name.change(df, "value", variable_name)
  df <- df.name.change(df,"Metric", "Submeasure")
  df <- df[,c(1:(names.length(df)-3),(names.length(df)-1),(names.length(df)-2),names.length(df))]
  toMatch <- c("Total", "total")
  df <- df[grep(paste(toMatch,collapse="|"), df$Submeasure , invert = T),]
  return(df)
}



TRAX_country_aggregation <- function(df){
  df$Trax <- f2n(df$Trax)
  df$value <- f2n(df$value)
  TRAX_1 <- df[df$Trax==1,]
  TRAX_0 <- df[df$Trax==0,]
  TRAX_0$Country <- paste0("Rest of ",TRAX_0$`Sub Code`)
  toMatch <- c("value","Trax","ISO")
  cols_to_keep <-grep(paste(toMatch,collapse="|"), names(TRAX_0), invert = T, value = T)
  TRAX_0 <- as.dt(TRAX_0)
  TRAX_0 <- TRAX_0[, list(value = sum(value)), by =cols_to_keep]
  TRAX_1 <- TRAX_1[,append(cols_to_keep,"value")]
  TRAX_0 <- as.df(TRAX_0)
  return(rbind.data.frame(TRAX_1, TRAX_0))
}
#~~~Functions End~~~#



# IRDs_wb <- loadWorkbook("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016_v2.xlsx") # yielded weird error, could not do it the quick way



#~~~IRDs Spend Start~~~#
IRDs_Spend <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016.xlsx", sheet = "Value", colNames = F)

IRDs_Spend_flat <- table.cleaning.IRDs(IRDs_Spend)#,"IRDs Spend")
IRDs_Spend_flat$Measure <- "IRDs Spend"
IRDs_Spend_flat <- df.name.change(IRDs_Spend_flat, "value", "Revenues (USD $m)")
IRDs_Spend_flat$Submeasure <- gsub(" (US$m)", "", IRDs_Spend_flat$Submeasure, fixed = T) 
#~~~IRDs Spend End~~~#



#~~~IRDs Shipments Start~~~#
IRDs_Shipment <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016.xlsx", sheet = "Streams Shipments", colNames = F)

IRDs_Shipment_flat <- table.cleaning.IRDs(IRDs_Shipment)#,"IRDs Shipment")
IRDs_Shipment_flat$Measure <- "Inhouse vs Outsource Shipments"
IRDs_Shipment_flat <- df.name.change(IRDs_Shipment_flat, "value", "Shipments (000s)")
#~~~IRDs Shipments End~~~#



#~~~IRDs Contribution Shipments Start~~~#
IRDs_Contri_Shipment <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016.xlsx", sheet = "Contribution shipments", colNames = F)

IRDs_Contri_Shipment_flat <- table.cleaning.IRDs(IRDs_Contri_Shipment)#,"Contribution Shipment")
IRDs_Contri_Shipment_flat$Submeasure <- "Contribution IRDs"
IRDs_Contri_Shipment_flat$Measure <- "Contribution vs Distribution IRD Shipments"
IRDs_Contri_Shipment_flat <- df.name.change(IRDs_Contri_Shipment_flat, "value", "Shipments (000s)")
#~~~IRDs Contribution Shipments End~~~#



#~~~IRDs Distribution Shipments Start~~~#
IRDs_Distri_Shipment <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016.xlsx", sheet = "Distribution shipments", colNames = F)

IRDs_Distri_Shipment_flat <- table.cleaning.IRDs(IRDs_Distri_Shipment)#,"Distribution Shipment")
IRDs_Distri_Shipment_flat$Submeasure <- "Distribution IRDs"
IRDs_Distri_Shipment_flat$Measure <- "Contribution vs Distribution IRD Shipments"
IRDs_Distri_Shipment_flat <- df.name.change(IRDs_Distri_Shipment_flat, "value", "Shipments (000s)")
#~~~IRDs Distribution Shipments End~~~#



big_table <- bind_rows(list(IRDs_Spend_flat, IRDs_Shipment_flat, IRDs_Contri_Shipment_flat, IRDs_Distri_Shipment_flat))

# fwrite(big_table, "M:/Technology/DATA/Media Management/Demand side/Models/Completed/Output/IRDs.csv", row.names = F)

save.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/Output/IRDs.xlsx", big_table)








