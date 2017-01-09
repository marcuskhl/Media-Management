if(is.na(match(c("devtools"),installed.packages()[,"Package"]))) install.packages(new.packages) else library(devtools)
suppressMessages(devtools::install_github("marcuskhl/BasicSettings"));library(BasicSettings)

#~~~Notes~~~#
# This script extracts outputs from M:/Technology/DATA/Media Management/Demand side/Models/Completed
#~~~Notes~~~#
# IRDs_wb <- loadWorkbook("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016_v2.xlsx") # yielded weird error, could not do it the quick way

#~~~Functions Start~~~#
table.cleaning.IRDs <- function(df,variable_name){ # added .IRDs because other models might look slightly different
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
  df <- df[!is.na(df$value),]
  df <- df.name.change(df, "value", variable_name)
  df <- df[,c(1:(names.length(df)-3),(names.length(df)-1),names.length(df),(names.length(df)-2))]
  return(df)
}
#~~~Functions End~~~#



#~~~IRDs Spend Start~~~#
IRDs_Spend <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016_v2.xlsx", sheet = "Value", colNames = F)

IRDs_Spend_flat <- table.cleaning.IRDs(IRDs_Spend,"IRDs Spend")
# IRDs_Spend_flat$Trax <- extraction # initially only Value sheet has TRAX column, need to extract and merge before aggreegating, but it is just quicker to do it in excel
# So now every relevant sheet has a TRAX column
#~~~IRDs Spend End~~~#



#~~~IRDs Shipments Start~~~#
IRDs_Shipment <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016_v2.xlsx", sheet = "Streams Shipments", colNames = F)

IRDs_Shipment_flat <- table.cleaning.IRDs(IRDs_Shipment,"IRDs Shipment")
#~~~IRDs Shipments End~~~#



#~~~IRDs Shipments Start~~~#
IRDs_Shipment <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016_v2.xlsx", sheet = "Streams Shipments", colNames = F)

IRDs_Shipment_flat <- table.cleaning.IRDs(IRDs_Shipment,"IRDs Shipment")
#~~~IRDs Shipments End~~~#



#~~~IRDs Shipments Start~~~#
IRDs_Shipment <- read.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016_v2.xlsx", sheet = "Streams Shipments", colNames = F)

IRDs_Shipment_flat <- table.cleaning.IRDs(IRDs_Shipment,"IRDs Shipment")
#~~~IRDs Shipments End~~~#


