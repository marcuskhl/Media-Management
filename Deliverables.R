if(is.na(match(c("devtools"),installed.packages()[,"Package"]))) install.packages("devtools") else library(devtools)
suppressMessages(devtools::install_github("marcuskhl/BasicSettings"));library(BasicSettings)

#~~~Notes~~~#
# This script extracts outputs from Demand Side Models in M:/Technology/DATA/Media Management/Demand side/Models/Completed
#~~~Notes~~~#



#~~~WB Read Start~~~#
IRDs_wb <- loadWorkbook("M:/Technology/DATA/Media Management/Demand side/Models/Completed/IRDs_2016.xlsx")
Multiplexer_wb <- loadWorkbook("M:/Technology/DATA/Media Management/Demand side/Models/Completed/Multiplexer_2016.xlsx")
Transcoder_wb <- loadWorkbook("M:/Technology/DATA/Media Management/Demand side/Models/Completed/Transcoding_2016.xlsx")
#~~~WB Read End~~~#



#~~~Functions Start~~~#
table.cleaning.IRDs <- function(df){ # added .IRDs because other models might look slightly different
  df <- df[-2,]
  colnames(df)[1:7] <- df[1,1:7]

  names.length <- fn.piper(length,names)
  # right fill the colnames
  for (i in 8:names.length(df)){
    if(is.na(df[1,i])|df[1,i]==0){
      df[1,i] <- df[1,i-1]
    }
  }

  #~~~Custom Manipulations Start~~~#
  if(Reduce("|", grepl("Hours", df[2,]))){
    df <- df[,-grep("Hours", df[2,], ignore.case = T)] # this column exists in some transcoder sheets
  }

  if(Reduce("|", grepl("(", df[1,], fixed = T))){
      df[1,] <- gsub("(", "", df[1,], fixed = T) # dont need units here 
  }    
  
  if(Reduce("|", grepl(")", df[1,], fixed = T))){
      df[1,] <- gsub(")", "", df[1,], fixed = T) # dont need units here 
  }    
  
  if(Reduce("|", grepl(" multiplexers US$m", df[1,], fixed = T))){
      df[1,] <- gsub(" multiplexers US$m", "", df[1,], fixed = T) # dont need units here 
  }  
  
  if(Reduce("|", grepl(" transcoders", df[1,], fixed = T))){
      df[1,] <- gsub(" transcoders", "", df[1,], fixed = T) # dont need units here 
  } 
  
      if(Reduce("|", grepl(" transcoders", df[1,], fixed = T))){
      df[1,] <- gsub(" transcoders", "", df[1,], fixed = T) # dont need units here 
  }  
    

  if(Reduce("|", grepl(" US$m", df[1,], fixed = T))){
    df[1,] <- gsub(" US$m", "", df[1,], fixed = T) # dont need units here 
  }
  
  if(Reduce("|", grepl(" US$m ", df[1,], fixed = T))){
    df[1,] <- gsub(" US$m ", "", df[1,], fixed = T) # dont need units here 
  }
  
  if(Reduce("|", grepl("inhouse", df[1,], fixed = T))){
    df[1,] <- gsub("inhouse", "On Prem", df[1,], fixed = T) # dont need units here 
  }
  
  if(Reduce("|", grepl("outsource", df[1,], fixed = T))){
    df[1,] <- gsub("outsource", "Services", df[1,], fixed = T) # dont need units here 
  }

  #~~~Custom Manipulations Start~~~#
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
  df <- df.name.change(df,"Metric", "Sub-Product")
  df <- df[,c(1:(names.length(df)-3),(names.length(df)-1),(names.length(df)-2),names.length(df))]
  toMatch <- c("Total", "total")
  df <- df[grep(paste(toMatch,collapse="|"), df$`Sub-Product` , invert = T),]
  df$value <- round(df$value, 3)
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
  TRAX_0 <- TRAX_0[, list(value = sum(value)), by = cols_to_keep]
  TRAX_1 <- TRAX_1[,append(cols_to_keep,"value")]
  TRAX_0 <- as.df(TRAX_0)
  return(rbind.data.frame(TRAX_1, TRAX_0))
}
#~~~Functions End~~~#



#~~~IRDs Spend Start~~~#
IRDs_Spend <- read.xlsx(IRDs_wb, sheet = "Value", colNames = F)

IRDs_Spend_flat <- table.cleaning.IRDs(IRDs_Spend)#,"IRDs Spend")
IRDs_Spend_flat$Product <- "Revenues by Compression Format"
IRDs_Spend_flat <- df.name.change(IRDs_Spend_flat, "value", "Revenues (USD $m)")
#IRDs_Spend_flat$`Sub-Product` <- gsub(" (US$m)", "", IRDs_Spend_flat$`Sub-Product`, fixed = T)
IRDs_Spend_flat <- df.name.change(IRDs_Spend_flat, "Sub-Product", "Compression Format")
IRDs_Spend_flat$`Compression Format` <- gsub(" IRDS", "", IRDs_Spend_flat$`Compression Format`)
IRDs_Spend_flat$`Compression Format` <- gsub(" IRDs", "", IRDs_Spend_flat$`Compression Format`)

IRDs_Spend_flat$Market <- "IRDs"
#~~~IRDs Spend End~~~#



#~~~IRDs Shipments Start~~~#
IRDs_Shipment <- read.xlsx(IRDs_wb, sheet = "Streams Shipments", colNames = F)

IRDs_Shipment_flat <- table.cleaning.IRDs(IRDs_Shipment)#,"IRDs Shipment")
IRDs_Shipment_flat$Product <- "Shipments by Business Type"
IRDs_Shipment_flat <- df.name.change(IRDs_Shipment_flat, "value", "Shipments (000s)")
IRDs_Shipment_flat <- df.name.change(IRDs_Shipment_flat, "Sub-Product", "Business Type")
IRDs_Shipment_flat$`Business Type` <- gsub(" IRDS", "", IRDs_Shipment_flat$`Business Type`)
IRDs_Shipment_flat$`Business Type` <- gsub(" IRDs", "", IRDs_Shipment_flat$`Business Type`)
IRDs_Shipment_flat$Market <- "IRDs"
#~~~IRDs Shipments End~~~#



#~~~IRDs Contribution Shipments Start~~~#
IRDs_Contri_Shipment <- read.xlsx(IRDs_wb, sheet = "Contribution shipments", colNames = F)

IRDs_Contri_Shipment_flat <- table.cleaning.IRDs(IRDs_Contri_Shipment)#,"Contribution Shipment")
IRDs_Contri_Shipment_flat$`Sub-Product` <- "Contribution"
IRDs_Contri_Shipment_flat$Product <- "Shipments by Operation Type"
IRDs_Contri_Shipment_flat <- df.name.change(IRDs_Contri_Shipment_flat, "value", "Shipments (000s)")
IRDs_Contri_Shipment_flat <- df.name.change(IRDs_Contri_Shipment_flat, "Sub-Product", "Operation Type")
IRDs_Contri_Shipment_flat$`Operation Type` <- gsub(" IRDs", "", IRDs_Contri_Shipment_flat$`Operation Type`)
IRDs_Contri_Shipment_flat$Market <- "IRDs"
#~~~IRDs Contribution Shipments End~~~#



#~~~IRDs Distribution Shipments Start~~~#
IRDs_Distri_Shipment <- read.xlsx(IRDs_wb, sheet = "Distribution shipments", colNames = F)

IRDs_Distri_Shipment_flat <- table.cleaning.IRDs(IRDs_Distri_Shipment)#,"Distribution Shipment")
IRDs_Distri_Shipment_flat$`Sub-Product` <- "Distribution"
IRDs_Distri_Shipment_flat$Product <- "Shipments by Operation Type"
IRDs_Distri_Shipment_flat <- df.name.change(IRDs_Distri_Shipment_flat, "value", "Shipments (000s)")
IRDs_Distri_Shipment_flat <- df.name.change(IRDs_Distri_Shipment_flat, "Sub-Product", "Operation Type")
IRDs_Distri_Shipment_flat$`Operation Type` <- gsub(" IRDs", "", IRDs_Distri_Shipment_flat$`Operation Type`)
IRDs_Distri_Shipment_flat$Market <- "IRDs"
#~~~IRDs Distribution Shipments End~~~#




#~~~Multiplexer Revenues Start~~~#
Multiplexer_rev <- read.xlsx(Multiplexer_wb, sheet = "Multiplexer revs", colNames = F)

Multiplexer_rev_flat <- table.cleaning.IRDs(Multiplexer_rev)#,"Distribution Shipment")
#Multiplexer_rev_flat$`Sub-Product` <- gsub(" multiplexers US$m", "", Multiplexer_rev_flat$`Sub-Product`, fixed = T)
Multiplexer_rev_flat$Product <- "Revenues by Platform"
Multiplexer_rev_flat <- df.name.change(Multiplexer_rev_flat, "value", "Revenues (USD $m)")
Multiplexer_rev_flat <- df.name.change(Multiplexer_rev_flat, "Sub-Product", "Platform")
Multiplexer_rev_flat$Market <- "Multiplexers"
#~~~Multiplexer Revenues End~~~#



#~~~Transcoding Live Inhouse Revenues Start~~~#
Transcoding_Live_Inhouse_rev <- read.xlsx(Transcoder_wb, sheet = "Live inhouse revs", colNames = F)

Transcoding_Live_Inhouse_rev_flat <- table.cleaning.IRDs(Transcoding_Live_Inhouse_rev)#,"Distribution Shipment")
#Transcoding_Live_Inhouse_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_Live_Inhouse_rev_flat$`Sub-Product`, fixed = T)
Transcoding_Live_Inhouse_rev_flat$Product <- "Revenues by Operation Type Resolution Format"
Transcoding_Live_Inhouse_rev_flat <- df.name.change(Transcoding_Live_Inhouse_rev_flat, "value", "Revenues (USD $m)")
Transcoding_Live_Inhouse_rev_flat <- df.name.change(Transcoding_Live_Inhouse_rev_flat, "Sub-Product", "Operation Type, Resolution Format")
Transcoding_Live_Inhouse_rev_flat$`Operation Type` <- "On Prem"
Transcoding_Live_Inhouse_rev_flat$`Operation Type, Resolution Format` <- gsub("On Prem ", "",Transcoding_Live_Inhouse_rev_flat$`Operation Type, Resolution Format`)
Transcoding_Live_Inhouse_rev_flat$`Live vs Non-Live` <- "Live"
Transcoding_Live_Inhouse_rev_flat$`Operation Type, Resolution Format` <- gsub("Live ", "",Transcoding_Live_Inhouse_rev_flat$`Operation Type, Resolution Format`)
Transcoding_Live_Inhouse_rev_flat <- df.name.change(Transcoding_Live_Inhouse_rev_flat,"Operation Type, Resolution Format", "Resolution Format" )
Transcoding_Live_Inhouse_rev_flat$Market <- "Transcoders"
#~~~Transcoding Live Inhouse Shipments End~~~#



#~~~Transcoding Non-Live Inhouse Revenues Start~~~#
Transcoding_NonLive_Inhouse_rev <- read.xlsx(Transcoder_wb, sheet = "Non-Live inhouse revs", colNames = F)

Transcoding_NonLive_Inhouse_rev_flat <- table.cleaning.IRDs(Transcoding_NonLive_Inhouse_rev)#,"Distribution Shipment")
#Transcoding_NonLive_Inhouse_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_NonLive_Inhouse_rev_flat$`Sub-Product`, fixed = T)
Transcoding_NonLive_Inhouse_rev_flat$Product <- "Revenues by Operation Type Resolution Format"
Transcoding_NonLive_Inhouse_rev_flat <- df.name.change(Transcoding_NonLive_Inhouse_rev_flat, "value", "Revenues (USD $m)")
Transcoding_NonLive_Inhouse_rev_flat <- df.name.change(Transcoding_NonLive_Inhouse_rev_flat, "Sub-Product", "Operation Type, Resolution Format")
Transcoding_NonLive_Inhouse_rev_flat$`Operation Type` <- "On Prem"
Transcoding_NonLive_Inhouse_rev_flat$`Operation Type, Resolution Format` <- gsub("On Prem ", "",Transcoding_NonLive_Inhouse_rev_flat$`Operation Type, Resolution Format`)
Transcoding_NonLive_Inhouse_rev_flat$`Live vs Non-Live` <- "Non-Live"
Transcoding_NonLive_Inhouse_rev_flat$`Operation Type, Resolution Format` <- gsub("Non-Live ", "",Transcoding_NonLive_Inhouse_rev_flat$`Operation Type, Resolution Format`)
Transcoding_NonLive_Inhouse_rev_flat <- df.name.change(Transcoding_NonLive_Inhouse_rev_flat,"Operation Type, Resolution Format", "Resolution Format" )
Transcoding_NonLive_Inhouse_rev_flat$Market <- "Transcoders"
#~~~Transcoding Non-Live Inhouse Revenues End~~~#



#~~~Transcoding Live Outsourced Revenues Start~~~#
Transcoding_Live_Outsource_rev <- read.xlsx(Transcoder_wb, sheet = "Live outsource revs", colNames = F)

Transcoding_Live_Outsource_rev_flat <- table.cleaning.IRDs(Transcoding_Live_Outsource_rev)#,"Distribution Shipment")
#Transcoding_Live_Outsource_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_Live_Outsource_rev_flat$`Sub-Product`, fixed = T)
Transcoding_Live_Outsource_rev_flat$Product <- "Revenues by Operation Type Resolution Format"
Transcoding_Live_Outsource_rev_flat <- df.name.change(Transcoding_Live_Outsource_rev_flat, "value", "Revenues (USD $m)")
Transcoding_Live_Outsource_rev_flat <- df.name.change(Transcoding_Live_Outsource_rev_flat, "Sub-Product", "Operation Type, Resolution Format")
Transcoding_Live_Outsource_rev_flat$`Operation Type` <- "Services"
Transcoding_Live_Outsource_rev_flat$`Operation Type, Resolution Format` <- gsub("Services ", "",Transcoding_Live_Outsource_rev_flat$`Operation Type, Resolution Format`)
Transcoding_Live_Outsource_rev_flat$`Live vs Non-Live` <- "Live"
Transcoding_Live_Outsource_rev_flat$`Operation Type, Resolution Format` <- gsub("Live ", "",Transcoding_Live_Outsource_rev_flat$`Operation Type, Resolution Format`)
Transcoding_Live_Outsource_rev_flat <- df.name.change(Transcoding_Live_Outsource_rev_flat,"Operation Type, Resolution Format", "Resolution Format" )
Transcoding_Live_Outsource_rev_flat$Market <- "Transcoders"
#~~~Transcoding Live Outsourced Revenues End~~~#



#~~~Transcoding Non-Live Outsourced Revenues Start~~~#
Transcoding_NonLive_Outsource_rev <- read.xlsx(Transcoder_wb, sheet = "Non-Live outsource revs", colNames = F)

Transcoding_NonLive_Outsource_rev_flat <- table.cleaning.IRDs(Transcoding_NonLive_Outsource_rev)#,"Distribution Shipment")
#Transcoding_NonLive_Outsource_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_NonLive_Outsource_rev_flat$`Sub-Product`, fixed = T)
Transcoding_NonLive_Outsource_rev_flat$Product <- "Revenues by Operation Type Resolution Format"
Transcoding_NonLive_Outsource_rev_flat <- df.name.change(Transcoding_NonLive_Outsource_rev_flat, "value", "Revenues (USD $m)")
Transcoding_NonLive_Outsource_rev_flat <- df.name.change(Transcoding_NonLive_Outsource_rev_flat, "Sub-Product", "Operation Type, Resolution Format")
Transcoding_NonLive_Outsource_rev_flat$`Operation Type` <- "Services"
Transcoding_NonLive_Outsource_rev_flat$`Operation Type, Resolution Format` <- gsub("Services ", "",Transcoding_NonLive_Outsource_rev_flat$`Operation Type, Resolution Format`)
Transcoding_NonLive_Outsource_rev_flat$`Live vs Non-Live` <- "Non-Live"
Transcoding_NonLive_Outsource_rev_flat$`Operation Type, Resolution Format` <- gsub("Non-Live ", "",Transcoding_NonLive_Outsource_rev_flat$`Operation Type, Resolution Format`)
Transcoding_NonLive_Outsource_rev_flat <- df.name.change(Transcoding_NonLive_Outsource_rev_flat,"Operation Type, Resolution Format", "Resolution Format" )
Transcoding_NonLive_Outsource_rev_flat$Market <- "Transcoders"
#~~~Transcoding Non-Live Outsourced Revenues End~~~#



#~~~Transcoding Broadcast Revenues Composition Start~~~#
Transcoding_Broadcast_rev_comp <- read.xlsx(Transcoder_wb, sheet = "Broadcast revs comp", colNames = F, rows = 1:238)

Transcoding_Broadcast_rev_comp_flat <- table.cleaning.IRDs(Transcoding_Broadcast_rev_comp)#,"Distribution Shipment")
#Transcoding_Broadcast_rev_comp_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_Broadcast_rev_comp_flat$`Sub-Product`, fixed = T)
Transcoding_Broadcast_rev_comp_flat$Product <- "Revenues by Buyer Type Resolution Format"
Transcoding_Broadcast_rev_comp_flat <- df.name.change(Transcoding_Broadcast_rev_comp_flat, "value", "Revenues (USD $m)")
Transcoding_Broadcast_rev_comp_flat <- df.name.change(Transcoding_Broadcast_rev_comp_flat, "Sub-Product", "Resolution Format")
Transcoding_Broadcast_rev_comp_flat$`Buyer Type` <- "Broadcast"
Transcoding_Broadcast_rev_comp_flat$`Resolution Format` <- gsub("Broadcast ", "",Transcoding_Broadcast_rev_comp_flat$`Resolution Format`)
Transcoding_Broadcast_rev_comp_flat$Market <- "Transcoders"
#~~~Transcoding Broadcast Revenues Composition End~~~#



#~~~Transcoding Broadcast Revenues Composition Start~~~#
Transcoding_Broadcast_rev <- read.xlsx(Transcoder_wb, sheet = "Broadcast revs", colNames = F, rows = 1:238)

Transcoding_Broadcast_rev_flat <- table.cleaning.IRDs(Transcoding_Broadcast_rev)#,"Distribution Shipment")
#Transcoding_Broadcast_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_Broadcast_rev_flat$`Sub-Product`, fixed = T)
Transcoding_Broadcast_rev_flat$Product <- "Revenues by Buyer Type Operation Type"
Transcoding_Broadcast_rev_flat <- df.name.change(Transcoding_Broadcast_rev_flat, "value", "Revenues (USD $m)")
Transcoding_Broadcast_rev_flat <- df.name.change(Transcoding_Broadcast_rev_flat, "Sub-Product", "Operation Type")
Transcoding_Broadcast_rev_flat$`Buyer Type` <- "Broadcast"
Transcoding_Broadcast_rev_flat$`Operation Type` <- gsub("Broadcast ", "",Transcoding_Broadcast_rev_flat$`Operation Type`)
Transcoding_Broadcast_rev_flat$Market <- "Transcoders"
#~~~Transcoding Broadcast Revenues Composition End~~~#



#~~~Transcoding Operators Revenues Composition Start~~~#
Transcoding_Operators_rev_comp <- read.xlsx(Transcoder_wb, sheet = "Operators revs comp", colNames = F, rows = 1:238)

Transcoding_Operators_rev_comp_flat <- table.cleaning.IRDs(Transcoding_Operators_rev_comp)#,"Distribution Shipment")
#Transcoding_Operators_rev_comp_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_Operators_rev_comp_flat$`Sub-Product`, fixed = T)
Transcoding_Operators_rev_comp_flat$Product <- "Revenues by Buyer Type Resolution Format"
Transcoding_Operators_rev_comp_flat <- df.name.change(Transcoding_Operators_rev_comp_flat, "value", "Revenues (USD $m)")
Transcoding_Operators_rev_comp_flat <- df.name.change(Transcoding_Operators_rev_comp_flat, "Sub-Product", "Resolution Format")
Transcoding_Operators_rev_comp_flat$`Buyer Type` <- "Operators"
Transcoding_Operators_rev_comp_flat$`Resolution Format` <- gsub("Operators ", "",Transcoding_Operators_rev_comp_flat$`Resolution Format`)
Transcoding_Operators_rev_comp_flat$Market <- "Transcoders"
#~~~Transcoding Operators Revenues Composition End~~~#



#~~~Transcoding Operators Revenues Start~~~#
Transcoding_Operators_rev <- read.xlsx(Transcoder_wb, sheet = "Operators revs", colNames = F, rows = 1:238)

Transcoding_Operators_rev_flat <- table.cleaning.IRDs(Transcoding_Operators_rev)#,"Distribution Shipment")
#Transcoding_Operators_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_Operators_rev_flat$`Sub-Product`, fixed = T)
Transcoding_Operators_rev_flat$Product <- "Revenues by Buyer Type Operation Type"
Transcoding_Operators_rev_flat <- df.name.change(Transcoding_Operators_rev_flat, "value", "Revenues (USD $m)")
Transcoding_Operators_rev_flat <- df.name.change(Transcoding_Operators_rev_flat, "Sub-Product", "Operation Type")
Transcoding_Operators_rev_flat$`Buyer Type` <- "Operators"
Transcoding_Operators_rev_flat$`Operation Type` <- gsub("Operators ", "",Transcoding_Operators_rev_flat$`Operation Type`)
Transcoding_Operators_rev_flat$Market <- "Transcoders"
#~~~Transcoding Operators Revenues End~~~#



#~~~Transcoding OTT Revenues Start~~~#
Transcoding_OTT_rev_comp <- read.xlsx(Transcoder_wb, sheet = "OTT revs comp", colNames = F, rows = 1:238)

Transcoding_OTT_rev_comp_flat <- table.cleaning.IRDs(Transcoding_OTT_rev_comp)#,"Distribution Shipment")
#Transcoding_OTT_rev_comp_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_OTT_rev_comp_flat$`Sub-Product`, fixed = T)
Transcoding_OTT_rev_comp_flat$Product <- "Revenues by Buyer Type Resolution Format"
Transcoding_OTT_rev_comp_flat <- df.name.change(Transcoding_OTT_rev_comp_flat, "value", "Revenues (USD $m)")
Transcoding_OTT_rev_comp_flat <- df.name.change(Transcoding_OTT_rev_comp_flat, "Sub-Product", "Resolution Format")
Transcoding_OTT_rev_comp_flat$`Buyer Type` <- "OTT"
Transcoding_OTT_rev_comp_flat$`Resolution Format` <- gsub("OTT ", "",Transcoding_OTT_rev_comp_flat$`Resolution Format`)
Transcoding_OTT_rev_comp_flat$`Resolution Format` <- gsub("Operators ", "",Transcoding_OTT_rev_comp_flat$`Resolution Format`) # manual error corrected here
Transcoding_OTT_rev_comp_flat$Market <- "Transcoders"
#~~~Transcoding OTT Revenues End~~~#



#~~~Transcoding OTT Revenues Start~~~#
Transcoding_OTT_rev <- read.xlsx(Transcoder_wb, sheet = "OTT revs", colNames = F, rows = 1:238)

Transcoding_OTT_rev_flat <- table.cleaning.IRDs(Transcoding_OTT_rev)#,"Distribution Shipment")
#Transcoding_OTT_rev_flat$`Sub-Product` <- gsub(" US$m", "", Transcoding_OTT_rev_flat$`Sub-Product`, fixed = T)
Transcoding_OTT_rev_flat$Product <- "Revenues by Buyer Type Operation Type"
Transcoding_OTT_rev_flat <- df.name.change(Transcoding_OTT_rev_flat, "value", "Revenues (USD $m)")
Transcoding_OTT_rev_flat <- df.name.change(Transcoding_OTT_rev_flat, "Sub-Product", "Operation Type")
Transcoding_OTT_rev_flat$`Buyer Type` <- "OTT"
Transcoding_OTT_rev_flat$`Operation Type` <- gsub("OTT ", "",Transcoding_OTT_rev_flat$`Operation Type`)
Transcoding_OTT_rev_flat$Market <- "Transcoders"
#~~~Transcoding OTT Revenues End~~~#



#~~~Shipment and Revenue Big_Table Building Start~~~#
names.length <- fn.piper(length,names)

shipment_table <- bind_rows(list( IRDs_Shipment_flat, IRDs_Contri_Shipment_flat, IRDs_Distri_Shipment_flat))



revenue_table <- bind_rows(list( 
  # IRDS:
  IRDs_Spend_flat, 
  
  # Multiplexers:
  Multiplexer_rev_flat,
  
  # Transcoders:
  Transcoding_Live_Inhouse_rev_flat, Transcoding_NonLive_Inhouse_rev_flat,
  Transcoding_Live_Outsource_rev_flat, Transcoding_NonLive_Outsource_rev_flat,
  Transcoding_Broadcast_rev_comp_flat, Transcoding_Broadcast_rev_flat,
  Transcoding_Operators_rev_comp_flat, Transcoding_Operators_rev_flat,
  Transcoding_OTT_rev_comp_flat, Transcoding_OTT_rev_flat
  
  ))


#~~~Makeup Sesction Start~~~#
shipment_table <- shipment_table[shipment_table$Year>2011,]
revenue_table <- revenue_table[revenue_table$Year>2011,]

shipment_table <- shipment_table[,c(1:(names.length(shipment_table)-2),(names.length(shipment_table)),(names.length(shipment_table)-1))]

tbl_makeup <- function(df) {
  df <- as.df(df)
  df <- df[!is.na(df[,1]),]
  df[is.na(df)] <- "N.A."
  df[df==" In House" | df==" In house" | df==" In-House" | df=="In-house" | df=="inhouse" | df=="Inhouse" ] <- "On Prem"
  
  df[df=="Outsourced" | df=="outsourced" | df=="Outsource" |df=="outsource" ] <- "Services"
  
  if(Reduce("|", grepl("Resolution Format", names(df)))){
    df$`Resolution Format` <- gsub(" ", "", df$`Resolution Format`, fixed = T)
  }
  
  
  return(df)
}

shipment_table <- tbl_makeup(shipment_table)
revenue_table <- tbl_makeup(revenue_table)


#~~~Makeup Sesction End~~~#


#~~~Shipment and Revenue Big_Table Building End~~~#


save.xlsx("M:/Technology/DATA/Media Management/Demand side/Models/Completed/Output/R_outputs.xlsx", shipment_table,revenue_table)


