library(xlsx)
library(dplyr)
library(stringr)
library(lubridate)

# Loading data
resource <- "/Users/Johnny/Downloads/Forex_History/report_1538067829_example.xls" ## You can change to your path
df <- read.xlsx2(resource, 1, stringsAsFactor= TRUE)
df <- data.frame(df)

# Order by time
df <- df[order(df$Time),]

# Change data format
df$Profit <- as.numeric(levels(df$Profit))[df$Profit]
df$Type <- as.character(df$Type)

# Remove informal expression
df$Type <- gsub("^balance","", df$Type) ## Omit original balance record
df$Type <- gsub("limit ","", df$Type) ## Omit limit orders

# Deal with raw data
tmp <- str_split_fixed(df$Type," ",3)
Type_split <- data.frame(tmp[,1:3])
colnames(Type_split) <- c("Buy_Sell", "Amount", "Pairs") ## Split Type to three col 
df <- cbind(df, Type_split)
df <- df[,c(1,2,11,12,4,5,13,8,6,7)]
df <- df[2:nrow(df),]
df$Time <- ymd_hms(df$Time)

# Raw data to result
df <- df %>% group_by(Year = year(df$Time), Month = month(df$Time), Pair = df$Pairs) %>% summarize(Mean=mean(Profit, na.rm = TRUE), Max = max(Profit),Min = min(Profit), Count = sum(!Buy_Sell == ""), Sum = sum(Profit))
df <- cbind(df, YM = paste(df$Year, df$Month, sep = "_"))
result <- df[df$Pair !="",]

# Setup file name
name <- paste(as.character(today()), "_Report.xlsx", sep = "")
name <- gsub("-","_", name)
path <- "/Users/Johnny/Downloads/Forex_History/"
dst <- paste(path, name, sep = "")

# Save
write.xlsx(as.data.frame(result), dst, sheetName="Report",col.names=TRUE, row.names=FALSE, append=TRUE)