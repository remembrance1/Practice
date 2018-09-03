library(xlsx)
options(warn=-1)

df <- read.xlsx("Data.xlsx", sheetName = "Input data", header = T)
df <- df[-1,]

#extracting first table
Listofmodels <- df[,1:2]
Listofmodels <- Listofmodels[complete.cases(Listofmodels), ]
colnames(Listofmodels) <- c("Model Name", "Threshold")

#extracting 2nd table
Listofrules <- df[,4:7]
Listofrules <- Listofrules[complete.cases(Listofrules), ]
colnames(Listofrules) <- c("Model Name", "Rule Name", "Score", "Customer Type")

#extracting 3rd table
Transactionhits <- df[,10:13]
Transactionhits <- Transactionhits[complete.cases(Transactionhits), ]
colnames(Transactionhits) <- c("Rule Name", "Customer Type", "Customer ID", "Hit Date")

#----Transactionhits data manipulation----#
#change date to extract day of week and week#
library(tibble)
library(janitor)

Transactionhits$`Hit Date` <- excel_numeric_to_date(as.numeric(as.character(Transactionhits$`Hit Date`)), date_system = "modern") 
Transactionhits$Day <- weekdays(as.Date(Transactionhits$`Hit Date`)) #extracting day
Transactionhits$Month <- months(as.Date(Transactionhits$`Hit Date`)) #extracting month

#getting week number
weekno <- data.frame(Dates = Transactionhits$`Hit Date`, Week = format(Transactionhits$`Hit Date`, format = "%W")) #extracted week number
weekno <- weekno[,2]
Transactionhits <- cbind(Transactionhits, weekno) #combined both dataframes together
names(Transactionhits)[names(Transactionhits) == 'weekno'] <- 'Week' #rename col name to week

#isolate individual and corporate
transcorp <- subset(Transactionhits, `Customer Type` == "Corporate")
transindv <- subset(Transactionhits, `Customer Type` == "Individual")

library(dplyr)
# extracting distinct entries based on week, customer ID and rule name, to avoid repeats
transcorp <- distinct(transcorp,`Customer ID`, `Rule Name`, `Week`, .keep_all = T) 
transindv <- distinct(transindv,`Customer ID`, `Rule Name`, `Week`, .keep_all = T) #Transactionhits data is ready

table(Transactionhits$`Rule Name`) #count the number of times each rule appears

#----Manipulating Listofmodels----#
#adding indv or corp to the right column for labelling now
Listofmodels$Type <- sub('.*\\_', '', Listofmodels$`Model Name`) #getting indv & corp
Listofmodels$`Model Name` <- sub('\\_.*', '', Listofmodels$`Model Name`) #extracting model name
Listofmodels$Type <- plyr::revalue(Listofmodels$Type, c("INV" = "Individual", "CP" = "Corporate"))

modindv <- subset(Listofmodels, Type == "Individual")
modcorp <- subset(Listofmodels, Type == 'Corporate')

#----Manipulating Listofrules----#
ruleindv <- subset(Listofrules, `Customer Type` == "Individual")
rulecorp <- subset(Listofrules, `Customer Type` == "Corporate")

#ignore the duplicated rules first? R01 R10 from indv and R02 from corp? Depends on business requirements
ruleindv <- ruleindv[!(ruleindv$`Rule Name`=="R01" | ruleindv$`Rule Name`=="R10"),] #removed R01 & R10
rulecorp <- rulecorp[!(rulecorp$`Rule Name`=="R02"),] #removed R02

#--------General Statistics--------#
#------Obtaining 3rd Table output-------#

#------Trying to find customer w max alerts------#
#aggregate the number of times a rule appears in a week
transindv$count <- 1 #create count column for each rule type 
transcorp$count <- 1 #create count column for each rule type 
testindv <- aggregate(count ~ Week + `Rule Name` + Month + `Customer ID`,transindv, sum) #i've obtained all the customer ID
testcorp <- aggregate(count ~ Week + `Rule Name` + Month + `Customer ID`,transcorp, sum) #i've obtained all the customer ID

#Adding rules' scores to each rule in the testindv table
library(sqldf) #showing SQL for the sake of this practice
testindv <- sqldf("SELECT Week, `Rule Name`, `Month`, `Customer ID`, count, `Model Name`, Score, `Customer Type` 
                  FROM testindv
                  LEFT JOIN ruleindv USING(`Rule Name`)")

testcorp <- sqldf("SELECT Week, `Rule Name`, `Month`, `Customer ID`, count, `Model Name`, Score, `Customer Type` 
                  FROM testcorp
                  LEFT JOIN rulecorp USING(`Rule Name`)")

testindv <- testindv[complete.cases(testindv),]
testcorp <- testcorp[complete.cases(testcorp),]

#testindv <- merge(testindv, ruleindv, by = "Rule Name", all.x = TRUE) #can use this tooo
#testcorp <- merge(testcorp, rulecorp, by = "Rule Name")

#summing up total score per week based on rule
as.numeric.factor <- function(x) {as.numeric(levels(x))[x]}
testindv$Score <- as.numeric(testindv$Score)
testcorp$Score <- as.numeric(testcorp$Score)

groupindv <- group_by(testindv, `Customer ID`, Week, `Model Name`) #indv
groupindv <- summarize(groupindv, `Total Score` = sum(Score))
groupindv <- ungroup(groupindv)

groupcorp <- group_by(testcorp, `Customer ID`, Week, `Model Name`) #corp
groupcorp <- summarize(groupcorp, `Total Score` = sum(Score))
groupcorp <- ungroup(groupcorp)

#adding threshold of each model type to groupindv from modindv
groupindv <- sqldf("SELECT Week,  `Customer ID`, `Total Score`, `Model Name`, Threshold, Type 
                   FROM groupindv
                   LEFT JOIN modindv USING(`Model Name`)")

#groupindv <- merge(groupindv, modindv, by = "Model Name", all.x = TRUE) 

#categorize if an alert is raised
groupindv$Threshold <- as.numeric.factor(groupindv$Threshold)
groupindv$Alert <- ifelse(groupindv$`Total Score` >= groupindv$Threshold, "1", "0")

#adding threshold of each model type to groupcorp from modcorp
groupcorp <- sqldf("SELECT Week,  `Customer ID`, `Total Score`, `Model Name`, Threshold, Type 
                   FROM groupcorp
                   LEFT JOIN modcorp USING(`Model Name`)")

#groupcorp <- merge(groupcorp, modcorp, by = "Model Name", all.x = TRUE) 

#categorize if an alert is raised
groupcorp$Threshold <- as.numeric.factor(groupcorp$Threshold)
groupcorp$Alert <- ifelse(groupcorp$`Total Score` >= groupcorp$Threshold, "1", "0")

groupcorp$Alert <- as.numeric(groupcorp$Alert)
groupindv$Alert <- as.numeric(groupindv$Alert)

#to obtain back the month ....
monthset <- testindv[, -c(2,4:8)]
monthset <- unique(monthset)
monthset <- monthset[order(monthset$Week),]
monthset <- monthset[-c(5,10,15, 25),]
monthset[nrow(monthset) + 1,] = list("00","January")

groupindv <- sqldf("SELECT Week,  `Customer ID`, `Total Score`, `Model Name`, Threshold, Type, Alert, Month 
                   FROM groupindv
                   LEFT JOIN monthset USING(Week)")

groupcorp <- sqldf("SELECT Week,  `Customer ID`, `Total Score`, `Model Name`, Threshold, Type, Alert, Month 
                   FROM groupcorp
                   LEFT JOIN monthset USING(Week)")

groupindv$Alert <- as.numeric(groupindv$Alert)
groupcorp$Alert <- as.numeric(groupcorp$Alert)

x <- aggregate(Alert ~ `Customer ID`, groupcorp, sum)
y <- aggregate(Alert ~ `Customer ID`, groupindv, sum)

subset(x, Alert == max(x$Alert)) #identify the max alert for corporate customer
subset(y, Alert == max(y$Alert)) #identify the max alert for individual customer

#obtaining total alerts for individual and corporate costumers
sum(groupindv$Alert) #indv
sum(groupcorp$Alert) #corp

#----------1st Table--------#
#alerts trend table 
#individual:
alertindv <- aggregate(Alert ~ `Model Name` + Month, groupindv, sum)
#adding back the _INV
alertindv$`Model Name` <- paste(alertindv$`Model Name`, "INV", sep = "_")
#trying to get to the same output format
library(reshape2)
alertindv <- melt(alertindv, c("Month", "Model Name"), "Alert")
alertindv <- dcast(alertindv, `Model Name` ~ Month) #output indv ready for alerts trends overview 

#corporate:
alertcorp <- aggregate(Alert ~ `Model Name` + Month, groupcorp, sum)
#adding back the _CP
alertcorp$`Model Name` <- paste(alertcorp$`Model Name`, "CP", sep = "_")
#trying to get to the same output format
library(reshape2)
alertcorp <- melt(alertcorp, c("Month", "Model Name"), "Alert")
alertcorp <- dcast(alertcorp, `Model Name` ~ Month) #output indv ready for alerts trends overview 

combalert <-rbind(alertcorp,alertindv) #alert trends
combalert[is.na(combalert)] <- 0 #revalue NA to 0
combalert <- combalert[,c(1,4,3,6,2,7,5)] #reorder columns

#--------------2nd Table-------------------# 

newgroupindv <- group_by(testindv, `Customer ID`, Week, `Model Name`, `Rule Name`) #indv 
newgroupindv <- summarize(newgroupindv, `Total Score` = sum(Score))
newgroupindv <- ungroup(newgroupindv)

newgroupcorp <- group_by(testcorp, `Customer ID`, Week, `Model Name`, `Rule Name`) #corp
newgroupcorp <- summarize(newgroupcorp, `Total Score` = sum(Score))
newgroupcorp <- ungroup(newgroupcorp)

#adding threshold of each model type to newgroupindv from modindv
newgroupindv <- sqldf("SELECT Week,  `Customer ID`, `Total Score`, `Rule Name`, `Model Name`, Threshold, Type 
                      FROM newgroupindv
                      LEFT JOIN modindv USING(`Model Name`)")

#categorize if an alert is raised
newgroupindv$Threshold <- as.numeric.factor(newgroupindv$Threshold)
newgroupindv$Alert <- ifelse(newgroupindv$`Total Score` >= newgroupindv$Threshold, "1", "0")

#adding threshold of each model type to newgroupcorp from modcorp
newgroupcorp  <- sqldf("SELECT Week,  `Customer ID`, `Total Score`, `Rule Name`, `Model Name`, Threshold, Type 
                       FROM newgroupcorp
                       LEFT JOIN modcorp USING(`Model Name`)")

#categorize if an alert is raised
newgroupcorp$Threshold <- as.numeric.factor(newgroupcorp$Threshold)
newgroupcorp$Alert <- ifelse(newgroupcorp$`Total Score` >= newgroupcorp$Threshold, "1", "0")

newgroupcorp$Alert <- as.numeric(newgroupcorp$Alert)
newgroupindv$Alert <- as.numeric(newgroupindv$Alert)

#adding number of hits column
newgroupcorp$hits <- 1
newgroupindv$hits <- 1

#obtain number of hits and alerts for indv
indvnoofhits <- aggregate(Alert ~ `Rule Name`, newgroupindv, sum)
indvnoofalerts <- aggregate(hits ~ `Rule Name`, newgroupindv, sum)

indvfull <- cbind(indvnoofalerts, indvnoofhits)
indvfull <- indvfull[,-3]

#obtain number of hits and alerts for corp
corpnoofhits <- aggregate(Alert ~ `Rule Name`, newgroupcorp, sum)
corpnoofalerts <- aggregate(hits ~ `Rule Name`, newgroupcorp, sum)

corpfull <- cbind(corpnoofalerts, corpnoofhits)
corpfull <- corpfull[,-3]

#combining both tables together
finalrule <- rbind(corpfull, indvfull)
finalrule1 <- aggregate(hits ~ `Rule Name`, finalrule, sum)
finalrule2 <- aggregate(Alert ~ `Rule Name`, finalrule, sum)
finalrule <- cbind(finalrule1, finalrule2)
finalrule <- finalrule[,-3]
colnames(finalrule) <- c("Rule Name", "Number of Hits", "Number of Alerts")

