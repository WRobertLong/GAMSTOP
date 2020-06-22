###################################################
#
# Download license data from gambling commission
# And determine which accounts should be integrated
# with GAMSTOP
# 
#  Robert Long
#  19 June 2020
#
###################################################

library(tidyverse)
library(openxlsx)

# access remote form
url1 <- "https://secure.gamblingcommission.gov.uk/PublicRegister/Search"
sesh1   <- html_session(url1)
form1   <- html_form(sesh1)[[1]]
form1   <- set_values(form1, Sector = "Remote", Status = "Granted" )
subform <- submit_form(sesh1, form1)

if(subform$response$status_code != 200) stop("accessing remote data")

url2 <- "https://secure.gamblingcommission.gov.uk/PublicRegister/Search/Download"

# create unique filename from date and time
# eg down-2020-06-19-11-55.xlsx

filename1 <- paste0("Acc-",format(Sys.time(), "%Y-%m-%d-%H-%M-%S"))
filename <- paste0(filename1, ".xlsx")
# download from remote site
download <- jump_to(subform, paste0(url2, "?format=xlsx"))
writeBin(download$response$content, filename)

# Load all accounts

dt.Account <- read.xlsx(filename, sheet = 1, startRow = 8)
# simple check for
if(names(dt.Account)[1] != "Account.Number") stop ("problem with format of first sheet in excel file") 

# Are there any duplicate account numbers ?
length(unique(dt.Account$Account.Number)) != nrow(dt.Account)
# no duplicates

# any duplicate account names ?
length(unique(dt.Account$Account.Name)) != nrow(dt.Account)
#no duplicates


# domain names:
dt.Domain <- read.xlsx(filename, sheet = 3)

# Retain only Domains with UK domains
dt.Domain <- dt.Domain[grepl(".uk", dt.Domain$Url), ]

# identify any domains with missing status and remove
dt.Domain.Status.NA <- dt.Domain[is.na(dt.Domain$Status), ]
dt.Domain <- dt.Domain[complete.cases(dt.Domain), ]

# remove those with Status Inactive
dt.Domain <- dt.Domain[dt.Domain$Status != "Inactive", ]

# Obtain list of unique accounts with uk domain names
domains <- unique(dt.Domain$Account.Number)


# load activities data
dt.Activities <- read.xlsx(filename, sheet = 4)

# Any missing data 
nrow(dt.Activities) != nrow(dt.Activities[complete.cases(dt.Activities), ])
# no missing data

# identify all levels of Remote Status
table(dt.Activities$Remote.Status)
# only Ancillary, NonRemote and Remote 

# retain only those with Remote status
dt.Activities <- subset(dt.Activities, Remote.Status == "Remote")

# identify all levels of Status
table(dt.Activities$Status)

# remove all "Surrendered"
# Note: this will include 6 that are "Revoked" and 33 "Pendind"

dt.Activities <- subset(dt.Activities, Status != "Surrendered")

# Identify all levels of Activity
table(dt.Activities$Activity)

# vecror of Activites covered:

Activity <- c(
  "Casino - R",
  "Bingo - R",
  "External Lottery Manager - R",
  "General Betting Standard - Virtual Event - R",
  "General Betting Standard - Real Event - R",
  "Pool Betting - R"
  )  %>% data.frame

names(Activity) = "Activity"

# retain only those activities that are specified
dt.Activities <- inner_join(dt.Activities, Activity)

# create vector of unique accounts
activities <- unique(dt.Activities$Account.Number)


# So now we have 3 data frames
# dt.Account: accounts and related info (address etc). Assume these are Active
# dt.Domain: accounts with UK domains after excluding missing values
# dt.Activities accounts with prescribed activies
# Note that fot domain and activites, these are not unique acounts
# unique accounts are in domains and activities respectively


# domains and activities each have Status variables so make them unique 
names(dt.Domain)[3] <- "Domain.Status"
names(dt.Activities)[4] <- "Activity.Status"

dt.tmp <- inner_join(dt.Account, dt.Activities, by = "Account.Number")
dt.Final <- inner_join(dt.tmp, dt.Domain, by = "Account.Number")

nrow(dt.Final)
# 1030

# data frame of accounts only (not including domain and activities)
dt.Final.Accounts <- unique(dt.Final[ , 1:length(dt.Account)])

nrow(dt.Final.Accounts)

# save as 2 sheets in a xlsx workbook

wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = c("Final.Accounts"))
writeData(wb, sheet = 1, dt.Final.Accounts)
addWorksheet(wb = wb, sheetName = c("Final.All"))
writeData(wb, sheet = 2, dt.Final)
filenameRes <- paste0(filename1, "-Results.xlsx")

# set column widths
setColWidths(wb,
             sheet = 1,
             1:length(names(dt.Final.Accounts)),
             widths = "auto"
)
setColWidths(wb,
             sheet = 2,
             1:length(names(dt.Final)),
             widths = "auto"
)

# freeze top row
for (i in 1:2){
  freezePane(
    wb,
    sheet = i,
    firstActiveRow = 2
  )
}

# make header row bold
style <- createStyle(textDecoration = "bold")
addStyle(wb = wb, sheet = 1, style = style, rows = 1, cols = 1:length(names(dt.Final.Accounts)))
addStyle(wb = wb, sheet = 2, style = style, rows = 1, cols = 1:length(names(dt.Final)))

# save workbook
saveWorkbook(wb, filenameRes , overwrite = TRUE)
