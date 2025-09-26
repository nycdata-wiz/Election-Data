# New York City Democratic Mayoral Primary Election (June 2025): Analysis

## methods 

### Data Sources

#### Ranked Choice Tabulations - Vote Records

Election data was downloaded from [the New York City Board of Elections website](https://vote.nyc/page/election-results-summary), specifically the PRIMARY - JUNE 24, 2025 - Cast Vote Record (CVR). Notably, data was pulled before finalization and certification of the election data was completed.

The following R code was used to import the Mayoral vote data for each vote cast in ranked choice voting from all the files downloaded:

````
#Set working directory
library(rstudioapi)
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
library(openxlsx)
library(stringr)

# List files in that directory
files <- list.files(dirname(getActiveDocumentContext()$path))
files = files[!str_detect(files, "~")]
files = files[!str_detect(files, "RCV analysis")]
files = files[!str_detect(files, "hash")]
files = files[!str_detect(files, "Candidacy")]
files = files[!str_detect(files, "Candidates")]

for (x in 1:length(files)){
  if (x == 1){
    df <- read.xlsx(files[x])
    df = df[,c("Cast.Vote.Record", "Precinct",
               "DEM.Mayor.Choice.1.of.5.Citywide.(026916)",
               "DEM.Mayor.Choice.2.of.5.Citywide.(226916)", 
               "DEM.Mayor.Choice.3.of.5.Citywide.(326916)", 
               "DEM.Mayor.Choice.4.of.5.Citywide.(426916)", 
               "DEM.Mayor.Choice.5.of.5.Citywide.(526916)")]
      }
  else{
    data <- read.xlsx(files[x])
    data = data[,c("Cast.Vote.Record", "Precinct",
               "DEM.Mayor.Choice.1.of.5.Citywide.(026916)",
               "DEM.Mayor.Choice.2.of.5.Citywide.(226916)", 
               "DEM.Mayor.Choice.3.of.5.Citywide.(326916)", 
               "DEM.Mayor.Choice.4.of.5.Citywide.(426916)", 
               "DEM.Mayor.Choice.5.of.5.Citywide.(526916)")]
    df = rbind(df, data)
    print(paste("completed file", x))
  }
}

#remove people who did not vote in this election (GOP voters)
df= subset(df, `DEM.Mayor.Choice.1.of.5.Citywide.(026916)`!="undervote")
````

Next, I calculated how many votes were attributed to Zohran Mamdani or Andrew Cuomo in the final round of ranked choice voting by converting their candidate codes to their names:

````
dfFinalVote = c("Invalid")
df$FinalVote[df$`DEM.Mayor.Choice.1.of.5.Citywide.(026916)` == 254052] = "Cuomo"
df$FinalVote[df$`DEM.Mayor.Choice.1.of.5.Citywide.(026916)` == 254286] = "Mamdani"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.2.of.5.Citywide.(226916)` == 254052] = "Cuomo"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.2.of.5.Citywide.(226916)` == 254286] = "Mamdani"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.3.of.5.Citywide.(326916)` == 254052] = "Cuomo"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.3.of.5.Citywide.(326916)` == 254286] = "Mamdani"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.4.of.5.Citywide.(426916)` == 254052] = "Cuomo"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.4.of.5.Citywide.(426916)` == 254286] = "Mamdani"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.5.of.5.Citywide.(526916)` == 254052] = "Cuomo"
df$FinalVote[df$FinalVote == "Invalid" & df$`DEM.Mayor.Choice.5.of.5.Citywide.(526916)` == 254286] = "Mamdani"
````

#### Maintaining a Record of Ranked Choice Tabulations

The following name-code mapping table was used to convert the provided candidate code format to a name format:
|CandidacyID|DefaultBallotName   |
|-----------|--------------------|
|254286     |Zohran Kwame Mamdani|
|254052     |Andrew M. Cuomo     |
|257465     |Adrienne E. Adams   |
|254393     |Scott M. Stringer   |
|254130     |Brad Lander         |
|255950     |Zellnor Myrie       |
|257441     |Whitney R. Tilson   |
|254607     |Michael Blake       |
|254365     |Jessica Ramos       |
|258719     |Paperboy Love Prince|
|258821     |Selma K. Bartholomew|

````
mapping = read.csv("Candidacy ID to name table.csv", sep = ",")
df$`DEM.Mayor.Choice.1.of.5.Citywide.(026916)` = as.character(df$`DEM.Mayor.Choice.1.of.5.Citywide.(026916)`)
df$`DEM.Mayor.Choice.2.of.5.Citywide.(226916)` = as.character(df$`DEM.Mayor.Choice.2.of.5.Citywide.(226916)`)
df$`DEM.Mayor.Choice.3.of.5.Citywide.(326916)` = as.character(df$`DEM.Mayor.Choice.3.of.5.Citywide.(326916)`)
df$`DEM.Mayor.Choice.4.of.5.Citywide.(426916)` = as.character(df$`DEM.Mayor.Choice.4.of.5.Citywide.(426916)`)
df$`DEM.Mayor.Choice.5.of.5.Citywide.(526916)` = as.character(df$`DEM.Mayor.Choice.5.of.5.Citywide.(526916)`)

for (x in 1:length(mapping$CandidacyID)){
  df$`DEM.Mayor.Choice.1.of.5.Citywide.(026916)`[df$`DEM.Mayor.Choice.1.of.5.Citywide.(026916)`==mapping$CandidacyID[x]]=mapping$DefaultBallotName[x]
  df$`DEM.Mayor.Choice.2.of.5.Citywide.(226916)`[df$`DEM.Mayor.Choice.2.of.5.Citywide.(226916)`==mapping$CandidacyID[x]]=mapping$DefaultBallotName[x]
  df$`DEM.Mayor.Choice.3.of.5.Citywide.(326916)`[df$`DEM.Mayor.Choice.3.of.5.Citywide.(326916)`==mapping$CandidacyID[x]]=mapping$DefaultBallotName[x]
  df$`DEM.Mayor.Choice.4.of.5.Citywide.(426916)`[df$`DEM.Mayor.Choice.4.of.5.Citywide.(426916)`==mapping$CandidacyID[x]]=mapping$DefaultBallotName[x]
  df$`DEM.Mayor.Choice.5.of.5.Citywide.(526916)`[df$`DEM.Mayor.Choice.5.of.5.Citywide.(526916)`==mapping$CandidacyID[x]]=mapping$DefaultBallotName[x]
}

````
The final dataframe containing all of the vote data for ranked choice voting submissions can be found here across two files: [NYC_2025_RCV_allvotes_sequences_1_of_2.xlsx](NYC Democratic Mayoral Primary '25/data/NYC_2025_RCV_allvotes_sequences_1_of_2.xlsx) and [NYC_2025_RCV_allvotes_sequences_2_of_2.xlsx][NYC_2025_RCV_allvotes_sequences_2_of_2.xlsx](NYC Democratic Mayoral Primary '25/data/NYC_2025_RCV_allvotes_sequences_1_of_2.xlsx)
