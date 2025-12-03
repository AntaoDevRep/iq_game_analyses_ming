
###### R Notebook that handles the preprocessing of the Gamification dataset ###############################
###### saved under the folder below: 
datasets.folder <- "D:/SynologyDriveWork/local_work_ming/study - gamification/IQ-Game Test Data V2/completed datasets v2"

workspace.root <- 'D:/local_dev/iq_game_analyses_ming'
local.workspace <- paste(workspace.root, "local_workspace", sep = "/")

## results of the validation script created by Istiyak 
data.validation.save.path <- paste(workspace.root, "data", "new_validation_results.xlsx", sep = "/")

JMIR.cohort.save.path <- paste(workspace.root, "data", "new_validation_results.xlsx", sep = "/")
ECliniMed.cohort.save.path <- paste(workspace.root, "data", "eCliniMed cohort gaming data sets.xlsx", sep = "/")

## output folder
output.folder <- paste(workspace.root, "outcomes", sep = "/")
if (!dir.exists(output.folder)){
  dir.create(output.folder, showWarnings = T)
}

## attach functional scripts
source("D:/SynologyDriveWork/R_workspace/R_customized_global_functions/my_functional_scripts/my_global_functions.R")
source("D:/SynologyDriveWork/R_workspace/R_customized_global_functions/my_functional_scripts/my_libraries.R")
source("D:/SynologyDriveWork/R_workspace/R_customized_global_functions/my_functional_scripts/R_intergroup_diff_test_functions.R")
source("D:/SynologyDriveWork/R_workspace/R_customized_global_functions/my_functional_scripts/iq_game_feature_extraction_v1.R")

## setup workspace
setwd(local.workspace)

m.font.size <- 24
m.theme <- theme(text = element_text(size = m.font.size,  family="sans", color = "black"),
                 plot.title = element_textbox_simple(size = m.font.size, face="bold"),
                 axis.text.y = element_text(size = m.font.size, color = "black"),
                 axis.title.y = element_text(size = m.font.size, color = "black", face="bold"),
                 axis.text.x =element_text(size = m.font.size, color = "black"),
                 axis.title.x = element_text(size = m.font.size, color = "black", face="bold"),
                 strip.text = element_text(size = m.font.size, color = "black", face="bold"),
                 legend.text = element_text(size = m.font.size, color = "black"),
                 panel.grid.major.x = element_blank(), # add horizontal grid
                 panel.grid.major.y = element_blank(), # remove vertical grid
                 panel.grid.minor = element_blank(), # remove grid
                 panel.background = element_rect(fill = "white"), # remove background
                 axis.line = element_line(colour = "black"), # draw axis line in black
                 legend.position = "right")


valida.df <- read.xlsx(file=data.validation.save.path, sheetIndex = 1)

ECliniMed.sit.df <- read.xlsx(file=ECliniMed.cohort.save.path, sheetIndex = 1)
ECliniMed.stand.df <- read.xlsx(file=ECliniMed.cohort.save.path, sheetIndex = 2)

valida.df$ECliniMed.Seated <- ""
valida.df[valida.df$testId_in_file %in% ECliniMed.sit.df$TestId, "ECliniMed.Seated"] <- "x"
table(valida.df$ECliniMed.Seated)

valida.df$ECliniMed.Stand <- ""
valida.df[valida.df$testId_in_file %in% ECliniMed.stand.df$TestId, "ECliniMed.Stand"] <- "x"
table(valida.df$ECliniMed.Stand)


## ------------ JMIR paper cohort ---------------------------------------------------
# get information of gaming data set
JMIR.datasets.folder <- "D:/SynologyDriveWork/local_work_ming/study - gamification/IQ-Game Test Data/completed datasets"
dataset.overview <- GET.IQGame.Dataset.Overview(JMIR.datasets.folder, subfolder.included = T)

# get patients of the entire cohort (n=329)
JMIR.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/JMIR paper/DM MS Cohort 1 Summary n=329.xlsx"
JMIR.pats.df <- read.xlsx(file=JMIR.pats.info.path, sheetIndex = 1)
JMIR.pats.game.df <- subset(dataset.overview, ProbandId %in% JMIR.pats.df$ProbandId & SessionNo==1)
JMIR.pats.df$ProbandId[!JMIR.pats.df$ProbandId %in% JMIR.pats.game.df$ProbandId]

# valida.df$JMIR.Overall <- ""
# valida.df[valida.df$testId_in_file %in% JMIR.pats.game.df$TestId, "JMIR.Overall"] <- "x"
# table(valida.df$JMIR.Overall)

# get patients of the subcohort2 (n=173, age>55, MoCA>25, first session)
JMIR.subcohort2.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/JMIR paper/DM MS Cohort 2 Summary n=173.xlsx"
JMIR.subcohort2.pats.df  <- read.xlsx(file=JMIR.subcohort2.pats.info.path, sheetIndex = 1)
JMIR.subcohort2.pats.game.df <- subset(dataset.overview, ProbandId %in% JMIR.subcohort2.pats.df$ProbandId & SessionNo==1)
JMIR.subcohort2.pats.df$ProbandId[!JMIR.subcohort2.pats.df$ProbandId %in% JMIR.subcohort2.pats.game.df$ProbandId]

valida.df$JMIR.Subcohort2 <- ""
valida.df[valida.df$testId_in_file %in% JMIR.subcohort2.pats.game.df$TestId, "JMIR.Subcohort2"] <- "x"
table(valida.df$JMIR.Subcohort2)
 
# get patients of the subcohort1 (n=37, with NCS, age>50, MoCA>25, first session)
JMIR.NCS.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/JMIR paper/NCS Cohort Summary n=37.xlsx"
JMIR.NCS.pats.df  <- read.xlsx(file=JMIR.NCS.pats.info.path, sheetIndex = 1)
JMIR.NCS.pats.game.df <- subset(dataset.overview, ProbandId %in% JMIR.NCS.pats.df$ProbandId & SessionNo==1)
JMIR.NCS.pats.df$ProbandId[!JMIR.NCS.pats.df$ProbandId %in% JMIR.NCS.pats.game.df$ProbandId]

valida.df$JMIR.NCScohort2 <- ""
valida.df[valida.df$testId_in_file %in% JMIR.NCS.pats.game.df$TestId, "JMIR.NCScohort2"] <- "x"
table(valida.df$JMIR.NCScohort2)

## ------------ Frontiers paper cohort ---------------------------------------------------
# get information of gaming data set
Frontiers.datasets.info.path <- "D:/local_dev/iq_game_analyses_ming/data/Frontiers paper/Feature.Extration.Report.xlsx"
Frontiers.datasets.df  <- read.xlsx(file=Frontiers.datasets.info.path, sheetIndex = 1)

# get patients with complete gaming data set  (n=247)
Frontiers.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/Frontiers paper/with game data Cohort3 DM probands.xlsx"
Frontiers.pats.df  <- read.xlsx(file=Frontiers.pats.info.path, sheetIndex = 2)
Frontiers.pats.game.df <- subset(dataset.overview, ProbandId %in% Frontiers.pats.df$ProbandId & SessionNo==1)

# select the first session and merge gaming data sets to the Fronties paper cohort (n=247)
Frontiers.pats.game.df <- subset(Frontiers.datasets.df, ProbandId %in% Frontiers.pats.df$ProbandId)
Frontiers.pats.game.df <- Frontiers.pats.game.df[order(Frontiers.pats.game.df$ProbandId, Frontiers.pats.game.df$TestDate, Frontiers.pats.game.df$TestTime), ]
sorted.Frontiers.pats.game.df <- data.frame()
for(n.pro in unique(Frontiers.pats.game.df$ProbandId)){
  n.pro.df <- subset(Frontiers.pats.game.df, ProbandId==n.pro)
  n.pro.df <- n.pro.df[order(n.pro.df$TestDate, n.pro.df$TestTime), ]
  if (nrow(n.pro.df)>1){
    #print(n.pro.df)
    sorted.Frontiers.pats.game.df <- rbind(sorted.Frontiers.pats.game.df, n.pro.df[1, ])
  } else {
    sorted.Frontiers.pats.game.df <- rbind(sorted.Frontiers.pats.game.df, n.pro.df)
  }
}
sorted.Frontiers.pats.game.df


valida.df$Frontiers.Cohort <- ""
valida.df[valida.df$testId_in_file %in% sorted.Frontiers.pats.game.df$TestId, "Frontiers.Cohort"] <- "x"
table(valida.df$Frontiers.Cohort)



## ------------ Frontiers paper cohort ---------------------------------------------------
write.xlsx(valida.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "Valid", append = F, row.names = FALSE)
write.xlsx(ECliniMed.sit.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "ECliniMed Seated", append = T, row.names = FALSE)
write.xlsx(ECliniMed.stand.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "ECliniMed Stand", append = T, row.names = FALSE)

write.xlsx(JMIR.pats.game.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "JMIR entire", append = T, row.names = FALSE)
write.xlsx(JMIR.NCS.pats.game.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "JMIR NCS", append = T, row.names = FALSE)
write.xlsx(JMIR.subcohort2.pats.game.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "JMIR Subcohort2", append = T, row.names = FALSE)


write.xlsx(sorted.Frontiers.pats.game.df, file = paste(output.folder, "validation_results.xlsx", sep = "/"), sheetName = "Frontiers", append = T, row.names = FALSE)


# get information of gaming data set
JMIR.datasets.folder <- "D:/SynologyDriveWork/local_work_ming/study - gamification/IQ-Game Test Data V2/completed datasets v2"
dataset.overview2 <- GET.IQGame.Dataset.Overview(JMIR.datasets.folder, subfolder.included = T)


write.xlsx(dataset.overview2, file = paste(output.folder, "antao_datasets_overview.xlsx", sep = "/"), sheetName = "Frontiers", append = T, row.names = FALSE)





