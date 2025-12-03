
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
output.folder <- paste(workspace.root, "outcomes_2024_06_17", "game_data_preprocessing_v3", sep = "/")
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
JMIR.datasets.folder <- "D:/SynologyDriveWork/local_work_ming/study - gamification/IQ-Game Test Data/completed datasets"
dataset.overview <- GET.IQGame.Dataset.Overview(JMIR.datasets.folder, subfolder.included = T)

JMIR.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/JMIR paper/DM MS Cohort 1 Summary n=329.xlsx"
JMIR.pats.df <- read.xlsx(file=JMIR.pats.info.path, sheetIndex = 1)
JMIR.pats.game.df <- subset(dataset.overview, ProbandId %in% JMIR.pats.df$ProbandId & SessionNo==1)
JMIR.pats.df$ProbandId[!JMIR.pats.df$ProbandId %in% JMIR.pats.game.df$ProbandId]


JMIR.subcohort2.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/JMIR paper/DM MS Cohort 2 Summary n=173.xlsx"
JMIR.subcohort2.pats.df  <- read.xlsx(file=JMIR.subcohort2.pats.info.path, sheetIndex = 1)
JMIR.subcohort2.pats.game.df <- subset(dataset.overview, ProbandId %in% JMIR.subcohort2.pats.df$ProbandId & SessionNo==1)
JMIR.subcohort2.pats.df$ProbandId[!JMIR.subcohort2.pats.df$ProbandId %in% JMIR.subcohort2.pats.game.df$ProbandId]
 
JMIR.NCS.pats.info.path <- "D:/local_dev/iq_game_analyses_ming/data/JMIR paper/NCS Cohort Summary n=37.xlsx"
JMIR.NCS.pats.df  <- read.xlsx(file=JMIR.NCS.pats.info.path, sheetIndex = 1)
JMIR.NCS.pats.game.df <- subset(dataset.overview, ProbandId %in% JMIR.NCS.pats.df$ProbandId & SessionNo==1)
JMIR.NCS.pats.df$ProbandId[!JMIR.NCS.pats.df$ProbandId %in% JMIR.NCS.pats.game.df$ProbandId]






