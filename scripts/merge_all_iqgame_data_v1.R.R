
###### R Notebook that handles the preprocessing of the Gamification dataset ###############################
###### saved under the folder below: 
datasets.folder <- "D:/SynologyDriveWork/local_work_ming/study - gamification/IQ-Game Test Data V2/completed datasets v2"

workspace.root <- 'D:/SynologyDriveWork/R_workspace/gamification_study_all_analysis_R_v2'
local.workspace <- paste(workspace.root, "local_workspace", sep = "/")

## results of the validation script created by Istiyak 
data.validation.save.path <- "D:\SynologyDriveWork\R_workspace\iq_game_analysis\data/new_validation_results"

JMIR.cohort.save.path <- 
  






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


sub_data_set_folders <- list.files(datasets.folder, include.dirs = T) # get a list that contains names of all sub folders of the "datasets.folder"
dataset.paths <- paste(datasets.folder, sub_data_set_folders, sep = "/")

dataset.overview <- GET.IQGame.Dataset.Overview(datasets.folder, subfolder.included = T)
head(dataset.overview)
dataset.overview[dataset.overview$AC==""&dataset.overview$BFL==""&dataset.overview$BFR==""&dataset.overview$CP==""&dataset.overview$IJ=="", ]


Extract.IQGame.Features(dataset.path=datasets.folder, subfolder.included=T, output.path=output.folder)





















