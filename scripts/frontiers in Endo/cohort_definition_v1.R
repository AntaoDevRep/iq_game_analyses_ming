###### -------------------------------------
## a script that merged probands' personal data, clinical findings and game session information together 
## and build analysis cohorts according to different predifined criterias.

workspace.root <- 'C:/local_work_ming/workspaces/r_workspace/gamification_study_all_analysis_R_v2'
local.workspace <- paste(workspace.root, "local_workspace", sep = "/")
functional.scripts.folder <- paste(workspace.root, "gami_scripts_2023_05_08", "my_functional_scripts", sep = "/")
game.overview.scripts.folder <- paste(workspace.root, "gami_scripts_2023_05_08", "draw_overview_figures", sep = "/")

## probands data save path
proband.data.save.path <- "C:/local_work_ming/workspaces/r_workspace/gamification_study_all_analysis_R_v2/outcomes/proband_data_preprocessing_v1/All probands 2023-05-08.xlsx"

## output folder
output.folder <- paste(workspace.root, "outcomes", "cohort_definition_v1", sep = "/")
if (!dir.exists(output.folder)){
  dir.create(output.folder, showWarnings = T)
}

### Load library, functions 
source("C:/local_work_ming/workspaces/r_workspace/R_customized_global_functions/my_functional_scripts/my_global_functions.R")
source("C:/local_work_ming/workspaces/r_workspace/R_customized_global_functions/my_functional_scripts/my_libraries.R")
source("C:/local_work_ming/workspaces/r_workspace/R_customized_global_functions/my_functional_scripts/R_intergroup_diff_test_functions.R")

## release unused momery
gc()
### set workspace
check.wd()



##### --------------------------------- load probands' personal information ------------------------------------------------------------------------------------------------------
## load proband personal information
probands.df <- read.xlsx(file=proband.data.save.path, sheetIndex = 1)

## data cleaning
probands.df[probands.df$GroupEng=="HC", "TypeDiabetes"] <- "No"
probands.df[!is.na(probands.df$TypeDiabetes)& probands.df$TypeDiabetes=="No", "DurationDiabetes"] <- 0

probands.df[probands.df$ProbandId=="P077", ]
probands.df[probands.df$ProbandId=="P111", ]
probands.df[probands.df$ProbandId=="P589", ]

## build a label for peripheral neuropathy
probands.df$PNP <- "NoPNP"
probands.df[is.na(probands.df$NSS)|is.na(probands.df$NDSAdjusted), "PNP"] <- "NA"
probands.df[probands.df$PNP!="NA"&(probands.df$NDSAdjusted>=6|(probands.df$NDSAdjusted>=3&probands.df$NSS>=5)), "PNP"] <- "PNP"
table(probands.df$NSS, probands.df$NDSAdjusted, probands.df$PNP)

## build a label for cognitive impairment
probands.df$CognitiveFunction <- "NoCI"
probands.df[is.na(probands.df$MoCA), "CognitiveFunction"] <- "NA"
probands.df[probands.df$CognitiveFunction!="NA"&probands.df$MoCA<=25, "CognitiveFunction"] <- "CI"
table(probands.df$MoCA, probands.df$CognitiveFunction)

## assign groups
probands.df$Group <- "Not Defined"
probands.df[probands.df$GroupEng=="HC"&probands.df$PNP=="NoPNP"&probands.df$CognitiveFunction=="NoCI", "Group"]<- "HC"
probands.df[probands.df$GroupEng=="DM"&probands.df$PNP=="NoPNP"&probands.df$CognitiveFunction=="NoCI", "Group"]<- "DM"
probands.df[probands.df$GroupEng=="DM"&probands.df$PNP=="PNP"&probands.df$CognitiveFunction=="NoCI", "Group"]<- "DM+PNP"
probands.df[probands.df$GroupEng=="DM"&probands.df$PNP=="NoPNP"&probands.df$CognitiveFunction=="CI", "Group"]<- "DM+CI"
probands.df[probands.df$GroupEng=="DM"&probands.df$PNP=="PNP"&probands.df$CognitiveFunction=="CI", "Group"]<- "DM+PNP+CI"
probands.df[probands.df$GroupEng=="MCI"&probands.df$PNP=="NoPNP"&probands.df$TypeDiabetes=="No", "Group"]<- "CI"
probands.df[probands.df$GroupEng=="SCI"&probands.df$PNP=="NoPNP"&probands.df$TypeDiabetes=="No", "Group"]<- "CI"
probands.df[probands.df$GroupEng=="AD"&probands.df$PNP=="NoPNP"&probands.df$TypeDiabetes=="No", "Group"]<- "CI"
table(probands.df$GroupEng, probands.df$Group)

## build cohort 1
names(probands.df)
selected.vars <- c("ProbandId", "Group", "Gender", "Age", "Weight", "BMI", "TypeDiabetes", "DurationDiabetes", "NSSLevel", "NDSLevel", "MoCA", "EducationYears", "VisuospatialSum", "Naming", "AttentionSum", "LanguageAll", "Abstraction", "Memory", "Orientation")
cohort1.df <- subset(probands.df, select=selected.vars, Group%in%c("HC", "DM", "DM+PNP", "DM+CI", "DM+PNP+CI", "CI"))
cohort1.df$Group <- factor(cohort1.df$Group, levels = c("HC", "DM", "DM+CI", "DM+PNP", "DM+PNP+CI", "CI"))
table(cohort1.df$Group)

cohort1.summary <- build.df.summary(target.df = cohort1.df, skip.column.names = c("ProbandId"), effective.digits = 1, debug.activated = F, only.mean.sd = F, group.test.enabled=F)
cohort1.summary

NA.df <- cohort1.df[!complete.cases(cohort1.df), ]
cohort1.df <- na.omit(cohort1.df)

write.xlsx(cohort1.summary, file = paste(output.folder, "cohort1 report.xlsx", sep = "/"),
           sheetName = "Report1", append = F, row.names = FALSE)
write.xlsx(cohort1.df, file = paste(output.folder, "cohort1 report.xlsx", sep = "/"),
           sheetName = "No NA Data", append = T, row.names = FALSE)
write.xlsx(NA.df, file = paste(output.folder, "cohort1 report.xlsx", sep = "/"),
           sheetName = "NA rows", append = T, row.names = FALSE)

# cohort1.df[is.na(cohort1.df$Weight), ]
# cohort1.df[is.na(cohort1.df$BMI), ]
# cohort1.df[is.na(cohort1.df$TypeDiabetes), ]
# cohort1.df[is.na(cohort1.df$DurationDiabetes), ]
# cohort1.df[is.na(cohort1.df$EducationYears), ]
# cohort1.df[is.na(cohort1.df$Abstraction), ]
# cohort1.df[is.na(cohort1.df$VisuospatialSum), ]
# cohort1.df[is.na(cohort1.df$Orientation), ]


#visualize.variables.betw.groups(target.df = cohort1.df, skip.column.names = c("ProbandId"), group.label.name="Group", fig.title ="Cohort 1 variables", figure.output.folder = output.folder)

table(probands.df$VibrationR)
table(probands.df$VibrationL)

head(probands.df)
probands.df$PNP
probands.df$CI <- probands.df$CognitiveFunction

selected.vars <- c("ProbandId", "Group", "PNP", "CI", "Gender", "Age", "Weight", "BMI", "TypeDiabetes", "DurationDiabetes", "NSS", "NDSAdjusted", "NSSLevel", "NDSLevel", "MoCA", "VisuospatialSum", "Naming", "AttentionSum", "LanguageAll", "Abstraction", "Memory", "Orientation",
                   "PinprickR", "PinprickL", "TempSensationR", "TempSensationL", "VibrationLScore", "VibrationRScore", "ReflexR", "ReflexL", "MonofilamentR", "MonofilamentL", "VibrationL", "VibrationR")
cohort2.df <- subset(probands.df, select=selected.vars, Group%in%c("HC", "DM"))
cohort2.df$Group <- factor(cohort2.df$Group, levels = c("HC", "DM"))
## check NA values
cohort2.df[!complete.cases(cohort2.df), ]
cohort2.df <- na.omit(cohort2.df)
table(cohort2.df$Group)


cohort3.df <- subset(probands.df, select=selected.vars, Group%in%c("DM", "DM+PNP", "DM+CI", "DM+PNP+CI"))
cohort3.df$Group <- factor(cohort3.df$Group, levels = c("DM", "DM+CI", "DM+PNP", "DM+PNP+CI"))
## check NA values
cohort3.df[!complete.cases(cohort3.df), ]
cohort3.df <- na.omit(cohort3.df)
table(cohort3.df$Group)

write.xlsx(cohort2.df, file = paste(output.folder, "cohort2 HC DM report.xlsx", sep = "/"),
           sheetName = "HC DM", append = F, row.names = FALSE)
write.xlsx(cohort3.df, file = paste(output.folder, "cohort3 DM PNP CI report.xlsx", sep = "/"),
           sheetName = "DM DM", append = F, row.names = FALSE)

