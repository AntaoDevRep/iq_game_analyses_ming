###### -------------------------------------
## a script that merged probands' personal data, clinical findings and game session information together 
## and build analysis cohorts according to different predifined criterias.
## following preprocessing steps were performed:
## 1. exclude patients with nerve system diseases
## 2. exclude patient P201 without clinical examination
## 3. exclude test persons
## 4. exclude probands with two Proband IDs, only keep one ID per person
## 5. add a new group label:"HC", "DM", "MetS", "MCI", "SCI", "AD", "CKD", "Other"
## 6. adjust vibration sensation and NDS according of calculated age of participants
## 7. format variables

workspace.root <- 'C:/local_work_ming/workspaces/r_workspace/gamification_study_all_analysis_R_v2'
local.workspace <- paste(workspace.root, "local_workspace", sep = "/")
functional.scripts.folder <- paste(workspace.root, "gami_scripts_2023_05_08", "my_functional_scripts", sep = "/")
game.overview.scripts.folder <- paste(workspace.root, "gami_scripts_2023_05_08", "draw_overview_figures", sep = "/")

## probands data save path
proband.data.save.path <- "C:/local_work_ming/study - gamification/Gamification Test Ordner/Masterdatei/2023-05-08/Gamification Probandendaten.xlsx"
proband.data.table.password <- "Gamification"
proband.diagnose.save.path <- "C:/local_work_ming/study - gamification/Gamification Test Ordner/Masterdatei/2023-05-08/Gamification_Masterdatei.xlsx"

## output folder
output.folder <- paste(workspace.root, "outcomes", "proband_data_preprocessing_v1", sep = "/")
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
proband.data.table <- read.xlsx(file=proband.data.save.path, sheetIndex = 1, password = proband.data.table.password)
proband.data.df <- subset(proband.data.table[, c(1:5, 10:13)], Name != "NA" & Name != "")
names(proband.data.df) <- c("ProbandId", "Surname", "Givenname", "Gender", "Group", "YearOfBirth", "Age", "Weight", "ShoeSize")
head(proband.data.df)

##### --------------------------------- load proband´s diagnostic findings -------------------------------------------------------------------------------------------------------
proband.diagnose.table <- read.xlsx(file=proband.diagnose.save.path, sheetIndex = 2)
names(proband.diagnose.table)[c(1:4, 25, 34, 39:48)]

proband.diagnose.df <- proband.diagnose.table[, c(1:4, 7, 9:18, 21, 24, 25, 34, 39:48)]
names(proband.diagnose.df) <- c("ProbandId", "TypeDiabetes", "DiabetesSince", "BMI", "Vision",  "Hyperlipidemia", "Retinopathy",
                                "Nephropathy", "Hypertension", "Gout", "NervousSystem", "Musculature", "SpineLegs", "FootUlcer",
                                "Handedness", "MovementRestrictions", "AppExperience", "NSS", "NDS", "ReflexR", "ReflexL", "VibrationR", "VibrationL",
                                "PinprickR", "PinprickL", "TempSensationR", "TempSensationL", "MonofilamentR", "MonofilamentL")
head(proband.diagnose.df)

##### --------------------------------- load MoCA test results -------------------------------------------------------------------------------------------------------------------
proband.MoCA.table <- read.xlsx(file=proband.diagnose.save.path, sheetIndex = 3, row.names=NULL)
names(proband.MoCA.table)[1:20] <- c("ProbandId", "Group", "ExecutiveTest", "CopyCube", "DrawClock", "VisuospatialSum", "Naming", 
                                     "DigitsList", "LettersList", "Subtraction", "AttentionSum", "SentenceRepeat", "FluencyNamingWords", "NumberWords",
                                     "LanguageAll", "Abstraction", "Memory", "Orientation", "EducationYears", "MoCA")
proband.MoCA.df <- proband.MoCA.table[, c(1, 3:20)]
head(proband.MoCA.df)
describe(proband.MoCA.table$MoCA)

##### --------------------------------- merge personal information and diagnostic findings ---------------------------------------------------------------------------------------
proband.full.df <- join(proband.data.df, proband.diagnose.df,  by="ProbandId")
proband.full.df <- join(proband.full.df, proband.MoCA.df,  by="ProbandId")
head(proband.full.df)

##### --------------------------------- check medical history ---------------------------------------------------------------------------------------
proband.full.df$Hyperlipidemia <- factor(proband.full.df$Hyperlipidemia, levels = c(0, 1), labels = c("No", "Yes"))
proband.full.df$Retinopathy <- factor(proband.full.df$Retinopathy, levels = c(0, 1), labels = c("No", "Yes"))
proband.full.df$Nephropathy <- factor(proband.full.df$Nephropathy, levels = c(0, 1), labels = c("No", "Yes"))
proband.full.df$Hypertension <- factor(proband.full.df$Hypertension, levels = c(0, 1), labels = c("No", "Yes"))
proband.full.df$Gout <- factor(proband.full.df$Gout, levels = c(0, 1), labels = c("No", "Yes"))

## Nervensystem (1:Polyneuropathie, 2: Depression, 3: Schädelhirntrauma, 4: Epilepsie, 5: Schlaganfall,6: Tremor, 7: transitorische Ischämie, 
##               8: Depressionen, 9: Fußheberschwäche, 10: Karpaltunnel, 11: Embolie Auge, 12: Hinblutung, 13: Sonstige)
# nerve.system.explanations <- c("No", "Polyneuropathy", "Depression", "Traumatic Brain Injury", "Epilepsy", "Stroke", "Tremor", "Transient Ischemia", 
#                                "Depression2", "Foot Dorsiflexion Weakness", "Carpal Tunnel", "Embolism Eye", "Outward Bleeding","Other")
nerve.system.explanations <- c("No", "Polyneuropathie", "Depression", "Schädelhirntrauma", "Epilepsie", "Schlaganfall", "Tremor", "transitorische Ischämie", 
                               "Depressionen", "Fußheberschwäche", "Karpaltunnel", "Embolie Auge", "Hinblutung", "Sonstige")


### a function that transfers data to clinical explanations
### for instance: "1,3,5" -> "Polyneuropathy", "TraumaticBrainInjury", "Stroke"
### the separator could be "," or ";"
split.strings <- function(record, explan.vectors){
  #print(paste("Convert the record", record, "according to the explanation:"))
  #print(explan.vectors)
  if (!is.na(record)&&is.character(record)&&record!=""&&record!="FALSE"&& !grepl("VLOOKUP", record, fixed = TRUE) ){
    if (grepl("(", record, fixed = TRUE) && grepl(")", record, fixed = TRUE)){
      record <- substr(record, 1, 1)
      print(record)
      print(explan.vectors)
    }
    
    strings.got <- c()
    new.value <- ""
    for (n.char in 1:nchar(record)) {
      one.char <- substr(record, n.char, n.char)
      if (one.char!=","&&one.char!=";"&&one.char!="."){
        new.value <- paste(new.value, one.char, sep = "")
      } else {
        strings.got <- c(strings.got, new.value)
        new.value <- ""
      }
    }
    if (nchar(new.value)>0){
      strings.got <- c(strings.got, new.value)
    }
    # print(paste("Split [", record, "] into multiple strings:"))
    # print(strings.got)
    
    if (length(strings.got)>0){
      record.indexes <- as.numeric(strings.got)+1
      if(max(record.indexes)<=length(explan.vectors)){
        output <- paste(explan.vectors[record.indexes], collapse=", ")
        # print(paste(record, "->", output))
        return(output)
      } else {
        over.range.indexes <- record.indexes[record.indexes>length(explan.vectors)]
        output<-paste(explan.vectors[setdiff(record.indexes, over.range.indexes)], collapse=", ")
        print(paste(record, "->", output))
        print(paste("No clinical explanation for the record:", over.range.indexes))
        return(output)
      }
    }
  }
  return("")
}

proband.full.df$NervousSystem <- sapply(proband.full.df$NervousSystem, split.strings, nerve.system.explanations)
proband.full.df$NervousSystem <- factor(proband.full.df$NervousSystem, levels = nerve.system.explanations)
table(proband.full.df$NervousSystem) 

# #Muskulatur (0:keine, 1: Krämpfe, 2: Quadrizepssehnenruptur, 3: Rotatorenmanschettenruptur, 4: Sonstiges)
# proband.full.df$Musculature <- sapply(proband.full.df$Musculature, split.strings, c("No", "Cramps", "Quadriceps Tendon Rupture", "Rotator Cuff Rupture", "Other"))
# describe(proband.full.df$Musculature)

#WS/Beine (0: keine, 1: Bandscheibenvorfälle, 2: Verformungen, 3: Veränderte Bandscheiben,4: Hüftgelenk, 5: Arthrose, 6: Spinalkanalstenose, 7: LWS-Syndrom, 8: Wirbelkörperversteifung, 9: Rückenmarksschäden, 10: Cervical-Syndrom, 11: Rheuma, 12: Osteoporose, 13: Ischias, 14: Hallux vagus, 15: Knie-TEP, 16: HWS-Syndrom, 17: Lumbalgie, 18: Skoliose, 19: Knie-TEP, 20: pAVK, 21 Hüft-TEP, 22: Bandscheibenimplantat, 23: chron. Rückenschmerzen,24: Lumboischalgiegelenk, 25: Fraktur, 26: Sonstige
SpineLegs.explainations <- c("No", "Bandscheibenvorfälle", "Verformungen", "Veränderte Bandscheiben", "Hüftgelenk", "Arthrose", "Spinalkanalstenose", "LWS-Syndrom", "Wirbelkörperversteifung", "Rückenmarksschäden", "Cervical-Syndrom", "Rheuma",
                             "Osteoporose", "Ischias", "Hallux vagus", "Knie-TEP2", "HWS-Syndrom", "Lumbalgie", "Skoliose", "Knie-TEP", "pAVK", "Hüft-TEP", "Bandscheibenimplantat", "chron. Rückenschmerzen", "Lumboischalgiegelenk", "Fraktur", "Sonstige", "Kreuzband-OP")
proband.full.df$SpineLegs <- sapply(proband.full.df$SpineLegs, split.strings, SpineLegs.explainations)
proband.full.df$SpineLegs <- factor(proband.full.df$SpineLegs, levels = SpineLegs.explainations)
table(proband.full.df$SpineLegs)


proband.full.df$FootUlcer <- factor(proband.full.df$FootUlcer, levels = c(0, 1), labels = c("No", "Yes"))
table(proband.full.df$FootUlcer)

proband.full.df$Handedness <- factor(proband.full.df$Handedness, levels = c(0, 1, 2), labels = c("Both", "Right", "Left"))
table(proband.full.df$Handedness)

#0: nein, 1: Adipositas, 2: Schwindel/Sturzgefahr, 3: Bandscheiben, 4: Gelenke, 5: Neurologisch, 6: Sonstige
MovementRestrictions.Explanations <- c("No", "Adipositas", "Schwindel/Sturzgefahr", "Bandscheiben", "Gelenke", "Neurologisch", "Sonstige")
proband.full.df$MovementRestrictions <- sapply(proband.full.df$MovementRestrictions, split.strings,
                                               MovementRestrictions.Explanations)
proband.full.df$MovementRestrictions <- factor(proband.full.df$MovementRestrictions, levels = MovementRestrictions.Explanations)
table(proband.full.df$MovementRestrictions)

##App-Erfahrungen (0: keine, 1: soziale Medien, Nachrichten, 2: Spiele (h/Monat), 3: Spiele (h/Woche), 4: Spiele tgl)
proband.full.df$AppExperience <- sapply(proband.full.df$AppExperience, split.strings,
                                        c("No", "SocialMedian", "Gaming Monthly", "Gaming Weekly", "Gaming Daily"))
proband.full.df$AppExperience <- factor(proband.full.df$AppExperience, levels = c("No", "SocialMedian", "Gaming Monthly", "Gaming Weekly", "Gaming Daily"))
table(proband.full.df$AppExperience)

proband.full.df[proband.full.df$ProbandId=="P618", ]

##### --------------------------------- exclude patients with nerve system diseases ---------------------------------------------------------------------------------------
table(proband.full.df$NervousSystem) 
exclude.proband.IDs <- proband.full.df[proband.full.df$NervousSystem %in% c("Schlaganfall", "Tremor", "transitorische Ischämie", "Schädelhirntrauma", "Epilepsie"), "ProbandId"]
print(paste("Exclude", length(exclude.proband.IDs), "patients with nerve system diseases:"))
print(exclude.proband.IDs)
proband.full.df <- subset(proband.full.df, !ProbandId %in% exclude.proband.IDs)
table(proband.full.df$NervousSystem)

##### --------------------------------- exclude probands without clinical examination or test persons ---------------------------------------------------------------------------------------
proband.full.df[proband.full.df$ProbandId=="P201", ]
proband.full.df <- subset(proband.full.df, ProbandId != "P201" & Group!="Testperson")
table(proband.full.df$Group)

##### --------------------------------- exclude probands with two Proband IDs, keep the second one ---------------------------------------------------------------------------------------
exclude.repeated.ID.df <- data.frame(FirstId=c("P039", "P040", "P045", "P046", "P070", "P078", "P079", "P091", "P103",  "P291", "P298", "P292", "P133", "P219", "P256", "P259"),
                                     SecondId=c("P255", "P254", "P321", "P320", "P309", "P314", "P238", "P560", "P234", "P107", "P114", "P116", "P281", "P269", "P290", "P297"))
proband.full.df <- subset(proband.full.df, !ProbandId %in% exclude.repeated.ID.df$FirstId)

##### --------------------------------- add a new group label ---------------------------------------------------------------------------------------
for (n.pro in 1:nrow(proband.full.df)) {
  #for (n.pro in 1:3) {
  if (is.na(proband.full.df$Group[n.pro])){
    proband.full.df$GroupEng[n.pro] <- "Other"
  } else if (grepl("Diabetiker", proband.full.df$Group[n.pro], fixed = T)){
    proband.full.df$GroupEng[n.pro] <- "DM"
  } else if (grepl("Metaboli", proband.full.df$Group[n.pro], fixed = T)){
    proband.full.df$GroupEng[n.pro] <- "MetS"
  } else if (grepl("dialyse", proband.full.df$Group[n.pro], fixed = T)){
    proband.full.df$GroupEng[n.pro] <- "CKD"
  } else if (proband.full.df$Group[n.pro] == "gesund"||proband.full.df$Group[n.pro] == "Gesund"){
    proband.full.df$GroupEng[n.pro] <- "HC"
  } else if (proband.full.df$Group[n.pro] == "Demenz"){
    proband.full.df$GroupEng[n.pro] <- "AD"
  } else if (proband.full.df$Group[n.pro] == "MCI"){
    proband.full.df$GroupEng[n.pro] <- "MCI"
  } else if (proband.full.df$Group[n.pro] == "SCI"){
    proband.full.df$GroupEng[n.pro] <- "SCI"
  } else {
    proband.full.df$GroupEng[n.pro] <- "Other"
  }
}
proband.full.df$GroupEng <- factor(proband.full.df$GroupEng, levels = c("HC", "DM", "MetS", "MCI", "SCI", "AD", "CKD", "Other"))
table(proband.full.df$Group, proband.full.df$GroupEng)
table(proband.full.df$GroupEng)


##### --------------------------------- format variables ---------------------------------------------------------------------------------------
proband.full.df$Gender <- factor(proband.full.df$Gender, levels = c("m", "w"), labels = c("Male", "Female"))
proband.full.df <- subset(proband.full.df, select=-c(Group))
# age ----------------------
proband.full.df$Age <- as.numeric(proband.full.df$Age)
proband.full.df$TestYear <- proband.full.df$YearOfBirth+as.numeric(proband.full.df$Age)
proband.full.df[is.na(proband.full.df$TestYear)|proband.full.df$TestYear>2023, c("ProbandId", "GroupEng", "YearOfBirth", "Age", "TestYear")]
proband.full.df <- subset(proband.full.df, select=-c(YearOfBirth))

proband.full.df$Weight <- as.numeric(proband.full.df$Weight)
proband.full.df$ShoeSize <- as.numeric(proband.full.df$ShoeSize)
proband.full.df$TypeDiabetes <- factor(proband.full.df$TypeDiabetes, levels = c(0, 1, 2, 3), labels = c("No", "Type 1", "Type 2", "Type 3"))

# duration of diabetes -----
proband.full.df$DiabetesSince <- as.numeric(proband.full.df$DiabetesSince)
proband.full.df[!is.na(proband.full.df$DiabetesSince)&proband.full.df$DiabetesSince==0, "DiabetesSince"] <- NA
proband.full.df$DurationDiabetes <- proband.full.df$TestYear - as.numeric(proband.full.df$DiabetesSince)
proband.full.df[!is.na(proband.full.df$DurationDiabetes)&proband.full.df$DurationDiabetes<0, "DurationDiabetes"] <-0
describe(proband.full.df$DurationDiabetes)
proband.full.df <- subset(proband.full.df, select=-c(DiabetesSince))

proband.full.df$BMI <- as.numeric(proband.full.df$BMI)
proband.full.df$BMI <- sapply(proband.full.df$BMI, round, 1)
proband.full.df$Vision <- as.numeric(proband.full.df$Vision)

##----------------------- adjust vibration sensation and NDS according of calculated age of participants ---------------------------------------
head(proband.full.df)
proband.full.df$VibrationLScore <- NA
proband.full.df$VibrationRScore <- NA
proband.full.df$NDScalculated <-NA
proband.full.df$NDSAdjusted <-NA

for (n.proband in 1:nrow(proband.full.df)) {
  if (!is.na(proband.full.df$Age[n.proband])){
    if (proband.full.df$Age[n.proband]>70){
      if (!is.na(proband.full.df$VibrationL[n.proband])){
        if (proband.full.df$VibrationL[n.proband]<4){
          proband.full.df$VibrationLScore[n.proband] <- 1
          print(paste("Adjust left foot vibration threshold for proband", proband.full.df$ProbandId[n.proband]))
        } else {
          proband.full.df$VibrationLScore[n.proband] <- 0
        }
      }
      
      if (!is.na(proband.full.df$VibrationR[n.proband])){
        if (proband.full.df$VibrationR[n.proband]<4){
          proband.full.df$VibrationRScore[n.proband] <- 1
          print(paste("Adjust right foot vibration threshold for proband", proband.full.df$ProbandId[n.proband]))
        } else {
          proband.full.df$VibrationRScore[n.proband] <- 0
        }        
      }
    } else {
      if (!is.na(proband.full.df$VibrationL[n.proband])){
        if (proband.full.df$VibrationL[n.proband]<5){
          proband.full.df$VibrationLScore[n.proband] <- 1
        } else {
          proband.full.df$VibrationLScore[n.proband] <- 0
        }
      }
      
      if (!is.na(proband.full.df$VibrationR[n.proband])){
        if (proband.full.df$VibrationR[n.proband]<5){
          proband.full.df$VibrationRScore[n.proband] <- 1
        } else {
          proband.full.df$VibrationRScore[n.proband] <- 0
        }        
      }
    }
    
    ## calculate NDS without adjustment of virbation thresholds
    if (!is.na(proband.full.df$VibrationL[n.proband])){
      if (proband.full.df$VibrationL[n.proband]<5){
        proband.full.df$VibrationLScoreNoAdj[n.proband] <- 1
      } else {
        proband.full.df$VibrationLScoreNoAdj[n.proband] <- 0
      }
    } else {
      proband.full.df$VibrationLScoreNoAdj[n.proband] <- NA
    }

    if (!is.na(proband.full.df$VibrationR[n.proband])){
      if (proband.full.df$VibrationR[n.proband]<5){
        proband.full.df$VibrationRScoreNoAdj[n.proband] <- 1
      } else {
        proband.full.df$VibrationRScoreNoAdj[n.proband] <- 0
      }
    } else {
      proband.full.df$VibrationRScoreNoAdj[n.proband] <- NA
    }
    
    proband.full.df$NDScalculated[n.proband] <- proband.full.df$VibrationLScoreNoAdj[n.proband] + proband.full.df$VibrationRScoreNoAdj[n.proband]+
      proband.full.df$TempSensationL[n.proband]+proband.full.df$TempSensationR[n.proband]+
      proband.full.df$PinprickL[n.proband]+proband.full.df$PinprickR[n.proband]+
      proband.full.df$ReflexL[n.proband]+proband.full.df$ReflexR[n.proband]
    
    proband.full.df$NDSAdjusted[n.proband] <- proband.full.df$VibrationRScore[n.proband] + proband.full.df$VibrationLScore[n.proband]+
      proband.full.df$TempSensationL[n.proband]+proband.full.df$TempSensationR[n.proband]+
      proband.full.df$PinprickL[n.proband]+proband.full.df$PinprickR[n.proband]+
      proband.full.df$ReflexL[n.proband]+proband.full.df$ReflexR[n.proband]
  }
}

NDS.diff.df <- proband.full.df[proband.full.df$NDS!=proband.full.df$NDSAdjusted|proband.full.df$NDS!=proband.full.df$NDScalculated, c("ProbandId", "Surname", "Givenname", "GroupEng", "NDS", "NDScalculated", "ReflexL", "ReflexR", "TempSensationL", "TempSensationR", "PinprickL", "PinprickR", "VibrationL", "VibrationR", "VibrationLScoreNoAdj", "VibrationRScoreNoAdj", "NDSAdjusted", "VibrationLScore", "VibrationRScore")]
NDS.diff.df$NDS.Diff <- NDS.diff.df$NDSAdjusted-NDS.diff.df$NDS
NDS.diff.df$NDS.Diff.Abs <- abs(NDS.diff.df$NDSAdjusted-NDS.diff.df$NDS)
write.xlsx(NDS.diff.df,
           file = paste(output.folder, "wrong NDS probands.xlsx", sep = "/"),
           sheetName = "All", append = F, row.names = FALSE)

### ----  neuropathy scores -----------------------
proband.full.df$NSSLevel <- sapply(proband.full.df$NSS, paper.NSS.four.levels)
proband.full.df$NDSLevel <- sapply(proband.full.df$NDSAdjusted, paper.NDS.four.levels)

### ----  neuropathy clinical examination -----------------------
proband.full.df$VibrationR <- sapply(proband.full.df$VibrationR, paper.vibration.to.factor)
proband.full.df$VibrationL <- sapply(proband.full.df$VibrationL, paper.vibration.to.factor)

proband.full.df$PinprickR <- sapply(proband.full.df$PinprickR, paper.pinprick.to.levels)
proband.full.df$PinprickL <- sapply(proband.full.df$PinprickL, paper.pinprick.to.levels)

proband.full.df$TempSensationR <- sapply(proband.full.df$TempSensationR, paper.temperature.to.levels)
proband.full.df$TempSensationL <- sapply(proband.full.df$TempSensationL, paper.temperature.to.levels)

proband.full.df$ReflexR <- sapply(proband.full.df$ReflexR, paper.reflex.to.levels)
proband.full.df$ReflexL <- sapply(proband.full.df$ReflexL, paper.reflex.to.levels)

proband.full.df$MonofilamentR <- sapply(proband.full.df$MonofilamentR, paper.monofilament.to.levels)
proband.full.df$MonofilamentL <- sapply(proband.full.df$MonofilamentL, paper.monofilament.to.levels)

proband.full.df$VibrationRScore <- factor(proband.full.df$VibrationRScore, levels = c(0, 1), labels = c("Present", "Reduced/Absent"))
proband.full.df$VibrationLScore <- factor(proband.full.df$VibrationLScore, levels = c(0, 1), labels = c("Present", "Reduced/Absent"))
proband.full.df$VibrationLScoreNoAdj <- factor(proband.full.df$VibrationLScoreNoAdj, levels = c(0, 1), labels = c("Present", "Reduced/Absent"))
proband.full.df$VibrationRScoreNoAdj <- factor(proband.full.df$VibrationRScoreNoAdj, levels = c(0, 1), labels = c("Present", "Reduced/Absent"))

delete.columns <- c("Group", "YearOfBirth", "TestYear", "DiabetesSince")
probands.summary <- build.df.summary(target.df = proband.full.df, skip.column.names = c("ProbandId", "Surname", "Givenname", "GroupNew", "Group"), effective.digits = 1, debug.activated = T, only.mean.sd = F)

write.xlsx(proband.full.df,
           file = paste(output.folder, "All probands 2023-05-08.xlsx", sep = "/"),
           sheetName = "All", append = F, row.names = FALSE)
write.xlsx(probands.summary,
           file = paste(output.folder, "All probands 2023-05-08.xlsx", sep = "/"),
           sheetName = "Summary", append = T, row.names = FALSE)



