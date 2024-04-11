library("readxl")
library("writexl")
library("stringr")
library("lubridate")

filename <- "~/Documents/Oxford/AndyApps/outcome new1.0justin.xlsx"
stemi_filename <- "~/Documents/Oxford/AndyApps/STEMIs in the time frame JL.xlsx"

# filename <- "~/Documents/Oxford/AndyApps/outcome new2.0justin.xlsx"
# stemi_filename <- "~/Documents/Oxford/AndyApps/STEMI Procedures - second batch 16 to 18.xlsx"

filestem <- tools::file_path_sans_ext(basename(filename))
file_ext <- tools::file_ext(basename(filename))
outfilename <- paste(dirname(filename), "/", filestem, "_output.", file_ext, sep = "")

# Read Excel sheets into data frames
Referral0 <- read_excel(filename, sheet = 1, skip = 1)
OtherReferrals <- read_excel(filename, sheet = 2, skip = 0)
Procedures <- read_excel(filename, sheet = 3, skip = 0)
MDROutcome <- read_excel(filename, sheet = 4, skip = 0)
BypassProc <- read_excel(filename, sheet = 5, skip = 0)
PCI <- read_excel(stemi_filename, skip = 4)

# Initialize output variables
out_chi <- c()
out_type <- c()
out_group <- c()
out_class <- c()
out_dates <- c()
out_ids <- c()
out_group <- c()

# Referrals
out_DOB <- c()
out_Gender <- c()
out_Age_op_ref <- c()
out_refType <- c()
out_Protocol <- c()
out_Referrer <- c()
out_RefHospital <- c()
out_GraceScore <- c()
out_TnlResult <- c()
out_TnTResult <- c()
out_TroponinResult <- c()
out_ECGResult <- c()
out_HBResult <- c()
out_CreatinineResult <- c()
out_Medications <- c()
out_OtherDrugs <- c()
out_Risks <- c()
out_MRCScore <- c()
out_DiabetesTreatment <- c()
out_GIBleedDetail <- c()
out_DrugAllergyDesc <- c()
out_OtherAllergyDesc <- c()
out_ContrastAllergyDesc <- c()
out_DialysisDetails <- c()
out_PatientOver175kg <- c()
out_DoD <- c()

# Procedures
out_AdmissionPathway <- c()
out_SmokingStatus <- c()
out_DiabeticStatus <- c()
out_FamilyHistoryOfCHD <- c()
out_PrevSurgery <- c()
out_PreviousPCI <- c()
out_Complications <- c()
out_DeviceCategory <- c()
out_OtherMedicalHist <- c()
out_CorStent_DEStent <- c()
out_OCT_IVUS <- c()
out_RotAtherectomy_CuttingBalloon <- c()
out_Vessels_territories_attempted <- c()
out_LatestProcedureType <- c()
out_CoronaryArteries <- c()

for (id in Referral0$chi) {
  # Pull out entries from all sheets for each CHI Number
  this_Referral0 <- Referral0[Referral0$chi == id,]
  this_OtherReferrals <- OtherReferrals[OtherReferrals$chi == id,]
  this_Procedures <- Procedures[Procedures$CHINumber == id,]
  this_MDROutcome <- MDROutcome[MDROutcome$ChiNumber == id,]
  this_BypassProc <- BypassProc[BypassProc$ChiNumber == id,]
  # print(this_MDROutcome)
  this_PCI <- PCI[PCI$CHI == id,]

  # Remove rows with blank (NA) chi
  this_Referral0 <- this_Referral0[!is.na(this_Referral0$chi),]
  this_OtherReferrals <- this_OtherReferrals[!is.na(this_OtherReferrals$chi),]
  this_Procedures <- this_Procedures[!is.na(this_Procedures$CHINumber),]
  this_MDROutcome <- this_MDROutcome[!is.na(this_MDROutcome$ChiNumber),]
  this_BypassProc <- this_BypassProc[!is.na(this_BypassProc$ChiNumber),]
  this_PCI <- this_PCI[!is.na(this_PCI$CHI),]
  
  # Extract dates for each CHI Number
  this_referral0_date <- this_Referral0$ReferralDate
  this_referral0_id <- this_Referral0$referralId
  this_DoD <- this_Referral0$DoD

  this_otherreferrals_dates <- this_OtherReferrals$ReferralDate
  this_otherreferrals_id <- this_OtherReferrals$referralId
  this_otherreferrals_refType <- this_OtherReferrals$refType

  this_procedure_dates <- c(this_Procedures$ProcedureDate, this_BypassProc$ProcedureDate)
  this_pci_dates <- this_PCI$'Procedure Date'

  # Calculate number of days from Referral 0
  days_referral_from_referral0 <- difftime(this_otherreferrals_dates, this_referral0_date, units = "days")
  days_procedure_from_referral0 <- difftime(this_procedure_dates, this_referral0_date, units = "days")
  days_procedure_from_pci <- difftime(this_procedure_dates, this_pci_dates, units = "days")
  days_pci_from_referral0 <- difftime(this_pci_dates, this_referral0_date, units = "days")
  days_procedure0_from_referral0 <- c()
  days_referral_from_procedure0 <- c()

  num_referrals <- 1 + length(days_referral_from_referral0)
  num_procedures <- length(this_Procedures$ProcedureDate)
  num_bypasses <- length(this_BypassProc$ProcedureDate)

  len_referral0 <- length(this_Referral0)
  len_otherref <- length(this_OtherReferrals)
  len_procedures <- length(this_Procedures)

  # Initialize Group for each CHI
  classification_referrals <- c()
  classification_procedures <- c()
  classification_group <- c()
  # classification_group <- c(rep("", num_referrals + num_procedures + num_bypasses))

  # Classify procedures based on the following rules:
  # Procedure -1 < referral 0
  # Procedure 0 >= Referral 0
  # Procedure 1>=procedure 0
  # Procedure 2 if PCI
  ### How to classify a procedure on the same day as Referral 0?

  if (length(days_procedure_from_referral0) > 0) {
    # Initialize vector of blank IDs
    procedure_ids <- rep("", length(days_procedure_from_referral0))
    # Start with a vector of 1's
    classification_procedures <- rep(1, length(days_procedure_from_referral0))
    # Classify procedures before Referral 0 (negative days difference) as -1
    classification_procedures[days_procedure_from_referral0 < 0] <- -1
    # Flag procedures on the same day as Referral 0 (zero days difference) as -0.5
    # Changed Procedure 0 definition to be >= Referral 0
    # classification_procedures[days_procedure_from_referral0 == 0] <- -0.5
    # Identify Procedure 0 as closest procedure on or after Referral 0
    pos_days <- days_procedure_from_referral0[days_procedure_from_referral0 >= 0]
    neg_days <- days_procedure_from_referral0[days_procedure_from_referral0 < 0]
    
    # Group one – patients who had a procedure following referral (>0 days after)
    # 1.5: Catch-all group
    classification_group <- "1.5"
    
    if (length(pos_days) > 0) {
      classification_procedures[days_procedure_from_referral0 == min(pos_days)] <- 0
      this_procedure0_date <- unique(this_procedure_dates[days_procedure_from_referral0 == min(pos_days)])
      days_procedure0_from_referral0 <- difftime(this_procedure0_date, this_referral0_date)
      days_referral_from_procedure0 <- difftime(this_otherreferrals_dates, this_procedure0_date, units = "days")
    }

    # 1.1: Those in which the procedure was done as an emergency (primary PCI) as a Procedure 0
    if (length(days_procedure_from_pci) > 0) {
      # Check if any of the PCI's were Procedure 0's
      if (any(days_pci_from_referral0 == min(pos_days))) {
        procedure_ids[days_procedure_from_pci == 0] <- "PCI"
        classification_group <- "1.1"
        # classification_group[num_referrals + which(days_referral_from_referral0 == days_procedure_from_pci)] <- "1.1"
        print("PCI")
      }
    }
    else if (length(days_referral_from_referral0) > 0) {
      # A procedure following a referral cancels the referral (resets the clock)
      days_unpaired_referrals_from_ref0 <- c()
      refs_within_90days_ref0 <- c()
      # Procedure-Referral pair annihilation
      days_unpaired_referrals_from_ref0 <- days_referral_from_referral0
      if ( length(neg_days) > 0 ) {
        refs_to_cancel <- c()
        # print(cat("neg_days: ", neg_days))
        # print(cat("days_ref_from_ref0: ", days_referral_from_referral0))
        for (this_proc in neg_days) {
          # this_proc is days from Referral 0 (negative)
          # Find referral immediately before procedure
          days_proc_from_referral <- this_proc - days_referral_from_referral0
          # Referral before this_proc is the smallest positive number in days_proc_from_referral
          date_ref_before_proc <- days_referral_from_referral0[days_proc_from_referral == min(days_proc_from_referral[days_proc_from_referral >= 0])]
          # print(cat("date_ref_before_proc: ", date_ref_before_proc))
          refs_to_cancel <- c(refs_to_cancel, date_ref_before_proc)
        }
        days_unpaired_referrals_from_ref0 <- setdiff(days_referral_from_referral0, refs_to_cancel)
      }
      
      # Use unpaired referrals list for calculations for grouping
      # Is there a second referral between referral 0 and procedure 0 (> ref0)?
      preceding_second_referral <- days_unpaired_referrals_from_ref0 > 0 & days_unpaired_referrals_from_ref0 < days_procedure0_from_referral0
      # Is there a referral within 90 days prior to referral 0, but with no procedure -1 between referral 0 / -1?
      refs_within_90days_ref0 <- days_unpaired_referrals_from_ref0 >= -90 & days_unpaired_referrals_from_ref0 < 0
      # print(cat("unpaired_refs: ", days_unpaired_referrals_from_ref0))

      
      # 1.2: Those In which there was a further ‘upgrade’ referral < the procedure date,  but > the initial referral date
      # upgrade_referral <- days_referral_from_procedure0 < 0 & days_referral_from_referral0 > 0
      upgrade_referral_days_from_ref0 <- days_referral_from_referral0[days_referral_from_referral0 > 0]
      upgrade_referral_days_from_ref0 <- upgrade_referral_days_from_ref0[upgrade_referral_days_from_ref0 < days_procedure0_from_referral0]
      if ( length(upgrade_referral_days_from_ref0) > 0 ) {
        classification_group <- "1.2"
        # Find refType of upgrade referral closest to Ref0
        this_upgrade_days <- min(upgrade_referral_days_from_ref0)
        this_index <- which(days_referral_from_referral0 == this_upgrade_days)
        this_upgrade_type <- this_otherreferrals_refType[this_index]
        if ( length(this_upgrade_type) > 0 ) {
          for (row in this_upgrade_type) {
            if ( str_detect(row, regex("inpatient", ignore_case = TRUE)) ) {
              classification_group <- "1.2.1"
            } else {
              classification_group <- "1.2.2"
            }
          }
        }
      }
      else if ( any(refs_within_90days_ref0) ) {
        # 1.3: Those in which a preceding second referral exists within 90 days from referral 0, but only if no procedure exists on or after this preceding referral and before the initial referral.
        # preceding_second_referral <- days_referral_from_referral0 >= -90 & days_referral_from_referral0 < 0
        # earliest_preceding_referral <- min(days_referral_from_referral0[which(preceding_second_referral)])
        # Procedures on or after preceding referral but before Referral 0
        # procedures_after_preceding_referral <- days_procedure_from_referral0 >= earliest_preceding_referral & days_procedure_from_referral0 < 0
        
        # Procedure-Referral pair annihilation should have removed all referrals preceding a procedure
        # So just see if any of these unpaired referrals are within 90 days from Referral 0

        # No procedures on or after preceding referral but before Referral 0
        classification_group <- "1.3"
        # Find refType of preceding referral closest to Ref0
        this_ref_days <- max(days_unpaired_referrals_from_ref0[refs_within_90days_ref0])
        this_index <- which(days_referral_from_referral0 == this_ref_days)
        this_ref_info <- this_OtherReferrals$refType[this_index]
        # which(days_referral_from_referral0 == this_ref_days)
        if ( str_detect(this_ref_info, regex("inpatient", ignore_case = TRUE)) ) {
          classification_group <- "1.3.1"
        } else {
          classification_group <- "1.3.2"
        }
      }
      else if ( length(days_unpaired_referrals_from_ref0[days_unpaired_referrals_from_ref0 < 0]) > 0 ) {
        if ( all(days_unpaired_referrals_from_ref0[days_unpaired_referrals_from_ref0 < 0] < -90) ) {
          # 1.4: Those in which no other referrals are listed within 90 days of referral 0
          classification_group <- "1.4"
        }
      }
    }

    if ( max(days_procedure_from_referral0) <= 0 ) {
      # Group two: patients who did not have a subsequent procedure
      classification_group <- "2"
      
      # 2.2: Procedure within the preceding 90days of referral
      procedures_90days_before_ref0 <- days_procedure_from_referral0 >= -90 & days_procedure_from_referral0 < 0
      if ( any(procedures_90days_before_ref0) ) { classification_group <- "2.2" }
      
      # 2.3: No procedure prior or >1 year before
      procedures_morethan90days_before_ref0 <- days_procedure_from_referral0 < -90
      if ( length(days_procedure_from_referral0) == 0 | all(procedures_morethan90days_before_ref0) ) {
        classification_group <- "2.3"
      }
      
      # 2.1: Dead within 90 days of referral
      if ( is.na(this_DoD) == FALSE ) {
        # Patient died, but is it within 90 days of referral?
        days_dead_from_ref0 <- difftime(this_DoD, this_referral0_date)
        print(days_dead_from_ref0)
        if ( abs(days_dead_from_ref0) <= 90 ) { classification_group <- "2.1" }
      }
    }
  }
  else {
    classification_procedures <- c()
    procedure_ids <- c()
    # 2.3: Group two: patients who did not have a subsequent procedure, No preceding procedure
    classification_group <- "2.3"
  }
  
  # 3.0: Referral 0 = Procedure 0
  if ( any(days_procedure_from_referral0 == 0) ) { classification_group <- "3" }

  # print(this_referral0_date)
  # print(this_procedure_dates)
  # print(classification_procedures)

  # Classify referrals based on the following rules:
  # Referral -1 < referral 0
  # Referral 1 > referral 0
  # Referral 1 < procedure 0
  # Referral 2 >= procedure 0 (operator driven subsequent referral)

  if (length(days_referral_from_referral0) > 0) {
    # Start with a vector of -2's
    classification_referrals <- rep(-2, length(days_referral_from_referral0))
    # Referral -1 < referral 0
    classification_referrals[days_referral_from_referral0 < 0] <- -1
    # Referral 0 == referral 0
    # classification_referrals[days_referral_from_referral0 == 0] <- 0.5
    # Referral 1 >= referral 0
    classification_referrals[days_referral_from_referral0 >= 0] <- 1
    # Referral 1 < procedure 0
    if (length(days_referral_from_procedure0) > 0) {
      classification_referrals[days_referral_from_procedure0 >= 0] <- 2
    }
  }
  else {
    classification_referrals <- c()
  }

  # print(this_referral0_date)
  # print(procedure0_date)
  # print(this_otherreferrals_dates)
  # print(classification_referrals)

  # Group classification

  # Append referrals into output variables
  out_chi <- c(out_chi, rep(id, num_referrals))
  out_type <- c(out_type, rep("Referral", num_referrals))
  out_class <- c(out_class, 0, classification_referrals)
  out_dates <- c(out_dates, this_referral0_date, this_otherreferrals_dates)
  out_ids <- c(out_ids, this_referral0_id, this_otherreferrals_id)
  
  out_DOB <- c(out_DOB, this_Referral0$DOB, this_OtherReferrals$dob)
  out_Gender <- c(out_Gender, this_Referral0$Gender, this_OtherReferrals$Gender)
  out_Age_op_ref <- c(out_Age_op_ref, this_Referral0$`Age (op ref)`, this_OtherReferrals$age)
  out_refType <- c(out_refType, this_Referral0$refType, this_OtherReferrals$refType)
  out_Protocol <- c(out_Protocol, this_Referral0$protocol, this_OtherReferrals$protocol)
  out_Referrer <- c(out_Referrer, this_Referral0$Referrer, this_OtherReferrals$referrer)
  out_RefHospital <- c(out_RefHospital, this_Referral0$`Ref Hospital`, this_OtherReferrals$`referring hospital`)
  # out_Referrer <- c(out_Referrer, this_Referral0$Referrer, this_OtherReferrals$Referrer)
  # out_RefHospital <- c(out_RefHospital, this_Referral0$`Ref Hospital`, this_OtherReferrals$`Ref Hospital`)
  out_GraceScore <- c(out_GraceScore, this_Referral0$GraceScore, this_OtherReferrals$GraceScore)
  out_TnlResult <- c(out_TnlResult, this_Referral0$TnlResult, this_OtherReferrals$TnlResult)
  out_TnTResult <- c(out_TnTResult, this_Referral0$TnTResult, this_OtherReferrals$TnTResult)
  out_TroponinResult <- c(out_TroponinResult, this_Referral0$TroponinResult, this_OtherReferrals$TroponinResult)
  out_ECGResult <- c(out_ECGResult, this_Referral0$ECGResult, this_OtherReferrals$ECGResult)
  out_HBResult <- c(out_HBResult, this_Referral0$HBResult, this_OtherReferrals$HBResult)
  out_CreatinineResult <- c(out_CreatinineResult, this_Referral0$CreatinineResult, this_OtherReferrals$CreatinineResult)
  out_Medications <- c(out_Medications, this_Referral0$Medications, this_OtherReferrals$Medications)
  out_OtherDrugs <- c(out_OtherDrugs, this_Referral0$OtherDrugs, this_OtherReferrals$OtherDrugs)
  out_Risks <- c(out_Risks, this_Referral0$Risks, this_OtherReferrals$Risks)
  out_MRCScore <- c(out_MRCScore, this_Referral0$MRCScore, this_OtherReferrals$MRCScore)
  out_DiabetesTreatment <- c(out_DiabetesTreatment, this_Referral0$DiabetesTreatment, this_OtherReferrals$DiabetesTreatment)
  out_GIBleedDetail <- c(out_GIBleedDetail, this_Referral0$GIBleedDetail, this_OtherReferrals$GIBleedDetail)
  out_DrugAllergyDesc <- c(out_DrugAllergyDesc, this_Referral0$DrugAllergyDesc, this_OtherReferrals$DrugAllergyDesc)
  out_OtherAllergyDesc <- c(out_OtherAllergyDesc, this_Referral0$OtherAllergyDesc, this_OtherReferrals$OtherAllergyDesc)
  out_ContrastAllergyDesc <- c(out_ContrastAllergyDesc, this_Referral0$ContrastAllergyDesc, this_OtherReferrals$ContrastAllergyDesc)
  out_DialysisDetails <- c(out_DialysisDetails, this_Referral0$DialysisDetails, this_OtherReferrals$DialysisDetails)
  out_PatientOver175kg <- c(out_PatientOver175kg, this_Referral0$PatientOver175kg, this_OtherReferrals$PatientOver175kg)
  out_DoD <- c(out_DoD, as.character(this_Referral0$DoD), as.character(this_OtherReferrals$dod))
  # print(this_Referral0$chi)
  # print(this_Referral0$ReferralDate)
  # print(this_Referral0$DoD)
  # print(this_OtherReferrals$dod)
  
  out_AdmissionPathway <- c(out_AdmissionPathway, rep("", num_referrals))
  out_SmokingStatus <- c(out_SmokingStatus, rep("", num_referrals))
  out_DiabeticStatus <- c(out_DiabeticStatus, rep("", num_referrals))
  out_FamilyHistoryOfCHD <- c(out_FamilyHistoryOfCHD, rep("", num_referrals))
  out_PrevSurgery <- c(out_PrevSurgery, rep("", num_referrals))
  out_PreviousPCI <- c(out_PreviousPCI, rep("", num_referrals))
  out_Complications <- c(out_Complications, rep("", num_referrals))
  out_DeviceCategory <- c(out_DeviceCategory, rep("", num_referrals))
  out_OtherMedicalHist <- c(out_OtherMedicalHist, rep("", num_referrals))
  out_CorStent_DEStent <- c(out_CorStent_DEStent, rep("", num_referrals))
  out_OCT_IVUS <- c(out_OCT_IVUS, rep("", num_referrals))
  out_RotAtherectomy_CuttingBalloon <- c(out_RotAtherectomy_CuttingBalloon, rep("", num_referrals))
  out_Vessels_territories_attempted <- c(out_Vessels_territories_attempted, rep("", num_referrals))
  out_LatestProcedureType <- c(out_LatestProcedureType, rep("", num_referrals))
  out_CoronaryArteries <- c(out_CoronaryArteries, rep("", num_referrals))
  
  # if ( length(unique(c(num_referrals, length(classification_referrals)+1, length(this_otherreferrals_dates)+1, length(this_otherreferrals_id)+1))) > 1 ) {
  #   print(c("REF: ", id))
  #   print(c(num_referrals, length(classification_referrals)+1, length(this_otherreferrals_dates)+1, length(this_otherreferrals_id)+1))
  #   print(classification_referrals)
  #   print(days_referral_from_procedure0)
  #   #      days_referral_from_procedure0 <- difftime(this_otherreferrals_dates, this_procedure0_date, units = "days")
  #   print(this_otherreferrals_dates)
  #   print(this_procedure0_date)
  #   print(difftime(this_otherreferrals_dates, this_procedure0_date, units = "days"))
  #   print("FER")
  # }
  
  # out_info <- c(out_info, this_Referral0, this_OtherReferrals)
  # out_info <- c(out_info, this_Referral0, rep(c(), len_otherref + len_procedures))
  # for (this_row in this_OtherReferrals) {
  #   out_info <- c(out_info, rep(c(), len_referral0), this_row, rep(c(), len_procedures))
  # }

  # Append procedures into output variables
  out_chi <- c(out_chi, rep(id, num_procedures + num_bypasses))
  out_type <- c(out_type, rep("Procedure", num_procedures + num_bypasses))
  out_class <- c(out_class, classification_procedures)
  out_dates <- c(out_dates, this_procedure_dates)
  out_ids <- c(out_ids, procedure_ids)
  out_group <- c(out_group, rep(classification_group, num_referrals + num_procedures + num_bypasses))
  
  # Blank for procedures
  out_DOB <- c(out_DOB, rep("", num_procedures + num_bypasses))
  out_Gender <- c(out_Gender, rep("", num_procedures + num_bypasses))
  out_Age_op_ref <- c(out_Age_op_ref, rep("", num_procedures + num_bypasses))
  out_refType <- c(out_refType, rep("", num_procedures + num_bypasses))
  out_Protocol <- c(out_Protocol, rep("", num_procedures + num_bypasses))
  out_Referrer <- c(out_Referrer, rep("", num_procedures + num_bypasses))
  out_RefHospital <- c(out_RefHospital, rep("", num_procedures + num_bypasses))
  out_GraceScore <- c(out_GraceScore, rep("", num_procedures + num_bypasses))
  out_TnlResult <- c(out_TnlResult, rep("", num_procedures + num_bypasses))
  out_TnTResult <- c(out_TnTResult, rep("", num_procedures + num_bypasses))
  out_TroponinResult <- c(out_TroponinResult, rep("", num_procedures + num_bypasses))
  out_ECGResult <- c(out_ECGResult, rep("", num_procedures + num_bypasses))
  out_HBResult <- c(out_HBResult, rep("", num_procedures + num_bypasses))
  out_CreatinineResult <- c(out_CreatinineResult, rep("", num_procedures + num_bypasses))
  out_Medications <- c(out_Medications, rep("", num_procedures + num_bypasses))
  out_OtherDrugs <- c(out_OtherDrugs, rep("", num_procedures + num_bypasses))
  out_Risks <- c(out_Risks, rep("", num_procedures + num_bypasses))
  out_MRCScore <- c(out_MRCScore, rep("", num_procedures + num_bypasses))
  out_DiabetesTreatment <- c(out_DiabetesTreatment, rep("", num_procedures + num_bypasses))
  out_GIBleedDetail <- c(out_GIBleedDetail, rep("", num_procedures + num_bypasses))
  out_DrugAllergyDesc <- c(out_DrugAllergyDesc, rep("", num_procedures + num_bypasses))
  out_OtherAllergyDesc <- c(out_OtherAllergyDesc, rep("", num_procedures + num_bypasses))
  out_ContrastAllergyDesc <- c(out_ContrastAllergyDesc, rep("", num_procedures + num_bypasses))
  out_DialysisDetails <- c(out_DialysisDetails, rep("", num_procedures + num_bypasses))
  out_PatientOver175kg <- c(out_PatientOver175kg, rep("", num_procedures + num_bypasses))
  out_DoD <- c(out_DoD, rep("", num_procedures + num_bypasses))
  
  out_AdmissionPathway <- c(out_AdmissionPathway, this_Procedures$AdmissionPathway, rep("", num_bypasses))
  out_SmokingStatus <- c(out_SmokingStatus, this_Procedures$`Smoking Status`, rep("", num_bypasses))
  out_DiabeticStatus <- c(out_DiabeticStatus, this_Procedures$`Diabetic Status`, rep("", num_bypasses))
  out_FamilyHistoryOfCHD <- c(out_FamilyHistoryOfCHD, this_Procedures$`Family History of CHD`, rep("", num_bypasses))
  out_PrevSurgery <- c(out_PrevSurgery, this_Procedures$PrevSurgery, rep("", num_bypasses))
  out_PreviousPCI <- c(out_PreviousPCI, this_Procedures$`Previous PCI`, rep("", num_bypasses))
  out_Complications <- c(out_Complications, this_Procedures$Complications, rep("", num_bypasses))
  out_DeviceCategory <- c(out_DeviceCategory, this_Procedures$DeviceCategory, rep("", num_bypasses))
  out_OtherMedicalHist <- c(out_OtherMedicalHist, this_Procedures$`Other Medical Hist`, rep("", num_bypasses))
  out_CorStent_DEStent <- c(out_CorStent_DEStent, this_Procedures$`Cor Stent_DE Stent`, rep("", num_bypasses))
  out_OCT_IVUS <- c(out_OCT_IVUS, this_Procedures$OCT_IVUS, rep("", num_bypasses))
  out_RotAtherectomy_CuttingBalloon <- c(out_RotAtherectomy_CuttingBalloon, this_Procedures$RotAtherectomy_CuttingBalloon, rep("", num_bypasses))
  out_Vessels_territories_attempted <- c(out_Vessels_territories_attempted, this_Procedures$`Vessels_territories attempted`, rep("", num_bypasses))
  out_LatestProcedureType <- c(out_LatestProcedureType, rep("", num_procedures), this_BypassProc$LatestProcedureType)
  out_CoronaryArteries <- c(out_CoronaryArteries, rep("", num_procedures), this_BypassProc$`Coronary arteries`)
  
  # out_info <- c(out_info, this_Procedures)
  # for (this_row in this_Procedures) {
  #   out_info <- c(out_info, rep(c(), len_referral0 + len_otherref), this_row)
  # }
  # if ( length(unique(c(num_procedures, length(classification_procedures), length(this_procedure_dates), length(procedure_ids)))) > 1 ) {
  #   print(c("PROC: ", id))
  #   print(c(num_procedures, length(classification_procedures), length(this_procedure_dates), length(procedure_ids)))
  #   print(procedure_ids)
  #   print(days_procedure_from_referral0)
  #   print(this_procedure_dates)
  #   print(this_referral0_date)
  #   print(days_procedure_from_pci)
  #   print("CORP")
  # }

  # Append groups into output variables
  # out_chi <- c(out_chi, id)
  # out_type <- c(out_type, "Group")
  # out_class <- c(out_class, classification_group)
  # out_dates <- c(out_dates, "")
  # out_ids <- c(out_ids, "")
}

# Write results to file
data <- data.frame(CHI = out_chi,
                   Type = out_type,
                   Group = out_group,
                   Classification = out_class,
                   Date = out_dates,
                   ID = out_ids,
                   # Referrals
                   DOB = out_DOB,
                   Gender = out_Gender,
                   Age_op_ref = out_Age_op_ref,
                   RefType = out_refType,
                   Protocol = out_Protocol,
                   Referrer = out_Referrer,
                   RefHospital = out_RefHospital,
                   GraceScore = out_GraceScore,
                   TnlResult = out_TnlResult,
                   TnTResult = out_TnTResult,
                   TroponinResult = out_TroponinResult,
                   ECGResult = out_ECGResult,
                   HBResult = out_HBResult,
                   CreatinineResult = out_CreatinineResult,
                   Medications = out_Medications,
                   OtherDrugs = out_OtherDrugs,
                   Risks = out_Risks,
                   MRCScore = out_MRCScore,
                   DiabetesTreatment = out_DiabetesTreatment,
                   GIBleedDetail = out_GIBleedDetail,
                   DrugAllergyDesc = out_DrugAllergyDesc,
                   OtherAllergyDesc = out_OtherAllergyDesc,
                   ContrastAllergyDesc = out_ContrastAllergyDesc,
                   DialysisDetails = out_DialysisDetails,
                   PatientOver175kg = out_PatientOver175kg,
                   DoD = out_DoD,
                   # Procedures
                   AdmissionPathway = out_AdmissionPathway,
                   SmokingStatus = out_SmokingStatus,
                   DiabeticStatus = out_DiabeticStatus,
                   FamilyHistoryOfCHD = out_FamilyHistoryOfCHD,
                   PrevSurgery = out_PrevSurgery,
                   PreviousPCI = out_PreviousPCI,
                   Complications = out_Complications,
                   DeviceCategory = out_DeviceCategory,
                   OtherMedicalHist = out_OtherMedicalHist,
                   CorStent_DEStent = out_CorStent_DEStent,
                   OCT_IVUS = out_OCT_IVUS,
                   RotAtherectomy_CuttingBalloon = out_RotAtherectomy_CuttingBalloon,
                   Vessels_territories_attempted = out_Vessels_territories_attempted,
                   LatestProcedureType = out_LatestProcedureType,
                   CoronaryArteries = out_CoronaryArteries
                   )

class(data$Date) <- "POSIXct"
# class(data$DoD) <- "POSIXct"

write_xlsx(data, outfilename)