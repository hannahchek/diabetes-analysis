library(dplyr)
library(ggplot2)
library(tidyr)
library(mice)
library(corrplot)
library(knitr)
library(kableExtra)
library(openxlsx)

# Load Data
data <- combined_data

# Merge Additional Variables
data <- data %>%
  left_join(newvars, by = 'RodamID') %>%
  left_join(Additional, by = 'RodamID')

# Convert haven_labelled Variables to Factors
data <- data %>%
  mutate(across(where(is.labelled), ~ set_label(as_factor(.x), label = get_label(.x)))) %>%
  mutate(across(where(~ inherits(.x, "haven_labelled") & is.numeric(.x)), 
                ~ set_label(as.numeric(.x), label = get_label(.x))))

# Ensure Numeric Variables are Properly Converted
numeric_vars <- c("R1_Chol", "R1_TG", "R1_HDLChol", "R1_LDLChol",
                  "R1_Glucose", "R1_Insulin", "R1_HbA1c",
                  "R2_Chol", "R2_TG", "R2_HDLChol", "R2_LDLChol",
                  "R2_Glucose", "R2_HbA1c", 
                  "R1_AlbUrine", "R2_AlbUrine", "R1_UricAcid", "R1_Calcium", 
                  "R1_Protein", "R1_Na", "R1_Ferritin",
                  "R1_GGT", "R1_ALAT", "R1_ASAT", "R1_Hb")

data <- data %>%
  mutate(across(all_of(numeric_vars), ~ as.numeric(as.character(.))))

# Remove Unnecessary Variables
data <- data %>%
  select(-starts_with("R1_Factor"))

# Visualize Missing Data
vis_miss(data)

# Generate Missing Data Summary
missing_summary <- data.frame(
  Variable = names(data),
  Missing = colSums(is.na(data)),
  Percentage = round((colSums(is.na(data)) / nrow(data) * 100), 2)
) %>%
  arrange(desc(Percentage))

# Display Missing Data Summary
print(missing_summary)
kable(missing_summary, format = "html", caption = "Summary of Missing Values") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed", "responsive"))

# Remove Participants from a Specific Site
data <- data[data$R1_Site != 'Dutch', ]

# Retain Variables with ≤ 50% Missing Data
data_reduced <- data[, colSums(is.na(data)) / nrow(data) <= 0.50]

# Identify Removed Variables
removed_vars <- names(colSums(is.na(data)) / nrow(data) > 0.50)
print(removed_vars)

# Create Predictor Matrix for Imputation
pred_reduced <- make.predictorMatrix(data_reduced)

# Remove Empty Factor Levels
data_reduced <- data_reduced %>%
  mutate(across(where(is.factor), droplevels))

# Handle Missing Categorical Values
data_reduced <- data_reduced %>%
  mutate(across(where(is.factor), ~ na_if(as.character(.), "Missing"))) %>%
  mutate(across(where(is.character), as.factor))

# Adjust Specific Variables
data_reduced <- data_reduced %>%
  mutate(R1_Active = ifelse(R1_Active == 0 | R1_Active == "0", NA, R1_Active)) %>%
  mutate(R1_Active = as.factor(R1_Active))

# Exclude High-Cardinality Variables from Predicting Others
high_cardinality_vars <- c("R2_Chol", "R2_TG", "R2_HDLChol", "R2_LDLChol", "R2_Glucose", "R2_HbA1c")
pred_reduced[, high_cardinality_vars] <- 0

# Convert Numeric-like Factors to Numeric
data_reduced <- data_reduced %>%
  mutate(across(all_of(high_cardinality_vars), ~ as.numeric(as.character(.))))

# Define Categorical and Numerical Variables
cat_vars_reduced <- names(data_reduced)[sapply(data_reduced, is.factor) | sapply(data_reduced, is.character)]
num_vars_reduced <- names(data_reduced)[sapply(data_reduced, is.numeric)]

# Define Binary and Multiclass Categorical Variables
binary_vars_reduced <- cat_vars_reduced[sapply(data_reduced[cat_vars_reduced], function(x) nlevels(as.factor(x)) == 2)]
multiclass_vars_reduced <- setdiff(cat_vars_reduced, binary_vars_reduced)

# Define Imputation Methods
method_reduced <- make.method(data_reduced)
method_reduced[binary_vars_reduced] <- 'logreg'
method_reduced[multiclass_vars_reduced] <- "polyreg"
method_reduced["RodamID"] <- "" # Prevent imputation
pred_reduced[, "RodamID"] <- 0  # Prevent as predictor

# Visualize Predictor Matrix
corrplot(pred_reduced, is.corr = FALSE, method = "color", tl.cex = 0.6, cl.cex = 0.6, tl.col = "black")

# Perform Multiple Imputation
imputed_data_reduced <- mice(data_reduced, method = method_reduced, predictorMatrix = pred_reduced, 
                             m = 2, maxit = 5, seed = 123)

# Extract Completed Dataset
imputed_data <- complete(imputed_data_reduced, 1)

# Save Imputed Data
save(imputed_data, file = 'imputed_data.RData')

# Confirm That All Missing Values Are Imputed
sum(is.na(imputed_data))

# Split Data into Baseline (R1) and Outcome (R2)
baseline_data <- cbind(imputed_data[, grepl("^R1", names(imputed_data))])
outcome_data <- cbind(imputed_data[, grepl("^R2", names(imputed_data))])

# Identify and Convert Categorical Variables
categorical_vars <- names(baseline_data)[sapply(baseline_data, is.factor) | sapply(baseline_data, is.character)]
baseline_data <- baseline_data %>%
  mutate(across(all_of(categorical_vars), as.factor))

# Convert Specific Variables to Numeric
baseline_data <- baseline_data %>%
  mutate(R1_CKDEPI_eGFR_2021 = as.numeric(as.character(R1_CKDEPI_eGFR_2021)),
         R1_ABI = as.numeric(as.character(R1_ABI)))

# Transform R1_alcohol_units_day to Binary Factor
baseline_data <- baseline_data %>%
  mutate(R1_alcohol_units_day = ifelse(R1_alcohol_units_day != 0, "Yes", "No"),
         R1_alcohol_units_day = as.factor(R1_alcohol_units_day))

# Create BMI Groups
baseline_data <- baseline_data %>%
  mutate(BMI_Group = ifelse(R1_BMI <= 25, "Lean", "Obese"))

# Create Summary Table 1
table1 <- CreateTableOne(vars = setdiff(names(baseline_data), c("RodamID", "BMI_Group")),
                         strata = "BMI_Group",
                         data = baseline_data,
                         factorVars = categorical_vars,
                         addOverall = TRUE)

# Print and Save Table 1
table1_df <- as.data.frame(print(table1, printToggle = FALSE))
write.xlsx(table1_df, file = "Table1_Scientific_Report.xlsx", row.names = FALSE)

# Save Processed Data
save(baseline_data, file = 'baseline_data.RData')
save(outcome_data, file = 'outcome_data.RData')

# Filter Outcome Data to Include Only Participants from Baseline Data
outcome_data_filtered <- outcome_data %>%
  filter(RodamID %in% baseline_data$RodamID)

# Merge BMI_Group from Baseline Data into Outcome Data
outcome_data_filtered <- outcome_data_filtered %>%
  left_join(baseline_data %>% select(RodamID, BMI_Group), by = "RodamID")

# Define Variables to Remove from Outcome Data
vars_to_remove_outcome <- c("RodamID", "R2_Site", "R2_Sex", "R2_Age", "R2_BMI", 
                            "R2_waist_calc", "R2_DiabetesMeds", "R2_Antilipidemics", 
                            "R2_DiabTreat", "R2_DiabTabl", "R2_DiabDiet", "R2_DiabIns", 
                            "R2_EyeDisease", "R2_DiabFam", "R2_Smoking", "R2_Active", 
                            "R2_DM_Dichot")

outcome_data_filtered <- outcome_data_filtered %>%
  select(-all_of(vars_to_remove_outcome))

# Identify Categorical and Numerical Variables in Outcome Data
outcome_numerical_vars <- names(outcome_data_filtered)[sapply(outcome_data_filtered, is.numeric)]
outcome_categorical_vars <- setdiff(names(outcome_data_filtered), outcome_numerical_vars)

# Define Composite Variable for Complications
outcome_data_filtered <- outcome_data_filtered %>%
  mutate(Has_Complications = ifelse(
    R2_Stroke == "Yes" |
      R2_HTN_MedBP == "Yes" |
      R2_CI_Rose == "Yes" |
      R2_possINF_Rose == "Yes" |
      R2_AP_Rose == "Yes" |
      R2_CVD_Rose == "Yes" |
      R2_CVA_Self == "Yes" |
      R2_Microalbuminuria == "Yes",
    "Yes", "No")) %>%
  mutate(Has_Complications = as.factor(Has_Complications))

# Define Variables for Outcome Table
vars_for_outcome <- c(outcome_numerical_vars, outcome_categorical_vars)

# Create Table for Outcome Data
outcome_table <- CreateTableOne(vars = vars_for_outcome,
                                strata = "BMI_Group",
                                data = outcome_data_filtered,
                                factorVars = outcome_categorical_vars,
                                addOverall = TRUE)

# Print and Save Outcome Table
outcome_table_df <- as.data.frame(print(outcome_table, printToggle = FALSE))
write.xlsx(outcome_table_df, file = "Outcome_Table.xlsx", row.names = FALSE)

# Save Processed Data
save(outcome_data_filtered, file = 'outcome_data.RData')

# Create dataset of outcome_data and baseline_data

merged_data <- innerjoin(outcome_data, baseline_data, by = 'RodamID')

# Function to Create Paired Boxplots
paired_boxplot <- function(data, baseline_var, followup_var, var_name) {
  plot_data <- data %>%
    select(RodamID, BMI_Group, all_of(baseline_var), all_of(followup_var)) %>%
    pivot_longer(cols = c(all_of(baseline_var), all_of(followup_var)),
                 names_to = "Timepoint",
                 values_to = var_name) %>%
    mutate(Timepoint = ifelse(Timepoint == baseline_var, "Baseline", "Follow-Up"))
  
  ggplot(plot_data, aes(x = BMI_Group, y = .data[[var_name]], fill = Timepoint)) +
    geom_boxplot(position = position_dodge(0.8), width = 0.6) +
    labs(title = paste(var_name, "Levels by BMI Group (Baseline vs. Follow-Up)"),
         x = "BMI Group",
         y = paste(var_name, "(mmol/L)")) +
    scale_fill_manual(values = c("Baseline" = "#1f77b4", "Follow-Up" = "#ff7f0e")) +
    theme_minimal()
}

# Generate Boxplots
paired_boxplot(merged_data, 'R1_HDLChol', 'R2_HDLChol', 'HDL Cholesterol')
paired_boxplot(merged_data, 'R1_Glucose', 'R2_Glucose', 'Glucose')
paired_boxplot(merged_data, 'R1_Chol', 'R2_Chol', 'Cholesterol')
paired_boxplot(merged_data, 'R1_TG', 'R2_TG', 'Triglycerides')
paired_boxplot(merged_data, 'R1_HbA1c', 'R2_HbA1c', 'HbA1c')
paired_boxplot(merged_data, 'R1_AlbUrine', 'R2_AlbUrine', 'Albumin in Urine')
paired_boxplot(merged_data, 'R1_BPsys_mean', 'R2_BPsys_mean', 'Systolic BP')
paired_boxplot(merged_data, 'R1_BPdia_mean', 'R2_BPdia_mean', 'Diastolic BP')

# Define Outcome Variables for Analysis
outcome_vars <- c("R2_HDLChol", "R2_TG", "R2_Chol", 
                  "R2_Glucose", "R2_HbA1c", "R2_CKDEPI_eGFR_2021", 
                  "R2_BPsys_mean", "R2_BPdia_mean", "R2_AlbUrine")

# Define Model Predictors
model_predictors <- list(
  "Model 1" = "BMI_Group",
  "Model 2" = "BMI_Group + R1_Age + R1_Sex",
  "Model 3" = "BMI_Group + R1_Age + R1_Sex + R1_Leptin + R1_HTN_MedBP + R1_HOMA_IR",
  "Model 4" = "BMI_Group + R1_Age + R1_Sex + R1_Leptin + R1_HTN_MedBP + R1_HOMA_IR + R1_Site + R1_Smoking"
)

# Function to Fit Regression Models
fit_models <- function(outcome) {
  results <- list()
  
  for (model_name in names(model_predictors)) {
    formula <- as.formula(paste(outcome, "~", model_predictors[[model_name]]))
    model <- glm(formula, data = merged_data, family = gaussian())
    
    # Extract Coefficients and Confidence Intervals
    tidy_model <- broom::tidy(model, conf.int = TRUE) %>%
      filter(term == "BMI_GroupObese") %>%
      mutate(Significance = ifelse(p.value < 0.05, "*", "")) %>%
      mutate(Result = paste0(round(estimate, 2), " (", round(conf.low, 2), ", ", round(conf.high, 2), ")", Significance)) %>%
      select(Result) %>%
      rename(!!model_name := Result)
    
    results[[model_name]] <- tidy_model
  }
  
  # Combine Results for the Outcome Variable
  combined_results <- bind_cols(results) %>%
    mutate(Outcome = outcome) %>%
    relocate(Outcome)
  
  return(combined_results)
}

# Apply Function to All Outcome Variables
all_model_results <- bind_rows(lapply(outcome_vars, fit_models))

# Save Regression Results to Excel
write.xlsx(all_model_results, file = "Regression_Results.xlsx", row.names = FALSE)

