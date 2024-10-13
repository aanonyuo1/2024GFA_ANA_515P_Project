library(readxl) 
library(ggplot2)      
library(dplyr)
library(naniar)
library(writexl)

#Loading data from both sheets
file_path <- "C:\\Users\\augus\\Downloads\\survey.xlsx"
df_sheet1 <- read_excel(file_path, sheet = "Sheet1")
df_sheet2 <- read_excel(file_path, sheet = "Sheet 2")

#Combining the two sheets into one dataset
combined_df <- rbind(df_sheet1, df_sheet2)

#Cleaning the Age column. Age values such as 0, 1, 2, etc., are unrealistic for a tech workplace and are therefore considered invalid
# These invalid age values will be replaced with NA, and then imputed using the median age.
combined_df$Age <- ifelse(combined_df$Age < 18 | combined_df$Age > 100, NA, combined_df$Age)
median_age <- median(combined_df$Age, na.rm = TRUE)  # Calculate median age
combined_df$Age[is.na(combined_df$Age)] <- median_age  # Replace NA ages with median

#Cleaning the Gender column by addressing inconsistent formatting. The values will be standardized to include only "Male", "Female", and "other"
combined_df$Gender <- tolower(combined_df$Gender)
combined_df$Gender <- ifelse(combined_df$Gender %in% c("male", "m"), "male", 
                             ifelse(combined_df$Gender %in% c("female", "f"), "female", "other"))

#Cleaning the self_employed column. Missing values will be replaced with "Unknown" or another placeholder
combined_df$self_employed[is.na(combined_df$self_employed)] <- "Unknown"

#Cleaning the work_interfere column by filling missing values with "unkown"
combined_df$work_interfere[is.na(combined_df$work_interfere)] <- "Unknown"

#Cleaning the no_employees column
# Replaced any erroneous values like negative or extremely large numbers with NA and cap values
combined_df$no_employees <- ifelse(combined_df$no_employees > 1000 | combined_df$no_employees < 0, NA, combined_df$no_employees)

#Handling missing values in categorical columns like 'Country', 'leave', etc.
combined_df$Country[is.na(combined_df$Country)] <- "Unknown"
combined_df$leave[is.na(combined_df$leave)] <- "Unknown"

#Visualizing missing data after cleaning and Checking if there are any remaining missing values
gg_miss_var(combined_df) +
  theme_minimal() +
  labs(title = "Missing Data After Cleaning")

#Checking for anomalies with histograms and boxplots
# Visualizing Age distribution after cleaning
ggplot(combined_df, aes(x = Age)) +
  geom_histogram(binwidth = 5, fill = "lightblue", color = "green") +
  theme_minimal() +
  labs(title = "Cleaned Age Distribution", x = "Age", y = "Frequency")

# Visualizing Gender distribution after cleaning
ggplot(combined_df, aes(x = Gender)) +
  geom_bar(fill = "lightyellow", color = "black") +
  theme_minimal() +
  labs(title = "Cleaned Gender Distribution", x = "Gender", y = "Count")

# Boxplot for Age to check for outliers after cleaning
ggplot(combined_df, aes(x = "", y = Age)) +
  geom_boxplot(fill = "red") +
  theme_minimal() +
  labs(title = "Boxplot of Cleaned Age Data")

# Visualize number of employees after cleaning
combined_df$no_employees <- as.factor(combined_df$no_employees)
ggplot(combined_df, aes(x = no_employees)) +
  geom_bar(fill = "yellow", color = "grey") +
  theme_minimal() +
  labs(title = "Cleaned Number of Employees Distribution", x = "Number of Employees (Categories)", y = "Frequency") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Rotate labels for better readability

#Final dataset summary
summary(combined_df)
write_xlsx(combined_df, "cleaned_survey_data.xlsx")