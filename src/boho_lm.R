rm(list=ls())
# Settings ----------------------------------------------------------------

library(zoo)
library(ggplot2)
library(readxl)
library(dplyr)
library(openxlsx)
library(car)
library(lmtest)

# Data --------------------------------------------------------------------

# Specify the path to your Excel file
file_path <- "./data/BOHO_data.xlsx"
df <- read_excel(file_path, sheet = "BOHO", skip = 4)
colnames(df) <- c('Date','BOHO','HO','BO')
df$Date <- as.Date(df$Date)

# Sort the DataFrame by Date in descending order
df <- df %>% arrange(Date)
df_date <- df$Date


# Remove unwanted columns
df <- df[, c('BOHO', 'HO', 'BO')]  # Keep only the necessary columns
df <- as.data.frame(df)

# Replace NA values with the value from the row below
df <- df %>%
  mutate(across(everything(), ~na.locf(., na.rm = FALSE, fromLast = TRUE)))

# Check if the columns exist and are numeric
if (!all(c("BOHO", "HO", "BO") %in% colnames(df))) {
  stop("One or more columns are missing.")
}

# Check if the data is numeric
df[] <- lapply(df, as.numeric)


# Descriptive Analysis ----------------------------------------------------

# Checking Correlation Between Variables
cor(df$BOHO, df$HO)
cor(df$BOHO, df$BO)

# Plot 1
plot(df$HO, df$BOHO, main = "BOHO vs HO", xlab = "HO", ylab = "BOHO")

# Plot 2
plot(df$BO, df$BOHO, main = "BOHO vs BO", xlab = "BO", ylab = "BOHO")

# LM ----------------------------------------------------------------------

# Linear Model
lm_boho <- lm(BOHO ~ HO + BO, data = df)
summary(lm_boho)

# Calculate Variance Inflation Factor (VIF)
vif_values <- vif(lm_boho)
vif_values

# Homogeneity of Variance
bptest(lm_boho)

# Shapiro-Wilk test for normality of residuals
shapiro_test <- shapiro.test(residuals(lm_boho))
cat("Shapiro-Wilk test p-value:", shapiro_test$p.value, "\n")

# Influential Plots
cooks_d <- cooks.distance(lm_boho)
plot(cooks_d, type = "h", main = "Cook's Distance", ylab = "Distance")
abline(h = 4 / length(cooks_d), col = "red")  # Threshold line

# Moving LM Model ---------------------------------------------------------
window_size <- 30  # Tamanho da janela (30 dias)

# Função para calcular os coeficientes da regressão
rolling_regression <- function(data, window_size) {
  coefs <- rollapply(data, width = window_size, FUN = function(x) {
    # Create a data frame using only the necessary columns
    x_df <- as.data.frame(x[, c("BOHO", "HO", "BO")])
    
    # Fit the linear model
    model <- lm(BOHO ~ HO + BO, data = x_df)
    
    # Return the coefficients
    return(coef(model))
  }, by.column = FALSE, fill = NA, align = "right")
  
  return(as.data.frame(coefs))
}

# Applying LM Function Over Time
rolling_results <- rolling_regression(df, window_size)
rolling_results$Date <- df_date

# Move Date to the first position
rolling_results <- rolling_results %>%
  select(Date, everything())

# Export to Excel file
write.xlsx(rolling_results, "rolling_results.xlsx")

# Convert to long format for ggplot
rolling_long <- reshape2::melt(rolling_results, id.vars = "Date", 
                               variable.name = "Coefficient", 
                               value.name = "Value")

# Plot the coefficients over time
ggplot(rolling_long, aes(x = Date, y = Value, color = Coefficient)) +
  geom_line() +
  labs(title = "Rolling Regression Coefficients Over Time",
       x = "Date",
       y = "Coefficient Value") +
  theme_minimal() +
  theme(legend.title = element_blank())

