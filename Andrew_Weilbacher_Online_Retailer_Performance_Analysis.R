library(tidyverse)
library(readr)
library(openxlsx)
library(readxl)
library(lubridate)
library(dplyr)
library(ggplot2)
library(reshape2)

# Load Data
sessionCounts <- read_csv("DataAnalyst_Ecom_data_sessionCounts.csv")
addsToCart <- read_csv("DataAnalyst_Ecom_data_addsToCart.csv")

##### Creating Needed Excel Sheet ####

# Data Cleaning
## No Missing Values, thus interpolation not needed
## No NA values for numerical data points
## No Cleaning Needed

# Worksheet 1: Month * Device Aggregation
worksheet1 <- sessionCounts %>%
  mutate(month = floor_date(mdy(dim_date), "month")) %>%  # Convert dim_date to month
  group_by(month, dim_deviceCategory) %>%
  summarise(Sessions = sum(sessions),
            Transactions = sum(transactions),
            QTY = sum(QTY),
            ECR = sum(transactions) / sum(sessions),
            .groups = 'drop')  # Drop grouping for the final dataframe

# Correcting the date creation for addsToCart dataset
addsToCart <- addsToCart %>%
  mutate(month = make_date(dim_year, dim_month, 1))  # Create a date with the first day of each month

# Aggregate sessionCounts to monthly level 
sessionCounts_monthly <- sessionCounts %>%
  mutate(month = floor_date(mdy(dim_date), "month")) %>%
  group_by(month) %>%
  summarize(sessions = sum(sessions),
            transactions = sum(transactions),
            quantity = sum(QTY)) 

# Join data sets
joinedData <- sessionCounts_monthly %>%
  left_join(addsToCart, by = "month")

# 1. Calculate Month-over-Month Metrics
# We first need to determine the most recent two months in the data
latest_month <- max(joinedData$month)
second_latest_month <- max(joinedData$month[joinedData$month < latest_month])

# Filter data for these two months and calculate differences
mom_comparison <- joinedData %>%
  filter(month %in% c(latest_month, second_latest_month)) %>%
  arrange(month) %>%
  mutate(
    sessions_previous = lag(sessions),
    transactions_previous = lag(transactions),
    quantity_previous = lag(quantity)
  ) %>%
  mutate(
    sessions_abs_diff = sessions - sessions_previous,
    sessions_rel_diff = (sessions - sessions_previous) / sessions_previous,
    transactions_abs_diff = transactions - transactions_previous,
    transactions_rel_diff = (transactions - transactions_previous) / transactions_previous,
    quantity_abs_diff = quantity - quantity_previous,
    quantity_rel_diff = (quantity - quantity_previous) / quantity_previous,
    addsToCart_previous = lag(addsToCart),
    addsToCart_abs_diff = addsToCart - addsToCart_previous,
    addsToCart_rel_diff = (addsToCart - addsToCart_previous) / addsToCart_previous
  ) %>%
  select(-c(sessions_previous, transactions_previous, quantity_previous, addsToCart_previous)) # Optionally remove the 'previous' columns

# 2. Merge with addsToCart Data
mom_comparison <- mom_comparison %>%
  left_join(select(addsToCart, month), by = "month")

# 3. Additional Calculations
mom_comparison <- mom_comparison %>%
  mutate(conversion_rate = transactions / sessions,
         conversion_rate_diff = conversion_rate - lag(conversion_rate))

# Output to Excel
wb <- createWorkbook()
addWorksheet(wb, "MonthDeviceAggregation")
writeData(wb, "MonthDeviceAggregation", worksheet1)
addWorksheet(wb, "MonthOverMonthComparison")
writeData(wb, "MonthOverMonthComparison", mom_comparison)
saveWorkbook(wb, "RetailerPerformanceAnalysis.xlsx", overwrite = TRUE)

#### Data Analysis ####

# Data Exploration
# Load Data
sessionCounts <- read_csv("DataAnalyst_Ecom_data_sessionCounts.csv")
addsToCart <- read_csv("DataAnalyst_Ecom_data_addsToCart.csv")

deviceCategory <- sessionCounts$dim_deviceCategory
transactions <- sessionCounts$transactions
sess_date <- sessionCounts$dim_date
sessions <- sessionCounts$sessions
ecr <- transactions/sessions

df <- data.frame(
  sess_date = sess_date,
  dim_deviceCategory = deviceCategory,
  transactions = transactions,
  sessions = sessions,
  ecr = ecr
)

### Simple Exploration of the Data ###

# Basic plot
ggplot(df, aes(x = dim_deviceCategory, y = transactions, fill = dim_deviceCategory, color = dim_deviceCategory)) +
  geom_bar(stat="identity") +
  theme_minimal() +
  labs(title="Transactions by Device Category",
       x="Device Category",
       y="Number of Transactions")

df$sess_date <- mdy(df$sess_date)

# Time Series Plot
ggplot(df, aes(x=sess_date, y=transactions, group=dim_deviceCategory, color=dim_deviceCategory)) +
  geom_line() +
  theme_minimal() +
  labs(title="Transactions Over Time by Device Category",
       x="Date",
       y="Number of Transactions")

# Replace Inf with NA
df[df == Inf] <- NA

# Drop all rows with NaN values
df_cleaned <- na.omit(df)

# View the cleaned dataframe
print(df_cleaned)

# Calculate average ECR by device category
df_avg_ecr <- df_cleaned %>%
  group_by(dim_deviceCategory) %>%
  summarise(avg_ecr = mean(ecr))

# View the resulting dataframe
print(df_avg_ecr)

# Create a bar plot
ggplot(df_avg_ecr, aes(x = dim_deviceCategory, y = avg_ecr, fill = dim_deviceCategory, color = dim_deviceCategory)) +
  geom_bar(stat = "identity") +
  theme_minimal() +
  labs(title = "Average ECR by Device Category",
       x = "Device Category",
       y = "Average ECR") +
  scale_fill_brewer(palette = "Set1")

### Time Series Plot of Original Worksheet Data ###

worksheet1 <- sessionCounts %>%
  mutate(month = floor_date(mdy(dim_date), "month")) %>%  # Convert dim_date to month
  group_by(month, dim_deviceCategory) %>%
  summarise(Sessions = sum(sessions),
            Transactions = sum(transactions),
            QTY = sum(QTY),
            ECR = sum(transactions) / sum(sessions),
            .groups = 'drop')  # Drop grouping for the final dataframe

# Plot Number of Transactions by Device Category by Month
ggplot(worksheet1, aes(x = month, y = Transactions, group = dim_deviceCategory, color = dim_deviceCategory)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Number of Transactions by Device Category by Month",
       x = "Month",
       y = "Number of Transactions") +
  scale_x_date(date_breaks = "1 month", date_labels = "%b %Y") 

# Plot Number of Transactions by Month
ggplot(sessionCounts_monthly, aes(x = month, y = transactions)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Number of Transactions by Month",
       x = "Month",
       y = "Number of Transactions") +
  scale_x_date(date_breaks = "1 month", date_labels = "%b %Y") 

# Plot ECR by Device Category by Month
ggplot(worksheet1, aes(x = month, y = ECR, group = dim_deviceCategory, color = dim_deviceCategory)) +
  geom_line() +
  theme_minimal() +
  labs(title = "ECR by Device Category by Month",
       x = "Month",
       y = "ECR") +
  scale_x_date(date_breaks = "1 month", date_labels = "%b %Y") 

### Times Series for Add to Cart Data ###

# Correcting the date creation for addsToCart dataset
addsToCart <- addsToCart %>%
  mutate(month = make_date(dim_year, dim_month, 1))  # Create a date with the first day of each month

# Plot Number of addsToCart by Month
ggplot(addsToCart, aes(x = month, y = addsToCart)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Number of addsToCart by Device Category by Month",
       x = "Month",
       y = "Number of addsToCart") +
  scale_x_date(date_breaks = "1 month", date_labels = "%b %Y") 

### Analysis of Combined Data Sets ###

# Reshape the data for plotting
longData <- melt(joinedData, id.vars = "month", 
                 measure.vars = c("transactions", "addsToCart", "sessions"))

# Plotting
ggplot(longData, aes(x = month, y = value, color = variable, group = variable)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Transactions, Add-to-Cart, and Sessions Over Time",
       x = "Month",
       y = "Count",
       color = "Metric") +
  scale_x_date(date_breaks = "1 month", date_labels = "%b %Y") # Formatting the x-axis labels

# Plotting log values
ggplot(longData, aes(x = month, y = log(value), color = variable, group = variable)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Transactions, Add-to-Cart, and Sessions Over Time",
       x = "Month",
       y = "Count",
       color = "Metric") +
  scale_x_date(date_breaks = "1 month", date_labels = "%b %Y") # Formatting the x-axis labels