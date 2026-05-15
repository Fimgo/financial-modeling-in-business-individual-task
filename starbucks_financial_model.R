# Starbucks Corporation financial model
# Course: Financial Modeling in Business
# Units: USD millions, except percentages

required_packages <- c("openxlsx", "ggplot2", "scales")
missing_packages <- required_packages[!required_packages %in% rownames(installed.packages())]

if (length(missing_packages) > 0) {
  install.packages(missing_packages)
}

library(openxlsx)
library(ggplot2)
library(scales)

output_file <- "Starbucks_Financial_Model.xlsx"

# Historical data from Starbucks Fiscal 2025 Annual Report, Form 10-K.
# Source:
# https://s203.q4cdn.com/326826266/files/doc_financials/2025/ar/Starbucks-Corporation_2025-Annual-Report-Web-Ready.pdf

historical <- data.frame(
  Year = c(2023, 2024, 2025),
  Revenue = c(35975.6, 36176.2, 37184.4),
  Product_Distribution_Costs = c(11409.1, 11180.6, 11658.2),
  Store_Operating_Expenses = c(14720.3, 15286.5, 17058.9),
  Other_Operating_Expenses = c(539.4, 565.6, 584.6),
  Depreciation_Amortization = c(1362.6, 1512.6, 1684.7),
  General_Admin_Expenses = c(2441.3, 2523.3, 2617.2),
  Restructuring_Impairments = c(21.8, 0.0, 892.0),
  Income_From_Equity_Investees = c(298.4, 301.2, 247.8)
)

historical$Total_Operating_Expenses <- with(
  historical,
  Product_Distribution_Costs +
    Store_Operating_Expenses +
    Other_Operating_Expenses +
    Depreciation_Amortization +
    General_Admin_Expenses +
    Restructuring_Impairments
)

historical$Operating_Income <- with(
  historical,
  Revenue - Total_Operating_Expenses + Income_From_Equity_Investees
)

historical$Operating_Margin <- historical$Operating_Income / historical$Revenue

# Forecast assumptions.
# Base case assumes modest revenue recovery, lower restructuring costs,
# and partial normalization of operating margin after the weak FY2025 year.
assumptions <- data.frame(
  Driver = c(
    "Revenue growth 2026E",
    "Revenue growth 2027E",
    "Revenue growth 2028E",
    "Variable cost ratio",
    "Fixed cost growth",
    "Equity income as % of revenue",
    "Restructuring costs 2026E",
    "Restructuring costs 2027E",
    "Restructuring costs 2028E"
  ),
  Value = c(0.035, 0.040, 0.040, 0.765, 0.030, 0.006, 300.0, 100.0, 50.0),
  Note = c(
    "Moderate recovery after FY2025",
    "Base-case normalized growth",
    "Base-case normalized growth",
    "Product, distribution, store, and other operating expenses as % of revenue",
    "D&A and G&A growth",
    "Based on recent historical range",
    "Lower than FY2025 due to restructuring normalization",
    "Continued decline in restructuring costs",
    "Near-normal level"
  )
)

forecast_years <- c(2026, 2027, 2028)
forecast <- data.frame(Year = forecast_years)

forecast$Revenue <- NA_real_
forecast$Product_Distribution_Costs <- NA_real_
forecast$Store_Operating_Expenses <- NA_real_
forecast$Other_Operating_Expenses <- NA_real_
forecast$Depreciation_Amortization <- NA_real_
forecast$General_Admin_Expenses <- NA_real_
forecast$Restructuring_Impairments <- c(300.0, 100.0, 50.0)
forecast$Income_From_Equity_Investees <- NA_real_

growth <- assumptions$Value[1:3]
variable_ratio <- assumptions$Value[4]
fixed_growth <- assumptions$Value[5]
equity_ratio <- assumptions$Value[6]

forecast$Revenue[1] <- historical$Revenue[3] * (1 + growth[1])
forecast$Revenue[2] <- forecast$Revenue[1] * (1 + growth[2])
forecast$Revenue[3] <- forecast$Revenue[2] * (1 + growth[3])

for (i in seq_along(forecast_years)) {
  variable_costs <- forecast$Revenue[i] * variable_ratio
  forecast$Product_Distribution_Costs[i] <- variable_costs * 0.385
  forecast$Store_Operating_Expenses[i] <- variable_costs * 0.595
  forecast$Other_Operating_Expenses[i] <- variable_costs * 0.020

  if (i == 1) {
    forecast$Depreciation_Amortization[i] <- historical$Depreciation_Amortization[3] * (1 + fixed_growth)
    forecast$General_Admin_Expenses[i] <- historical$General_Admin_Expenses[3] * (1 + fixed_growth)
  } else {
    forecast$Depreciation_Amortization[i] <- forecast$Depreciation_Amortization[i - 1] * (1 + fixed_growth)
    forecast$General_Admin_Expenses[i] <- forecast$General_Admin_Expenses[i - 1] * (1 + fixed_growth)
  }

  forecast$Income_From_Equity_Investees[i] <- forecast$Revenue[i] * equity_ratio
}

forecast$Total_Operating_Expenses <- with(
  forecast,
  Product_Distribution_Costs +
    Store_Operating_Expenses +
    Other_Operating_Expenses +
    Depreciation_Amortization +
    General_Admin_Expenses +
    Restructuring_Impairments
)

forecast$Operating_Income <- with(
  forecast,
  Revenue - Total_Operating_Expenses + Income_From_Equity_Investees
)

forecast$Operating_Margin <- forecast$Operating_Income / forecast$Revenue

model <- rbind(historical, forecast)

fixed_costs_2026 <- forecast$Depreciation_Amortization[1] +
  forecast$General_Admin_Expenses[1] +
  forecast$Restructuring_Impairments[1]
contribution_margin <- 1 - variable_ratio
break_even_revenue <- fixed_costs_2026 / contribution_margin

sensitivity <- expand.grid(
  Revenue_Growth = c(-0.05, 0.00, 0.05),
  Variable_Cost_Ratio = c(variable_ratio - 0.02, variable_ratio, variable_ratio + 0.02)
)

sensitivity$Revenue_2026E <- historical$Revenue[3] * (1 + assumptions$Value[1] + sensitivity$Revenue_Growth)
sensitivity$Operating_Income_2026E <- with(
  sensitivity,
  Revenue_2026E * (1 - Variable_Cost_Ratio) -
    fixed_costs_2026 +
    Revenue_2026E * equity_ratio
)
sensitivity$Operating_Margin_2026E <- sensitivity$Operating_Income_2026E / sensitivity$Revenue_2026E

wb <- createWorkbook()
addWorksheet(wb, "Cover")
addWorksheet(wb, "Data")
addWorksheet(wb, "Assumptions")
addWorksheet(wb, "Forecast")
addWorksheet(wb, "P&L")
addWorksheet(wb, "Break-even")
addWorksheet(wb, "Sensitivity")
addWorksheet(wb, "Sources")

writeData(wb, "Cover", "Starbucks Corporation Financial Model", startCol = 1, startRow = 1)
writeData(wb, "Cover", "Currency: USD millions. Historical period: FY2023-FY2025. Forecast period: FY2026E-FY2028E.", startCol = 1, startRow = 3)
writeData(wb, "Cover", "Main source: Starbucks Fiscal 2025 Annual Report, Form 10-K.", startCol = 1, startRow = 4)

writeData(wb, "Data", historical)
writeData(wb, "Assumptions", assumptions)
writeData(wb, "Forecast", forecast)
writeData(wb, "P&L", model)

break_even <- data.frame(
  Metric = c("Variable cost ratio", "Contribution margin", "Fixed costs 2026E", "Break-even revenue 2026E"),
  Value = c(variable_ratio, contribution_margin, fixed_costs_2026, break_even_revenue)
)
writeData(wb, "Break-even", break_even)
writeData(wb, "Sensitivity", sensitivity)

sources <- data.frame(
  Source_ID = c("S1", "S2", "S3"),
  Source = c(
    "Starbucks Fiscal 2025 Annual Report PDF",
    "Starbucks Investor Relations - Annual Reports",
    "Starbucks SEC Filing Details - Form 10-K filed 2025-11-14"
  ),
  URL = c(
    "https://s203.q4cdn.com/326826266/files/doc_financials/2025/ar/Starbucks-Corporation_2025-Annual-Report-Web-Ready.pdf",
    "https://investor.starbucks.com/financials/annual-reports/",
    "https://investor.starbucks.com/financials/sec-filings/sec-filings-details/default.aspx?FilingId=18927461"
  ),
  Notes = c(
    "Used for consolidated revenues, costs, operating income, and business description.",
    "Official page for annual reports.",
    "Official 10-K filing page with PDF, Excel, and XBRL downloads."
  )
)
writeData(wb, "Sources", sources)

header_style <- createStyle(textDecoration = "bold", fgFill = "#1F4E78", fontColour = "#FFFFFF")
input_style <- createStyle(fontColour = "#0000FF")
percent_style <- createStyle(numFmt = "0.0%")
currency_style <- createStyle(numFmt = "$#,##0.0;[Red]($#,##0.0);-")

for (sheet in names(wb)) {
  addStyle(wb, sheet, header_style, rows = 1, cols = 1:20, gridExpand = TRUE)
  freezePane(wb, sheet, firstRow = TRUE)
  setColWidths(wb, sheet, cols = 1:20, widths = "auto")
}

addStyle(wb, "Assumptions", input_style, rows = 2:(nrow(assumptions) + 1), cols = 2)
addStyle(wb, "Assumptions", percent_style, rows = 2:7, cols = 2)
addStyle(wb, "Break-even", percent_style, rows = 2:3, cols = 2)
addStyle(wb, "Break-even", currency_style, rows = 4:5, cols = 2)

saveWorkbook(wb, output_file, overwrite = TRUE)

revenue_chart <- ggplot(model, aes(x = Year, y = Revenue)) +
  geom_line(color = "#1F4E78", linewidth = 1.2) +
  geom_point(color = "#1F4E78", size = 2.5) +
  scale_y_continuous(labels = dollar_format(suffix = "m")) +
  labs(title = "Starbucks Revenue Dynamics", x = NULL, y = "USD millions") +
  theme_minimal()

profit_chart <- ggplot(model, aes(x = Year, y = Operating_Income)) +
  geom_col(fill = "#70AD47") +
  scale_y_continuous(labels = dollar_format(suffix = "m")) +
  labs(title = "Starbucks Operating Income", x = NULL, y = "USD millions") +
  theme_minimal()

ggsave("starbucks_revenue_chart.png", revenue_chart, width = 8, height = 4.5, dpi = 160)
ggsave("starbucks_operating_income_chart.png", profit_chart, width = 8, height = 4.5, dpi = 160)

message("Done: ", output_file)
