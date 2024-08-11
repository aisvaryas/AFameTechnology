import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r'C:\Users\sry91\Downloads\ECOMM DATA.xlsx'  # Update with your file path
xls = pd.ExcelFile(file_path)

# Load all sheets into dataframes
orders_df = pd.read_excel(xls, 'Orders')
returns_df = pd.read_excel(xls, 'Returns')
people_df = pd.read_excel(xls, 'People')

# Data Cleaning: Convert columns to appropriate data types
orders_df['Order Date'] = pd.to_datetime(orders_df['Order Date'])
orders_df['Ship Date'] = pd.to_datetime(orders_df['Ship Date'])

# Add Year and Month columns for time-based analysis
orders_df['Year'] = orders_df['Order Date'].dt.year
orders_df['Month'] = orders_df['Order Date'].dt.month

# 1. Sales and Profit Analysis Over Time
monthly_sales = orders_df.groupby(['Year', 'Month'])[['Sales', 'Profit']].sum().reset_index()

# Plot Sales and Profit Trends Over Time
plt.figure(figsize=(14, 7))
sns.lineplot(data=monthly_sales, x='Month', y='Sales', hue='Year', marker='o')
plt.title('Monthly Sales Trend Over Years')
plt.ylabel('Sales')
plt.xlabel('Month')
plt.show()

plt.figure(figsize=(14, 7))
sns.lineplot(data=monthly_sales, x='Month', y='Profit', hue='Year', marker='o')
plt.title('Monthly Profit Trend Over Years')
plt.ylabel('Profit')
plt.xlabel('Month')
plt.show()

# 2. Performance by Product Category and Sub-Category
category_performance = orders_df.groupby('Category')[['Sales', 'Profit']].sum().sort_values(by='Sales', ascending=False).reset_index()
subcategory_performance = orders_df.groupby(['Category', 'Sub-Category'])[['Sales', 'Profit']].sum().sort_values(by='Sales', ascending=False).reset_index()

# Plot Top Performing Product Categories
plt.figure(figsize=(10, 6))
sns.barplot(data=category_performance, x='Sales', y='Category', palette='viridis')
plt.title('Top Performing Product Categories by Sales')
plt.xlabel('Sales')
plt.ylabel('Category')
plt.show()

# 3. Returns Analysis
returns_df = returns_df.rename(columns={"Order ID": "Order ID"})
returned_orders = pd.merge(orders_df, returns_df, on='Order ID', how='left')
returned_orders['Returned'] = returned_orders['Returned'].fillna('No')

return_rate = returned_orders['Returned'].value_counts(normalize=True) * 100
return_impact = returned_orders[returned_orders['Returned'] == 'Yes'][['Sales', 'Profit']].sum()

# Plot Return Rates
plt.figure(figsize=(7, 5))
sns.barplot(x=return_rate.index, y=return_rate.values, palette='coolwarm')
plt.title('Return Rates')
plt.ylabel('Percentage')
plt.xlabel('Returned')
plt.show()

# 4. Regional Analysis
regional_sales = orders_df.groupby('Region')[['Sales', 'Profit']].sum().sort_values(by='Sales', ascending=False).reset_index()

# Plot Sales and Profit by Region
plt.figure(figsize=(12, 7))
sns.barplot(data=regional_sales, x='Sales', y='Region', palette='magma')
plt.title('Sales by Region')
plt.xlabel('Sales')
plt.ylabel('Region')
plt.show()

plt.figure(figsize=(12, 7))
sns.barplot(data=regional_sales, x='Profit', y='Region', palette='magma')
plt.title('Profit by Region')
plt.xlabel('Profit')
plt.ylabel('Region')
plt.show()

# Example of identifying any new pattern or trend
category_return_rate = returned_orders.groupby(['Category', 'Returned'])['Order ID'].count().unstack().fillna(0)
category_return_rate['Return Rate (%)'] = (category_return_rate['Yes'] / (category_return_rate['Yes'] + category_return_rate['No'])) * 100

# Plot Category Return Rates
plt.figure(figsize=(10, 6))
category_return_rate['Return Rate (%)'].sort_values(ascending=False).plot(kind='bar', color='teal')
plt.title('Category Return Rates')
plt.ylabel('Return Rate (%)')
plt.xlabel('Category')
plt.show()

# Identifying if any new trend is found
new_trend = category_return_rate.sort_values(by='Return Rate (%)', ascending=False).head(1)
print("\nNew Pattern Found: Highest return rate observed in the category:", new_trend.index[0])
