import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import FuncFormatter
import calendar
import matplotlib.patheffects as path_effects

sheets = pd.read_excel('/content/Regional Sales Dataset.xlsx', sheet_name = None)

df_sales = sheets['Sales Orders']
df_customers = sheets['Customers']
df_products = sheets['Products']
df_regions = sheets['Regions']
df_state_reg = sheets['State Regions']
df_budgets = sheets['2017 Budgets']

print("df_sales shape: ",(df_sales.shape))
print("df_customers shape: ",(df_customers.shape))
print("df_products shape: ",(df_products.shape))
print("df_regions shape: ",(df_regions.shape))
print("df_state_reg shape: ",(df_state_reg.shape))
print("df_budgets shape: ",(df_budgets.shape))

df_sales.head(5)

df_customers.head(5)

new_header = df_state_reg.iloc[0]
df_state_reg.columns = new_header
df_state_reg = df_state_reg[1 :].reset_index(drop=True)

df_state_reg.head(5)

df_sales.isnull().sum()

df_customers.isnull().sum()

## Merge with Customer
df = df_sales.merge(
    df_customers,
    how = 'left',
    left_on='Customer Name Index',
    right_on='Customer Index'
)

df.head(2)

## Merge With Prodcuts
df = df.merge(
    df_products,
    how='left',
    left_on='Product Description Index',
    right_on='Index'
)

df.head(2)

## Merge With Regions
df = df.merge(
    df_regions,
    how='left',
    left_on='Delivery Region Index',
    right_on='id'
)

df.head(2)

## Merge With State Region
df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)

df.head(2)

cols_to_drop =['Customer Index','Index','id','State Code']
df = df.drop(columns=cols_to_drop,errors='ignore')

df.head(5)

# The next line contains a typo `coloums` which was corrected to `columns` later in the notebook.
# To generate a functional script, we will use the corrected version from the `uvkcXPOfG-IX` cell.
# df.coloums = df.columns.str.lower()
# df.columns.values

## Merge With Budgets
df = df.merge(
    df_budgets,
    how='left',
    on ='Product Name'
)

df.columns = df.columns.str.lower()
df.columns.values

 ## Keeping the important colums and deleting the colums we dont need

cols_to_keep = [

 'ordernumber',
    'orderdate',
    'customer names',
    'channel',
    'product name',
    'order quantity',
    'unit price',
    'line total',
    'total unit cost',
    'state_code',
    'county',
    'state',
    'region',
    'latitude',
    'longitude',
    '2017 budgets'
]

df = df[cols_to_keep]

df.head(2)

# Rename the columns

# Rename the columns
df = df.rename(columns={
    'ordernumber': 'order_number',
    'orderdate': 'order_date',
    'customer names': 'customer_name',
    'product name': 'product_name',
    'order quantity': 'order_quantity',
    'unit price': 'unit_price',
    'line total': 'revenue',
    'total unit cost': 'cost',
    'state_code': 'state',
    'state': 'state_name',
    'latitude': 'lat',
    'longitude': 'lon',
    '2017 budgets': 'budget'
})

df.head(2)

# Blank out budget for non-2017 orders
df.loc[df['order_date'].dt.year != 2017, 'budget'] = pd.NA

# Line total is revenue
df[['order_date', 'product_name', 'revenue', 'budget']].tail(2)

df.info()

# df.info() # This was a duplicate call, not including in the script

## filter the dataset to include only the record for year 2017
df_2017 = df[df['order_date'].dt.year == 2017]

df_2017.head(2)

df['total_cost'] = df['order_quantity'] * df['cost']

df['profit'] = df['revenue'] - df['total_cost']
df['profit_margin_pct']=df['profit']/df['revenue']*100

df.head(2)

### EDA

## MONTHLY SALES TRENDS

df['order_month'] = df['order_date'].dt.to_period('M')
monthly_sales = df.groupby('order_month')['revenue'].sum()
plt.figure(figsize=(12, 6))
formatter = FuncFormatter(lambda x, pos: f'{x/1e6:.1f}M')  # fixed le6 → 1e6
plt.gca().yaxis.set_major_formatter(formatter)
plt.plot(monthly_sales.index.astype(str), monthly_sales.values, marker='o')
plt.title('Monthly Sales Trends')
plt.xlabel('Month')
plt.ylabel('Total Revenue (Millions)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

df['order_month'] = df['order_date'].dt.to_period('M')
monthly_sales = df.groupby('order_month')['revenue'].sum()
months = monthly_sales.index.astype(str)
values = monthly_sales.values
plt.figure(figsize=(12, 6))
plt.plot(months, values, color='#1f77b4', linewidth=2.5, marker='o', markersize=6)
formatter = FuncFormatter(lambda x, pos: f'{x/1e6:.1f}M')
plt.gca().yaxis.set_major_formatter(formatter)
plt.grid(True, linestyle='--', alpha=0.6)
plt.title('Monthly Sales Trends (Revenue by Month)', fontsize=16, fontweight='bold', pad=15)
plt.xlabel('Month', fontsize=12)
plt.ylabel('Total Revenue (in Millions)', fontsize=12)
plt.xticks(rotation=45, ha='right')
for x, y in zip(months, values):
    plt.text(x, y + (max(values)*0.01), f'{y/1e6:.1f}M', ha='center', va='bottom', fontsize=9, color='#333')
plt.tight_layout()
plt.show()

# Ensure order_date is in datetime format
df['order_date'] = pd.to_datetime(df['order_date'])

# Filter out records from February 2018
df_new = df[~((df['order_date'].dt.year == 2018) & (df['order_date'].dt.month == 2))]

df['order_month_num'] = df_new['order_date'].dt.month
df['order_month_name'] = df_new['order_date'].dt.month_name()

monthly_seasonality = (
    df.groupby('order_month_num')['revenue']
      .sum()
      .reset_index()
      .sort_values('order_month_num')
)

monthly_seasonality['month'] = monthly_seasonality['order_month_num'].dropna().apply(lambda x: calendar.month_abbr[int(x)])

plt.figure(figsize=(12, 6))
plt.plot(
    monthly_seasonality['month'],
    monthly_seasonality['revenue'],
    color='#ff7f0e',
    linewidth=2.5,
    marker='o',
    markersize=6
)

formatter = FuncFormatter(lambda x, pos: f'{x/1e6:.1f}M')
plt.gca().yaxis.set_major_formatter(formatter)

plt.grid(True, linestyle='--', alpha=0.6)

plt.title('Overall Monthly Sales Trend (All Years Combined)', fontsize=16, fontweight='bold', pad=15)
plt.xlabel('Month', fontsize=12)
plt.ylabel('Total Revenue (in Millions)', fontsize=12)

plt.xticks(rotation=45, ha='right')

for x, y in zip(monthly_seasonality['month'], monthly_seasonality['revenue']):
    plt.text(x, y + (max(monthly_seasonality['revenue']) * 0.01), f'{y/1e6:.1f}M',
             ha='center', va='bottom', fontsize=9, color='#333')

plt.tight_layout()
plt.show()

top_products = (
    df_new.groupby('product_name')['revenue']
          .sum()
          .reset_index()
          .sort_values('revenue', ascending=False)
          .head(10)
)

plt.figure(figsize=(12, 7))
sns.set_style("whitegrid")

colors = sns.color_palette("viridis", len(top_products))

bars = plt.barh(
    top_products['product_name'],
    top_products['revenue'],
    color=colors,
    edgecolor='none',
    height=0.6,
    alpha=0.9
)

plt.gca().invert_yaxis()

for bar in bars:
    width = bar.get_width()
    plt.text(
        width + (max(top_products['revenue']) * 0.01),
        bar.get_y() + bar.get_height()/2,
        f'{width/1e6:.1f}M',
        va='center',
        ha='left',
        fontsize=11,
        fontweight='bold',
        color='#333',
        path_effects=[path_effects.withStroke(linewidth=2, foreground='white')]
    )

formatter = FuncFormatter(lambda x, pos: f'{x/1e6:.0f}M')
plt.gca().xaxis.set_major_formatter(formatter)

plt.title('Top 10 Products by Total Revenue', fontsize=18, fontweight='bold', pad=15)
plt.xlabel('Total Revenue (in Millions)', fontsize=13)
plt.ylabel('Product Name', fontsize=13)

plt.grid(axis='x', linestyle='--', alpha=0.5)
sns.despine(left=True, bottom=True)

plt.gca().set_facecolor('#f9f9f9')

plt.tight_layout()
plt.show()

bottom_products = (
    df_new.groupby('product_name')['revenue']
          .sum()
          .reset_index()
          .sort_values('revenue', ascending=True)
          .head(10)
)

plt.figure(figsize=(12, 7))
sns.set_style("whitegrid")

colors = sns.color_palette("mako", len(bottom_products))

bars = plt.barh(
    bottom_products['product_name'],
    bottom_products['revenue'],
    color=colors,
    edgecolor='none',
    height=0.6,
    alpha=0.9
)

plt.gca().invert_yaxis()

for bar in bars:
    width = bar.get_width()
    plt.text(
        width + (max(bottom_products['revenue']) * 0.01),
        bar.get_y() + bar.get_height()/2,
        f'{width/1e6:.1f}M',
        va='center',
        ha='left',
        fontsize=11,
        fontweight='bold',
        color='#333',
        path_effects=[path_effects.withStroke(linewidth=2, foreground='white')]
    )

formatter = FuncFormatter(lambda x, pos: f'{x/1e6:.0f}M')
plt.gca().xaxis.set_major_formatter(formatter)

plt.title('Bottom 10 Products by Total Revenue', fontsize=18, fontweight='bold', pad=15)
plt.xlabel('Total Revenue (in Millions)', fontsize=13)
plt.ylabel('Product Name', fontsize=13)

plt.grid(axis='x', linestyle='--', alpha=0.5)
sns.despine(left=True, bottom=True)

plt.gca().set_facecolor('#f9f9f9')

plt.tight_layout()
plt.show()

chan_sales = (
    df_new.groupby('channel')['revenue']
          .sum()
          .sort_values(ascending=True)
)

plt.figure(figsize=(9, 6))

plt.pie(
    chan_sales.values,
    labels=chan_sales.index,
    autopct='%1.1f%%',
    startangle=90,
    colors=sns.color_palette('coolwarm', len(chan_sales)),
    wedgeprops={'edgecolor': 'white', 'linewidth': 1}
)

plt.title('Total Sales by Channel', fontsize=16, fontweight='bold', pad=15)
plt.tight_layout()
plt.show()

# Calculate AOV per order (group by order_number)
aov = df.groupby('order_number')['revenue'].sum()

# Plot histogram
plt.figure(figsize=(12, 4))
plt.hist(
    aov,
    bins=50,
    color='orange',
    edgecolor='black',
    alpha=0.8
)

# Add title and labels
plt.title('Distribution of Average Order Value (AOV)', fontsize=16, fontweight='bold', pad=15)
plt.xlabel('Order Value (USD)', fontsize=12)
plt.ylabel('Number of Orders', fontsize=12)

# Add gridlines for readability
plt.grid(True, linestyle='--', alpha=0.5)

plt.tight_layout()
plt.show()

###unit price distribution per product
sns.set_theme(style="whitegrid")

plt.figure(figsize=(14, 7))

sns.boxplot(
    data=df,
    x='product_name',
    y='unit_price',
    hue='product_name',
    dodge=False,
    palette='viridis',
    legend=False
)


plt.title('Unit Price Distribution per Product', fontsize=16, fontweight='bold', pad=15)
plt.xlabel('Product', fontsize=12)
plt.ylabel('Unit Price (USD)', fontsize=12)
plt.xticks(rotation=45, ha='right')

plt.grid(True, linestyle='--', alpha=0.4)
plt.tight_layout()
plt.show()

##Top 10 States by Revenue and Order Count
state_summary = (
    df.groupby('state_name')
      .agg({'revenue': 'sum', 'order_number': 'nunique'})
      .sort_values('revenue', ascending=False)
      .head(10)
      .reset_index()
)


fig, ax1 = plt.subplots(figsize=(12, 6))
sns.barplot(data=state_summary, x='state_name', y='revenue', color='#4c72b0', ax=ax1, label='Revenue')

ax2 = ax1.twinx()
sns.lineplot(data=state_summary, x='state_name', y='order_number', color='#dd8452', marker='o', ax=ax2, label='Order Count')

ax1.set_title('Top 10 States by Revenue and Order Count', fontsize=16, fontweight='bold')
ax1.set_xlabel('State')
ax1.set_ylabel('Total Revenue ($)')
ax2.set_ylabel('Order Count')

ax1.tick_params(axis='x', rotation=45)
fig.legend(loc='upper left', bbox_to_anchor=(0.1,0.9))
plt.tight_layout()
plt.show()

# Calculate average profit margin by channel

channel_profit = (
    df.groupby('channel')['profit_margin_pct']
      .mean()
      .reset_index()
      .sort_values('profit_margin_pct', ascending=False)
)

plt.figure(figsize=(10, 6))

sns.barplot(
    data=channel_profit,
    x='channel',
    y='profit_margin_pct',
    hue='channel',
    palette='coolwarm',
    dodge=False,
    legend=False
)


plt.title('Average Profit Margin by Channel', fontsize=16, fontweight='bold', pad=15)
plt.xlabel('Channel', fontsize=12)
plt.ylabel('Average Profit Margin (%)', fontsize=12)
plt.grid(True, linestyle='--', alpha=0.5)


for i, row in channel_profit.iterrows():
    plt.text(i, row['profit_margin_pct'] + 0.5, f"{row['profit_margin_pct']:.1f}%",
             ha='center', va='bottom', fontsize=10, fontweight='bold', color='#333')

plt.tight_layout()
plt.show()

customer_revenue = (
    df.groupby('customer_name')['revenue']
      .sum()
      .sort_values(ascending=False)
      .reset_index()
)

top_10_customers = customer_revenue.head(10)
bottom_10_customers = customer_revenue.tail(10)

fig, axes = plt.subplots(1, 2, figsize=(16, 6))

sns.barplot(
    data=top_10_customers,
    y='customer_name',
    x='revenue',
    hue='customer_name',
    palette='crest',
    dodge=False,
    legend=False,
    ax=axes[0]
)
axes[0].set_title('Top 10 Customers by Revenue', fontsize=14, fontweight='bold')
axes[0].set_xlabel('Revenue ($)')
axes[0].set_ylabel('Customer')
axes[0].grid(True, linestyle='--', alpha=0.4)


sns.barplot(
    data=bottom_10_customers,
    y='customer_name',
    x='revenue',
    hue='customer_name',
    palette='flare',
    dodge=False,
    legend=False,
    ax=axes[1]
)
axes[1].set_title('Bottom 10 Customers by Revenue', fontsize=14, fontweight='bold')
axes[1].set_xlabel('Revenue ($)')
axes[1].set_ylabel('')
axes[1].grid(True, linestyle='--', alpha=0.4)


plt.tight_layout()
plt.show()

###Customer Segmentation: Revenue vs Profit Margin
customer_seg = (
    df.groupby('customer_name')
      .agg({'revenue': 'sum', 'profit_margin_pct': 'mean'})
      .reset_index()
)

plt.figure(figsize=(10, 6))
sns.scatterplot(
    data=customer_seg,
    x='revenue',
    y='profit_margin_pct',
    hue='profit_margin_pct',
    size='revenue',
    sizes=(40, 400),
    palette='viridis',
    alpha=0.8,
    edgecolor='black'
)

plt.title('Customer Segmentation: Revenue vs Profit Margin', fontsize=16, fontweight='bold')
plt.xlabel('Revenue ($)')
plt.ylabel('Profit Margin (%)')
plt.grid(True, linestyle='--', alpha=0.4)
plt.legend(title='Profit Margin (%)')
plt.tight_layout()
plt.show()

###Correlation Heatmap

numeric_cols = df.select_dtypes(include=['float64', 'int64'])

plt.figure(figsize=(10, 6))
sns.heatmap(
    numeric_cols.corr(),
    annot=True,
    fmt=".2f",
    cmap='coolwarm',
    linewidths=0.5
)
plt.title('Correlation Heatmap', fontsize=16, fontweight='bold')
plt.tight_layout()
plt.show()

df.to_csv('Sales Analysis (After EDA).csv')

from google.colab import files
# files.download('Sales Analysis (After EDA).csv') # The user asked to download a specific file, so commenting this out.
