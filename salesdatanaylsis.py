# Data Manipulation
import matplotlib as matplotlib
import pandas as pd
from openpyxl import Workbook
# Data Visualisation
import matplotlib.pyplot as plt
from jedi.api.refactoring import inline
#%matplotlib inline

import seaborn as sns
df=pd.read_excel('Example1//superstore_sales.xlsx')
#df = pd.read_csv("C://Users//Agiliad//Downloads//Example1//superstore_sales.csv")
#print(df)
"""Error:
ImportError: Missing optional dependency 'openpyxl'. Use pip or conda to install openpyxl.

Resolution:
import workbook from openpyxl"""
# First five rows of the dataset
#print(df.head())

# Last five rows of the dataset
#print(df.tail())

# Shape of the dataset  no of rows * no of columns
#print(df.shape)
"""Error:
TypeError: 'tuple' object is not callable
resolve it by using solution: df.shape 
"""

# Columns present in the dataset
#print(df.columns)
"""TypeError: 'Index' object is not callable
use only df.columns"""

# A concise summary of the dataset
#print(df.info())

#Check missing values in dataset
#print(df.isna())

# Checking missing values sum  if no missing value it will show as 0 column wise
#print(df.isna().sum())

# Generating descriptive statistics summary
#print(df.describe().round())

#WHAT IS THE OVERALL SALES TREND?
# Getting month year from order_date
df['month_year'] = df['order_date'].apply(lambda x: x.strftime('%Y-%m'))


# grouping month_year by sales
#b=df_temp = df.groupby('month_year').sum()['sales'].reset_index()

# Setting the figure size
"""plt.figure(figsize=(16, 5))
plt.plot(df_temp['month_year'], df_temp['sales'], color='#b80045')
plt.xticks(rotation='vertical', size=4)
plt.show()"""

# Grouping products by sales
prod_sales = pd.DataFrame(df.groupby('product_name').sum()['sales'])
#print(prod_sales)

# Sorting the dataframe in descending order
a=prod_sales.sort_values(by=['sales'], inplace=True, ascending=False)
#print(a)

# Top 10 products by sales
prod_sales[:10]

# Grouping products by Quantity
best_selling_prods = pd.DataFrame(df.groupby('product_name').sum()['quantity'])

# Sorting the dataframe in descending order
best_selling_prods.sort_values(by=['quantity'], inplace=True, ascending=False)

# Most selling products
best_selling_prods[:10]

# Setting the figure size
"""plt.figure(figsize=(5, 4))

# countplot: Show the counts of observations in each categorical bin using bars
sns.countplot(x='ship_mode', data=df)

# Display the figure
plt.show()
"""
# Grouping products by Category and Sub-Category
cat_subcat = pd.DataFrame(df.groupby(['category', 'sub_category']).sum()['profit'])

# Sorting the values
cat_subcat.sort_values(['category','profit'], ascending=False)

"""**Priject link github others**
https://github.com/EDSOfficial/Sales-Analysis/blob/master/SalesAnalysis.ipynb"""