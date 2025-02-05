import os
import csv
import pandas as pd
import re
import openpyxl



class Workbook:
	"""
	Stores multiple Sheet object in a dictionay for excel export as multiple sheets
	"""
	
	def __init__(self, sheets):
		self.sheets = sheets
		print("Workbook created! Use .export(filename='example' to export excel Workbook to Workbook Accounting folder)")
		
	def export(self, filename='example'):
		output_file = f"Workbooks Accounting/{filename}.xlsx"
		with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
			for sheet_name, sheet_data in self.sheets.items():
				df = pd.DataFrame(sheet_data.subset_accounts)
				df.to_excel(writer, sheet_name=sheet_name, index=False) 		
				
			print(f"Excel file '{output_file}' created succesfully!")
						



class Sheet:
	def __init__(self, year, df):
		self.tax_year = year
		self.accounts = df
		self.subset_accounts = self.accounts
## Doesn't work
	def subset_dates(self, start, end):
		"""
		Method to subset dates
		"""
		self.subset_accounts = self.subset_accounts[end >= ['Date'] >= start]
		return self
		
	def subset_credit(self):
		"""
		Method to subset deposits (often taxable income) 
		"""
		self.subset_accounts = self.subset_accounts[self.subset_accounts['Transaction'] > 0]
		return self

	def subset_transaction(self, lower=-1000000, upper=1000000):
		"""
		Method to subset transactions of a certain size(e.g > £100)
		"""
		self.subset_accounts = self.subset_accounts[(self.subset_accounts['Transaction'] > lower) & (self.subset_accounts['Transaction'] < upper)]
		return self

	def subset_ref(self, word, notword=False):
		

		"""
		Subset for "text" in desciption (e.g. to look for certain expenses ('tfl'))
		"""
		# use self.accounts for subsetting, as want to be inclusive i.e. select from all of accounts when notword = False
		if notword == False:
			self.subset_accounts = self.accounts[self.accounts['Ref'].str.contains(word, case=False)]
			return self
		# Can use self.subset_accounts for subsetting, when narrowing down account with multiple exlusions i.e. when notword = True
		else:
			self.subset_accounts = self.subset_accounts[~self.subset_accounts['Ref'].str.contains(word, case=False)]
			return self
			

	def subset_refs(self, *args, notword=False):
		"""
		Subset for multiple "text" in description (e.g. to look for certain expenses ('tfl'))
		"""
		
		# Where I want to concatenate multiple DataFrames with matiching references		
		if notword == False:
			subset = pd.DataFrame()

			for word in args:
				self.subset_ref(word, notword=False)
				subset = pd.concat([subset, self.subset_accounts])
			self.subset_accounts = subset
			return self	
			
		# Where I want to subset in an iterative manner on the same DataFrame 		
		else:
			for word in args:
				self.subset_ref(word, notword=True)
			return self


	def subset_reset(self):
		"""
		Subset to reset
		"""
		self.subset_accounts = self.accounts
		return self

	def subset_total(self):
		
		self.subset_accounts.loc['Total'] = {'Date': "Total", 'Transaction': self.subset_accounts['Transaction'].sum()}
		return self
		
		
		
		
# Plotting 

"""
Time series bar plot of subset data (e.g. monthly, weekly income and expenses)
"""


# Output in some way. Excel? CSV? Better way to visualise this?



"""
method to export to excel for viewing
"""


# E.g. could output
# 1. Whole of data in tax year
# 2. All deposits
# 3. Subset of deposits that are income. Sum total!
# 4. Expenses. Sum Total!

# Curious as to whether can automate driving expenses e.g. here's a list of destinations...go calculate mileage'




# Program starts here...

# print files in dir and check working directory
files = os.listdir("HSBC Transaction Data (Tax)")
print(files)
print("\n")
print(os.getcwd())
print("\n" * 3)

# upload csv data to DataFrame (df) using a path variable, print path
path = os.path.join("HSBC Transaction Data (Tax)", "TransactionHistory.csv")
print(path)
df = pd.read_csv(path)

# Check shape of data 
print(df.shape)

# Add column names
df.columns = ['Date', 'Ref', 'Transaction']

# Change datatypes to aid calculations and check head and dtypes to confirm
##              df['Date'].astype('datetime')
df["Transaction"] = df["Transaction"].str.replace(",", "").astype(float)
print('\n')
print(df.dtypes)









# create instance of Sheet for income, 2024 tax year end
income = Sheet(2024, df)

# Store incoming transactions as credits variable using one of Tax class methods 
credits = income.subset_credit()
income.subset_credit()
print(credits)
print("\n")


# Further subset credits to exclude transaction involving myself or family
##credits_not_Dowling = tax_2024.subset_ref('Dowling', notword=True)
income.subset_refs('Gregory', 'Dowling', 'wix', 'Alex', '400810', 'EUI', notword=True)
print(income.subset_accounts)
print("\n")
print(income.subset_accounts['Transaction'].sum())

# Further subset over a certain transaction size
income.subset_transaction(lower=100)
# Calculate sum
income.subset_total()

print(income.subset_accounts)



# potential expenses
expenses = Sheet(2024, df)

# subset refs
expenses.subset_refs('presto', 'Arwel', 'premier inn')
#'trainline', 'presto', 'amazon', 'wix', 'Arwel', 'Roberts', 'trl'
expenses.subset_transaction(lower=-1000000, upper=0)
##expenses.subset_refs('Gregory', 'Dowling', notword=True)
expenses.subset_total()

print("\n", expenses.subset_accounts)




# money to and from apple
apple = Sheet(2024, df)

apple.subset_refs('apple')
apple.subset_transaction()
apple.subset_total()

print("\n", apple.subset_accounts, apple.subset_accounts['Transaction'].sum())


# money to and from amazon
amazon = Sheet(2024, df)
amazon.subset_refs('amazon').subset_transaction().subset_total()

print("\n", amazon.subset_accounts)


other = Sheet(2024, df)
other.subset_refs("EUI", 'wix', "CHQ", "booking.com", "dart")
print(other.subset_accounts)

full_accounts = Sheet(2024, df)
full_accounts.subset_accounts = full_accounts.accounts

positive = Sheet(2024, df).subset_credit()
positive.subset_accounts = positive.subset_accounts.sort_values(by="Transaction", ascending=False)

tfl = Sheet(2024, df)
tfl.subset_refs('tfl').subset_total()
print(tfl.subset_accounts)

# Create Workbook from Sheet objects and export
sheets={'income':income, 'expenses': expenses, 'apple': apple, 'amazon': amazon, 'other':other, 'positive':positive, 'tfl': tfl, 'full_accounts':full_accounts}
workbook = Workbook(sheets)
print(workbook.sheets)
workbook.export(filename='example7')


