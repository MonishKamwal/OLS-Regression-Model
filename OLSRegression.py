import pandas as pd 
import statsmodels.api as sm 

# Load the Excel file
file_path = "data.xlsx"
sheet_name = 'Data'  # Update this if your data is in a different sheet

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Specify the columns you want to extract
columns_to_extract = ["Year", "Sectors", "RE_Share", "FCEO", "VCEO", "Board Size", "LEV","Firm Size"] 

# Extract the specific columns
extracted_data = df[columns_to_extract]
# Convert columns to lists
column_lists = {col: extracted_data[col].tolist() for col in columns_to_extract}

# Example DataFrame (replace with actual data)
# Assign column "Year" to 'Year' variable and so on..
data = pd.DataFrame({
    'Year': column_lists["Year"], 
    'Sectors': column_lists["Sectors"],  
    'RE_Share': column_lists["RE_Share"],  
    'FCEO': column_lists["FCEO"],   
    'VCEO': column_lists["VCEO"],
    #'TCEO' : column_lists["TCEO"],
    'Board Size' : column_lists["Board Size"],
    'LEV' : column_lists["LEV"],  
    'Firm Size' : column_lists["Firm Size"],
})

# Create dummy variables for the 'Sectors' column
data = pd.get_dummies(data, columns=['Sectors','Year'], drop_first=True)

# Dummy variables are in the form of "TRUE"/"FALSE" so we convert them to 0/1

# Convert boolean (True/False) dummy variables to integers (0/1)
for column in data.select_dtypes(include=['bool']).columns:
    data[column] = data[column].astype(int)

# Check for any non-numeric columns (after creating dummies, all should be numeric)
print(data.dtypes)

# Ensure all data is numeric and no missing values
data = data.apply(pd.to_numeric, errors='coerce').dropna()

# Add a constant term for the intercept
data['Intercept'] = 1

# Define the dependent and independent variables
#X = data[["Year", "Sectors", "FCEO", "VCEO", "TCEO", "Board Size", "LEV", "Sectors_FMCG", "Sectos_Consumer Durables", "Sectors_Healthcare" and so on...]]
X = data.drop(columns=['RE_Share']) # everything except RE_Share
y = data[ "RE_Share"]

# Perform the regression
model = sm.OLS(y, X).fit()

# Print the regression results
print(model.summary())
