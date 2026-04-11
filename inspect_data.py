import pandas as pd

# Load the Excel file
df = pd.read_excel('data/taxi_groups_airtransat.xlsx')

# Display data info
print("=== DATA INFO ===")
print(df.info())

print("\n=== FIRST 10 ROWS ===")
print(df.head(10))

print("\n=== TAXI NUMBER DISTRIBUTION (TOP 20) ===")
print(df['TAXI'].value_counts().head(20))

print("\n=== BASIC STATISTICS ===")
print(df.describe(include='all'))

print("\n=== UNIQUE VALUES IN CATEGORICAL COLUMNS ===")
for col in df.select_dtypes(include=['object']).columns:
    print(f"{col}: {df[col].nunique()} unique values")
    print(df[col].value_counts().head(5))
    print()