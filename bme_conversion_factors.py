import pandas as pd
print("hello world")

file_path = "C:/Users/m-i025352/Documents/PyPowerBI/MaterialConversionData.xlsx"
conversion_path = "C:/Users/m-i025352/Documents/PyPowerBI/ConversionValues.xlsx"
output_file_path = "C:/Users/m-i025352/Documents/PyPowerBI/processed_data.xlsx"

df = pd.read_excel(file_path)

print("Initial DataFrame:")
print(df.head())

# Initialize a dictionary to store conversion factors
conversion_factors = []
conversion_targets = []

# Define a function to perform the conversions
def convert_units(material_df):
    bme = material_df['BME'].iloc[0]
    
    if bme == 'ST':
           for _, row in material_df.iterrows():
            if row['AME'] != 'ST':
                conversion_factors.append({
                    'Material': row['Material'],
                    'AME': row['AME'],
                    'Conversion_Factor': row['Nenner'] / row['Zähler'],
                    'Conversion_Target': 'ST'
                })
    else:
        if 'ST' in material_df['AME'].values:
            st_row = material_df[material_df['AME'] == 'ST'].iloc[0]
            conversion_factors.append({
                'Material': material_df['Material'].iloc[0],
                'AME': material_df['AME'].iloc[0],
                'Conversion_Factor': st_row['Nenner'] / st_row['Zähler'],
                'Conversion_Target': 'ST'
            })
        else:
            for _, row in material_df.iterrows():
                if row['Nenner'] == 1 and row['Zähler'] == 1:
                    conversion_factors.append({
                        'Material': material_df['Material'].iloc[0],
                        'AME': material_df['AME'].iloc[0],
                        'Conversion_Factor': 1,
                        'Conversion_Target': bme
                    })
                else:
                    conversion_factors.append({
                        'Material': material_df['Material'].iloc[0],
                        'AME': material_df['AME'].iloc[0],
                        'Conversion_Factor': row['Nenner'] / row['Zähler'],
                        'Conversion_Target': bme
                    })
# Process each material
for material in df['Material'].unique():
    material_df = df[df['Material'] == material]
    
    if len(material_df) > 1:
        convert_units(material_df)

# # Print the conversion factors for inspection
# print("\nConversion Factors:")
# for k, v in conversion_factors.items():
#     print(f"{k}: {v}")

# Convert the conversion_factors dictionary to a DataFrame
conversion_df = pd.DataFrame(conversion_factors)

# Print the conversion factors for inspection
print("\nConversion Factors DataFrame:")
print(conversion_df.head())

# Save the conversion factors to a new Excel file
conversion_df.to_excel(conversion_path, index=False)
print("\nConversion factors have been saved to:", conversion_path)


# # Apply conversion factors to the DataFrame
# df['Converted Value'] = df.apply(lambda row: row['L/B'] * conversion_factors.get((row['Material'], row['AME']), 1), axis=1)

# # Print the updated DataFrame
# print("\nUpdated DataFrame:")
# print(df.head())

# # Save the processed data to a new Excel file
# df.to_excel(output_file_path, index=False)

# print("\nProcessed data has been saved to:", output_file_path)