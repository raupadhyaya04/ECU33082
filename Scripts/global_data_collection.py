import pandas as pd
import wbgapi as wb

def fetch_and_clean_global_data():
    print("Fetching Global Data from World Bank API...")
    
    # Define indicators
    # FR.INR.RINR: Real interest rate (%)
    # EG.ELC.ACCS.ZS: Access to electricity (% of population)
    # EG.CFT.ACCS.ZS: Access to clean fuels and technologies for cooking (% of population)
    
    print("Processing Interest Rate Data...")
    df_ir = wb.data.DataFrame('FR.INR.RINR', time=range(2007, 2024), numericTimeKeys=True).reset_index()
    df_ir = df_ir.melt(id_vars=['economy'], var_name='Year', value_name='Interest_Rate')
    
    print("Processing Electricity Access Data...")
    df_elec = wb.data.DataFrame('EG.ELC.ACCS.ZS', time=range(2007, 2024), numericTimeKeys=True).reset_index()
    df_elec = df_elec.melt(id_vars=['economy'], var_name='Year', value_name='Electricity_Access')
    
    print("Processing Clean Cooking Access Data...")
    df_cook = wb.data.DataFrame('EG.CFT.ACCS.ZS', time=range(2007, 2024), numericTimeKeys=True).reset_index()
    df_cook = df_cook.melt(id_vars=['economy'], var_name='Year', value_name='Clean_Cooking_Access')
    
    # Merge datasets
    merged_df = pd.merge(df_ir, df_elec, on=['economy', 'Year'], how='inner')
    merged_df = pd.merge(merged_df, df_cook, on=['economy', 'Year'], how='inner')
    merged_df.rename(columns={'economy': 'Country_Code'}, inplace=True)
    
    # Drop rows with missing values to ensure quality transitions
    merged_df.dropna(inplace=True)
    
    # Energy Poverty proxy: Inverse of both electricity and clean cooking access
    merged_df['Energy_Poverty_Elec'] = 100 - merged_df['Electricity_Access']
    merged_df['Energy_Poverty_Cook'] = 100 - merged_df['Clean_Cooking_Access']
    
    # Create a composite Energy Poverty score by averaging the lack of access to both
    merged_df['Composite_Energy_Poverty'] = (merged_df['Energy_Poverty_Elec'] + merged_df['Energy_Poverty_Cook']) / 2
    
    # Discretize into Low, Medium, High states using quantiles (tertiles)
    print("Categorizing states into Low, Medium, High...")
    
    merged_df['Interest_Rate_State'] = pd.qcut(
        merged_df['Interest_Rate'], 
        q=3, 
        labels=['Low', 'Medium', 'High']
    )
    
    merged_df['Energy_Poverty_State'] = pd.qcut(
        merged_df['Composite_Energy_Poverty'], 
        q=3, 
        labels=['Low', 'Medium', 'High']
    )
    
    output_path = "Cleaned Data/Global_Interest_Energy_Poverty_2007_2023.csv"
    merged_df.to_csv(output_path, index=False)
    print(f"Global dataset saved to {output_path}")

if __name__ == "__main__":
    fetch_and_clean_global_data()
