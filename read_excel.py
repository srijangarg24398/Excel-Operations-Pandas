import pandas as pd
import numpy as np

df=pd.read_excel("excel-comp-data.xlsx")
df.head()
df['Total']=df['Jan']+df['Feb']+df['Mar']
# print(df.head(10))

# Performing column level analysis
print (df["Jan"].sum(), df["Jan"].mean(),df["Jan"].min(),df["Jan"].max() )

# Calculating the sum of the columns
sum_row = df[["Jan","Feb","Mar","Total"]].sum()
print(sum_row)

# Convert data from row-based to column-based so that it can be easily integrated into the excel
sum_col = pd.DataFrame(data=sum_row).T
sum_col = sum_col.reindex(columns=df.columns)
print(sum_col)

df_final=df.append(sum_col,ignore_index=True)
print (df_final.tail() )

from fuzzywuzzy import fuzz
from fuzzywuzzy import process
state_to_code = {"VERMONT": "VT", "GEORGIA": "GA", "IOWA": "IA", "Armed Forces Pacific": "AP", "GUAM": "GU",
                 "KANSAS": "KS", "FLORIDA": "FL", "AMERICAN SAMOA": "AS", "NORTH CAROLINA": "NC", "HAWAII": "HI",
                 "NEW YORK": "NY", "CALIFORNIA": "CA", "ALABAMA": "AL", "IDAHO": "ID", "FEDERATED STATES OF MICRONESIA": "FM",
                 "Armed Forces Americas": "AA", "DELAWARE": "DE", "ALASKA": "AK", "ILLINOIS": "IL",
                 "Armed Forces Africa": "AE", "SOUTH DAKOTA": "SD", "CONNECTICUT": "CT", "MONTANA": "MT", "MASSACHUSETTS": "MA",
                 "PUERTO RICO": "PR", "Armed Forces Canada": "AE", "NEW HAMPSHIRE": "NH", "MARYLAND": "MD", "NEW MEXICO": "NM",
                 "MISSISSIPPI": "MS", "TENNESSEE": "TN", "PALAU": "PW", "COLORADO": "CO", "Armed Forces Middle East": "AE",
                 "NEW JERSEY": "NJ", "UTAH": "UT", "MICHIGAN": "MI", "WEST VIRGINIA": "WV", "WASHINGTON": "WA",
                 "MINNESOTA": "MN", "OREGON": "OR", "VIRGINIA": "VA", "VIRGIN ISLANDS": "VI", "MARSHALL ISLANDS": "MH",
                 "WYOMING": "WY", "OHIO": "OH", "SOUTH CAROLINA": "SC", "INDIANA": "IN", "NEVADA": "NV", "LOUISIANA": "LA",
                 "NORTHERN MARIANA ISLANDS": "MP", "NEBRASKA": "NE", "ARIZONA": "AZ", "WISCONSIN": "WI", "NORTH DAKOTA": "ND",
                 "Armed Forces Europe": "AE", "PENNSYLVANIA": "PA", "OKLAHOMA": "OK", "KENTUCKY": "KY", "RHODE ISLAND": "RI",
                 "DISTRICT OF COLUMBIA": "DC", "ARKANSAS": "AR", "MISSOURI": "MO", "TEXAS": "TX", "MAINE": "ME"}

sample1 = process.extractOne("Minnesotta",choices=state_to_code.keys())
print(sample1)
sample2 = process.extractOne("AlaBAMMazzz",choices=state_to_code.keys(),score_cutoff=80)
print(sample2)

def convert_state(row):
    if pd.notnull(row['state']): 
        abbrev = process.extractOne(row["state"],choices=state_to_code.keys(),score_cutoff=80)
        if abbrev:
            return state_to_code[abbrev[0]]
            return np.nan
        else:
            return np.nan

df_final.insert(6,"abbrev",np.nan)
print(df_final.head())
df_final['abbrev'] = df_final.apply(convert_state, axis=1)
print(df_final.tail())

# Calculate sub totals
df_sub=df_final[["abbrev","Jan","Feb","Mar","Total"]].groupby('abbrev').sum()
print(df_sub)

def money(x):
    return "${:,.0F}".format(x)

formatted_df=df_sub.applymap(money)
print(formatted_df)

sum_row=df_sub[["Jan","Feb","Mar","Total"]].sum()
print(sum_row)
df_sub_sum = pd.DataFrame(data=sum_row).T
df_sub_sum = df_sub_sum.applymap(money)
print(df_sub_sum)

final_table=formatted_df.append(df_sub_sum)
final_table=final_table.rename(index={0:"Total"})
print(final_table)
