import pandas as pd
import numpy as np
import re

# Create a list of months to delete
months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]


# Define a function to clean and process each sheet
def clean_power_outage_data(sheet_index, header_row, months_to_exclude):
    # Load the Excel sheet
    df_clean = pd.read_excel(
        "C:/Users/Robert Kuyer/Data analytics/Portfolio/Power outage/power.xlsx",
        sheet_name=sheet_index,
        header=header_row,
    )

    # strip columns of whitespace
    df_clean.columns = df_clean.columns.str.strip()

    # Define the mapping of old column names to new column names
    column_mapping = {
        "Date Event Began": "Date",
        "Time Event Began": "Time",
        "Date of Restoration": "Restoration_date",
        "Time of Restoration": "Restoration_time",
        "Area Affected": "Area Affected",
        "NERC Region": "NERC Region",
        "Event Type": "Type of Disturbance",
        "Demand Loss (MW)": "Loss (megawatts)",
        "Number of Customers Affected 1": "Number of Customers Affected",
        "Number of Customers Affected 1[1]": "Number of Customers Affected",
        "Event Month": "Month",
        "Restoration Time": "Restoration_date",
        "Restoration": "Restoration_date",
        "Area": "Area Affected",
    }

    # Apply the column renaming
    df_clean = df_clean.rename(columns=column_mapping)

    # Clean the DataFrame for all missing rows and cluttered months in between rows
    df_clean = (
        df_clean[~df_clean["Date"].isin(months_to_exclude)]
        .dropna(how="all")
        .reset_index(drop=True)
    )

    return df_clean


# List of configurations for each sheet
sheet_configs = [
    {"sheet_index": 0, "header_row": 0},  # 2002
    {"sheet_index": 1, "header_row": 1},  # 2003
    {"sheet_index": 2, "header_row": 0},  # 2004
    {"sheet_index": 3, "header_row": 1},  # 2005
    {"sheet_index": 4, "header_row": 1},  # 2006
    {"sheet_index": 5, "header_row": 1},  # 2007
    {"sheet_index": 6, "header_row": 2},  # 2008
    {"sheet_index": 7, "header_row": 1},  # 2009
    {"sheet_index": 8, "header_row": 1},  # 2010
    {"sheet_index": 9, "header_row": 1},  # 2011
    {"sheet_index": 10, "header_row": 1},  # 2012
    {"sheet_index": 11, "header_row": 1},  # 2013
    {"sheet_index": 12, "header_row": 1},  # 2014
    {"sheet_index": 13, "header_row": 1},  # 2015
    {"sheet_index": 14, "header_row": 1},  # 2016
    {"sheet_index": 15, "header_row": 1},  # 2017
    {"sheet_index": 16, "header_row": 1},  # 2018
    {"sheet_index": 17, "header_row": 1},  # 2019
    {"sheet_index": 18, "header_row": 1},  # 2020
    {"sheet_index": 19, "header_row": 1},  # 2021
    {"sheet_index": 20, "header_row": 1},  # 2022
    {"sheet_index": 21, "header_row": 1},  # 2023
    # Add configurations for additional sheets as needed
]

# Process all sheets using the defined function
cleaned_dfs = {}
for config in sheet_configs:
    sheet_name = f"df_{config['sheet_index'] + 2002}"  # Naming convention, e.g., df_2002, df_2003
    cleaned_dfs[sheet_name] = clean_power_outage_data(
        config["sheet_index"], config["header_row"], months
    )
    # Change Date to type date and delete all missings
    cleaned_dfs[sheet_name]["Date"] = pd.to_datetime(
        cleaned_dfs[sheet_name]["Date"], errors="coerce"
    )
    cleaned_dfs[sheet_name] = cleaned_dfs[sheet_name].dropna(subset=["Date"])

# Add year column to each DataFrame before concatenating
for sheet_name, df in cleaned_dfs.items():
    year = sheet_name.split("_")[1]  # Extract the year from 'df_2002', 'df_2003', etc.
    df["year"] = year

# Now concatenate all DataFrames
combined_df = pd.concat(cleaned_dfs.values(), ignore_index=True)
combined_df = combined_df.drop(columns=["Event Year"])
combined_df.columns

# Change data types for restoration date and year
combined_df["Restoration_date"] = pd.to_datetime(
    combined_df["Restoration_date"], errors="coerce"
)
combined_df["year"] = combined_df["year"].astype("int")

# Change year of restoration date for wrong values
mask_year = (combined_df["year"] > 2005) & (combined_df["year"] < 2011)

# Apply the year replacement only to the filtered rows
combined_df.loc[mask_year, "Restoration_date"] = combined_df.loc[mask_year].apply(
    lambda row: row["Restoration_date"].replace(year=row["year"])
    if pd.notnull(row["Restoration_date"])
    else pd.NaT,
    axis=1,
)

# update restoration time for year before 2011
combined_df.loc[combined_df["year"] < 2011, "Restoration_time"] = combined_df[
    "Restoration_date"
].dt.time

# update restoration time for year after 2011
combined_df.loc[combined_df["year"] > 2011, "Restoration_time"] = pd.to_datetime(
    combined_df["Restoration_time"], errors="coerce", format="%H:%M:%S"
).dt.time

# update restoration date data type
combined_df["Restoration_date"] = combined_df["Restoration_date"].dt.date

# update time for 2003 - 2010
combined_df.loc[(combined_df["year"] > 2002) & (combined_df["year"] < 2011), "Time"] = (
    combined_df["Time"].str.replace(".", "").str.lower()
)
combined_df.loc[(combined_df["year"] > 2002) & (combined_df["year"] < 2011), "Time"] = (
    pd.to_datetime(combined_df["Time"], errors="coerce").dt.time
)

# Change time to type datetime time
combined_df["Time"] = pd.to_datetime(
    combined_df["Time"], errors="coerce", format="%H:%M:%S"
).dt.time

# Change years for restoration_date for 2006 - 2010
combined_df["Month"] = combined_df["Date"].dt.month
combined_df["Day"] = combined_df["Date"].dt.day
combined_df["Date"] = pd.to_datetime(combined_df[["year", "Month", "Day"]])


# Clean NERC Region
combined_df["NERC Region"] = (
    combined_df["NERC Region"]
    .str.strip()
    .str.replace(" ", "")
    .str.replace("/", ",")
    .str.replace(";", ",")
)

# Clean categories of Disturbance type
# Capitalizing the first letter of each word in the 'Type of Disturbance' column
combined_df["Type of Disturbance"] = combined_df["Type of Disturbance"].str.title()

# Define the new grouped categories
weather_related = "Thunderstorms|Heat Wave|Tropical|Highwinds|Wind|Storm|Weather|Lightning|Ice|Natural Disaster|Flood|Wildfire"
security_incidents = "Vandalism|Sabotage|Actual Physical|Physical Attack|Suspicious|Suspected Cyber Attack|Cyber Attack|Cyber Event|Cyber Threat"
equipment_operational_failures = "Equipment Failure|Faulted|Failure|Malfunction|Fault|Faulty|Unit Trip|Tripped|Transmission Interruption|System Operations|Islanding|Generation Inadequacy|Voltage Reduction|Fuel Supply|Fuel Shortage|Resource Availability"

# Conditions based on the new categories
conditions = [
    (combined_df["Type of Disturbance"].str.contains(weather_related)),
    (combined_df["Type of Disturbance"].str.contains(security_incidents)),
    (combined_df["Type of Disturbance"].str.contains(equipment_operational_failures)),
]

# New Disturbance Categories
Disturbance_categories = [
    "Weather-Related Issues",
    "Security Incidents",
    "Equipment and Operational Failures",
]

# Apply the conditions to create a new column for categorized disturbances
combined_df["Disturbance Category"] = np.select(
    conditions, Disturbance_categories, default="Other"
)

###
# Clean number of customers affected
# Create flags for wrong inputs
combined_df["million"] = combined_df["Number of Customers Affected"].str.contains(
    "million", case=False
)
combined_df["date_added"] = combined_df["Number of Customers Affected"].str.contains(
    "/"
)

# Clean
combined_df.loc[
    combined_df["Number of Customers Affected"] == "1 PG&E",
    "Number of Customers Affected",
] = "not a digit"
combined_df["Number of Customers Affected"] = (
    combined_df["Number of Customers Affected"].astype(str).str.strip()
)

# remove all text
combined_df["Number of Customers Affected"] = (
    combined_df["Number of Customers Affected"]
    .str.replace(r"[^\d.]+", "", regex=True)  # Keep digits and decimal points
    .str.replace(r"^\.+", "", regex=True)  # Remove leading dots
    .str.replace(
        r"\.(?=\.)", "", regex=True
    )  # Remove any dots that are immediately followed by another dot
    .str.replace(r"(?<=\d)\.+$", "", regex=True)  # Remove trailing dots
    .replace("", np.nan)  # Replace empty strings with NaN
)

# manually adjust failures ('million', date included in number, range of number added)
adjustments_dict = {
    # million in text
    19: 1500000,
    59: 1800000,
    228: 1100000,
    # Date included in number
    60: 320000,
    64: 530000,
    73: 104195,
    # range got added
    56: 133000,
    283: 8000,
    304: 2500000,
    413: 1175,
}

for index, adjustment in adjustments_dict.items():
    combined_df.loc[index, "Number of Customers Affected"] = adjustment

# change data type of number affected
combined_df.loc[
    combined_df["Number of Customers Affected"] == "", "Number of Customers Affected"
] = np.nan
combined_df["Number of Customers Affected"] = pd.to_numeric(
    combined_df["Number of Customers Affected"]
)
combined_df["Number of Customers Affected"] = combined_df[
    "Number of Customers Affected"
].astype("Int64")

# drop unnecessary columns
combined_df.drop(columns=["million", "date_added"], inplace=True)

###
# Clean loss of power
combined_df["Loss (megawatts)"] = combined_df["Loss (megawatts)"].astype(str)
combined_df["Loss (megawatts)"] = combined_df["Loss (megawatts)"].str.replace("to", "-")
mask = combined_df["Loss (megawatts)"].str.contains("-", na=False)
combined_df.loc[mask, "Loss (megawatts)"] = (
    combined_df.loc[mask, "Loss (megawatts)"].str.split("-").str[0]
)

# remove all text
combined_df["Loss (megawatts)"] = (
    combined_df["Loss (megawatts)"]
    .str.replace(r"[^\d.]+", "", regex=True)  # Keep digits and decimal points
    .str.replace(r"^\.+", "", regex=True)  # Remove leading dots
    .str.replace(
        r"\.(?=\.)", "", regex=True
    )  # Remove any dots that are immediately followed by another dot
    .str.replace(r"(?<=\d)\.+$", "", regex=True)  # Remove trailing dots
    .replace("", np.nan)  # Replace empty strings with NaN
)

# Change missings and data type
combined_df["Loss (megawatts)"] = pd.to_numeric(
    combined_df["Loss (megawatts)"], errors="coerce"
)
combined_df.loc[combined_df["Loss (megawatts)"] < 0, "Loss (megawatts)"] *= -1
combined_df.loc[combined_df["Loss (megawatts)"] > 1000000, "Loss (megawatts)"] = np.nan
combined_df["Loss (megawatts)"] = combined_df["Loss (megawatts)"].astype(
    "Int64", errors="ignore"
)
###
# Change NERC regions with mulitple regions covered to 'multiple'
combined_df["NERC Region"].where(
    combined_df["NERC Region"].str.split(",").str.len() == 1, "Multiple", inplace=True
)

# Clean Area Affected
states = [
    "Alabama",
    "Alaska",
    "Arizona",
    "Arkansas",
    "California",
    "Colorado",
    "Connecticut",
    "Delaware",
    "Florida",
    "Georgia",
    "Hawaii",
    "Idaho",
    "Illinois",
    "Indiana",
    "Iowa",
    "Kansas",
    "Kentucky",
    "Louisiana",
    "Maine",
    "Maryland",
    "Massachusetts",
    "Michigan",
    "Minnesota",
    "Mississippi",
    "Missouri",
    "Montana",
    "Nebraska",
    "Nevada",
    "New Hampshire",
    "New Jersey",
    "New Mexico",
    "New York",
    "North Carolina",
    "North Dakota",
    "Ohio",
    "Oklahoma",
    "Oregon",
    "Pennsylvania",
    "Rhode Island",
    "South Carolina",
    "South Dakota",
    "Tennessee",
    "Texas",
    "Utah",
    "Vermont",
    "Virginia",
    "Washington",
    "West Virginia",
    "Wisconsin",
    "Wyoming",
]

# State abbreviations
abbreviations = [
    "AL",
    "AK",
    "AZ",
    "AR",
    "CA",
    "CO",
    "CT",
    "DE",
    "FL",
    "GA",
    "HI",
    "ID",
    "IL",
    "IN",
    "IA",
    "KS",
    "KY",
    "LA",
    "ME",
    "MD",
    "MA",
    "MI",
    "MN",
    "MS",
    "MO",
    "MT",
    "NE",
    "NV",
    "NH",
    "NJ",
    "NM",
    "NY",
    "NC",
    "ND",
    "OH",
    "OK",
    "OR",
    "PA",
    "RI",
    "SC",
    "SD",
    "TN",
    "TX",
    "UT",
    "VT",
    "VA",
    "WA",
    "WV",
    "WI",
    "WY",
]

# Create a mapping of state abbreviations to full state names
abbreviation_to_full = dict(zip(abbreviations, states))

# Clean 'Area Affected'
combined_df["Area Affected"] = combined_df["Area Affected"].str.strip()

# Create a regex pattern from the states list and abbreviations
full_pattern = "|".join(states) + "|" + "|".join(abbreviations)

# Create a new column 'State Affected' to extract matched states
combined_df["State Affected"] = combined_df["Area Affected"].str.extract(
    f"({full_pattern})", expand=False
)

# Map abbreviations to full names
combined_df["State Affected"] = combined_df["State Affected"].replace(
    abbreviation_to_full
)

# Fill NaN values with 'Other' or any placeholder if needed
combined_df["State Affected"] = combined_df["State Affected"].fillna("Other")

# Create a new column 'State Affected List' to extract all matched states as a list
combined_df["State Affected List"] = combined_df["Area Affected"].str.findall(
    full_pattern
)

# Join the matched states into a single string or handle None
combined_df["State Affected List"] = combined_df["State Affected List"].apply(
    lambda x: ", ".join(x) if isinstance(x, list) and x else "Other"
)

combined_df.drop(columns=["Area Affected", "Type of Disturbance"], inplace=True)

# drop rows where restoration date is before event
rows_to_drop = combined_df[combined_df["Restoration_date"] < combined_df["Date"]].index
combined_df.drop(index=rows_to_drop, inplace=True)


# Clean NERC region based on site: https://www.nerc.com/AboutNERC/keyplayers/Pages/default.aspx
# Mapping of states to NERC regions
state_to_nerc_region = {
    "Alabama": "SERC",
    "Alaska": "WECC",
    "Arizona": "WECC",
    "Arkansas": "SERC",
    "California": "WECC",
    "Colorado": "WECC",
    "Connecticut": "NPCC",
    "Delaware": "RF",
    "Florida": "SERC",
    "Georgia": "SERC",
    "Hawaii": "WECC",
    "Idaho": "WECC",
    "Illinois": "MRO",
    "Indiana": "MRO",
    "Iowa": "MRO",
    "Kansas": "MRO",
    "Kentucky": "SERC",
    "Louisiana": "SERC",
    "Maine": "NPCC",
    "Maryland": "RF",
    "Massachusetts": "NPCC",
    "Michigan": "MRO",
    "Minnesota": "MRO",
    "Mississippi": "SERC",
    "Missouri": "MRO",
    "Montana": "WECC",
    "Nebraska": "MRO",
    "Nevada": "WECC",
    "New Hampshire": "NPCC",
    "New Jersey": "RF",
    "New Mexico": "WECC",
    "New York": "NPCC",
    "North Carolina": "SERC",
    "North Dakota": "MRO",
    "Ohio": "MRO",
    "Oklahoma": "SERC",
    "Oregon": "WECC",
    "Pennsylvania": "RF",
    "Rhode Island": "NPCC",
    "South Carolina": "SERC",
    "South Dakota": "MRO",
    "Tennessee": "SERC",
    "Texas": "TRE",
    "Utah": "WECC",
    "Vermont": "NPCC",
    "Virginia": "RF",
    "Washington": "WECC",
    "West Virginia": "RF",
    "Wisconsin": "MRO",
    "Wyoming": "WECC",
    "Other": "Other",
}

# Create a new column in the DataFrame for the NERC region
combined_df["NERC Region"] = combined_df["State Affected"].map(state_to_nerc_region)

###
# Missings
###
# impute missing Restoration datesby average grouped by distrubance category
combined_df["Restoration_date"] = pd.to_datetime(combined_df["Restoration_date"])
combined_df["Difference"] = (
    combined_df["Restoration_date"] - combined_df["Date"]
).dt.days
combined_df["median_days"] = (
    combined_df.groupby("Disturbance Category")["Difference"]
    .transform("median")
    .round()
    .astype(int)
)

combined_df.loc[combined_df["Restoration_date"].isnull(), "Restoration_date"] = (
    combined_df["Date"] + pd.to_timedelta(combined_df["median_days"], unit="D")
)

# Drop the mean_days and difference column
combined_df = combined_df.drop(columns=["median_days", "Difference"])

# Fill missing in Time variable 12:00
combined_df["Time"] = combined_df["Time"].fillna(pd.to_datetime("12:00:00").time())

# Fill missings in Restoration Time with 00:00:00 to ensure its earlier than start time, to capture these values in following step to impute average
combined_df["Restoration_time"] = combined_df["Restoration_time"].fillna(
    pd.to_datetime("01:00:00").time()
)

# Create datetime variable for start and end date to calculate duration column
combined_df["start_datetime"] = pd.to_datetime(
    combined_df["Date"].astype(str) + " " + combined_df["Time"].astype(str)
)
combined_df["end_datetime"] = pd.to_datetime(
    combined_df["Restoration_date"].astype(str)
    + " "
    + combined_df["Restoration_time"].astype(str)
)
combined_df["duration"] = combined_df["end_datetime"] - combined_df["start_datetime"]

# Convert duration to hours
combined_df["duration_hours"] = (
    combined_df["duration"].dt.total_seconds() / 3600
).round(2)

# replace all negative duration hours with nan
combined_df.loc[combined_df["duration_hours"] < 0, "duration_hours"] = np.nan

# Create median hours to impute missing values
combined_df["median_hours"] = (
    combined_df.groupby("Disturbance Category")["duration_hours"]
    .transform("median")
    .round(2)
)

# impute missings and adjust the end datetime column for values where end datettime was smaller than start datetime
combined_df.loc[combined_df["duration_hours"].isnull(), "end_datetime"] = (
    pd.to_datetime(
        combined_df["start_datetime"]
        + pd.to_timedelta(combined_df["median_hours"], unit="h")
    )
)
combined_df.loc[combined_df["duration_hours"].isnull(), "duration_hours"] = combined_df[
    "median_hours"
]

# drop unnecessary columns
combined_df.drop(columns=["duration"], inplace=True)

###
# Primary keys
# NERC id
combined_df["NERC_id"] = combined_df.groupby("NERC Region").ngroup() + 1

# State key
combined_df["state_id"] = combined_df.groupby("State Affected").ngroup() + 1

# Disturbance key
combined_df["Disturbance_id"] = combined_df.groupby("Disturbance Category").ngroup() + 1

# More than 20 days all some to be errors, thus delete
combined_df = combined_df.drop(combined_df[combined_df["duration_hours"] > 480].index)

###
# Sanity checks
# is end date smaller than start date?
assert (
    combined_df["end_datetime"] >= combined_df["start_datetime"]
).all(), "Some end_datetime values are less than start_datetime!"

# is duration always larger than 0?
assert (combined_df["duration_hours"] >= 0).all(), "Some duration hours is negative"


combined_df.groupby("State Affected")["NERC Region"].unique()

# Save cleaned dataset
combined_df.to_excel(
    "C:/Users/Robert Kuyer/Data analytics/Portfolio/Power outage/power_cleaned.xlsx",
    index=False,
)
