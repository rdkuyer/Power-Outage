# Power-Outage
## Objective  
This project was created as part of Maven's Power Outage Challenge. The objective was to act as a Senior Analytics Consultant hired by the U.S. Department of Energy to clean data and create a report/dashboard showcasing trends and key insights from 20 years' worth of event-level power outage data. The data provided was raw, making data cleaning a key challenge.
## Data Cleaning
-	Event date:
    -	Change to date type, correct for dates that do not include a year.
-	Type of Disturbance: 
    -	Categorize in 3 major groups: Weather-related issues, Security incidents, Equipment and Operational failures, and assign the rest to others.
-	Number of Customers Affected: 
    -Change text like ‘thousands’ and ‘millions’ to correct numbers. 
    -	Remove ranges  assume lower bound.
    - Remove text.
    - Change to numeric.
    -	Sanity check outliers, change numbers that included the date or the range got added.
-	Loss of Power:
    -	Change text like ‘thousands’ and ‘millions’ to correct numbers. 
    -	Remove ranges  assume lower bound.
    -	Remove text.
    -	Replace loss lower than 0 by absolute number.
    -	Sanity check, replace all values above 1 million by NA
-	Area affected:
    -	Change all abbreviations to full name of state.
    -	Remove town/cities
    -	Remove missing values with ‘Other’.
    -	Assumption: If multiple states listed  choose first in list.
-	NERC region: 
    -	Correct the NERC column for typos and to contain regions consistent with: https://www.nerc.com/AboutNERC/keyplayers/Pages/default.aspx. (based on the states variables created as mentioned above.
    -	When list of values, change to ‘multiple’.
-	Event Time: 
    -	Replace al missing values by 12 pm.
-	Restoration Time:
    -	Impute missing values by the median duration of an event grouped by the disturbance type and add to event add.
    -	Sanity check: replace all values where restoration occurred before event, impute as discussed above.
-	Duration:
    -	Create variable by taking difference restoration and event datetime. 
    -	Sanity check outliers: remove any duration larger than 20 days to avoid skewness.
-	Create primary keys for NERC Region, State, Disturbance Group.

## Model
I created a date table, and a table for disturbance category, NERC region and states, to be able to model a star schema for optimal performance in Power BI.
 

## Analysis
Three main branches of analysis: trends over time, regional differences, causes of event types.
### Trends over time
Show how number of reported outages, duration of events and number of customers affected per event have developed over time. 
### Regional differences
The filled map shows the states that have the highest cumulative amount of customers affected over the entire period. Texas, Nevada and California are the top three sates with approximately 14, 12 and 7.5 million customers affected respectively. Further regional analysis is provided by the utilizing quadrant analysis. Plotting average demand loss against average number of customers affected and plotting average lines divides the scatter plot into 4 quadrants where the upper right quadrant shows the states that are highest in both average demand loss and average number of customers affected. We identify seven states in this quadrant: North Carolina, Nevada, Florida, Hawaii, Michigan, Arizona, and South Carolina. Cross-filtering the report reveals that, with the exception of Nevada, all these states are disproportionately affected by weather-related incidents, exceeding the average impact.
### Event Cause
Weather-related incidents are the highest cause of power outage reports. Moreover, weather-related incidents are also the most disturbing cause of power outages with the most customers affected per event and longest duration per event, compared to the other categories. As mentioned before, the states that are most affected by power outages are also states that are heavily affected by weather incidents.
