import pandas as pd
import re
from collections import Counter

# Read Excel file:
submissions = pd.read_excel('data/abyss_submissions.xlsx')
print("Submissions loaded.")

# Create dictionary to store First Half and Second Half counts and constellations
unit_data = {'First Half': {}, 'Second Half': {}}

# Create lists to store teams
first_half_teams = []
second_half_teams = []
all_teams = []

# Function to pull character strings and constellations:
def extract_units(text, half):
    lines = text.strip().split('\n')
    team = []
    for line in lines:
        if '`' in line:  # Ensure there is a unit in the line
            match = re.search(r'C(\d+) `([^`]+)`', line)
            if match:
                constellation = int(match.group(1))
                unit = match.group(2)
                if unit in unit_data[half]:
                    unit_data[half][unit]['count'] += 1
                    unit_data[half][unit]['constellations'].append(constellation)
                else:
                    unit_data[half][unit] = {'count': 1, 'constellations': [constellation]}
                team.append(unit)
    if len(team) == 4:
        if half == 'First Half':
            first_half_teams.append(tuple(sorted(team)))  # Sort units to ignore order and store as a tuple
        elif half == 'Second Half':
            second_half_teams.append(tuple(sorted(team)))  # Sort units to ignore order and store as a tuple
        all_teams.append(tuple(sorted(team)))  # Add to all_teams as well

# Iterate through submissions:
for cell in submissions.values.flatten():
    cell_str = str(cell)
    first_half_idx = cell_str.find('**First Half**')
    second_half_idx = cell_str.find('**Second Half**')
    if first_half_idx != -1 and second_half_idx != -1:
        first_half_text = cell_str[first_half_idx + len('**First Half**'):second_half_idx].strip()
        second_half_text = cell_str[second_half_idx + len('**Second Half**'):].strip()
        extract_units(first_half_text, 'First Half')
        extract_units(second_half_text, 'Second Half')

# Function to calculate average constellation and round to nearest whole number
def calculate_average(constellations):
    return round(sum(constellations) / len(constellations)) if constellations else 0

# Create dataframes for each half
first_half_data = []
second_half_data = []
overall_data = {}

for unit, data in unit_data['First Half'].items():
    avg_constellation = calculate_average(data['constellations'])
    first_half_data.append([unit, data['count'], avg_constellation])
    if unit in overall_data:
        overall_data[unit]['count'] += data['count']
        overall_data[unit]['constellations'].extend(data['constellations'])
    else:
        overall_data[unit] = {'count': data['count'], 'constellations': data['constellations']}

for unit, data in unit_data['Second Half'].items():
    avg_constellation = calculate_average(data['constellations'])
    second_half_data.append([unit, data['count'], avg_constellation])
    if unit in overall_data:
        overall_data[unit]['count'] += data['count']
        overall_data[unit]['constellations'].extend(data['constellations'])
    else:
        overall_data[unit] = {'count': data['count'], 'constellations': data['constellations']}

first_half_df = pd.DataFrame(first_half_data, columns=['Unit', 'Count', 'Average Constellation']).sort_values(by='Count', ascending=False)
second_half_df = pd.DataFrame(second_half_data, columns=['Unit', 'Count', 'Average Constellation']).sort_values(by='Count', ascending=False)

# Prepare overall data for the new DataFrame
overall_data_list = []
for unit, data in overall_data.items():
    avg_constellation = calculate_average(data['constellations'])
    overall_data_list.append([unit, data['count'], avg_constellation])

overall_df = pd.DataFrame(overall_data_list, columns=['Unit', 'Count', 'Average Constellation']).sort_values(by='Count', ascending=False)

# Count the occurrences of each unique team and create DataFrames
first_half_team_counts = Counter(first_half_teams)
first_half_team_data = [{'Team': ', '.join(team), 'Count': count} for team, count in first_half_team_counts.items()]
first_half_team_df = pd.DataFrame(first_half_team_data).sort_values(by='Count', ascending=False)

second_half_team_counts = Counter(second_half_teams)
second_half_team_data = [{'Team': ', '.join(team), 'Count': count} for team, count in second_half_team_counts.items()]
second_half_team_df = pd.DataFrame(second_half_team_data).sort_values(by='Count', ascending=False)

all_team_counts = Counter(all_teams)
all_team_data = [{'Team': ', '.join(team), 'Count': count} for team, count in all_team_counts.items()]
all_team_df = pd.DataFrame(all_team_data).sort_values(by='Count', ascending=False)

# Save to CSV
first_half_df.to_csv('output/first_half.csv', index=False)
second_half_df.to_csv('output/second_half.csv', index=False)
overall_df.to_csv('output/overall_usage.csv', index=False)
first_half_team_df.to_csv('output/first_half_teams.csv', index=False)
second_half_team_df.to_csv('output/second_half_teams.csv', index=False)
all_team_df.to_csv('output/all_teams.csv', index=False)

print("First Half CSV created: first_half.csv")
print("Second Half CSV created: second_half.csv")
print("Overall Usage CSV created: overall_usage.csv")
print("First Half Teams CSV created: first_half_teams.csv")
print("Second Half Teams CSV created: second_half_teams.csv")
print("All Teams CSV created: all_teams.csv")

## CREATE SUMMARY FILE:
# Load the CSV files
overall_df = pd.read_csv('output/overall_usage.csv')
first_half_df = pd.read_csv('output/first_half.csv')
second_half_df = pd.read_csv('output/second_half.csv')
all_teams_df = pd.read_csv('output/all_teams.csv')
first_half_teams_df = pd.read_csv('output/first_half_teams.csv')
second_half_teams_df = pd.read_csv('output/second_half_teams.csv')

# Open a text file to write the summary
with open('output/abyss_highlights.txt', 'w') as file:
    # Write the header
    file.write("**Abyss Submission Highlights**\n")
    file.write("Note this may be inaccurate due to improper submissions, small sample size, and error in code. \n")
    file.write("Constellation values are average based on submissions. \n \n" )
    
    # Top 4 overall units
    file.write("**Top 4 units**\n")
    top_4_overall = overall_df.head(4)
    for index, row in top_4_overall.iterrows():
        file.write(f"{row['Unit']} - C{row['Average Constellation']}\n")
    file.write("\n")
    
    # Top 4 First-Half units
    file.write("**Top 4 First-Half units**\n")
    top_4_first_half = first_half_df.head(4)
    for index, row in top_4_first_half.iterrows():
        file.write(f"{row['Unit']} - C{row['Average Constellation']}\n")
    file.write("\n")
    
    # Top 4 Second-Half units
    file.write("**Top 4 Second-Half units**\n")
    top_4_second_half = second_half_df.head(4)
    for index, row in top_4_second_half.iterrows():
        file.write(f"{row['Unit']} - C{row['Average Constellation']}\n")
    file.write("\n")
    
    # Top 3 teams overall
    file.write("**Top 3 teams overall**\n")
    top_3_teams_overall = all_teams_df.head(3)
    for index, row in top_3_teams_overall.iterrows():
        file.write(f"{row['Team']} \n")
    file.write("\n")
    
    # Top 3 teams First-Half
    file.write("**Top 3 teams First-Half**\n")
    top_3_teams_first_half = first_half_teams_df.head(3)
    for index, row in top_3_teams_first_half.iterrows():
        file.write(f"{row['Team']} \n")
    file.write("\n")
    
    # Top 3 teams Second-Half
    file.write("**Top 3 teams Second-Half**\n")
    top_3_teams_second_half = second_half_teams_df.head(3)
    for index, row in top_3_teams_second_half.iterrows():
        file.write(f"{row['Team']} \n")

print("Summary text file created: abyss_highlights.txt")
