# This is a sample Python script.
import openpyxl as openpyxl
# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd
import pip

# Install openpyxl for reading .xlsx files
# !pip install openpyxl

# Read the .xlsx file
file_path = "exterminator_data.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')



print("Column names:")
print(df.columns)

# Fill missing 'Tech Assigned' values with the most recent non-missing value
df['Tech Assigned'] = df['Tech Assigned'].fillna(method='ffill')

# Create a dictionary to store the count of new regular accounts per tech
tech_new_regular_accounts = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    regular_service = row['Customer Elected Regular Service']

    # Check if the customer elected regular service
    if regular_service == 'Yes':
        # Increment the count of new regular accounts for the technician
        if technician in tech_new_regular_accounts:
            tech_new_regular_accounts[technician] += 1
        else:
            tech_new_regular_accounts[technician] = 1

# Print the new regular accounts sold per tech
print("New Regular Accounts Sold per Tech:")
for technician, count in tech_new_regular_accounts.items():
    print(f"{technician}: {count}")

tech_completed_regular_accounts = {}


for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']
    service_type = row['Service Type']

    # Check if the lead status is "Completed" and the service type is "Recurring"
    if lead_status == '8-Completed' and service_type == 'Recurring':
        # Increment the count of completed regular accounts for the technician
        if technician in tech_completed_regular_accounts:
            tech_completed_regular_accounts[technician] += 1
        else:
            tech_completed_regular_accounts[technician] = 1

# Print the completed regular accounts per tech
print("Completed Reg Accounts per Tech:")
for technician, count in tech_completed_regular_accounts.items():
    print(f"{technician}: {count}")

# Create a dictionary to store the count of canceled regular accounts per tech
tech_canceled_regular_accounts = {}

# Create a dictionary to store the name and address of canceled customers per tech
tech_canceled_customer_info = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    service_type = row['Service Type']
    lead_status = row['Lead Status']
    first_name = row['First Name']
    last_name = row['Last Name']
    street = row['Street']

    # Check if the service type is "Recurring" and the lead status is "Canceled"
    if service_type == 'Recurring' and lead_status == '11-Canceled':
        # Increment the count of canceled regular accounts for the technician
        if technician in tech_canceled_regular_accounts:
            tech_canceled_regular_accounts[technician] += 1
        else:
            tech_canceled_regular_accounts[technician] = 1

        # Add the name and address of the canceled customer to the dictionary
        if technician in tech_canceled_customer_info:
            tech_canceled_customer_info[technician].append((f"{first_name} {last_name}", street))
        else:
            tech_canceled_customer_info[technician] = [(f"{first_name} {last_name}", street)]

# Print the canceled regular accounts per tech and the corresponding customer info
print("Canceled Reg Accounts per Tech:")
for technician, count in tech_canceled_regular_accounts.items():
    print(f"{technician}: {count}")
    print("Canceled customer information:")
    for name, address in tech_canceled_customer_info[technician]:
        print(f"  Name: {name}, Address: {address}")

# Create a dictionary to store the count of canceled jobs per tech
tech_canceled_jobs = {}

# Create a dictionary to store the name and address of canceled customers per tech
tech_canceled_job_customer_info = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    service_type = row['Service Type']
    lead_status = row['Lead Status']
    first_name = row['First Name']
    last_name = row['Last Name']
    street = row['Street']

    # Check if the lead status is "Canceled" and the service type is not "Recurring"
    if lead_status == '11-Canceled' and service_type != 'Recurring':
        # Increment the count of canceled jobs for the technician
        if technician in tech_canceled_jobs:
            tech_canceled_jobs[technician] += 1
        else:
            tech_canceled_jobs[technician] = 1

        # Add the name and address of the canceled customer to the dictionary
        if technician in tech_canceled_job_customer_info:
            tech_canceled_job_customer_info[technician].append((f"{first_name} {last_name}", street))
        else:
            tech_canceled_job_customer_info[technician] = [(f"{first_name} {last_name}", street)]

# Print the canceled jobs per tech and the corresponding customer info
print("Canceled Jobs per Tech:")
for technician, count in tech_canceled_jobs.items():
    print(f"{technician}: {count}")
    print("Canceled customer information:")
    for name, address in tech_canceled_job_customer_info[technician]:
        print(f"  Name: {name}, Address: {address}")

