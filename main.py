# This is a sample Python script.
import openpyxl as openpyxl


import pandas as pd
import pip

# Install openpyxl for reading .xlsx files

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

# Create a dictionary to store the total number of Lps/RTU used per tech
tech_total_lps_rtu = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    product_used = row['Product Used']
    amount_used = row['Amount Used']

    # Check if the product used is "Rodent RTU" or "Rodent LP"
    if product_used == 'Rodent RTU' or product_used == 'Rodent LP':
        # Initialize the technician's dictionary if not present
        if technician not in tech_total_lps_rtu:
            tech_total_lps_rtu[technician] = {'LP': 0, 'RTU': 0}

        # Add the amount used to the appropriate product count
        if product_used == 'Rodent RTU':
            tech_total_lps_rtu[technician]['RTU'] += amount_used
        else:
            tech_total_lps_rtu[technician]['LP'] += amount_used

# Print the total number of Lps/RTU used per tech
print("Total # of Lps/RTU Used per Tech:")
for technician, products in tech_total_lps_rtu.items():
    lp_count = products['LP']
    rtu_count = products['RTU']
    print(f"{technician}: {lp_count}LP, {rtu_count}RTU")

# Create a dictionary to store the count of product used for pest jobs per tech
tech_pest_job_product_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    main_pest_problem = row['Main Pest Problem']
    product_used = row['Product Used']

    # Check if the main pest problem is not "Wildlife" and the product used is not empty
    if main_pest_problem != 'Wildlife' and pd.notna(product_used):
        # Increment the count of product used for pest jobs for the technician
        if technician in tech_pest_job_product_count:
            tech_pest_job_product_count[technician] += 1
        else:
            tech_pest_job_product_count[technician] = 1

# Print the count of product used for pest jobs per tech
print("Product Used for Pest Job per Tech:")
for technician, count in tech_pest_job_product_count.items():
    print(f"{technician}: {count}")

# Create a dictionary to store the count of product used for wildlife jobs per tech
tech_wildlife_job_product_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    main_pest_problem = row['Main Pest Problem']
    product_used = row['Product Used']

    # Check if the main pest problem is "Wildlife" and the product used is not empty
    if main_pest_problem == 'Wildlife' and pd.notna(product_used):
        # Increment the count of product used for wildlife jobs for the technician
        if technician in tech_wildlife_job_product_count:
            tech_wildlife_job_product_count[technician] += 1
        else:
            tech_wildlife_job_product_count[technician] = 1

# Print the count of product used for wildlife jobs per tech
print("Product Used for Wildlife Job per Tech:")
for technician, count in tech_wildlife_job_product_count.items():
    print(f"{technician}: {count}")

# Create a dictionary to store the count of new calls per tech
tech_new_calls_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    service_type = row['Service Type']
    lead_status = row['Lead Status']
    hear_about_us = row['How did you hear about us']

    # Check if the criteria for a new call are met
    if (service_type == 'One-Time'
            and lead_status not in ['4- Dave assist', '7-Reg Account', '7C- Inactive Regs', '7H- Reg On Hold',
                                    '7T- Transferred', '7X- Owes $$ for Regs', '10-NQC']
            and hear_about_us != 'Other Source Jobs (not from SF)'
            and technician not in ['Kenny', 'Epcs Commercial']):

        # Increment the count of new calls for the technician
        if technician in tech_new_calls_count:
            tech_new_calls_count[technician] += 1
        else:
            tech_new_calls_count[technician] = 1

# Print the count of new calls per tech
print("Number of New Calls per Tech:")
for technician, count in tech_new_calls_count.items():
    print(f"{technician}: {count}")

# Create dictionaries to store the count of Follow Ups and Retreats per tech
tech_follow_ups_count = {}
tech_retreats_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']
    service_type = row['Service Type']

    # Check if the criteria for a Follow Up or Retreat are met
    if lead_status == '8-Completed' and service_type in ['Follow-Up', 'Retreat']:
        if service_type == 'Follow-Up':
            # Increment the count of Follow Ups for the technician
            if technician in tech_follow_ups_count:
                tech_follow_ups_count[technician] += 1
            else:
                tech_follow_ups_count[technician] = 1
        elif service_type == 'Retreat':
            # Increment the count of Retreats for the technician
            if technician in tech_retreats_count:
                tech_retreats_count[technician] += 1
            else:
                tech_retreats_count[technician] = 1

# Print the count of Follow Ups and Retreats per tech
print("Number of Follow Ups and Retreats per Tech:")
for technician in set(tech_follow_ups_count.keys()).union(tech_retreats_count.keys()):
    follow_ups_count = tech_follow_ups_count.get(technician, 0)
    retreats_count = tech_retreats_count.get(technician, 0)
    print(f"{technician}: {follow_ups_count} Follow Ups / {retreats_count} Retreats")

# Create a dictionary to store the count of scheduled jobs per tech
tech_scheduled_jobs_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']
    service_type = row['Service Type']

    # Check if the criteria for a scheduled job are met
    if lead_status == 'Scheduled' and service_type != 'Recurring':
        # Increment the count of scheduled jobs for the technician
        if technician in tech_scheduled_jobs_count:
            tech_scheduled_jobs_count[technician] += 1
        else:
            tech_scheduled_jobs_count[technician] = 1

# Print the count of scheduled jobs per tech
print("Number of Scheduled Jobs per Tech:")
for technician, count in tech_scheduled_jobs_count.items():
    print(f"{technician}: {count}")
