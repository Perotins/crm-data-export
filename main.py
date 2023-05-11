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

# Create a dictionary to store the count of proposals per tech
tech_proposals_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']

    # Check if the criteria for a proposal are met
    if lead_status == '3P - Pending Prop':
        # Increment the count of proposals for the technician
        if technician in tech_proposals_count:
            tech_proposals_count[technician] += 1
        else:
            tech_proposals_count[technician] = 1

# Print the count of proposals per tech
print("Number of Proposals per Tech:")
for technician, count in tech_proposals_count.items():
    print(f"{technician}: {count}")

# Create dictionaries to store the count of No Contacts and Passes per tech
tech_no_contacts_count = {}
tech_passes_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']

    # Check if the criteria for a No Contact or Pass are met
    if lead_status in ['2-No Contact', '9-Pass']:
        if lead_status == '2-No Contact':
            # Increment the count of No Contacts for the technician
            if technician in tech_no_contacts_count:
                tech_no_contacts_count[technician] += 1
            else:
                tech_no_contacts_count[technician] = 1
        elif lead_status == '9-Pass':
            # Increment the count of Passes for the technician
            if technician in tech_passes_count:
                tech_passes_count[technician] += 1
            else:
                tech_passes_count[technician] = 1

# Print the count of No Contacts and Passes per tech
print("Number of No Contacts and Passes per Tech:")
for technician in set(tech_no_contacts_count.keys()).union(tech_passes_count.keys()):
    no_contacts_count = tech_no_contacts_count.get(technician, 0)
    passes_count = tech_passes_count.get(technician, 0)
    print(f"{technician}: {no_contacts_count} No Contacts / {passes_count} Passes")

# Create a dictionary to store the count of other source calls per tech
tech_other_source_calls_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    how_heard = row['How did you hear about us']

    # Check if the criteria for an other source call is met
    if pd.isna(how_heard):
        # Increment the count of other source calls for the technician
        if technician in tech_other_source_calls_count:
            tech_other_source_calls_count[technician] += 1
        else:
            tech_other_source_calls_count[technician] = 1

# Print the count of other source calls per tech
print("Number of Other Source Calls per Tech:")
for technician, count in tech_other_source_calls_count.items():
    print(f"{technician}: {count}")

# Create a dictionary to store the count of completed non-recurring, non-zero-charge leads per tech
tech_completed_non_recurring_non_zero_charge_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']
    service_type = row['Service Type']
    service_charge = row['Service Charge']

    # Check if the criteria for a completed non-recurring, non-zero-charge lead are met
    if lead_status == '8-Completed' and service_type != 'Recurring' and service_charge != 0:
        # Increment the count for the technician
        if technician in tech_completed_non_recurring_non_zero_charge_count:
            tech_completed_non_recurring_non_zero_charge_count[technician] += 1
        else:
            tech_completed_non_recurring_non_zero_charge_count[technician] = 1

# Create a dictionary to store the number of completed leads per tech
tech_completed_leads_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']

    # Check if the lead status is "8 - Complete"
    if lead_status == '8-Completed':
        # Increment the count of completed leads for the technician
        if technician in tech_completed_leads_count:
            tech_completed_leads_count[technician] += 1
        else:
            tech_completed_leads_count[technician] = 1


# Create a dictionary to store the percentage sold per tech
tech_sold_percentage = {}

# Iterate over the technicians
for technician in tech_new_calls_count.keys():
    # Calculate the percentage sold per tech
   # print("%d %d new calls for %s", tech_new_calls_count[technician], tech_completed_leads_count[technician])
    tech_sold_percentage[technician] = (tech_new_calls_count[technician] / tech_completed_leads_count[technician]) * 100

# Print the percentage sold per tech
for tech, percentage in tech_sold_percentage.items():
    print(f'Technician: {tech}, % sold: {percentage}%')

    # # Calculate the percentage of calls sold per tech
    # print("Percentage of Calls Sold per Tech:")
    # for technician in set(tech_new_calls_count.keys()).union(tech_completed_non_recurring_non_zero_charge_count.keys()):
    #     new_calls_count = tech_new_calls_count.get(technician, 0)
    #     completed_count = tech_completed_non_recurring_non_zero_charge_count.get(technician, 0)
    #     if new_calls_count == 0:  # Avoid division by zero
    #         percentage_sold = 0
    #     else:
    #         #percentage_sold = (new_calls_count / completed_count) * 100
    #         percentage_sold = (completed_count / new_calls_count) * 100
    #     print(f"{technician}: {percentage_sold:.2f}%")

    # Create a dictionary to store the count of completed jobs per tech and the total service charge collected per tech
    tech_completed_jobs_count = {}
    tech_total_service_charge = {}

    # Iterate over the rows of the DataFrame
    for index, row in df.iterrows():
        technician = row['Tech Assigned']
        lead_status = row['Lead Status']
        service_charge = row['Service Charge']

        # Check if the lead status is "Completed"
        if lead_status == '8-Completed':
            # Increment the count of completed jobs for the technician
            if technician in tech_completed_jobs_count:
                tech_completed_jobs_count[technician] += 1
            else:
                tech_completed_jobs_count[technician] = 1

            # Add the service charge to the total service charge collected by the technician
            if technician in tech_total_service_charge:
                tech_total_service_charge[technician] += service_charge
            else:
                tech_total_service_charge[technician] = service_charge

    # Print the count of completed jobs per tech and the total service charge collected per tech
print("Total Number of Jobs Completed per Tech:")
for technician, count in tech_completed_jobs_count.items():
    print(f"{technician}: {count}")

print("\nTotal/Gross Amount Collected per Tech:")
for technician, total in tech_total_service_charge.items():
    print(f"{technician}: ${total:.2f}")

    # Calculate and print the total number of completed jobs by summing up the counts for each technician
total_jobs_completed = sum(tech_completed_jobs_count.values())
print(f"\nTotal Number of Jobs Completed (by all technicians): {total_jobs_completed}")

# Create dictionaries to store the total amount billed/bartered per tech and the count of leads with status "1-Initial" per tech
tech_billed_amount = {}
tech_initial_leads_count = {}

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    technician = row['Tech Assigned']
    lead_status = row['Lead Status']
    service_charge = row['Service Charge']
    payment_method = row['Check, Invoice, or Reference Info']

    # Check if the payment method is "Invoice"
    if payment_method == 'To be invoiced':
        # Add the service charge to the total amount billed/bartered by the technician
        if technician in tech_billed_amount:
            tech_billed_amount[technician] += service_charge
        else:
            tech_billed_amount[technician] = service_charge

    # Check if the lead status is "1-Initial"
    if lead_status == '1-Initial':
        # Increment the count of leads with status "1-Initial" for the technician
        if technician in tech_initial_leads_count:
            tech_initial_leads_count[technician] += 1
        else:
            tech_initial_leads_count[technician] = 1

# Print the total amount billed/bartered per tech
for tech, amount in tech_billed_amount.items():
    print(f'Technician: {tech}, Total amount billed/bartered: ${amount}')

# Print the count of leads with status "1-Initial" per tech
for tech, count in tech_initial_leads_count.items():
    print(f'Technician: {tech}, # of leads as initials: {count}')
