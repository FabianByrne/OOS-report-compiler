import pandas as pd
import os

dictname_list = list()
OOS_dict = dict()


def run_program():
    """ This starts the application by:
    1. Obtaining the file paths of the OOS result folder and the Contact details spreadsheet, then storing them in variables
    2. Passing the output of 'spreadsheet_iterator()' into 'criteria_check()'
    3. Passing the output of 'criteria check' into the 'user_interface()' function's 'returned_data' parameter """

    excelfile_paths()
    user_interface(criteria_check(spreadsheet_iterator()))

def excelfile_paths():
    # Specify the folder path containing the OOS spreadsheets and the contact info
    global OOS_file_input
    global contact_file_input
    OOS_file_input = input("\nInsert the file path for the 'OOS results folder': ")
    contact_file_input = input("Insert a file path for the 'Contact details' spreadsheet: ")

def spreadsheet_iterator():
    # Specify the folder path containing the Excel files

    folder_path = OOS_file_input

    # Iterate through all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            # Build the full file path
            file_path = os.path.join(folder_path, filename)

            # Read the Excel file into a DataFrame
            excel_file = pd.read_excel(file_path)

            main_concerns(excel_file)

    """ This 'list_and_dict' tuple stores the dictname_list and the OOS_dict as separate items in the same tuple, allowing
    them to be easily passed to the 'criteria_check' function, where they can be deconstructed again."""

    list_and_dict = (dictname_list, OOS_dict)

    return list_and_dict


def main_concerns(excel_file):
    """
    This function applies the test_checker function to each Excel file in the OOS folder.
    For each excel file, any sample with a test that is being checked for (see below) will be tacked onto the OOS_dict.
    This is because many OOS tests will never require a client to be contacted (e.g: Total Viable Counts)

    This thinned collection of samples with (potentially) important results will next have the 'criteria_check()' function
    run on them to organize the contents of the OOS_dict.

    Removing or adding fields here determine whether they will show up on the call list page, escalation page etc.

    """

    test_checker(excel_file, "Salmonella SOLUS")
    test_checker(excel_file, "Listeria SOLUS")
    test_checker(excel_file, "E.coli")
    test_checker(excel_file, "Listeria enumeration")
    test_checker(excel_file, "Listeria MALDITOF")
    test_checker(excel_file, "Salmonella MALDITOF")
    test_checker(excel_file, "Clostridium perfringens")


def test_checker(excel_file, test):
    """
    There are two items being returned from this function:

    OOS_dict:

    Here, every OOS result is recorded as a key:value pair between a 'OOS sample number + the test that failed' key and
    a nested dictionary with all the sample's information as a value
    (e.g: 'CCY1234567 Salmonella SOLUS' : {'Customer': 'CheeseCo', 'Sample Number': 'CCY1234567', 'Sample Name': 'Floor swab 1' , etc.})

    dictname_list:

    A list with all of the 'keys' (a.k.a: 'CCY1234567 Salmonella SOLUS' in the example above) is stored in a separate list.

    """

    filtered_data = excel_file[excel_file['Test'] == (test)]

    row_numbers = filtered_data.index.tolist()

    for row in row_numbers:
        customer_data = (filtered_data.at[row, "Customer"])
        sample_number = (filtered_data.at[row, "Number"])
        sampletype_data = (filtered_data.at[row, "Type"])
        name_data = (filtered_data.at[row, "Name"])
        description_data = (filtered_data.at[row, "Description"])
        test_data = (filtered_data.at[row, "Test"])
        result_data = (filtered_data.at[row, "Result"])

        nested_dict_name = str(sample_number + ' ' + test_data)

        dictname_list.append(nested_dict_name)

        OOS_dict[nested_dict_name] = {
            "Customer": customer_data,
            "Sample Number": sample_number,
            "Sample Name": name_data,
            "Sample Type": sampletype_data,
            "Description": description_data,
            "Test": test_data,
            "Result": result_data,
        }
    return (dictname_list), (OOS_dict)


def criteria_check(list_and_dict):
    """ This is where the 'name / sample number' in the dictname_list is organized into new categories based on criteria outlined below.
    Later, if more information is needed about a sample, the dictname_list entry can be used to search the disorganised OOS_dict for a
    value (the nested dictionary) corresponding to a key of the same name

    Any changing requirements can be reflected by altering the code below..."""

    dictname_list = list_and_dict[0]
    OOS_dict = list_and_dict[1]

    phone_call = []
    escalation_needed = []
    daily_path_positives = []
    daily_path_negatives = []
    IQC_EQA_list = []
    local_environmental_failures = []

    for item in dictname_list:
        name = OOS_dict.get(item).get("Sample Name")
        result = OOS_dict.get(item).get("Result")
        customer = OOS_dict.get(item).get("Customer")
        description = OOS_dict.get(item).get("Description")
        if "Salmonella SOLUS" in item:
            if customer == "Trow_lab" and ("Positive control") in description and result == "Presumptive":
                daily_path_positives.append(item)
            elif customer == "Trow_lab" and ("Positive control") in description and result == "Not Detected":
                daily_path_negatives.append(item)
            elif ("Finished product" or "finished product") in description:
                escalation_needed.append(item)
            elif customer == "Trow_lab" and ("IQC" or "EQA") in name:
                IQC_EQA_list.append(item)
            elif customer == "Trow_lab_QA":
                local_environmental_failures.append(item)
            else:
                phone_call.append(item)
        elif "Listeria SOLUS" in item:
            if customer == "Trow_lab" and ("Positive control") in description and result == "Presumptive":
                daily_path_positives.append(item)
            elif customer == "Trow_lab" and ("Positive control") in description and result == "Not Detected":
                daily_path_negatives.append(item)
            elif customer == "Trow_lab" and ("IQC" or "EQA") in name:
                IQC_EQA_list.append(item)
            elif customer == "Trow_lab_QA":
                local_environmental_failures.append(item)
            else:
                phone_call.append(item)

        elif "Listeria MALDITOF" in item:
            if customer == "Trow_lab" and ("Positive control") in description and result == "Listeria monocytogenes":
                daily_path_positives.append(item)
            elif customer == "Trow_lab" and ("Positive control") in description and result != "Listeria monocytogenes":
                daily_path_negatives.append(item)
            elif customer == "Trow_lab" and ("IQC" or "EQA") in name:
                IQC_EQA_list.append(item)
            elif customer == "Trow_lab_QA":
                local_environmental_failures.append(item)
            else:
                phone_call.append(item)
        elif "Salmonella MALDITOF" in item:
            if customer == "Trow_lab" and ("Positive control") in description and result == "Presumptive":
                daily_path_positives.append(item)
            elif customer == "Trow_lab" and ("Positive control") in description and result == "Not Detected":
                daily_path_negatives.append(item)
            elif customer == "Trow_lab" and ("IQC" or "EQA") in name:
                IQC_EQA_list.append(item)
            elif ("Finished product" or "finished product") in description:
                escalation_needed.append(item)
            elif customer == "Trow_lab_QA":
                local_environmental_failures.append(item)
            else:
                phone_call.append(item)
        elif "E.coli" in item:
            value = (result.split(' ')[0])  # Split the string and convert the first part to an integer
            if value[0] == '>':
                value = int(value[1:])
            else:
                value = int(value)
            if value > 1000:
                phone_call.append(item)
            else:
                pass
        elif "Listeria enumeration" in item:
            value = (result.split(' ')[0])
            if value[0] == '>':
                value = value[1:]
            else:
                value = int(value)

            if value > 100:
                phone_call.append(item)
            elif value > 100 and ("SOL" or "Finished Product") in OOS_dict.get(item).get("Description"):
                escalation_needed.append(item)
            else:
                pass
        elif "Clostridium perfringens" in item:
            if ("RSA" or "rsa") in description:
                phone_call.append(item)
            elif customer == "Trow_lab" and ("Positive control") in description and "<10" not in result:
                daily_path_positives.append(item)
            elif customer == "Trow_lab" and "Positive control" in description and "<10" in result:
                daily_path_negatives.append(item)
            elif customer == "Trow_lab" and ("IQC" or "EQA") in name:
                IQC_EQA_list.append(item)
            else:
                pass

    print(
        f"\nTask complete! There are {len(phone_call)} results out of specification, and {len(escalation_needed)} samples requiring escalation...\n")


    return phone_call, escalation_needed, daily_path_positives, daily_path_negatives, IQC_EQA_list, local_environmental_failures


def user_interface(returned_data):
    """ This function processes the user's input - obtained through the nested user_choice() function"""


    escape = False
    phone_call = returned_data[0]

    escalation_needed = returned_data[1]
    daily_path_positives = returned_data[2]
    daily_path_negatives = returned_data[3]
    IQC_EQA_list = returned_data[4]
    local_environmental_failures = returned_data[5]

    while escape == False:
        # user choice is the function to take user input
        user_choice()

        if user_input == 1:
            phonecall_list = []

            for item in phone_call:
                phonecall_list.append(OOS_dict.get(item).get("Customer"))

            # quickly turn it to a set and back to get only unique IDs
            phonecall_set = set(phonecall_list)
            print(f"A total of {len(phonecall_set)} customers require telephone contact...")
            phonecall_list = list(phonecall_set)

            print("Phone calls required for :")
            counter = 0
            customer_num_reference = dict()
            while counter < len(phonecall_list):
                print(f"{counter + 1}. {phonecall_list[counter]}")
                customer_num_reference[counter] = phonecall_list[counter]
                counter += 1
            # print(customer_num_reference)
            customerexpand = input(f"Select a customer to expand (1-{len(phonecall_set)}):\n ")
            customerexpand = int(customerexpand)

            if customerexpand not in range(1, len(phonecall_list) + 1):
                print("Invalid input!")
            else:
                print(f"Expanding {phonecall_list[customerexpand - 1]}...")
                currentTest = ''
                contact_info(customer_num_reference.get(
                    customerexpand - 1))  # This calls the contact_info function with the expanded customer as an argument
                for item in phone_call:
                    if OOS_dict.get(item).get("Customer") == (customer_num_reference.get(customerexpand - 1)):
                        currentItem = (OOS_dict.get(item))
                        if currentTest != currentItem.get('Test'):
                            currentTest = currentItem.get('Test')
                            print(f"{currentTest}: ")
                        else:
                            pass

                        print(
                            f"-> Customer: {currentItem.get('Customer')} , Sample_No.: {currentItem.get('Sample Number')} , Sample_Name: {currentItem.get('Sample Name')} , Sample Type: {currentItem.get('Sample Name')} , Description: {currentItem.get('Description')} , Test: {currentItem.get('Test')} , Result: {currentItem.get('Result')}")
                    else:
                        pass
            pressEnter = input("\nPress enter to continue... ")
        elif user_input == 2:
            if (escalation_needed == 0):
                print("No escalations required.")
            else:
                print("Escalation needed for: ")
                for item in escalation_needed:
                    currentItem = (OOS_dict.get(item))
                    print(
                        f"-> Customer: {currentItem.get('Customer')} , Sample_No.: {currentItem.get('Sample Number')} , Sample_Name: {currentItem.get('Sample Name')} , Sample Type: {currentItem.get('Sample Name')} , Description: {currentItem.get('Description')} , Test: {currentItem.get('Test')} , Result: {currentItem.get('Result')}")
            pressEnter = input("\nPress enter to continue... ")

        elif user_input == 3:
            if len(local_environmental_failures) == 0:
                print("No local environmental failures found.")
            else:
                print("Environmental failures found:")
                for item in local_environmental_failures:
                    currentItem = (OOS_dict.get(item))
                    print(
                        f"-> Customer: {currentItem.get('Customer')} , Sample_No.: {currentItem.get('Sample Number')} , Sample_Name: {currentItem.get('Sample Name')} , Sample Type: {currentItem.get('Sample Name')} , Description: {currentItem.get('Description')} , Test: {currentItem.get('Test')} , Result: {currentItem.get('Result')}")
            pressEnter = input("\nPress enter to continue... ")

        elif user_input == 4:
            if len(IQC_EQA_list) == 0:
                print("No IQC or EQA samples found.")
            else:
                print("IQC / EQA samples found:")
                for item in IQC_EQA_list:
                    currentItem = (OOS_dict.get(item))
                    print(
                        f"-> Customer: {currentItem.get('Customer')} , Sample_No.: {currentItem.get('Sample Number')} , Sample_Name: {currentItem.get('Sample Name')} , Sample Type: {currentItem.get('Sample Name')} , Description: {currentItem.get('Description')} , Test: {currentItem.get('Test')} , Result: {currentItem.get('Result')}")
            pressEnter = input("\nPress enter to continue... ")

        elif user_input == 5:
            if len(daily_path_positives) == 0:
                print("No successful positive controls found!")
            else:
                print(f"Successful positive controls:")
                for item in daily_path_positives:
                    currentItem = (OOS_dict.get(item))
                    print(
                        f"-> Customer: {currentItem.get('Customer')} , Sample_No.: {currentItem.get('Sample Number')} , Sample_Name: {currentItem.get('Sample Name')} , Sample Type: {currentItem.get('Sample Name')} , Description: {currentItem.get('Description')} , Test: {currentItem.get('Test')} , Result: {currentItem.get('Result')}")
            pressEnter = input("\nPress enter to continue... ")

        elif user_input == 6:
            print("Shutting down program...\n.\n.\n.")
            escape = True


        else:
            print("Invalid input")


def user_choice():
    """ This is the function creates the 'menu' the user will interact with, and accepts a choice between (1 and 6). This
    function is used within the 'User_interface()' function, to display the menu and request user input"""

    global user_input
    print(
        "----------------\n     MENU\n----------------\n1. Client call list\n2. Escalation\n3. Environmental failures\n4. IQC/ EQA samples\n5. Positive Controls\n6. Exit program\n ----------------")

    user_input = input("Choose an option (1-6): ")

    try:
        user_input = int(user_input)
    except ValueError:
        print(f"Not an integer!")
    else:
        print("...")
        return user_input


def contact_info(customer):
    """ This function is called when expanding a client's OOS list on the 'Client call list' option. It searches the
    'Contact info' excel file for the relevant details and prints them. """

    contact_file = pd.read_excel(contact_file_input)

    filtered_data = contact_file[contact_file['Customer'] == customer]

    row_numbers = filtered_data.index.tolist()

    try:
        row = (row_numbers[0])
    except:
        print(f"\'{customer}\' not found in contacts!")
    else:
        customer_name = (filtered_data.at[row, "Customer"])
        contact_name = (filtered_data.at[row, "Contact Name"])
        telephone_num = (filtered_data.at[row, "Telephone Number"])
        email_address = (filtered_data.at[row, "Email Address"])
        notes = (filtered_data.at[row, "Notes"])
        print(
            f"---------------\n{customer_name} || Contact Name: {contact_name} , Telephone: {telephone_num} , Email: {email_address}, Notes: {notes} \n---------------")

