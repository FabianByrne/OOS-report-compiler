from Functions import run_program

"""All of the functions are listed in Functions.py"""
"""The order of operations appears like this:
1. The excelfile_paths() function obtains the file paths for the OOS results folder and the Contact details spreadsheet, storing them in variables
2. The 'spreadsheet_iterator()' function passes runs through each OOS (out of specification) spreadsheet in the chosen folder,
    extracting the useful data row by row, then executing the 'main_concerns()' function against the contents of each spreadsheet.
    3. The main_concerns() function effectively filters the out any OOS results that will not be of interest, and compiles a 
        a) a list of in the format of '{sample_number} {test}'
        b) a dictionary with the '{sample_number} {test}' as a key, and a nested dictionary containing rest of sample info as the value.

            4. The main_concerns() function does this by executing the 'test_checker()' function (which only records data for a specified test 
                type) for a number of prioritised tests.

5. Using the output of the 'spreadsheet_iterator()' as an argument (the list and dictionary), the 'criteria_checker()'
    function further filters and organises the samples into new lists.

6. The 'user_interface' function now has satisfied it's required parameters. It uses the function 'user_choice()' to request 
    input, then processes the result and displays information relevant to the user's choice.

7. The 'contact_info()' function will search through an excel file containing contact information if the user wishes to 
    view a client's contact details. 
"""
run_program()
