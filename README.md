# outlook_contact_compiler
Compiles a list of all the contacts (email, first/last name, company name) that you've emailed or have emailed you, or been CC'd on emails sent to you. 

### Export list a contacts from all incoming and outgoing Outlook emails:
File > Open & Export > Import/Export > Export to a file > CSV > Select Inbox > Clear Map > Add From: Name, From: Address, To: Name, To: Address, CC: Name, CC: Address > resave as .xlsx in root directory

### Create output file
- Make a blank .xlsx file with "Email", "First Name", "Last Name", "Company" as headings in cells A1, A2, A3, A4 (or use existing "compiled_contacts.xlsx")
- Save in root directory

### Use above filenames to consolidate a list of unique contact names, emails and companies
- type in the above filesnames in the `consolidate_raw_list()` function
- this will output a file named "compiled_contacts.xlsx"
- e.g. `consolidate_raw_list('heather_raw_contacts.xlsx', 'compiled_contacts.xlsx')` will compile a list of unique contacts from "heather_raw_contacts" and save them in a file named "compiled_contacts"

### Compare the output (compiled_contacts.xlsx) to another sheet to eliminate any shared contacts
- compare the compiled contacts to an already complete list of contacts (the "master" sheet) by typing there file names into the `compare_and_remove_duplicates()` function
- this will outpul a file named "compiled_noDuplicates.xlsx" of only the contacts not already contained in the master sheet
- e.g. `compare_and_remove_duplicates('josh_compiled_contacts.xlsx','compiled_contacts.xlsx')` will compare the "compiled_contacts" sheet to a master sheet of contacts previously compiled (in this case, "josh_compiled_contacts") and remove any contacts already contained in the master.  It will save the new list in "compiled_noDuplicates.xlsx"

### How to Run
- run program by typing python3 outlook_contacts.py
