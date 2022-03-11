# outlook_contact_compiler
Compiles a list of all the contacts (email, first/last name, company name) that you've emailed or have emailed you, or been CC'd on emails sent to you. 

### Export list a contacts from all incoming and outgoing Outlook emails:
File > Open & Export > Import/Export > Export to a file > CSV > Select Inbox > Clear Map > Add From: Name, From: Address, To: Name, To: Address, CC: Name, CC: Address > resave as .xlsx in root directory

### Create output file
- Make a blank .xlsx file with "Email", "First Name", "Last Name", "Company" as headings in cells A1, A2, A3, A4
- Save in root directory

### Use above filenames
- type in the above filesnames in the open_input and open_output function calls

### Run
- run program by typing python3 outlook_contacts.py
