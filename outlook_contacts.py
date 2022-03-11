import openpyxl # must use version 2.6.2
import re

def open_input(name):
    try:
        input_wb = openpyxl.load_workbook(name)
        print('Successfully loaded input worksheet')
        return input_wb
    except Exception as e:
        print(e)
        return

def open_output(name):
    try:
        output_wb = openpyxl.load_workbook(name)
        print('Successfully loaded output worksheet')
        return output_wb
    except Exception as e:
        print(e)
        return

def compile_name(full_name):
    first = full_name[0].strip('\'').strip(',') if len(full_name) > 0 and "@" not in full_name[0] else ""
    last = full_name[1].strip('\'').strip(',') if len(full_name) > 1 else ""
    return {'first': first, 'last': last}

def compile_contacts(input_wb, output_wb):    
    emails_found = []
    current_row_in_results = 2
    input_worksheet = input_wb.active
    output_worksheet = output_wb.active

    for i in range(2, input_worksheet.max_row + 1):
        from_email = input_worksheet['B' + str(i)].value.lower()
        if from_email[0] != '/':
            try:
                emails_found.index(from_email)
            except:
                emails_found.append(from_email)

                full_name = input_worksheet['A' + str(i)].value.split(' ')
                first_name = compile_name(full_name)['first']
                last_name = compile_name(full_name)['last']

                try:
                    company = re.search('@(.*)\.[com|ca|org|net]', from_email).group(1)
                except:
                    company = ''

                output_worksheet['A' + str(current_row_in_results)].value = from_email
                output_worksheet['B' + str(current_row_in_results)].value = first_name
                output_worksheet['C' + str(current_row_in_results)].value = last_name
                output_worksheet['D' + str(current_row_in_results)].value = company
                
                current_row_in_results += 1

        to_emails = input_worksheet['D' + str(i)].value.split(';') if input_worksheet['D' + str(i)].value else []
        to_full_names = input_worksheet['C' + str(i)].value.split(';') if input_worksheet['C' + str(i)].value else []
        if len(to_emails) > 0:
            for j in range(0, len(to_emails)):
                if len(to_emails[j]) > 0 and to_emails[j][0] != '/':
                    to_emails[j] = to_emails[j].lower()
                    try:
                        emails_found.index(to_emails[j])
                    except:
                        emails_found.append(to_emails[j])

                        full_name = to_full_names[j].split(' ') if len(to_full_names) == len(to_emails) else []
                        first_name = compile_name(full_name)['first']
                        last_name = compile_name(full_name)['last']

                        try:
                            company = re.search('@(.*)\.[com|ca|org|net]', to_emails[j]).group(1)
                        except:
                            company = ''

                        output_worksheet['A' + str(current_row_in_results)].value = to_emails[j]
                        output_worksheet['B' + str(current_row_in_results)].value = first_name
                        output_worksheet['C' + str(current_row_in_results)].value = last_name
                        output_worksheet['D' + str(current_row_in_results)].value = company

                        current_row_in_results += 1

        cc_emails = input_worksheet['F' + str(i)].value.split(';') if input_worksheet['F' + str(i)].value else []
        cc_full_names = input_worksheet['E' + str(i)].value.split(';') if input_worksheet['E' + str(i)].value else []
        if len(cc_emails) > 0:
            for j in range(0, len(cc_emails)):
                if cc_emails[j][0] != '/':
                    cc_emails[j] = cc_emails[j].lower()
                    try:
                        emails_found.index(cc_emails[j])
                    except:
                        emails_found.append(cc_emails[j])

                        full_name = cc_full_names[j].split(' ') if len(cc_full_names) == len(cc_emails) else []
                        first_name = compile_name(full_name)['first']
                        last_name = compile_name(full_name)['last']

                        try:
                            company = re.search('@(.*)\.[com|ca|org|net]', cc_emails[j]).group(1)
                        except:
                            company = ''

                        output_worksheet['A' + str(current_row_in_results)].value = cc_emails[j]
                        output_worksheet['B' + str(current_row_in_results)].value = first_name
                        output_worksheet['C' + str(current_row_in_results)].value = last_name
                        output_worksheet['D' + str(current_row_in_results)].value = company

                        current_row_in_results += 1

    print("Completed contact compilation")
    return output_wb.save('compiled_contacts.xlsx')


if __name__ == '__main__':
    input_wb = open_input('Outlook Contacts_Mar 2022.xlsx')
    output_wb = open_output('compiled_contacts.xlsx')
    compile_contacts(input_wb, output_wb)
