import re
import xlsxwriter



# Text files containing details for all new papers captured by pubmed for a particular day/days.
text_files = ['pubmed-230226[edat].txt', 'pubmed-230227[edat].txt', 'pubmed-230228[edat].txt', 'pubmed-230301[edat].txt', 'pubmed-230302_230303[edat].txt', 'pubmed-230304_230305[edat].txt', 'pubmed-230306_230307[edat].txt', 'pubmed-230308_230309[edat].txt']


# Loop for each file in text_files list
for name in text_files:
    
    f = open(name, encoding="utf8")

    # Read the text file into object conatining file content as long single string 
    raw_text = f.read()

    # Individual reseach paper records are separated by a blank line.
    # Split the file at blank line points to create a list of individual paper records 
    complete_entries = [entry for entry in re.split(r"(?m)^\s*$\s*", raw_text) if entry.strip()]
    

    # Create empty lists (size = number of complete entries) to hold ids, titles and abstracts for each paper.
    id_text = [None] * len(complete_entries)
    ti_text = [None] * len(complete_entries)
    ab_text = [None] * len(complete_entries)


    # For each of the paper complete entries
    for i in range(len(complete_entries)):
        
        # Extract paper ID
        id_text[i] = re.findall('PMID- (.+?)\n', complete_entries[i], re.DOTALL)
    
        # Extract paper Title, add blank space if no title exists
        temp_ti = re.findall('TI  - (.*?) - ', complete_entries[i], re.DOTALL)
        if temp_ti:
            ti_text[i] = temp_ti
        else:
            ti_text[i] = [' ']

        # Extract paper Abstract, add blank space if no abstarct exists
        temp_ab =  re.findall('AB  - (.*?) - ', complete_entries[i], re.DOTALL)
        if temp_ab:
            ab_text[i] = temp_ab
        else:
            ab_text[i] = [' ']





    # Remove the trailing characters for each abstract
    for i in range(len(ab_text)):
        str = ab_text[i][0].split()[:-1]
        str_join = ' '.join(str)
        ab_text[i] = str_join
    
    # Remove the trailing characters for each title
    print(ti_text[3])
    for i in range(len(ti_text)):
        str = ti_text[i][0].split()[:-1]
        str_join = ' '.join(str)
        ti_text[i] = str_join



    # Remove the .txt exntension from file name
    new_name = name.split('.')[0]

    # Create xlsx workbook using file name
    workbook = xlsxwriter.Workbook(new_name+'.xlsx')

    # Add worksheet to xlsx file
    worksheet = workbook.add_worksheet()


    # Names for columns to be written to the xlxs worksheet
    col_names = ['PMID', 'TI', 'AB']

    # Wrtie column names to the xlxs worksheet
    for col, item in enumerate(col_names):
        worksheet.write(0, col, item)

    # Wrtie paper id's  to the xlxs worksheet in col 0
    for row, item in enumerate(id_text):
        worksheet.write(row+1, 0, item[0])

    # Write paper titles to the xlxs worksheet in col 1
    for row, item in enumerate(ti_text):
        worksheet.write(row+1, 1, item)

    # Wrtie paper abstracts to the xlxs worksheet in col 2
    for row, item in enumerate(ab_text):
        worksheet.write(row+1, 2, item)


    workbook.close()