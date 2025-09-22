import pandas as pd
from bs4 import BeautifulSoup
from io import StringIO
import re
def hlo(html_file,output_excel):
    filename=html_file.split('/')[-1]
    filevalues=filename.replace('.html','').split('_')
    if len(filevalues)<6:
        filevalues=filevalues+['']*(6-len(filevalues))
    elif len(filevalues)>6:
        filevalues=filevalues[-6:]
    
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, 'html.parser')
    columns = ['SYSTEM NAME','DEPARTMENT','EMPLOYEE NAME','LOCATION','WORK PLACE','PORT NUMBER','COMPUTER NAME', 'OPERATING SYSTEM', 'SYSTEM MODEL', 'SERVICE TAG', 'PROCESSOR', 'BOARD',
               'HDD(IN GB)', 'MEMORY(IN MB)','NO OF RAM SLOTS', 'DISPLAY', 'MONITOR']
    otpt = []
    otpt.extend(filevalues)
    target = 'reportHeader'
    table = soup.find('table', class_=target)
    if table:
        table_str = str(table)
        table_io = StringIO(table_str)
        df = pd.read_html(table_io)[0]
        otpt.append(df.iloc[1, 1])

    div = soup.find('div', class_='reportSection rsLeft')
    if div:
        table = div.find('table')
        if table:
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                for cell in cells:
                    cell_texts = cell.decode_contents().split('<br/>')
                    cell_texts = [text.strip() for text in cell_texts]
                    otpt.append(cell_texts[0])

    div = soup.find('div', class_='reportSection rsRight')
    if div:
        table = div.find('table')
        if table:
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                for cell in cells:
                    cell_texts = cell.decode_contents().split('<br/>')
                    cell_texts = [text.strip() for text in cell_texts]
                    otpt.append(cell_texts[0])

    div = soup.find('div', class_='reportSection rsRight')
    if div:
        table = div.find('table')
    if table:
        rows = table.find_all('tr')
        for row in rows:
            cells = row.find_all('td')
            for cell in cells:
                cell_texts = cell.decode_contents().split('<br/>')
                cell_texts = [text.strip() for text in cell_texts]
                otpt.append(cell_texts[1][22:])

    divs = soup.find_all('div', class_='reportSection rsLeft')
    specific_div = divs[1]
    table = specific_div.find('table')
    rows = table.find_all('tr')
    first_row = rows[0]
    first_row_data = []
    for cell in first_row.find_all(['td', 'th']):
        cell_lines = cell.decode_contents().split('<br/>')
        first_line = BeautifulSoup(cell_lines[0], 'html.parser').get_text(strip=True)
        first_row_data.append(first_line)
    otpt.append(first_row_data[0])

    divs = soup.find_all('div', class_='reportSection rsRight')
    specific_div = divs[1]
    table = specific_div.find('table')
    rows = table.find_all('tr')
    first_row = rows[0]
    first_row_data = []
    for cell in first_row.find_all(['td', 'th']):
        cell_lines = cell.decode_contents().split('<br/>')
        first_line = BeautifulSoup(cell_lines[0], 'html.parser').get_text(strip=True)
        first_row_data.append(first_line)
    otpt.append(first_row_data[0][6:])

    divs = soup.find_all('div', class_='reportSection rsLeft')
    specific_div = divs[2]
    table = specific_div.find('table')
    rows = table.find_all('tr')
    first_row = rows[0]
    first_row_data = []
    for cell in first_row.find_all(['td', 'th']):
        cell_lines = cell.decode_contents().split('<br/>')
        first_line = BeautifulSoup(cell_lines[0], 'html.parser').get_text(strip=True)
        first_row_data.append(first_line)
        statement = first_row_data[0]
        match = re.search(r'\b(\d+)\.\d+\s+Gigabytes', statement)
    if match:
        integer_part = match.group(1)
        otpt.append(integer_part)

    divs = soup.find_all('div', class_='reportSection rsRight')
    specific_div = divs[2]
    table = specific_div.find('table')
    rows = table.find_all('tr')
    first_row = rows[0]
    first_row_data = []
    for cell in first_row.find_all(['td', 'th']):
        cell_lines = cell.decode_contents().split('<br/>')
        first_line = BeautifulSoup(cell_lines[0], 'html.parser').get_text(strip=True)
        first_row_data.append(first_line)
        statement = first_row_data[0]
        match = re.search(r'\d+', statement)
    if match:
        otpt.append(match.group())
    divs = soup.find_all('div', class_='reportSection rsRight')
    specific_div = divs[2]  
    table = specific_div.find('table')
    rows = table.find_all('tr')
    line=0
    for row in rows:
       tds=row.find_all('td')
       for td in tds:
        line=len(td.find_all('br'))
        
    otpt.append(line-1)
    
    

    divs = soup.find_all('div', class_='reportSection rsRight')
    specific_div = divs[4]
    table = specific_div.find('table')
    rows = table.find_all('tr')
    first_row = rows[0]
    first_row_data = []
    for cell in first_row.find_all(['td', 'th']):
        cell_lines = cell.decode_contents().split('<br/>')
        first_line = BeautifulSoup(cell_lines[0], 'html.parser').get_text(strip=True)
        first_row_data.append(first_line)
    otpt.append(first_row_data[0])

    divs = soup.find_all('div', class_='reportSection rsRight')
    specific_div = divs[4]
    table = specific_div.find('table')
    rows = table.find_all('tr')
    second_row = rows[0]
    second_row_data = []
    for cell in second_row.find_all(['td', 'th']):
        cell_lines = cell.decode_contents().split('<br/>')
        if len(cell_lines) > 1:
            second_line = BeautifulSoup(cell_lines[1], 'html.parser').get_text(strip=True)
            second_row_data.append(second_line)
    if second_row_data:
        statement = {'Description': [second_row_data[0]]}
        df = pd.DataFrame(statement)

        def extract_name(description):
            match = re.search(r'^(.*?)\[Monitor\]', description)
            if match:
                return match.group(1)
            else:
                return None

        df['Name'] = df['Description'].apply(extract_name)
        otpt.append(df['Name'].values[0])
       
    df_otpt = pd.DataFrame([otpt], columns=columns)
    
    df_otpt.to_excel('output12.xlsx', index=False)
    df_otpt.to_excel("C:\\Users\\sudar\\OneDrive\\Desktop\\projects\\.ipynb_checkpoints\\output12.xlsx")
    print(df_otpt)
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_otpt.to_excel(writer, index=False)
        # Make the Excel sheet visible
        writer.sheets['Sheet1'].sheet_view.showGridLines = True
        # Auto adjust column widths
        for column in writer.sheets['Sheet1'].columns:
            max_length = 0
            column = column[0].column_letter
            for cell in writer.sheets['Sheet1'][column]:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            writer.sheets['Sheet1'].column_dimensions[column].width = adjusted_width
hlo('SYMXC039_ESD_K.SURESH_VSPITP_1B_D69.html','output.xlsx')