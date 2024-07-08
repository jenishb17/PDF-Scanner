import PyPDF2
import os
import sys
import xlwings as xw 

# Define the path to your PDF file


def extract_form_values(path,reader):
    with open(path , 'rb') as file:
        if not reader.get_fields():
            return
        
        form_fields = reader.get_fields()
        field_values = {}
        for field_name, field_data in form_fields.items():
            field_values[field_name] =  field_data.get('/V',None)

        return field_values
    
def extract_Money(form_values):
    Cash = "Group2"

    Iscash = False

    if(form_values[Cash] == "/Choice6"):
        Iscash = True
    if(Iscash == True and form_values["Other"] != None):
        Iscash =0
    Money = "Approximate amount you wish to invest in a mortgage investment vehicle"
    Amount = form_values[Money]

    return Amount,Iscash
def Extract_Time(form_values):
    Frame = "Group1a"
    Choice = form_values[Frame]
    Pick =""
    if(Choice == "/Choice1"):
        Pick = "Less than 1 year"
    elif(Choice == "/Choice2"):
        Pick = "1-3 years"
    elif(Choice == "/Choice3"):
        Pick = "4-5 years"
    elif(Choice == "/Choice4"):
        Pick = "6-9 years"
    else:
        Pick = "10+ years"
    
    return Pick
def Extract_Source(form_values):
    referred = None
    keys_to_extract = ["Print Media", "Online", "undefined","Word of Mouth","undefined_2","undefined_3"]
    i=0
    On=True
    source = None
    while(i<len(keys_to_extract) and On == True ):
        if keys_to_extract[i] in form_values:
            if(form_values[keys_to_extract[i]] == "/On"):
                source = keys_to_extract[i]
                if(keys_to_extract[i] == "undefined"):
                    referred = form_values["Referred By"]
                    source = "Referred"
                elif(keys_to_extract[i] == "undefined_2"):
                    source = "TV Program"
                    referred = form_values["TV Program"]
                elif(keys_to_extract[i] == "undefined_3"):
                    source = "Other"
                    referred = form_values["Other"]

                On=False
                break
        i+=1

    return source,referred
def extract_specific_values(form_values):
    # Define the specific keys we want to extract values for
    keys_to_extract = [
        "Last Name",
        "First Name",
        "Address",
        "CityTown",
        "ProvinceTerritory",
        "Postal Code",
        "Phone Number",
        "Email",
        "undefined_6",
        "3KRQH 1XPEHU  BBBBBBBBBBBBBBBBBBBBB   PDLO"
    ]
    
    # Create a dictionary to store the extracted values
    extracted_values = {}
    
    # Loop through each key in the list of keys to extract
    for key in keys_to_extract:
        # Check if the key exists in the form_values dictionary
        if key in form_values:
            # Add the key-value pair to the extracted_values dictionary
            extracted_values[key] = form_values[key]
        else:
            # Optionally, add a None value or a default value if the key is not found
            extracted_values[key] = None
    
    return extracted_values

def Extract_Designation(form_values):
    designated = "Not Designated"

    if((form_values["Investor annual income"]) == "/Choice5" or (form_values["Investor annual income"]) == "/Choice6"):
        designated = "Designated"
    elif(((form_values["Investor annual income"]) == "/Choice5" or (form_values["Investor annual income"]) == "/Choice6" or (form_values["Investor annual income"]) == "/Choice4") and ((form_values["Spouse annual icome"]) == "/Choice5" or (form_values["Spouse annual icome"]) == "/Choice6" or (form_values["Spouse annual icome"]) == "/Choice4")):
        designated = "Designated"
    elif(form_values["Investor Assets"] == "/Choice4" or form_values["Spouse assets"] == "/Choice4"):
        designated = "Designated"
    elif(form_values["Investor Assets"] == "/Choice3" and form_values["Spouse assets"] == "/Choice3"):
        designated = "Designated"
    elif(form_values["Investor Net Assets"] == "/Choice4" or form_values["Sopouse Net Assets"] == "/Choice4"):
        designated = "Designated"
    elif(form_values["Investor Net Assets"] == "/Choice3" and form_values["Sopouse Net Assets"] == "/Choice3"):
        designated = "Designated"

    return designated

def Alreadys_Exists(email):
    wb = xw.books.active
    ws = wb.sheets[0]
    count =1
    while(ws.range(f'D{count}').value != None):
        count +=1
        
    i =1
    while(i<count):
        if(ws.range(f'D{i}').value == email):         
            return True,i
        i+=1
    return False, 2
def Into_excel(values,Change):
    LastName= values[0]
    FirstName=values[1]
    Address = values[2]
    City =values[3]
    Province = values[4]
    Pcode = values[5]
    Pnum = values[6]
    email = values[7]
    date = values[8]
    altEmail = values[9]
    Amount,Iscash = extract_Money(form_values)

    TimeFrame = Extract_Time(form_values)

    source,referred = Extract_Source(form_values)

    designated = Extract_Designation(form_values)

    # Write to the active Excel workbook using xlwings
    try:
            
        wb = xw.books.active
        ws = wb.sheets[0]
        ws.api.Rows("2:2").Insert(Shift=-4161)  
        rowExist =2
        row = ws.range(f"A2:S2")
        if(Change == False):
            row.color = (0, 255, 0) 
        else:
            row.color = (255, 255, 0) 
        ws.range(f'A{rowExist}').value = FirstName
        ws.range(f'B{rowExist}').value = LastName
        FullName = f"{FirstName}, {LastName}"
        ws.range(f'C{rowExist}').value = FullName
        ws.range(f'D{rowExist}').value = email
        ws.range(f'E{rowExist}').value = altEmail
        ws.range(f'F{rowExist}').value = source
        ws.range(f'G{rowExist}').value = designated
        if(Iscash == True):
            ws.range(f'H{rowExist}').value = Amount
        elif(Iscash == False):
            ws.range(f'I{rowExist}').value = Amount
        else: 
            ws.range(f'H{rowExist}:I{rowExist}').value = Amount

        ws.range(f'J{rowExist}').value = TimeFrame
        ws.range(f'K{rowExist}').value = Address
        ws.range(f'L{rowExist}').value = City
        ws.range(f'M{rowExist}').value = Province
        ws.range(f'N{rowExist}').value = Pcode
        ws.range(f'O{rowExist}').value = Pnum
        ws.range(f'P{rowExist}').value = date
        ws.range(f'Q{rowExist}').value = referred
        ws.range(f'R{rowExist}').value = None
        wb.save()
    except Exception as e:      
        print(f"Error writing to Excel: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_pdf_data.py <pdf_path>")
    else:
        path = sys.argv[1]
        #path = 'C:/Users/jbharucha/OneDrive - Valour Management Inc/Desktop/ProgramCOde/Jenish.pdf'
        if not os.path.exists(path):
            print(f"Error: PDF file not found at {path}")

        pdf = open(path, 'rb')

        reader = PyPDF2.PdfReader(pdf)
        
        info = reader.metadata
        form_values = extract_form_values(path,reader)

        extracted_form_values = extract_specific_values(form_values)
        values = []
        
        for i,j in extracted_form_values.items():
            values.append(j)
        email = values[7]
        Exists,rowExist = Alreadys_Exists(email)


        # Write to the active Excel workbook using xlwings
        try:
            
            wb = xw.books.active
            ws = wb.sheets[0]
            if(Exists == False):
                Into_excel(values,False)
            else:
                ws.range(f"A{rowExist}").api.EntireRow.Delete()
                Into_excel(values,True)
                wb.save()
                    
        except Exception as e:      
            print(f"Error writing to Excel: {e}")
    





