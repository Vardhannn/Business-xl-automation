import openpyxl as xl


def excel_edit():
    wb = xl.load_workbook("Functions.xlsx")
    sheet = wb["Sheet1"]
    lease = "d5"
    Electricity = "d6"
    Employee_salary = "d7"
    maintance = "d8"
    print("Enter the value you want to edit.")
    print("1 lease ")
    print("2 Electricity " )
    print("3 Employee_salary ")
    print("4 maintance ")
    a = int(input())
    if a == 1:
        a1 = sheet[lease].value
        if a1 is None:
            print("upaid")
            update = input("Due is upaid would you like to change it: ")
            if update == "yes":
                update_value = sheet["e5"].value
                sheet[lease].value = update_value
                wb.save("Functions.xlsx")
            elif update == "no ":
                    print("Great current status is unpaid")
        else:
            print("paid")
            update_change = input("Due is already paid would you like to change it: ")
            if update_change == "yes":
                update_change_value = None
                sheet[lease].value = update_change_value
                wb.save("Functions.xlsx")
            elif update_change == "no":
                    print("Great current status is paid")
    if a==2:
        a2 = sheet[Electricity].value
        if a2==None:
            print("unpaid")
            update = input("Due is upaid would you like to change it: ")
            if update == "yes":
                update_value = sheet["e6"].value
                sheet[Electricity].value = update_value
                wb.save("Functions.xlsx")
            elif update == "no":
                    print("Great current status is unpaid")
        else:
            print("paid")
            update_change = input("Due is already paid would you like to change it: ")
            if update_change == "yes":
                update_change_value = None
                sheet[Electricity].value = update_change_value
                wb.save("Functions.xlsx")
            elif update_change == "no":
                    print("Great current status is paid")

    if a==3:
        a3 = sheet[Employee_salary].value
        if a3==None:
            print("unpaid")
            update = input("Due is upaid would you like to change it: ")
            if update == "yes":
                update_value = sheet["e7"].value
                sheet[Employee_salary].value = update_value
                wb.save("Functions.xlsx")
            elif update == "no":
                    print("Great current status is unpaid")
        else:
            print("paid")
            update_change = input("Due is already paid would you like to change it: ")
            if update_change == "yes":
                update_change_value = None
                sheet[Employee_salary].value = update_change_value
                wb.save("Functions.xlsx")
            elif update_change == "no":
                    print("Great current status is paid")

    if a==4:
        a4 = sheet[maintance].value
        if a4==None:
            print("unpaid")
            update = input("Due is upaid would you like to change it: ")
            if update == "yes":
                update_value = sheet["e7"].value
                sheet[maintance].value = update_value
                wb.save("Functions.xlsx")
            elif update == "no":
                    print("Great current status is unpaid")
        else:
            print("paid")
            update_change = input("Due is already paid would you like to change it: ")
            if update_change == "yes":
                update_change_value = None
                sheet[maintance].value = update_change_value
                wb.save("Functions.xlsx")
            elif update_change == "no":
                    print("Great current status is unpaid")

excel_edit()

