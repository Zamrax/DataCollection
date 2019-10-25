import openpyxl

def add_data(sheet, dsa):
    if dsa == "Yes":
        print('What type of cigarette? (Cigarettes, Vape, Juul, Iqos, Other)')
        type = input()
        if type == "Cigarettes":
            sheet['B2'].value += 1
        elif type == "Vape":
            sheet['B3'].value += 1
        elif type == "Juul":
            sheet['B4'].value += 1
        elif type == "Iqos":
            sheet['B5'].value += 1
        elif type == "Other":
            sheet['B6'].value += 1
        else:
            return 1
        return 0
    elif dsa == "No":
        print('What type of cigarette? (Cigarettes, Vape, Juul, Iqos, Other)')
        type = input()
        if type == "Cigarettes":
            sheet['C2'].value += 1
        elif type == "Vape":
            sheet['C3'].value += 1
        elif type == "Juul":
            sheet['C4'].value += 1
        elif type == "Iqos":
            sheet['C5'].value += 1
        elif type == "Other":
            sheet['C6'].value += 1
        else:
            return 1
        return 0
    else:
        return 1

def del_data(sheet, dsa):
    if dsa == "Yes":
        print('What type of cigarette? (Cigarettes, Vape, Juul, Iqos, Other)')
        type = input()
        if type == "Cigarettes":
            sheet['B2'] -= 1
        elif type == "Vape":
            sheet['B3'] -= 1
        elif type == "Juul":
            sheet['B4'] -= 1
        elif type == "Iqos":
            sheet['B5'] -= 1
        elif type == "Other":
            sheet['B6'] -= 1
        else:
            return 1
        return 0
    elif dsa == "No":
        print('What type of cigarette? (Cigarettes, Vape, Juul, Iqos, Other)')
        type = input()
        if type == "Cigarettes":
            sheet['C2'] -= 1
        elif type == "Vape":
            sheet['C3'] -= 1
        elif type == "Juul":
            sheet['C4'] -= 1
        elif type == "Iqos":
            sheet['C5'] -= 1
        elif type == "Other":
            sheet['C6'] -= 1
        else:
            return 1
        return 0
    else:
        return 1

if __name__ == "__main__":
    print("What file do you want to work with?:")
    file = input()
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    while True:
        print('What do you want to do? (Add new)')
        command = input()
        if command == "Add new data" or command == "add new data" or command == "Add" or command == "add":
            print('Designated smoking area?')
            dsa = input()
            result = add_data(sheet, dsa)
            if result == 1:
                print('Try again')
        elif command == "Delete the data" or command == "delete the data" or command == "Delete" or command == "delete":
            print('Designated smoking area?')
            dsa = input()
            result = del_data(sheet, dsa)
            if result == 1:
                print('Try again')
        elif command == "Print data" or command == "print" or command = "Print":
            for i in range(1, sheet.max_row + 1):
                for j in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row = i, column = j)
                    print(cell.value,end=' | ')
                print('\n')
        elif command == "Exit" or command == "exit":
            break
        else:
            print('Try again!')