import openpyxl


if __name__ == "__main__":
    print("What file do you want to work with?:")
    file = input()
    openpyxl.load_workbook(file)
    while true:
        print('What do you want to do?\n')
        command = input()
        if command == "Add new data":
            print('Designated smoking area?')
            dsa = input()
            if dsa == "Yes":
                print('What type of cigarette? (Cigarettes, Vape, Juul, Iqos, Other)')
                type = input()
                if type == "Cigarettes":
                    
                else if type == "Vape":

                else if type == "Juul":

                else if type == "Iqos":

                else if type == "Other":

                else:
                    print("Try again!")
            else if dsa == "No":
                print('What type of cigarette? (Cigarettes, Vape, Juul, Iqos, Other)')
                type = input()
                if type == "Cigarettes":
                    
                else if type == "Vape":

                else if type == "Juul":

                else if type == "Iqos":

                else if type == "Other":

                else:
                    print("Try again!")
            else:
                print("Try again")
        else if command == "Delete the data":

        else if command == "Calculate mean":

        else if command == "Calculate variance":

        else if command == "Calculate estimated mean":

        else if command == "Calculate estimated variance":

        else if command == "Print data":

        else if command == "Exit":
            break
        else:
            print('Try again!')