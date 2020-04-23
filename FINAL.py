import xlrd

def user_input():
    """
    Goal = to have the following pattern : "3, Kal. July -> 3 days before the kalendes of July
    """
    try:
        number = int(input("Please enter the number of days before the feast. If it's the day of a feast, write '1'; if it's the day before a feast, write '2': "))
    except ValueError:
        number = int(input("Please enter a number, not a string: "))
    while number < 0 or number > 31:
        number = int(input("Your number seems wrong. It should be between 1 and 31. Please check your number: "))

    feast = str.capitalize(input("Please insert your feast: "))
    while feast not in list_feast:
        feast = str.capitalize(input("Your feast seems wrong. Please insert 'kal', 'nones', or 'ides': "))
    if feast in ["Nones", "Non", "None", "Non."]:
        if number > 7:
            print("It seems that what you asked is not possible since the", str(number) + "th of Nones does not exist in the Roman Calendar")
            return user_input()
    if feast in ["Ides", "Eid", "Eid."]:
        if number > 9:
            print("It seems that what you asked is not possible since the", str(number) + "th of Ides does not exist in the Roman Calendar")
            return user_input()

    month = str.capitalize(input("Please insert your month either in Latin either in English: "))
    while month not in list_month:
        month = str.capitalize(input("Your month seems wrong. Please check your month: "))

    return number, feast, month


def feast_():
    """
    Goal = Input dates of feasts
    """
    if feast in ["Kal", "Kalendes", "Kal."]:
        return 1
    if feast in ["Nones", "Non", "Non.", "None"]:
        if month in ["March", "May", "July", "October"]:
            return 7
        else:
            return 5
    if feast in ["Ides", "Eid", "Eid."]:
        if month in ["March", "May", "July", "October"]:
            return 15
        else:
            return 13


def roman_month_kal():
    """
    Goal = To provide nbre of days in months for the required in the calcul_kal()
    """
    if month in ["January", "February", "April", "June", "Augustus", "September", "November"]:
        return 31
    if month in ["May", "July", "October", "December"]:
        return 30
    if month in ["March"]:
        return 28

def roman_month_else():
    """
    Goal = To provide nbre of days in months for ides and nones required in the calcul_else()
    """
    if month in ["January", "February", "March", "May", "July", "August", "October", "December"]:
        return 31
    if month in ["April", "June", "September", "November"]:
        return 30
    if month in ["February"]:
        return 28

def month_translation():
    """
    Give EN translation of the month in case the user wrote in Latin
    """
    if month == "Ianuarius":
        return "January"
    if month == "Februarius":
        return "February"
    if month == "Martius":
        return "March"
    if month == "Aprilis":
        return "April"
    if month == "Maius":
        return "May"
    if month == "Iunius":
        return "June"
    if month == "Quinctilis" or month == "Iulius":
        return "July"
    if month == "Sextilis" or month == "Augustus":
        return "August"
    else:
        return month

def month_kal():
    """
    Function required to provide the good month in answer for the user since x of Kal. means x days before the 1st of a month
    """
    if month == "January":
        return "December"
    if month == "February":
        return "January"
    if month == "March":
        return "February"
    if month == "April":
        return "March"
    if month == "May":
        return "April"
    if month == "June":
        return "May"
    if month == "July":
        return "June"
    if month == "August":
        return "July"
    if month == "September":
        return "August"
    if month == "October":
        return "September"
    if month == "November":
        return "October"
    if month == "December":
        return "November"


def calcul_kal():
    """
    calcul for Kalendes
    """
    a = roman_month_kal() - number + 2
    return a

def calcul_else():
    """
    calcul for Ides and Nones
    """
    a = feast_() - number + 1
    return a

def open_file(path):
    """
    def to find the year of the calcul regarding the consul
    """
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    while True:
        agreement = str.capitalize(input("Do you want to search a specific year (with our consul database)? Please note it's still in beta and that result should be verified. Type yes or no: "))
        if agreement == "Yes":
            break
        if agreement == "No":
            return False

    while True:
        global name
        name = input("Please enter the name of the consul: ")
        global nbre_consulate
        try:
            nbre_consulate = int(input("Please check the number of times he is consul: "))

        except ValueError:
            nbre_consulate = int(input("Please enter the number of times he is consul, not a string: "))

        for row_num in range(sheet.nrows):
            row_value = sheet.row_values(row_num)

            if row_value[0] == name and row_value[1] == nbre_consulate:
                return int(row_value[2])
                break
        if row_value[0] != name or row_value[1] != nbre_consulate:
            print("It seems that your entries didn't match our database. You can check our list and try again: ")


def year_final():
    """
    in case user wanted consul database, mute the number between BC and AD
    """
    if result_year >= 0:
        x = str(result_year) + " A.D"
        return x
    else:
        x = str(abs(result_year)) + " B.C"
        return x

"""
Lists are to check if user's input is good
"""
list_feast = ["Kal", "Kalendes", "Kal.", "Nones", "Non", "None", "Non.", "Ides", "Eid", "Eid."]
list_month = ["Ianuarius", "January", "Februarius", "February", "Martius", "March", "Aprilis", "April", "Maius", "May", "Iunius", "June", "Quinctilis", "July", "Iulius", "Sextilis", "Augustus", "August", "September", "October", "November", "December"]


"""
beginning of the execution
"""
number, feast, month = user_input() #User's input
month = month_translation() #Translation of the Month in English in case the user wrote in Latin

path = "Consuls_list.xlsx" #Set the path to the consul database
result_year = open_file(path) #Open the def for consul database



if result_year == False: #If the user doesn't want the consul database
    if feast_() == 1 and number == 1 :
        print("The date you asked is the Kalendes of", month + ". In our calendar, it means on the 1st", month_translation() + ".")
    if feast_() == 1 and number == 2:
        print("The date you asked is the day just before the Kalendes of", month + ". In our calendar, it means on the", calcul_kal(), month_kal() + ".")
    if feast_() == 1 and number > 2:
        print("The date you asked is the", str(number) + "th before the Kalendes of", month + ". In our calendar, it means on the", calcul_kal(), month_kal() + ".")
    if feast_() != 1:
        print("The date you asked is the", str(number) + "th before the", feast, "of", month + ". In our calendar, it means on the", calcul_else(), month_translation() + ".")

else: #if the user wants consul database
    if feast_() == 1 and number == 1 :
        print("The date you asked is the Kalendes of", month, "under the consulate number", nbre_consulate, "of", name + ". In our calendar, it means on the 1st", month_translation(), year_final() + ".")
    if feast_() == 1 and number == 2:
        print("The date you asked is the day just before the Kalendes of", month, "under the consulate number", nbre_consulate, "of", name + ". In our calendar, it means on the", calcul_kal(), month_kal(), year_final() + ".")
    if feast_() == 1 and number > 2:
        print("The date you asked is the", str(number) + "th before the Kalendes of", month, "under the consulate number", nbre_consulate, "of", name + ". In our calendar, it means on the", calcul_kal(), month_kal(), year_final() + ".")
    if feast_() != 1:
        print("The date you asked is the", str(number) + "th before the", feast, "of", month, "under the consulate number", nbre_consulate, "of", name + ". In our calendar, it means on the", calcul_else(), month_translation(), year_final() + ".")
