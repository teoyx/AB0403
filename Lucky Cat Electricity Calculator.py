# -*- coding: utf-8 -*-
"""
Created on Sat Jan 29 15:44:09 2022

@author: Group 1
"""

import csv
import xlsxwriter as xls
import datetime

session = 0


title = open("dragon.txt","r")
for each_line in title:
    print(each_line,end="")
title.close()
    
name=input("""Welcome to the electricity calculator! 
           
           \nWhat's your name? """)
#fix the throwback 
def main():
    global session
    session += 1    
    def homecheck():
        category=0
        while True:
            try:
                print("="*100)
                menu=""" 
    Welcome {}! 
        
        [1] Home
        [2] Quit
    
    Please enter the appropriate number: """.format(name)
                category = int(input(menu))
                
                if category <=0 or category >=3:
                   print("please enter a valid number")
                   continue
         
            except ValueError:
                   print("please enter a valid option")
                   continue
            return(category)
    
    with open("pricelist.csv") as pricelist:
        pricelist_list= csv.reader(pricelist)
        allprice={}
        benchmark={}
        for each in pricelist_list:
            allprice[each[0]]= float(each[1])  # create dictionary for electricity providers
            benchmark[each[3]] =float(each[4]) # create dictionary for home benchmarks 
    
    with open("electricity consumption.csv") as consumption:
        appliance_list= csv.reader(consumption)
        appliancekwh={}
        for each in appliance_list:
            appliancekwh[each[0]]= float(each[1])/1000  # create dictionary for electricity providers
            #print(appliancekwh)
    
    def homedetails():
        homedict = {"A":"1-Room / 2 Room","B":"3-Room","C":"4-Room","D":"5-Room and Executive",
                    "E":"Private Apartments and Condominiums","F":"Landed Properties"}
        while True:
            print("="*100)
            print("""              
    Public Housing
       [A] 1-Room / 2 Room
       [B] 3-Room
       [C] 4-Room
       [D] 5-Room and Executive
    
    Private Housing
       [E] Private Apartments and Condominiums
       [F] Landed Properties""")
        
            home = input("Please enter the letter corresponding to your home type: ")
            try:
                if home in "ABCDEFabcdef":
                    True
                else:
                    print("Please enter a valid letter. ")
                    continue
                home = home.upper()
                print("")
                print("The home type you have entered is",homedict[home])
                print("="*100)
                return(homedict[home])
            except KeyError:
                print("Please enter a valid letter. ")
                continue
                
        
    def provider_select():   #to obtain service provider from user
        while True:
            selection = input("""
    List of Singapore Electricity Service Providers:
    ------------------------------------------------
        [TP] Tuas Power Supply Pte Ltd 
        [SP] SP Services Ltd 
        [KE] Keppel Electric Pte Ltd 
        [SE] Seraya Energy Pte Ltd (Geneco) 
        [SCP] Sembcorp Power Pte Ltd 
        [SES] Senoko Energy Supply Pte Ltd 
        [PLE] PacificLight Energy Pte Ltd 
    ------------------------------------------------
    Please enter the initials of your current service provider: """)
            
            if selection == "TP" or selection.upper() == "TP":
                rate = allprice["Tuas Power Supply Pte Ltd"]
                return rate
            elif selection == "SP" or selection.upper() == "SP":
                rate = allprice["SP Services Ltd"]
                return rate
            elif selection == "KE" or selection.upper() == "KE":
                rate = allprice["Keppel Electric Pte Ltd"]
                return rate
            elif selection == "SE" or selection.upper() == "SE":
                rate = allprice["Seraya Energy Pte Ltd (Geneco)"]
                return rate
            elif selection == "SCP" or selection.upper() == "SCP":
                rate = allprice["Sembcorp Power Pte Ltd"]
                return rate
            elif selection == "SES" or selection.upper() == "SES":
                rate = allprice["Senoko Energy Supply Pte Ltd"]
                return rate
            elif selection == "PLE" or selection.upper() == "PLE":
                rate = allprice["PacificLight Energy Pte Ltd"]
                return rate
            else:
                print("Vendor not found. Please enter from the following: ")
                selection = input("""
    List of Singapore Electricity Service Providers:
    ------------------------------------------------    
        [TP] Tuas Power Supply Pte Ltd 
        [SP] SP Services Ltd 
        [KE] Keppel Electric Pte Ltd 
        [SE] Seraya Energy Pte Ltd (Geneco) 
        [SCP]Sembcorp Power Pte Ltd 
        [SES] Senoko Energy Supply Pte Ltd 
        [PLE] PacificLight Energy Pte Ltd 
    ------------------------------------------------
    Please enter the initials of your current service provider: """)
    
    
    def last_mth_bill():    
        def input_bill():
            last_bill=0
            while True:
                try: 
                    last_bill=float(input("How much was your electricity bill last month?: "))
                    if last_bill <=0:
                        print("Please enter a valid amount")
                        continue
                    else:
                        break 
                except ValueError:
                    print("Please enter a valid amount ")
                    continue
            print("")
            print("Last month, your electricity bill amounted to: ${:.2f}".format(last_bill))
            print("="*100)
            return(last_bill)
        previous_bill = input_bill()
        choice = input("Please verify if last month's electricity bill is correct: Y for Yes, N for No: ")
        choice_up = choice.upper()
        while choice_up != "Y" and choice_up != "N":
            print("Please enter Y for Yes or N for No")
            choice = input("Please verify if last month's electricity bill is correct: Y for Yes, N for No: ")
            choice_up = choice.upper()
        if choice_up == "Y":
            return previous_bill
        elif choice_up == "N":
            previous_bill = last_mth_bill()
            return previous_bill
    
    
    def input_budget():
        budget=0
        while True:
            try: 
                budget=float(input("Enter your current monthly budget for your electricity bill: "))
                if budget <=0:
                    print("Please enter a valid amount")
                    continue
                else:
                    break 
            except ValueError:
                print("Please enter a valid amount ")
                continue
        print("The budget you have entered in is: ${:.2f}".format(budget))
        print("="*100)
        return(budget)
        
    appliance_dict = {"A": ["Aircon", 0],
    "B": ["Desktop",0],
    "C": ["Electric Stove",0] ,
    "D": ["Fan",0], 
    "E": ["Hairdryer",0], 
    "F": ["Laptop",0],
    "G": ["Light Bulbs",0], 
    "H": ["Oven",0], 
    "I": ["Phone Charger",0], 
    "J": ["Refrigerator",0], 
    "K": ["Rice Cooker",0], 
    "L": ["Television",0], 
    "M": ["Toaster",0], 
    "N": ["Water Heater",0],
    "O": ["Other Appliances",0],
    "P": ["I'm done, generate the report.",0] }
    
    
    
    appliancekey = []
    hourskey = []
    numitemskey = []
    def hoursinput():
        while True:
            print("Appliance                         Total Hours")
            print("---------                         -----------")
            for key, value in appliance_dict.items():
                print("{:s}: {:<30s}: {:>10.2f}". format(key, value[0],value[1]))
            print("="*100)
            print("""Now, we will collect data on the number and hours used of each appliance. 
    The table above will be updated with each entry. 
     """)
            appliance = input("Enter letter that symbolises the electric appliance: ").upper()
            if appliance == "P":
                return list(dict.fromkeys(appliancekey))
                break
            elif appliance in appliance_dict.keys():
                appliancekey.append(appliance)
                while True:
                    num_items = input("Enter number of such electricity appliance that you have: ")
                    try:
                        num_items = float(num_items)
                        if num_items <= 0:
                            print("Number needs to be more than 0! Try again.")
                            continue
                        else:
                            numitemskey.append(num_items)
                            break
                    except ValueError: 
                        print("Invalid input. Try again.")
                        continue    
                while True:
                    hours = input("Enter number of hours the electric appliance has been used per day: ")
                    try:
                        hours= float(hours)
                        if hours <= 0 or hours > 24:
                            print("Number needs to be within 24 hours! Try again.")
                            continue
                        else:
                            hourskey.append(hours)
                            appliance_dict[appliance][1] += num_items*hours
                            print("Record saved.")
                            print("="*100)
                            break
                    except ValueError: 
                        print("Invalid input. Try again.")
                        continue
            else:
                print("Invalid appliance key. Try again")
    
    
    
    
    def calc():
        for each in appliancekey:
            appliancekey[appliancekey.index(each)] = appliance_dict[each]
        
        wattskey = [item[0] for item in appliancekey]
        for each in wattskey:
            wattskey[wattskey.index(each)]=appliancekwh[each]  
        print("=" * 100)
        print("")
        print("Total amount of kW of electricity used in a month:{:.2f} kW".format(sum(a*b*c for a,b,c in zip(wattskey,hourskey,numitemskey))*30))
        print("")
        print("Total price of Electricity consumed in the month: ${:.2f}".format(sum(a*b*c for a,b,c in zip(wattskey,hourskey,numitemskey))*30*providerprice/100))
        return wattskey, [a*b*c*30 for a,b,c in zip(wattskey,hourskey,numitemskey)], sum(a*b*c for a,b,c in zip(wattskey,hourskey,numitemskey))*30*providerprice/100
    
    
    def topthree():
        topusage = sorted(zip(totalkwh, [item[0] for item in appliancekey]), reverse=True)[:3]
        print("="*100)
        print("Your top 3 consumptions are:")
        for each in topusage:
            print("{:.2f} kWh of electricity for {}, costing ${:.2f} and takes up {:.2f}% of your total bill\n".format(each[0],each[1],each[0]*providerprice/100, each[0]*providerprice/currentbill))
        return topusage
    
    def comparison_kwh(totalkwh):
        mark = benchmark[houseinfo]
        consumption = sum(totalkwh)
        if consumption < mark:
            difference = float(mark - consumption)
            result = "Good job!! You have managed to consume below the national average by {:,.2f} kWh".format(difference)
        elif consumption == mark:
            result = "Good job!! You have managed to consume within the national average"
        elif consumption > mark:
            difference = float(consumption - mark)
            result = "You have exceeded the national average by {:,.2f}kWh!".format(difference)
        return result
        
    def comparison_budget(bill_calculated, budget):
        if bill_calculated > budget:
            difference = bill_calculated - budget
            budget_result = "You have exceeded the current monthly budget by ${:,.2f}!".format(difference)
        elif bill_calculated == budget:
            budget_result = "You have met your current monthly budget. Good Job!!!"
        elif bill_calculated < budget:
            difference = budget - bill_calculated
            budget_result = "Good job! You have met the current monthly budget by spending ${:,.2f} less".format(difference)
        return budget_result
    
    def comparison_month(this_month, last_month):
        if this_month > last_month:
            difference = this_month - last_month
            month_result = "You have consumed more electricity in this month than last month! This month's bill increased by ${:,.2f}".format(difference)
        elif this_month == last_month:
            month_result = "This month's consumption is the same as last month's!"
        elif this_month < last_month:
            difference = last_month - this_month
            month_result = "Good job! You have reduced your electricity bill for this month by ${:,.2f}".format(difference)
        return month_result
    
    def feedback(topusage):
        counta=0
        countb=0
        countc=0
    
        todays_date = datetime.date.today()
        with open("electricity consumption.csv") as suggestion:
            suggestion= csv.reader(suggestion)
            allsuggestion={}
            appliancekey2={}
            for each in suggestion:
                allsuggestion[each[3]] = each[2]
                appliancekey2[each[3]] = each[0]
    
    
      
        totalvol=len(topusage)
        conversionlist=[]
        for countc in range(totalvol):
            conversionlist.append(topusage[countc][1])
            countc +=1
    
        userusage=[]
    
        for countb in range(totalvol):
            userusage.append(topusage[countb][0])
            countb +=1
            
        result=[]
    
        for each in conversionlist:
            keys = [k for k, v in appliancekey2.items() if v == each]
            result.append(keys[0])
    
         
      
           
        for x in result:
            print("\nSuggestion for {} : {}".format(appliancekey2[x],allsuggestion[x]))
    
        print("Suggestions are printed to : User Advice {}.xlsx for your reference".format(session))
    
        outbook=xls.Workbook("{} User Advice {}.xlsx".format(name,session))
        outsheet=outbook.add_worksheet("Instructions")
    
    
        outsheet.write("A1","Appliance")
        outsheet.write("B1","Advice")
        outsheet.write("L1","Top 3 Usage")
        outsheet.write("M1",todays_date.year-1)
        outsheet.write("N1","Projected")
        outsheet.write("A7","Summary of electricity budgeting")
        outsheet.write("A8","Current Electricity Bill")
        outsheet.write("D8",currentbill)
        outsheet.write("A9","Budgeted Electricity Bill")
        outsheet.write("D9",budgettotal)
        outsheet.write("A10","Budget Result")
        outsheet.write("D10",budget_result)
        outsheet.write("A11","Status")
        outsheet.write("B11",kwh_compare)
        outsheet.write("A13","General Electricity Saving Tips")
        outsheet.write("A14","1. Remember to turn off appliances when not in use")
        outsheet.write("A15","2. Off Your Mains when not using")
        outsheet.write("A16","3. Switch to Eco-Friendly Devices")
        outsheet.write("T1","Consumption of electricity measured in kWh")
        outsheet.write("T2","Projections based on uptake of advice")
        outsheet.write("L13","Why must we save electricity?")
        outsheet.write("L14","Electricity wastage contributes to climate change, which directly leads to the rising sea levels and global warming. T_T")
        outsheet.write("L15","Start saving electricity today! :)")
    
        
    
        for year in range(0,5):
            excel="MOPQR"
            outsheet.write(0,14+year,todays_date.year+year)
            outsheet.write_formula(1,14+year,"={}2*.90".format(excel[year]))
            outsheet.write_formula(2,14+year,"={}3*.90".format(excel[year]))
            outsheet.write_formula(3,14+year,"={}4*.90".format(excel[year]))
    
        for item in result:
            counta +=1
            outsheet.write(counta,0,appliancekey2[item])
            outsheet.write(counta,11,appliancekey2[item])
            outsheet.write(counta,12,userusage[counta-1])
            outsheet.write(counta,1,allsuggestion[item])
        
        chart1 = outbook.add_chart({'type': 'line'})
        
        chart1.add_series({
            'name':       '=Instructions!$L$2',
            'categories': '=Instructions!$O$1:$S$1',
            'values':     '=Instructions!$O$2:$S$2',
        })
    
        chart1.add_series({
            'name':       '=Instructions!$L$3',
            'categories': '=Instructions!$O$1:$S$1',
            'values':     '=Instructions!$O$3:$S$3',
        })
        chart1.add_series({
            'name':       '=Instructions!$L$4',
            'categories': '=Instructions!$O$1:$S$1',
            'values':     '=Instructions!$O$4:$S$4',
        })
    
        chart1.set_title ({'name': 'Top 3 Usage Projections'})
        chart1.set_x_axis({'name': 'Year'})
        chart1.set_y_axis({'name': 'Usage (KWH)'})
    
        chart1.set_style(10)
    
    
        outsheet.insert_chart('L18', chart1, {'x_offset': 5, 'y_offset': 5})
    
    
        print("="*100)    
        print("\nGeneral Electricity Saving  \n 1. Remember to turn off appliances when not in use \n 2. Off Your Mains when not using \n 3. Switch to Eco-Friendly Devices")
    
        outbook.close()
    
    def cont():
        con = input("Would you like to restart? Y/N:").upper()
        while con != "Y" and con != "N":
            print("Please enter Y for Yes or N for No:")
            con = input("Would you like to restart? Y/N:").upper()
        return con



    
    
    house=homecheck()
    if house == 1:
        houseinfo= homedetails()
        providerprice = provider_select()
        lastbill = last_mth_bill()   
        budgettotal= input_budget()
        appliancekey = hoursinput()   
        wattskey, totalkwh, currentbill = calc()
        topusage = topthree()
        kwh_compare = comparison_kwh(totalkwh)
        print("="*100)
        print(f"\n{kwh_compare}")
        budget_result = comparison_budget(currentbill,budgettotal)
        print(f"\n{budget_result}")
        month_compare = comparison_month(currentbill, lastbill)
        print(f"\n{month_compare}")
        feedback(topusage)
        con = cont()
        if con == "Y":
            main()
        else:
            print("Goodbye  [^._.^]ﾉ ")
        
        
                
        
    
    if house == 2:
        print("Goodbye  [^._.^]ﾉ ")

    
main()        
    

