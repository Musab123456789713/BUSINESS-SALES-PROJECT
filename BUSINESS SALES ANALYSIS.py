import os
import pandas as pd
import matplotlib.pyplot as plt
import pyttsx3 
import random

os.chdir("c:/Users/HP/Documents/PYTHON PROJECTS")

data = pd.read_excel("BUSINESS SALES DATA.xlsx")

def Robo_Speaker(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)
    engine.say(audio)
    engine.runAndWait()

def data_cleaning():
    data.drop_duplicates()
    Transaction = data['nsaction_ID']
    ID = []
    for i in Transaction:
        ID.append(f"Transaction {i}.")
    data.index = ID
    data.drop('nsaction_ID', axis=1, inplace=True)
    data.sort_values(by='Price', ascending=False, inplace=True)
    Robo_Speaker("Your Data is clean now")

def data_transformation():
    Total_Transaction = data['Quantity']*data['Price']
    data['Transaction_Amount'] = Total_Transaction
    Robo_Speaker("Your Data is Transform now")
    
def data_analysis():
    total_sales_of_each_Product = data.groupby('Product')['Transaction_Amount'].sum()
    Product = pd.DataFrame(total_sales_of_each_Product)
    Product.sort_values(by='Transaction_Amount', ascending=False, inplace=True)
    print (Product,'\n')
    Top_3_Product = Product.head(3)
    Robo_Speaker("Following are the Top 3 products")
    print ("Following are the Top 3 products")
    print (Top_3_Product, '\n')

    Robo_Speaker("Do you want to analyze cities")
    user = str(input("Do you want to analyze cities: ")).lower()
    if 'yes' in user:
        City = data.groupby('City')['Transaction_Amount'].sum()
        City_data = pd.DataFrame(City)
        City_data.sort_values(by='Transaction_Amount', ascending=False, inplace=True)
        Top_Cities = City_data.head(3)
        Robo_Speaker("Following are the Top 3 cities with highest sales")
        print ("Following are the Top 3 cities with highest sales")
        print (Top_Cities, '\n')
    else:
        print ("OK, Hava a good day.")
        Robo_Speaker("OK, Hava a good day.")
    
    Robo_Speaker("Do you want to perform Descriptive Statistics")
    user = str(input("Do you want to perform Descriptive Statistics: ")).lower()
    if 'yes' in user:
        total_sales = data['Transaction_Amount'].sum()
        print ("\n Your Total Sale is ", total_sales, '\n')
        average_sales = data['Transaction_Amount'].mean()
        print ("Your Average Sale is ", average_sales, '\n')
        common_payment = data['Payment_Method'].mode()
        print ("The Most Common Payment Method is ", common_payment, '\n')
    else:
        print ("OK, Hava a good day.")
        Robo_Speaker("OK, Hava a good day.")

def data_visualization():
    Sales_over_payment_method = data.groupby('Payment_Method')['Transaction_Amount'].sum()
    barchart_data = pd.DataFrame(Sales_over_payment_method)
    plt.subplot(1,2,1)
    plt.bar(barchart_data.index, barchart_data['Transaction_Amount'], width=0.5, color='blue')
    plt.xlabel("Payment Method")
    plt.ylabel("Amount")
    
    plt.subplot(1,2,2)
    City = data.groupby('City')['Transaction_Amount'].sum()
    City_data = pd.DataFrame(City)
    City_data.sort_values(by='Transaction_Amount', ascending=False, inplace=True)
    Top_Cities = City_data.head(5)
    plt.pie(Top_Cities['Transaction_Amount'], labels=Top_Cities.index, autopct='%0.1f%%', colors=['red','blue','green','yellow','orange'])
    
    plt.show()

Robo_Speaker('WELCOME TO BUSINESS SALES ANAYLSIS SYSTEM')
print ('WELCOME TO REAL ESTATE SALES ANAYLSIS SYSTEM')
options = ['Clean the data', 'Give insights', 'Visualize the data']
Robo_Speaker('Following are the analytical option we provide')
print ('Following are the analytical option we provide:\n')
for i in options:
    print (i)
while True:
    try:
        Robo_Speaker("Enter a menue option or exit")
        user = input('\nEnter a menue option or exit: ').lower()
        if 'clean' in user:
            data_cleaning()
        
        elif 'insights' in user:
            data_transformation()
            data_analysis()
            print ()

        elif 'visualize' in user:
            data_visualization()

        elif 'exit' in user:
            break
        else:
            print ('Wrong Input')
            greeting = ['Have a good day', 'Take care of your self', 'Take it easy Bro', 'Have a nice day']
            greeting = random.choice(greeting)
            print (greeting)   
    except ValueError:
        print ("Please enter a string value")

