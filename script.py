import speech_recognition as sr
import openpyxl
import pyttsx3

engine = pyttsx3.init("sapi5")
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    

def getInput(prompt):
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print(prompt)
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
        try:
            text = recognizer.recognize_google(audio)
            return text
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand. Please try again.")
            return None
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return None

def main():
    
    try:
        wb = openpyxl.load_workbook("supermarket_bills.xlsx")
        sheet = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Company Name", "Bill No", "Total Amount"])

    speak("Please enter company name.")
    company_name = getInput("Please enter company name:")
    print(company_name)
    while not company_name:
        speak("Please say the company name again.")
        company_name = getInput("Please say the company name again:")

    speak("Please enter the bill number.")
    bill_no = getInput("Please enter the bill number:")
    print(bill_no)
    while not bill_no:
        speak("Please enter the bill number again")
        bill_no = getInput("Please enter the bill number again:")

    speak("Please enter the total amount")
    total_amount = getInput("Please enter the total amount:")
    print(total_amount)
    while not total_amount:
        speak("Please enter the total amount again.")
        total_amount = getInput("Please enter the total amount again:")

    speak("confirm the data you have entered.")
    confirmation = input("is the data that you've entered correct? (yes/no) ")
    if(confirmation == "yes"):
        sheet.append([company_name, bill_no, total_amount])
        wb.save("supermarket_bills.xlsx")
        speak("Data saved successfully!")
        print("Data saved successfully!")

    elif confirmation == "no":
        print('enter the data manually.')
        company_name = input("Please enter company name:")
        bill_no = int(input("Please enter the bill number:"))
        total_amount = int(input("Please enter the total amount:"))
        sheet.append([company_name, bill_no, total_amount])
        wb.save("supermarket_bills.xlsx")
        speak("Data saved successfully!")
        print("Data saved successfully!")

    else:
        print(f"please enter a valid input.")


if __name__ == "__main__":
    speak("enter the number of entries.")
    a = int(input("enter the number of entries."))
    for item in range(0,a):
        main()