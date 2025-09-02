import win32com.client, pyautogui,webbrowser,datetime,time,os

# store your grade check template here with other email things
template="Dear Mother and Father,\n Thank you for taking the time to read this email. I hope your day is well. "
template2 = "\nSincerely,\n name" #replace with name
emailTo="email" #replace with parent emails
emailCC="PLP teacher" #replace with PLP teacher email
print("How are you're grades/What's something that happened this week")
weeklyChange = input(">>")#Type the thing you change each week ie why your grades are bad or something interesting about the week
emailBody=(template+weeklyChange+template2)
url = "https://paccess.mveca.org/Student/Grades"
emailURL = "https://outlook.office365.com/mail/"
#goes to grades
webbrowser.get('windows-default').open(url)

# # waits a second and takes screenshot using pyautogui
time.sleep(3)
image = pyautogui.screenshot()

#gets the date
now = datetime.datetime.now()
date = str(now.strftime("%m-%d"))
#Names the subject with todays date
emailSubject=("Weekly Grade Check " + date)
#Names the file with the date in the image folder (Make sure you have a folder named gradeCheckImages)
fileName = os.path.join("C:/Users/axthelm.2/OneDrive - Dayton Regional STEM School/Courses 2025-2026/Professionism/gradeCheckImages", date + ".png")
#Saves the file to where you set in fileName
image.save(fileName)

#opens outlook app (Not website)
ol=win32com.client.Dispatch('Outlook.Application')
olmailitem=0x0 #size of the new email
#creates a new email
newmail=ol.CreateItem(olmailitem)
#Email info (Variables in the beginning)
newmail.Subject= emailSubject
newmail.To=emailTo
newmail.CC=emailCC
newmail.Body= emailBody
#adds the image
newmail.Attachments.Add(fileName)
#shows the email in the app (Not website)
newmail.Display() 
time.sleep(1)
#opens website
webbrowser.get('windows-default').open(emailURL)
#Check your drafts and it should be in there where you can check that everything looks right and then you can send it
