import win32com.client

#Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = inbox folder

#Retrieve emails from the inbox
emails = inbox.Items

#Iterate over each email
for email in emails:
    subject = email.Subject
    body = email.Body

    #Create a file
    filename = f"{subject}.txt"
    file_path = f"C:/Users/canti/OneDrive/Desktop/emails/{filename}"  # Set your desired file path here

    #Write the email content to the file
    with open(file_path, "w") as file:
        file.write(body)

    print(f"File '{filename}' created.")

#Disconnect
namespace = None
outlook = None