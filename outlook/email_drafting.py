import win32com.client as win32

try:
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = Mail item

    mail.To = "Prajwal.Gowda@buildersacademy.com.au"
    mail.Subject = "Test Draft from Python"
    mail.Body = "Hi,\n\nThis is a test draft created using Python.\n\nRegards,\nPrajwal"

    mail.Save()

    print("Success: Draft created in Outlook Drafts.")
except Exception as e:
    print("Error:", e)