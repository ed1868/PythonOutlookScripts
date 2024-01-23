import win32com.client
import datetime

def mark_emails_as_unread(days):
    # Calculate the date from the specified number of days ago
    date_from = datetime.datetime.now() - datetime.timedelta(days=days)

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox

    # Filter emails received within the specified date range and mark them as unread
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    for message in messages:
        if message.ReceivedTime >= date_from:
            message.UnRead = True
            print(f"Marked as unread: {message.Subject}")
        else:
            break

def main():
    # Prompt the user to enter the number of days
    days = int(input("Enter the number of days to go back from: "))
    mark_emails_as_unread(days)

if __name__ == "__main__":
    main()
