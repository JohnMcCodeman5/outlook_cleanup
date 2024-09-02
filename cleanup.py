import win32com.client
from datetime import datetime, timedelta

def clean_outlook(email, folder_name, days_old, subject_keyword):
    # Use gencache.EnsureDispatch to ensure proper COM interface is used
    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"Failed to connect to Outlook: {e}")
        return 0
    
    # Navigate to the specified folder
    try:
        folder_path = folder_name.split('/')
        folder = namespace.Folders.Item(email)
        for subfolder in folder_path:
            folder = folder.Folders[subfolder]
    except Exception as e:
        print(f"Failed to access the specified folder '{folder_name}': {e}")
        return 0
    
    # Calculate the date threshold
    date_threshold = datetime.now() - timedelta(days=days_old)
    
    # Loop through the emails in the specified folder
    messages = folder.Items
    deleted_count = 0
    try:
        for message in list(messages):
            # Convert to Python datetime
            received_time = message.ReceivedTime
            received_time = datetime(received_time.year, received_time.month, received_time.day,
                                     received_time.hour, received_time.minute, received_time.second)
            
            # Check if the email meets the criteria
            if received_time < date_threshold or subject_keyword.lower() in message.Subject.lower():
                message.Delete()  # Delete the email
                deleted_count += 1
    except Exception as e:
        print(f"Error while processing messages: {e}")
        return deleted_count
    
    return deleted_count
