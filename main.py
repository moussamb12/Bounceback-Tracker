import win32com.client  # Import the library to interact with Outlook
import os  # Import the library for operating system interactions (not used yet, but can be useful later)

def connect_to_outlook():
    """
    Connects to the Outlook application and returns the namespace object.
    """
    try:
        # Create an instance of the Outlook application and get the MAPI namespace
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        print("Connected to Outlook.")  # Confirm successful connection
        return outlook  # Return the Outlook namespace object
    except Exception as e:
        # Print an error message if the connection fails
        print(f"Error connecting to Outlook: {e}")
        return None  # Return None if the connection fails

def create_bounceback_folder(outlook):
    """
    Creates a 'bounceback' folder in the primary inbox if it doesn't already exist.
    """
    try:
        # Get the default Inbox folder (6 is the constant for the Inbox in Outlook)
        inbox = outlook.GetDefaultFolder(6)
        bounceback_folder = None  # Initialize the variable for the bounceback folder

        # Loop through all subfolders in the Inbox to check if 'bounceback' already exists
        for folder in inbox.Folders:
            if folder.Name.lower() == "bounceback":  # Check if the folder name matches (case-insensitive)
                bounceback_folder = folder  # Assign the existing folder to the variable
                break  # Exit the loop as the folder is found

        # If the folder doesn't exist, create it
        if not bounceback_folder:
            bounceback_folder = inbox.Folders.Add("bounceback")  # Create a new folder named 'bounceback'
            print("Created 'bounceback' folder in Outlook.")  # Confirm folder creation
        else:
            print("'bounceback' folder already exists.")  # Inform that the folder already exists

        return bounceback_folder  # Return the bounceback folder object
    except Exception as e:
        # Print an error message if folder creation fails
        print(f"Error creating 'bounceback' folder: {e}")
        return None  # Return None if folder creation fails

def move_bounceback_emails(outlook, bounceback_folder):
    """
    Searches for bounceback emails in the primary inbox and moves them to the 'bounceback' folder.
    """
    try:
        # Get the default Inbox folder
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items  # Get all the messages in the Inbox
        bounceback_count = 0  # Initialize a counter for the number of bounceback emails

        # Loop through all emails in the Inbox
        for message in messages:
            # Check if the email subject contains common bounceback indicators
            if "Undeliverable" in message.Subject or "Delivery Status Notification" in message.Subject:
                message.Move(bounceback_folder)  # Move the email to the 'bounceback' folder
                bounceback_count += 1  # Increment the counter

        # Print the total number of emails moved
        print(f"Moved {bounceback_count} bounceback emails to the 'bounceback' folder.")
    except Exception as e:
        # Print an error message if email processing fails
        print(f"Error moving bounceback emails: {e}")

def main():
    """
    Main function to execute the email bounceback tracker.
    """
    # Step 1: Connect to Outlook
    outlook = connect_to_outlook()
    if not outlook:  # If the connection fails, exit the program
        return

    # Step 2: Create the 'bounceback' folder
    bounceback_folder = create_bounceback_folder(outlook)
    if not bounceback_folder:  # If folder creation fails, exit the program
        return

    # Step 3: Move bounceback emails to the 'bounceback' folder
    move_bounceback_emails(outlook, bounceback_folder)

# Entry point of the script
if __name__ == "__main__":
    main()  # Call the main function to start the program