from fetch_emails import fetch_emails
#from write_to_csv import write_to_csv
from send_lark import send_lark

def main():
    # Fetch emails and extract data
    email_data = fetch_emails()
    
    # Write data to CSV file
    #write_to_csv('emails.csv', email_data)
    
    # Call the send_lark function
    send_lark()

if __name__ == "__main__":
    main()
