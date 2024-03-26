import csv
import requests

def search_email_by_content(content):
    matching_emails = []
    with open('fetched_emails.csv', 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if content in row[3]:  # Check if the content is present in the email body (assuming it's in the fourth column)
                matching_emails.append(row)
    return matching_emails

def send_email_to_webhook(email_data, webhook_url):
    email_text = ', '.join(email_data)
    payload = {
        "msg_type": "text",
        "content": {
            "text": f"Email found: {email_text}"
        }
    }
    response = requests.post(webhook_url, json=payload)
    if response.status_code == 200:
        print("Email sent to webhook successfully.")
    else:
        print(f"Failed to send email to webhook. Status code: {response.status_code}")

# Example usage
if __name__ == "__main__":
    query_content = "SCQ362KKJV"ANY UNIQUE IDENTIFIER
    webhook_url = ""#WEBHOOK URL
    
    matching_emails = search_email_by_content(query_content)
    if matching_emails:
        for email in matching_emails:
            send_email_to_webhook(email, webhook_url)
    else:
        print("No emails found matching the specified content.")

