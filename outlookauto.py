import os
import win32com.client
from docx import Document

options = set([
    "OPTIONS FOR SUBJECT TEMPLATE: ", "INFORMATION IN THIS SECTION IS CONFIDENTIAL AND HAD TO BE REMOVED"
])

def read_word_file(file_path):
    doc = Document(file_path)
    data = {}
    
    print("Options: " + ", ".join(options))
    email_subject = input("Email Subject: ")
    email_addresses = input("Email addresses (comma-separated): ").split(',')
    
    # Debugging print statements
    print(f"Email Subject: {email_subject}")
    print(f"Email Addresses: {email_addresses}")
    
    paragraphs = doc.paragraphs
    start_collecting = False
    collected_text = []
    
    for para in paragraphs:
        text = para.text.strip()
        print(f"r: {text}")
        if email_subject in text.lower(): 
            start_collecting = True 
        if start_collecting:
            collected_text.append(text)
            if text.endswith("strictly prohibited."):
                break
    
    data[email_subject] = "\n".join(collected_text)
    
    return data, email_subject, email_addresses

def prepare_email(data, email_subject, email_addresses):
    print("Preparing emails...")
    ol = win32com.client.Dispatch('Outlook.Application')
    drafts = []
    print("Preparing emails2...")
    for email_address in email_addresses:
        email_address = email_address.strip()
        if not email_address:
            continue
        
        mail = ol.CreateItem(0)
        mail.Subject = email_subject
        mail.To = email_address
        mail.Body = data[email_subject]
        
        save_path = os.path.join(os.getcwd(), f'email_draft_{email_address}.msg')
        mail.SaveAs(save_path, 3)  # Save as .msg file
        drafts.append(save_path)
        print(f"Email draft saved for {email_address} at {save_path}. Please review and send manually.")
        
    return drafts

def send_all_drafts(drafts):
    ol = win32com.client.Dispatch('Outlook.Application')
    
    for draft_path in drafts:
        try:
            # Open the draft email from the saved .msg file
            mail = ol.CreateItemFromTemplate(draft_path)
            mail.Send()
            print(f"Draft sent from {draft_path}")
        except Exception as e:
            print(f"Failed to send draft from {draft_path}: {e}")

def main():
    file_path = 'Outlook Signatures for MyChart 06242023.docx'
    data, email_subject, email_addresses = read_word_file(file_path)
    drafts = prepare_email(data, email_subject, email_addresses)
    
    send_option = input("Do you want to send all drafts automatically? (yes/no): ").strip().lower()
    
    if send_option == "yes":
        send_all_drafts(drafts)
    else:
        print("Drafts prepared. Please review and send manually.")
    
if __name__ == "__main__":
    main()
