import datetime
from tqdm import tqdm

def process_messages(messages, service):
    filtered_data = []

    for msg in tqdm(messages, desc="Downloading emails", unit="email"):
        msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
        headers = msg_data['payload']['headers']

        def get_header_value(header_name, default_value="No Value"):
            return next((header['value'] for header in headers if header['name'] == header_name), default_value)
        
        sender = get_header_value('From', "Unknown Sender")
        subject = get_header_value('Subject', "No Subject")
        date = get_header_value('Date', "No Date")

        filtered_data.append({
            'From': sender,
            'Subject': subject,
            'Date': date,
        })
    return filtered_data