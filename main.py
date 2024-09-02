import creds
import xmlxexp
from googleapiclient.discovery import build

from tqdm import tqdm

def get_message_details(service, user_id, msg_id):
    msg = service.users().messages().get(userId=user_id, id=msg_id, format='full').execute()
    
    headers = msg['payload']['headers']
    
    details = {}
    
    for header in headers:
        if header['name'] == 'From':
            details['From'] = header['value']
        elif header['name'] == 'To':
            details['To'] = header['value']
        elif header['name'] == 'Subject':
            details['Subject'] = header['value']
        elif header['name'] == 'Date':
            details['Date'] = header['value']
    
    details['Snippet'] = msg.get('snippet', '')
    
    return details

def main():
    creds = creds.load_credentials()
    service = build('gmail', 'v1', credentials=creds)

    results = []
    next_page_token = None

    pbar = tqdm(unit='request')

    while True:
        request = service.users().messages().list(userId='me', maxResults=500, pageToken=next_page_token)
        response = request.execute()

        if 'messages' in response:
            for message in response['messages']:
                msg_details = get_message_details(service, 'me', message['id'])
                results.append(msg_details)  # Сохраняем детали письма
                pbar.update(1)

        next_page_token = response.get('nextPageToken')
        
        if not next_page_token:
            break

    pbar.close()
    
    xmlxexp.save_to_excel(results)
    print("Загрузка завершена.")


if __name__ == '__main__':
    main()
