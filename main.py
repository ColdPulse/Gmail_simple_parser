import creds
import xmlxexp
import mailget
from googleapiclient.discovery import build
from tqdm import tqdm

def main():
    cred = creds.load_credentials()
    service = build('gmail', 'v1', credentials=cred)

    results = []
    next_page_token = None

    pbar = tqdm(desc='Requests from GMail',unit=' rq')

    while True:
        request = service.users().messages().list(userId='me', maxResults=100, pageToken=next_page_token)
        response = request.execute()

        if 'messages' in response:
            for message in response['messages']:
                msg_details = mailget.get_message_details(service, 'me', message['id'])
                results.append(msg_details)
                pbar.update(1)

        next_page_token = response.get('nextPageToken')

        if not next_page_token:
            break

    pbar.close()
    
    xmlxexp.save_to_excel(results)
    print("Загрузка завершена.")


if __name__ == '__main__':
    main()
