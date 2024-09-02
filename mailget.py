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
