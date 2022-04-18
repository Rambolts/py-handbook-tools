from docusign import DocusignConnector

docusign_data = {
    'integration_key' : '',
    'user_id' : '',
    'authorization_server' : '',
    'base_path' : ''
}

api_client, account_id = DocusignConnector(
    docusign_data['integration_key'],
    docusign_data['user_id'],
    docusign_data['authorization_server'],
    docusign_data['base_path']
).get_connection()

print(f"api_client: {api_client}\naccount_id: {account_id}")