list_tables_excel = https://graph.microsoft.com/v1.0/me/drive/items/{Item-id}/workbook/worksheets/Sheet1/tables

add table to excel after session created

worksheet_header = {
        "content-type": "application/json",
        "Authorization": f"Bearer {token['access_token']}"
    }
    worksheet_url = "https://graph.microsoft.com/v1.0/me/drive/items/{Item-id}/workbook/worksheets/Sheet1/tables/add"    

    worksheet_body = {
        "address": "A1:D8",
        "hasHeaders": True,
    }
# create session
graph_session_url = "https://graph.microsoft.com/v1.0/me/drive/root:/graph-api-project.xlsx:/workbook/"
    graph_session_body = {
        "persistChanges": True
    }
    token = auth.get_token_for_user(app_config.SCOPE)
    #print(token['access_token'])
    graph_session_header = {'Authorization': 'Bearer ' + token['access_token'],
                            'content-type': 'application/json'}
    graph_session_response = requests.post(graph_session_url, headers=graph_session_header, json=graph_session_body)
    
    
    


    
    

    # add rows to table
    # obtain table ID through graph explorer
    add_rows_url = "https://graph.microsoft.com/v1.0/me/drive/items/{Item-id}/workbook/worksheets/Sheet1/tables/{table-id}/rows"   

    add_rows_header = {
        "content-type": "application/json",
        "Authorization": f"Bearer {token['access_token']}",
        "Prefer" : "responsd-async"
    }

    add_rows_body = {
        # will append the rows to the table, thus making the number of rows longer to the table
        # the order of values in the array below needs to match the order of the columns in the table
        "values": [
            [12, 22, 33, 1],
            [44, 54, 64, 1]
        ]
   }
    add_rows_response = requests.post(url=add_rows_url, headers=add_rows_header, json=add_rows_body)
    print(add_rows_response.json())
