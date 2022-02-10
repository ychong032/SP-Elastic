import json
from elastic_enterprise_search import WorkplaceSearch
import os
import requests
import sys
from requests_ntlm3 import HttpNtlmAuth

# Initialise Workplace Search URL, content source ID and access token. Change these as necessary.
WORKPLACE_SEARCH_URL = "http://192.168.100.5:3002/"
CONTENT_SOURCE_ID = "61f7ab68a861d849f72239f8"
CONTENT_SOURCE_TOKEN = "9c9045303b6fca01d1ca6fb55c24e336ecbda8744ad576f821c212f362338e35"

# Initialise the Workplace Search client.
workplace_search = WorkplaceSearch(WORKPLACE_SEARCH_URL)


# Assign permissions to item by using group names as permission names.
def define_item_permissions(item_properties):
    dict_list = []
    item_id, item_name, list_name, item_url, attached_content, item_modified, item_modified_by, item_content_text = item_properties["details"]
    group_list = item_properties["permissions"]
    temp_dict = {"id": item_id, "_allow_permissions": group_list, "name": item_name, "list_name": list_name, "item_url": item_url, "attached_content": attached_content,
                 "modified": item_modified, "modified_by": item_modified_by, "content": item_content_text}
    dict_list.append(temp_dict)
        
    response = workplace_search.index_documents(content_source_id=CONTENT_SOURCE_ID,
                                                http_auth=CONTENT_SOURCE_TOKEN,
                                                documents=dict_list)
    print("Define item permissions: ", json.dumps(response, indent=4)) 


# Clear existing user permissions.
def clear_user_permissions(groups):
    try:
        all_permissions = workplace_search.list_permissions(content_source_id=CONTENT_SOURCE_ID, http_auth=CONTENT_SOURCE_TOKEN)['results']
        for dictionary in all_permissions:
            response = workplace_search.remove_user_permissions(content_source_id=CONTENT_SOURCE_ID,
                                                                http_auth=CONTENT_SOURCE_TOKEN,
                                                                user=dictionary['user'],
                                                                body={
                                                                    "permissions": dictionary['permissions']
                                                                })
            print(response)
    except Exception as e:
        print(e)
            

# Assign permissions to users by using their SharePoint group name as the permission name.  
def assign_user_permissions(groups):
    print("First list item --> clearing permissions from all SharePoint Groups...")
    clear_user_permissions(groups)
        
    print("Assigning permissions to SharePoint users...")
    for group in groups:
        for username in groups[group]:
            response = workplace_search.add_user_permissions(content_source_id=CONTENT_SOURCE_ID,
                                                             http_auth=CONTENT_SOURCE_TOKEN,
                                                             user=username,
                                                             body={
                                                                 "permissions": [group]
                                                             })
            print("Assigned user permissions:", json.dumps(response, indent=4))


# Convert the JSON string arguments into dictionaries.
groups = json.loads(sys.argv[1])
item_properties = json.loads(sys.argv[2])
is_first_item = sys.argv[3]

# Display the current list item's details and the groups that have access to the item.
print("Item name:", item_properties["details"][1])
print("Permitted groups:", item_properties["permissions"])

# Clear and assign permissions to SharePoint users.
if is_first_item == "True":
    assign_user_permissions(groups)

define_item_permissions(item_properties)


