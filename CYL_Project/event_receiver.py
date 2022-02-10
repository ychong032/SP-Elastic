import json
from elastic_enterprise_search import WorkplaceSearch
import os
import requests
import sys
from requests_ntlm3 import HttpNtlmAuth

# Initialise Workplace Search URL, custom content source ID and access token.
WORKPLACE_SEARCH_URL = "http://192.168.100.5:3002/"
CONTENT_SOURCE_ID = "61f7ab68a861d849f72239f8"
CONTENT_SOURCE_TOKEN = "9c9045303b6fca01d1ca6fb55c24e336ecbda8744ad576f821c212f362338e35"

# Initialise the Workplace Search client.
workplace_search = WorkplaceSearch(WORKPLACE_SEARCH_URL)


def added_item():
    dict_list = []
    my_dict = {"name": item_name, "list_name": list_name, "item_url": rel_url, "id": item_id, "attached_content": attached_content,
               "modified": item_modified, "modified_by": item_modified_by, "content": item_content_text}
    dict_list.append(my_dict)
    response = workplace_search.index_documents(content_source_id=CONTENT_SOURCE_ID,
                                                http_auth=CONTENT_SOURCE_TOKEN,
                                                documents=dict_list)
    return response


def deleting_item():
    response = workplace_search.delete_documents(content_source_id=CONTENT_SOURCE_ID,
                                                 http_auth=CONTENT_SOURCE_TOKEN,
                                                 document_ids=[item_id])
    return response


# This is the same as added_item(). Can be deleted.
def added_attachment():
    dict_list = []
    my_dict = {"name": item_name, "list_name": list_name, "item_url": rel_url, "id": item_id, "attached_content": attached_content,
               "modified": item_modified, "modified_by": item_modified_by, "content": item_content_text}
    dict_list.append(my_dict)
    response = workplace_search.index_documents(content_source_id=CONTENT_SOURCE_ID,
                                 http_auth=CONTENT_SOURCE_TOKEN,
                                 documents=dict_list)
    return response

action, item_name, list_name, rel_url, item_id, attached_content, item_modified, item_modified_by, item_content_text = sys.argv[1:]

# It is not possible to update individual fields of existing documents in Workplace Search using Python.
# In other words, to update a document with changes, the document must be re-indexed entirely.
if action == "added" or action == "updated":
    response = added_item()
elif action == "deleting":
    response = deleting_item()
elif action == "attachment_added": # Currently, this is never triggered. Refer to the code comments for ItemAttachmentAdded in C#.
    response = added_attachment()

print(response)




    
