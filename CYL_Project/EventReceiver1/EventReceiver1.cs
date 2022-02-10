using Microsoft.SharePoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Text;
using TikaOnDotNet.TextExtraction;

namespace CYL_Project.EventReceiver1
{
    /// <summary>
    /// This class handles events related to list items.
    /// </summary>
    public class EventReceiver1 : SPItemEventReceiver
    {
        
        /// <summary>
        /// Executes when an item is being deleted.
        /// </summary>
        /// <remarks>
        /// Only the item's ID is needed. However, for logging purposes, get the details as per usual.
        /// </remarks>
        public override void ItemDeleting(SPItemEventProperties properties)
        {
            base.ItemDeleting(properties);
            String path = @"C:\Users\Administrator\Documents\Logs\item_deleting.txt"; // File for logging and debugging.

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----Item Being Deleted on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm") + "-----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    // Get item details. Modify to obtain desired details.
                    string action = Action.deleting.ToString();
                    string itemName = properties.ListItem.Name;
                    string listName = properties.ListTitle;
                    string itemId = properties.ListItem.ID.ToString();
                    // string itemUrl = web.Site.MakeFullUrl(properties.List.DefaultDisplayFormUrl + "?ID=" + itemId);       // This is the more proper way of getting the item URL.
                    string itemUrl = "http://192.168.100.7:3877/" + properties.List.DefaultDisplayFormUrl + "?ID=" + itemId; // Replace the URL prefix as necessary.
                    
                    DateTime itemModifiedDate = (DateTime)properties.ListItem["Modified"];
                    string itemModified = itemModifiedDate.ToString("dd/MM/yyyy, HH:mm");
                    
                    string strUserValue = (string)properties.ListItem["Modified By"]; // Returns a string with the format ID#User_Display_Name
                    int intIndex = strUserValue.IndexOf("#");
                    string itemModifiedBy = strUserValue.Substring(intIndex + 1); // Get only the display name of the user

                    // The "Content" field is in HTML format by default. Convert it to plain text.
                    // Additionally, escape all double quotes as they will otherwise be removed by the command line when passed to the Python script.
                    var itemContentField = properties.List.Fields.GetField("Content");
                    var itemContent = properties.ListItem[itemContentField.Id];
                    string itemContentText = itemContentField.GetFieldValueAsText(itemContent).Replace(@"""", @"\""");

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Name: " + itemName);
                        sw.WriteLine("List Name: " + listName);
                        sw.WriteLine("Item URL: " + itemUrl);
                        sw.WriteLine("Item ID: " + itemId);
                        sw.WriteLine("Item Modified: " + itemModified);
                        sw.WriteLine("Item Modified By: " + itemModifiedBy);
                        sw.WriteLine("Attachments: ");
                    }

                    SPAttachmentCollection attachments = properties.ListItem.Attachments;
                    if (attachments.Count > 0)
                    {
                        for (int i = 0; i < attachments.Count; i++)
                        {
                            string attachmentUrl = attachments.UrlPrefix + attachments[i];
                            SPFile attachedFile = web.GetFile(attachmentUrl);
                            string attachedName = attachedFile.Name;
                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("\t" + (i + 1) + ". " + attachedName);
                            }
                        }
                    }

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine();
                    }

                    CallPython(action, itemName, listName, itemUrl, itemId, "", itemModified, itemModifiedBy, itemContentText);
                }
                catch (Exception ex)
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Deleting Exception!");
                        sw.WriteLine(ex);
                        sw.WriteLine("");
                    }
                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Something went wrong when deleting this item.";
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Executes when an item was added.
        /// </summary>
        /// <remarks>
        /// This function essentially does the same thing as ItemUpdated, except both functions are triggered by different events.
        /// </remarks>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
            String path = @"C:\Users\Administrator\Documents\Logs\item_added.txt";

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----Item Added on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm") + "-----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    // Get item details. Modify to obtain desired details.
                    string action = Action.added.ToString();
                    string itemName = properties.ListItem.Name;
                    string listName = properties.ListTitle;
                    string itemId = properties.ListItem.ID.ToString();
                    // string itemUrl = web.Site.MakeFullUrl(properties.List.DefaultDisplayFormUrl + "?ID=" + itemId);
                    string itemUrl = "http://192.168.100.7:3877/" + properties.List.DefaultDisplayFormUrl + "?ID=" + itemId;

                    DateTime itemModifiedDate = (DateTime)properties.ListItem["Modified"];
                    string itemModified = itemModifiedDate.ToString("dd/MM/yyyy, HH:mm");

                    string strUserValue = (string)properties.ListItem["Modified By"];
                    int intIndex = strUserValue.IndexOf("#");
                    string itemModifiedBy = strUserValue.Substring(intIndex + 1);

                    // The "Content" field is in HTML format by default. Convert it to plain text.
                    // Additionally, escape all double quotes as they will otherwise be removed by the command line when passed to the Python script.
                    var itemContentField = properties.List.Fields.GetField("Content");
                    var itemContent = properties.ListItem[itemContentField.Id];
                    string itemContentText = itemContentField.GetFieldValueAsText(itemContent).Replace(@"""", @"\""");

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Name: " + itemName);
                        sw.WriteLine("List Name: " + listName);
                        sw.WriteLine("Item URL: " + itemUrl);
                        sw.WriteLine("Item ID: " + itemId);
                        sw.WriteLine("Item Modified: " + itemModified);
                        sw.WriteLine("Item Modified By: " + itemModifiedBy);
                        sw.WriteLine("Attachments: ");
                    }

                    SPAttachmentCollection attachments = properties.ListItem.Attachments;
                    TextExtractor textExtractor = new TextExtractor();
                    StringBuilder stringBuilder = new StringBuilder();
                    if (attachments.Count > 0)
                    {
                        // For each attachment, extract the contents and append it to a string.
                        for (int i = 0; i < attachments.Count; i++)
                        {
                            string attachmentUrl = attachments.UrlPrefix + attachments[i];
                            SPFile attachedFile = web.GetFile(attachmentUrl);
                            string attachedName = attachedFile.Name;
                            byte[] fileBytes = attachedFile.OpenBinary();
                            string fileString = textExtractor.Extract(fileBytes).Text.Trim();

                            stringBuilder.AppendLine(fileString);
                            stringBuilder.AppendLine("---------"); // Indicates the end of 1 attachment.
                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("\t" + (i + 1) + ". " + attachedName);
                            }
                        }
                    }

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine();
                    }

                    CallPython(action, itemName, listName, itemUrl, itemId, stringBuilder.ToString(), itemModified, itemModifiedBy, itemContentText);
                }
                catch (Exception ex)
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Added Exception!");
                        sw.WriteLine(ex);
                        sw.WriteLine("");
                    }
                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Something went wrong after adding this item.";
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Executes when an attachment is added to an item.
        /// </summary>
        /// <remarks>
        /// There may not be a need for this function as clicking "Save" after attaching a file to an item on SharePoint triggers the ItemUpdated function too.
        /// ItemUpdated seems to override ItemAttachmentAdded as a result.
        /// </remarks>
        public override void ItemAttachmentAdded(SPItemEventProperties properties)
        {
            base.ItemAttachmentAdded(properties);

            String path = @"C:\Users\Administrator\Documents\Logs\item_attachment_added.txt";

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----Attachment Added on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "-----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    string action = Action.attachment_added.ToString();
                    string itemName = properties.ListItem.Name;
                    string listName = properties.ListTitle;
                    string itemId = properties.ListItem.ID.ToString();
                    // string itemUrl = web.Site.MakeFullUrl(properties.List.DefaultDisplayFormUrl + "?ID=" + itemId);
                    string itemUrl = "http://192.168.100.7:3877/" + properties.List.DefaultDisplayFormUrl + "?ID=" + itemId;

                    DateTime itemModifiedDate = (DateTime)properties.ListItem["Modified"];
                    string itemModified = itemModifiedDate.ToString("dd/MM/yyyy, HH:mm");

                    string strUserValue = (string)properties.ListItem["Modified By"];
                    int intIndex = strUserValue.IndexOf("#");
                    string itemModifiedBy = strUserValue.Substring(intIndex + 1);

                    // The "Content" field is in HTML format by default. Convert it to plain text.
                    // Additionally, escape all double quotes as they will otherwise be removed by the command line when passed to the Python script.
                    var itemContentField = properties.List.Fields.GetField("Content");
                    var itemContent = properties.ListItem[itemContentField.Id];
                    string itemContentText = itemContentField.GetFieldValueAsText(itemContent);

                    // Get attached file and convert to bytes.
                    SPAttachmentCollection attachments = properties.ListItem.Attachments;
                    string attachmentUrl = attachments.UrlPrefix + attachments[attachments.Count - 1];
                    SPFile attachedFile = web.GetFile(attachmentUrl);
                    string attachedName = attachedFile.Name;

                    /*
                     * The code below is commented out because it is already triggered by ItemUpdated.
                     * See the remarks of ItemAttachmentAdded.
                     */

                    // byte[] fileBytes = attachedFile.OpenBinary();

                    // TextExtractor textExtractor = new TextExtractor();
                    // string fileString = textExtractor.Extract(fileBytes).Text.Trim();

                    // CallPython(action, itemName, itemTitle, listName, relUrl, itemId, fileString);

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Name: " + itemName);
                        sw.WriteLine("List Name: " + listName);
                        sw.WriteLine("Item URL: " + itemUrl);
                        sw.WriteLine("Item ID: " + itemId);
                        sw.WriteLine("Item Modified: " + itemModified);
                        sw.WriteLine("Item Modified By: " + itemModifiedBy);
                        sw.WriteLine("Attachment Name: " + attachedName);
                        sw.WriteLine("Attachment URL: " + attachmentUrl);
                        sw.WriteLine();
                    }
                }
                catch (Exception ex)
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Attachment Added Exception!");
                        sw.WriteLine(ex);
                        sw.WriteLine("");
                    }
                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Something went wrong after attaching something to this item.";
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Executes when an item was updated.
        /// </summary>
        /// <remarks>
        /// This function executes whenever the "Save" button is clicked after editing a list item on SharePoint.
        /// This includes adding attachments to an item.
        /// This function seems to overwrite any changes made by ItemAttachmentAdded.
        /// </remarks>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);

            String path = @"C:\Users\Administrator\Documents\Logs\item_updated.txt"; // File for logging and debugging.

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----Item Updated on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "-----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    // Get the list item's details. Add or remove details as desired.
                    string action = Action.updated.ToString();
                    string itemName = properties.ListItem.DisplayName;
                    string listName = properties.ListTitle;
                    string itemId = properties.ListItem.ID.ToString();
                    // string itemUrl = web.Site.MakeFullUrl(properties.List.DefaultDisplayFormUrl + "?ID=" + itemId);       // Proper way of obtaining item URL.
                    string itemUrl = "http://192.168.100.7:3877/" + properties.List.DefaultDisplayFormUrl + "?ID=" + itemId; // Replace the URL prefix as necessary.

                    DateTime itemModifiedDate = (DateTime)properties.ListItem["Modified"];
                    string itemModified = itemModifiedDate.ToString("dd/MM/yyyy, HH:mm");

                    string strUserValue = (string)properties.ListItem["Modified By"]; // Gets a string that consists of the user's ID and display name.
                    int intIndex = strUserValue.IndexOf("#");
                    string itemModifiedBy = strUserValue.Substring(intIndex + 1); // Gets the user's display name without the ID.

                    // The "Content" field is in HTML format by default. Convert it to plain text.
                    // Additionally, escape all double quotes as they will otherwise be removed by the command line when passed to the Python script.
                    var itemContentField = properties.List.Fields.GetField("Content");
                    var itemContent = properties.ListItem[itemContentField.Id];
                    string itemContentText = itemContentField.GetFieldValueAsText(itemContent).Replace(@"""", @"\""");

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Name: " + itemName);
                        sw.WriteLine("List Name: " + listName);
                        sw.WriteLine("Item URL: " + itemUrl);
                        sw.WriteLine("Item ID: " + itemId);
                        sw.WriteLine("Item Modified: " + itemModified);
                        sw.WriteLine("Item Modified By: " + itemModifiedBy);
                        sw.WriteLine("Attachments: ");
                    }

                    SPAttachmentCollection attachments = properties.ListItem.Attachments;
                    TextExtractor textExtractor = new TextExtractor();
                    StringBuilder stringBuilder = new StringBuilder();
                    if (attachments.Count > 0)
                    {
                        // For each attachment, extract the contents and append it to a string.
                        for (int i = 0; i < attachments.Count; i++)
                        {
                            string attachmentUrl = attachments.UrlPrefix + attachments[i];
                            SPFile attachedFile = web.GetFile(attachmentUrl);
                            string attachedName = attachedFile.Name;
                            byte[] fileBytes = attachedFile.OpenBinary();
                            string fileString = textExtractor.Extract(fileBytes).Text.Trim();

                            stringBuilder.AppendLine(fileString);
                            stringBuilder.AppendLine("---------"); // Indicates the end of 1 attachment.
                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("\t" + (i + 1) + ". " + attachedName);
                            }
                        }
                    }

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine();
                    }

                    CallPython(action, itemName, listName, itemUrl, itemId, stringBuilder.ToString(), itemModified, itemModifiedBy, itemContentText);                    
                }
                catch (Exception ex)
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Updated Exception!");
                        sw.WriteLine(ex);
                        sw.WriteLine("");
                    }
                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Something went wrong after updating this item. Check if the file format of any attachments are unsupported.";
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Calls and executes a Python script for indexing documents to Workplace Search.
        /// </summary>
        private void CallPython(string action, string itemName, string listName, string itemUrl, string itemId, string attachedContent, 
            string itemModified, string itemModifiedBy, string itemContentText)
        {
            // 1) Create process info
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = @"C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe"; // Full path to Python installation. Change this as necessary.

            // 2) Provide script and arguments
            var script = @"Z:\CYL_Project\Python\event_receiver.py"; // Path to Python script. Change this as necessary.
            start.Arguments = $"\"{script}\" \"{action}\" \"{itemName}\" \"{listName}\" \"{itemUrl}\" \"{itemId}\" \"{attachedContent}\" " +
                $"\"{itemModified}\" \"{itemModifiedBy}\" \"{itemContentText}\"";

            // 3) Process configuration
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            start.CreateNoWindow = true;
            start.RedirectStandardError = true;
            using (Process process = Process.Start(start))
            {
                string errors = process.StandardError.ReadToEnd();
                string result = process.StandardOutput.ReadToEnd();

                String path = @"C:\Users\Administrator\Documents\Logs\elastic_log.txt"; // File for logging and debugging.

                if (!File.Exists(path))
                {
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.WriteLine("-----New File-----");
                    }
                }

                using (StreamWriter sw = File.AppendText(path))
                {
                    string eventType = "item " + action;
                    TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                    sw.WriteLine("----" + myTI.ToTitleCase(eventType) + " occurred on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm") + "-----");
                    sw.WriteLine("Item Name: " + itemName);
                    sw.WriteLine("List Name: " + listName);
                    sw.WriteLine("Item URL: " + itemUrl);
                    sw.WriteLine("Item ID: " + itemId);
                    sw.WriteLine("Errors: ");
                    sw.WriteLine(errors);
                    sw.WriteLine();
                    sw.WriteLine("Results: ");
                    sw.WriteLine(result);
                }
            }
        }

        /// <summary>
        /// Denotes the type of event/action that occurred.
        /// </summary>
        private enum Action
        {
            deleting,
            added,
            attachment_added,
            updated
        }
    }
}