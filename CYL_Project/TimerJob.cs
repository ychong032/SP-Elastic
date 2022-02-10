using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint;
using System.IO;
using Newtonsoft.Json;
using System.Diagnostics;
using TikaOnDotNet.TextExtraction; // Used for extracting text from attachments.

// NB: every time this TimerJob is deployed, the SharePoint Timer Service in services.msc on Windows must be restarted.

namespace TimerJobTest
{
    class TimerJob : SPJobDefinition
    {   
        public TimerJob() : base()
        {
        }
        public TimerJob(string jobName, SPService service, SPServer server, SPJobLockType targetType) : base(jobName, service, server, targetType)
        {
            this.Title = "TimerJob";
        }
        public TimerJob(string jobName, SPWebApplication webApplication) : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "TimerJob";
        }
        
        /// <summary>
        /// This function executes whenever the timer job runs.
        /// It iterates through the SharePoint list and gets each item's properties, the list of groups who have access to the item, and the users belonging to each group.
        /// These are then passed to a Python script for indexing access control to Workplace Search.
        /// </summary>
        /// <param name="ContentDatabaseID"></param>
        public override void Execute(Guid ContentDatabaseID)
        {
            // Create a text file for logging and debugging purposes.
            String path = @"C:\Users\Administrator\Documents\Logs\access_control.txt";
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }
            
            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("======Access Control Sync executed on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "=====");
            }

            try
            {
                using (SPSite oSite = new SPSite("http://server-2k16:3877")) // Replace site url as necessary.
                {
                    using (SPWeb oWeb = oSite.OpenWeb())
                    {
                        SPList oList = oWeb.Lists["Policies and Procedures"]; // Replace with desired list name.
                        var isFirstItem = true; // Used to clear user permissions before iterating through SP List.
                        Dictionary<string, List<string>> usersByGroup = new Dictionary<string, List<string>>();

                        foreach (SPGroup group in oWeb.Groups)
                        {
                            // When passing arguments to the Python script later on, they are treated as command line arguments.
                            // Hence, quotation marks will automatically be removed.
                            // To preserve the quotation marks so that the JSON format is maintained in the Python script, add escaped quotation marks.
                            // This is done for every string value that will be passed to the Python script.
                            string groupName = "\"" + group.Name + "\"";
                            HashSet<string> groupMembers = new HashSet<string>();

                            foreach (SPUser user in group.Users)
                            {
                                string userWithoutDomain = user.LoginName.Substring(user.LoginName.IndexOf('\\') + 1); // Remove the domain from the login name. E.g., remove "SHAREPOINT\" from "SHAREPOINT\admin"
                                string doubleQuoted = "\"" + userWithoutDomain + "\"";
                                groupMembers.Add(doubleQuoted);
                            }

                            usersByGroup.Add(groupName, groupMembers.ToList());
                        }

                        string groupsJson = JsonConvert.SerializeObject(usersByGroup); // Store all groups of users as a JSON string.

                        foreach (SPListItem item in oList.Items)
                        {
                            Dictionary<string, List<string>> itemProperties = new Dictionary<string, List<string>>(); // Hold all of the item's properties.
                            HashSet<string> groupList = new HashSet<string>(); // Contain groups that have access to the item.
                            HashSet<string> details = new HashSet<string>();   // Hold the item's details like name, list name, attachment contents etc.

                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("Groups with access to the list item " + item.Name);
                            }

                            foreach (SPRoleAssignment role in item.RoleAssignments) // Iterate through each group that has access to the item and get the users belonging to that group.
                            {

                                if ((role.Member as SPGroup) != null) // Check if role.Member is a group, not a user.
                                {
                                    SPGroup group = role.Member as SPGroup;
                                    string groupName = "\"" + group.Name + "\"";
                                    groupList.Add(groupName);
                                }
                            }

                            using (StreamWriter sw = File.AppendText(path))
                            {
                                foreach (string name in groupList)
                                {
                                    sw.WriteLine("- " + name);
                                }
                                sw.WriteLine();
                            }

                            // Get all of the item's details.
                            string doubleQuoteItemID = "\"" + item.ID + "\"";
                            string doubleQuoteName = "\"" + item.Name + "\"";
                            string doubleQuoteListName = "\"" + item.ParentList.Title + "\"";
                            string doubleQuoteItemUrl = "\"" + "http://192.168.100.7:3877/" + item.ParentList.DefaultDisplayFormUrl + "?ID=" + item.ID + "\""; // URL for displaying the list item. Replace where necessary.

                            DateTime itemModifiedDate = (DateTime)item["Modified"];
                            string doubleQuoteItemModified = "\"" + itemModifiedDate.ToString("dd/MM/yyyy, HH:mm") + "\"";

                            string strUserValue = (string)item["Modified By"]; // Returns a string with the format ID#User_Display_Name
                            int intIndex = strUserValue.IndexOf("#");
                            string doubleQuoteItemModifiedBy = "\"" + strUserValue.Substring(intIndex + 1) + "\""; // Get only the display name of the user

                            // The "Content" field is in HTML format by default. Convert it to plain text.
                            // Additionally, escape all double quotes as they will otherwise be removed by the command line when passed to the Python script.
                            var itemContentField = item.ParentList.Fields.GetField("Content");
                            var itemContent = item[itemContentField.Id];
                            string doubleQuoteItemContentText = "\"" + itemContentField.GetFieldValueAsText(itemContent).Replace(@"""", @"\""") + "\"";

                            SPAttachmentCollection attachments = item.Attachments;
                            TextExtractor textExtractor = new TextExtractor();
                            StringBuilder stringBuilder = new StringBuilder("\"");
                            if (attachments.Count > 0)
                            {
                                // For each attachment, extract the contents and append it to a string.
                                for (int i = 0; i < attachments.Count; i++)
                                {
                                    string attachmentUrl = attachments.UrlPrefix + attachments[i];
                                    SPFile attachedFile = oWeb.GetFile(attachmentUrl);
                                    string attachedName = attachedFile.Name;
                                    byte[] fileBytes = attachedFile.OpenBinary();
                                    string fileString = textExtractor.Extract(fileBytes).Text.Trim();
                                    stringBuilder.AppendLine(fileString);
                                    stringBuilder.AppendLine("---------");
                                }
                            }
                            stringBuilder.Append("\"");
                            string doubleQuoteAttachedContent = stringBuilder.ToString();

                            details.Add(doubleQuoteItemID);
                            details.Add(doubleQuoteName);
                            details.Add(doubleQuoteListName);
                            details.Add(doubleQuoteItemUrl);
                            details.Add(doubleQuoteAttachedContent);
                            details.Add(doubleQuoteItemModified);
                            details.Add(doubleQuoteItemModifiedBy);
                            details.Add(doubleQuoteItemContentText);

                            itemProperties.Add("\"" + "details" + "\"", details.ToList());
                            itemProperties.Add("\"" + "permissions" + "\"", groupList.ToList());

                            string itemJson = JsonConvert.SerializeObject(itemProperties); // Store properties of the items as a JSON string.

                            ExecutePythonScript(groupsJson, itemJson, isFirstItem.ToString());

                            isFirstItem = false;
                        }  
                    }
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("Something went wrong: " + ex.Message + "\n");
                }
            }
        }

        /// <summary>
        /// Passes arguments from this C# class to a Python script and executes it.
        /// </summary>
        /// <param name="groupsJson">Groups of SharePoint users</param>
        /// <param name="itemJson">Properties of an item</param>
        private void ExecutePythonScript(string groupsJson, string itemJson, string isFirstItem)
        {
            // 1) Create process info
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = @"C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe"; // Full path to Python installation. Change this as necessary.

            // 2) Provide script and arguments
            var script = @"Z:\CYL_Project\Python\access_control_sync.py"; // Path to Python script. Change this as necessary.
            start.Arguments = String.Format("{0} {1} {2} {3}", script, groupsJson, itemJson, isFirstItem);

            // 3) Process configuration
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            start.CreateNoWindow = true;
            start.RedirectStandardError = true;
            using (Process process = Process.Start(start))
            {
                string errors = process.StandardError.ReadToEnd();
                string result = process.StandardOutput.ReadToEnd();

                string path = @"C:\Users\Administrator\Documents\Logs\elastic_permissions_log.txt"; // File for logging and debugging when indexing to Workplace Search.

                if (!File.Exists(path))
                {
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.WriteLine("-----New File-----");
                    }
                }

                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("-----Document Permissions Ported to Workplace Search on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm") + "-----");
                    sw.WriteLine("Errors: ");
                    sw.WriteLine(errors);
                    sw.WriteLine();
                    sw.WriteLine("Results: ");
                    sw.WriteLine(result);
                }
            }
        }
    }
}