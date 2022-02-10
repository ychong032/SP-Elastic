using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Net;
using System.Net.Http;
using System.Text;
using System.IO;
using System.Diagnostics;
using Microsoft.SharePoint.Taxonomy;
using System.Net.Mail;
using Microsoft.SharePoint.Administration;

namespace IDSCircular.IDSCircularER
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class IDSCircularER : SPItemEventReceiver
    {
        /// <summary>
        /// An item is being deleted.
        /// </summary>
        public override void ItemDeleting(SPItemEventProperties properties)
        {
            base.ItemDeleting(properties);
            String path = @"D:\IDS\EventReceiver\Logs\Circular.txt";

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----- Item Deleted Triggered -----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    SPListItem currentItem = properties.ListItem;
                    String url = String.Format("{0}/{1}?ID={2}", properties.Web.Url, properties.List.Forms[PAGETYPE.PAGE_DISPLAYFORM].Url, properties.ListItem.ID);
                    System.Guid listID = properties.ListId;
                    Int32 itemID = properties.ListItem.ID;
                    String listName = properties.ListTitle;
                    String relUrl = properties.RelativeWebUrl;

                    String siteCollectionName = "circular";
                    String subSite = "";
                    String topicTitle = "";
                    String allowRead = "";
                    
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("URL: " + url + "; Item ID: " + itemID + "; List Name: " + listName + "; Site Collection Name: " + siteCollectionName);
                    }

                    try
                    {
                        SPFieldUserValueCollection objUserFieldValueCol =
                            new SPFieldUserValueCollection(properties.Web, properties.ListItem["AllowRead"].ToString());

                        for (int i = 0; i < objUserFieldValueCol.Count; i++)
                        {
                            SPFieldUserValue singlevalue = objUserFieldValueCol[i];

                            if (singlevalue.User == null)
                            {
                                SPGroup group = web.Groups[singlevalue.LookupValue];

                                using (StreamWriter sw = File.AppendText(path))
                                {
                                    sw.WriteLine("Allow Read Field: " + group.ToString());
                                    sw.WriteLine("Allow Read Field: " + group.ToString().Split('#')[1]);
                                }

                                if (group.ToString().Split('#')[1] == "Everyone_In_SID")
                                {
                                    allowRead = "Everyone";
                                }
                                else
                                {
                                    allowRead = "NotEveryone";
                                }
                            }
                            else
                            {
                                using (StreamWriter sw = File.AppendText(path))
                                {
                                    sw.WriteLine("Allow Read Field: " + singlevalue.ToString());
                                    sw.WriteLine("Allow Read Field: " + singlevalue.ToString().Split('#')[1]);
                                }

                                if (singlevalue.ToString().Split('#')[1] == "Everyone_In_SID")
                                {
                                    allowRead = "Everyone";
                                }
                                else
                                {
                                    allowRead = "NotEveryone";
                                }
                            }
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Deleted - Permission Assignment");
                            sw.WriteLine("Item Deleted - Allow Read: " + allowRead);
                        }

                    }
                    catch (Exception ex)
                    {
                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Deleted - Permission Assignment Exception");
                            sw.WriteLine(ex);
                        }

                        throw ex;
                    }

                    try
                    {
                        String fieldValue = properties.ListItem["Topic"].ToString();

                        if (fieldValue != null && fieldValue != "")
                        {
                            int topicIndexLeft = fieldValue.IndexOf("#");
                            int topicIndexRight = fieldValue.IndexOf("|");
                            topicTitle = fieldValue.Substring((topicIndexLeft + 1), (topicIndexRight - topicIndexLeft - 1));
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Deleted - Get Topic");
                            sw.WriteLine("Item Deleted - Topic: " + topicTitle);
                        }

                    }
                    catch (Exception ex)
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Deleted - Get Topic Exception");
                            sw.WriteLine(ex);
                        }

                        throw ex;
                    }

                    var postUrl = "http://SGOMSCWFE01.sid.gov.sg:8009/api/v1/ids/delete?id=" + itemID + "&siteCollectionName=" + siteCollectionName + "&subsite=" + subSite + "&topic='" + topicTitle + "'";

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Deleted - Sending Post Request..");
                        sw.WriteLine("Item Deleted - Post URL: " + postUrl.ToString());
                    }

                    if (allowRead == "Everyone")
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Deleted - Sending Post Request..");
                        }

                        try
                        {
                            var request = (HttpWebRequest)WebRequest.Create(postUrl);
                            var response = (HttpWebResponse)request.GetResponse();
                            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                        }
                        catch (Exception ex)
                        {

                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("Item Deleted - Send Post Request Exception");
                                sw.WriteLine(ex);
                            }

                            throw ex;
                        }
                    }
                }
                catch (Exception ex)
                {

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Deleted Exception");
                        sw.WriteLine(ex);
                    }

                    throw ex;
                }
            }
        }        

        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
            String path = @"D:\IDS\EventReceiver\Logs\Circular.txt";

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----- Item Added Triggered -----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    SPListItem currentItem = properties.ListItem;
                    String url = String.Format("{0}/{1}?ID={2}", properties.Web.Url, properties.List.Forms[PAGETYPE.PAGE_DISPLAYFORM].Url, properties.ListItem.ID);
                    System.Guid listID = properties.ListId;
                    Int32 itemID = properties.ListItem.ID;
                    String listName = properties.ListTitle;
                    String relUrl = properties.RelativeWebUrl;

                    String siteCollectionName = "circular";
                    String subSite = "";
                    String topicTitle = "";
                    String expiryDate = "";
                    String effectiveDate = "";
                    String allowRead = "";

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("URL: " + url + "; Item ID: " + itemID + "; List Name: " + listName + "; Site Collection Name: " + siteCollectionName);
                    }

                    try
                    {
                        SPFieldUserValueCollection objUserFieldValueCol =
                            new SPFieldUserValueCollection(properties.Web, properties.ListItem["AllowRead"].ToString());

                        for (int i = 0; i < objUserFieldValueCol.Count; i++)
                        {
                            SPFieldUserValue singlevalue = objUserFieldValueCol[i];

                            if (singlevalue.User == null)
                            {
                                SPGroup group = web.Groups[singlevalue.LookupValue];

                                using (StreamWriter sw = File.AppendText(path))
                                {
                                    sw.WriteLine("Allow Read Field: " + group.ToString());
                                    sw.WriteLine("Allow Read Field: " + group.ToString().Split('#')[1]);
                                }

                                if (group.ToString().Split('#')[1] == "Everyone_In_SID")
                                {
                                    allowRead = "Everyone";
                                }
                                else
                                {
                                    allowRead = "NotEveryone";
                                }
                            }
                            else
                            {
                                using (StreamWriter sw = File.AppendText(path))
                                {
                                    sw.WriteLine("Allow Read Field: " + singlevalue.ToString());
                                    sw.WriteLine("Allow Read Field: " + singlevalue.ToString().Split('#')[1]);
                                }

                                if (singlevalue.ToString().Split('#')[1] == "Everyone_In_SID")
                                {
                                    allowRead = "Everyone";
                                }
                                else
                                {
                                    allowRead = "NotEveryone";
                                }
                            }
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Added - Permission Assignment");
                            sw.WriteLine("Item Added - Allow Read: " + allowRead);
                        }
                    }
                    catch (Exception ex)
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Added - Permission Assignment Exception");
                            sw.WriteLine(ex);
                        }

                        throw ex;
                    }

                    try
                    {
                        String fieldValue = properties.ListItem["Topic"].ToString();

                        if (fieldValue != null && fieldValue != "")
                        {
                            int topicIndexLeft = fieldValue.IndexOf("#");
                            int topicIndexRight = fieldValue.IndexOf("|");
                            topicTitle = fieldValue.Substring((topicIndexLeft + 1), (topicIndexRight - topicIndexLeft - 1));
                        }

                        var date = properties.ListItem["Expiry Date"].ToString();

                        if (date != null && date != "")
                        {
                            expiryDate = date.Split(' ')[0];
                            expiryDate = expiryDate.Replace("/", "-");
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Added - Get Topic and Expiry Date");
                            sw.WriteLine("Item Added - Topic: " + topicTitle + "; Expiry Date: " + expiryDate);
                        }
                    }
                    catch (Exception ex)
                    {
                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Added - Get Topic/Expiry Date Exception");
                            sw.WriteLine(ex);
                        }

                        throw ex;
                    }

                    var postUrl = "http://SGOMSCWFE01.sid.gov.sg:8009/api/v1/ids/create?id=" + itemID + "&listName=" + "Circular" + "&listId=" + listID + "&displayListName=" + listName + "&status=ItemAdded" + "&siteCollectionName=" + siteCollectionName + "&subsite=" + subSite + "&topic='" + topicTitle + "'" + "&expiryDate=" + expiryDate + "&effectiveDate=" + effectiveDate;

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Added - Sending Post Request..");
                        sw.WriteLine("Item Added - Post URL: " + postUrl.ToString());
                    }

                    if (allowRead == "Everyone")
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Added - Sending Post Request..");
                        }

                        try
                        {
                            var request = (HttpWebRequest)WebRequest.Create(postUrl);
                            var response = (HttpWebResponse)request.GetResponse();
                            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                        }
                        catch (Exception ex)
                        {

                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("Item Added - Send Post Request Exception");
                                sw.WriteLine(ex);
                            }

                            throw ex;
                        }
                    }

                    /*
                    SPUser allowReadUserName = null;
                    SPUser allowEditUserName = null;
                    SPUser emailUserName = null;
                    String toAddress = null;

                    SPFieldUserValueCollection userCollection = new SPFieldUserValueCollection();

                    //Email Initialization
                    bool appendHtmlTag = false;
                    bool htmlEncode = false;
                    string subject = properties.ListItem["Title"].ToString();
                    string message = properties.ListItem["EmailBody"].ToString();

                    //get usernames 
                    string tempAllowReadValues = currentItem["AllowRead"].ToString();
                    string[] allowReadArray = currentItem["AllowRead"].ToString().Split(';');

                    string tempAllowEditValues = currentItem["AllowEdit"].ToString();
                    string[] allowEditArray = currentItem["AllowEdit"].ToString().Split(';');

                    string tempEmailRecipientValues = currentItem["EmailRecipient"].ToString();
                    string[] emailRecipientArray = currentItem["EmailRecipient"].ToString().Split(';');

                    //Remove permission first 
                    web.AllowUnsafeUpdates = true;
                    currentItem.BreakRoleInheritance(false);
                    SPRoleAssignmentCollection raCollection = currentItem.RoleAssignments;
                    //remove exising permissions one by one 
                    for (int a=raCollection.Count-1; a>=0; a--)
                    {
                        raCollection.Remove(a);
                    }

                    for (int i = 1; i < emailRecipientArray.Length; i++)
                    {
                        tempEmailRecipientValues = emailRecipientArray[i].Substring(emailRecipientArray[i].LastIndexOf('#')+1);

                        //currentItem["Status"] = "Sent EMail" + tempEmailRecipientValues;
                        //currentItem.Update();

                        emailUserName = web.EnsureUser(emailRecipientArray[i]);                                               

                        //currentItem["Status"] = "EmailUserName" + emailUserName;
                        //currentItem.Update();

                        toAddress = emailUserName.Email;

                        //currentItem["Status"] = "Address" + toAddress;
                        //currentItem.Update();

                        SPSecurity.RunWithElevatedPrivileges(delegate(){

                            //Email user
                            bool result = SPUtility.SendEmail(web, appendHtmlTag, htmlEncode, toAddress, subject, message);
                        });

                        //currentItem["Content"] = "Sent EMail" + emailUserName.Email; 
                        //currentItem.Update();
                    }
                        
                    for(int i = 1; i<allowReadArray.Length; i++)
                    {
                        tempAllowReadValues = allowReadArray[i].Substring(allowReadArray[i].LastIndexOf('#') + 1);
                        //allowReadUserName = web.EnsureUser(@"JUPITER\" + tempAllowReadValues.ToString());
                        allowReadUserName = web.EnsureUser(allowReadArray[i]);

                        SPSecurity.RunWithElevatedPrivileges(delegate()
                        {
                            //Permissions 
                            //grant permissions for specific list item
                            SPRoleDefinition roleDefinition = web.RoleDefinitions.GetByType(SPRoleType.Reader);
                            SPRoleAssignment roleAssignment = new SPRoleAssignment(allowReadUserName);

                            roleAssignment.RoleDefinitionBindings.Add(roleDefinition);
                            currentItem.RoleAssignments.Add(roleAssignment);
                            currentItem.Update();
                        });
                    }

                    for (int i = 1; i < allowEditArray.Length; i++)
                    {

                        tempAllowEditValues = allowEditArray[i].Substring(allowEditArray[i].LastIndexOf('#') + 1);
                        allowEditUserName = web.EnsureUser(allowEditArray[i]);

                        SPSecurity.RunWithElevatedPrivileges(delegate()
                        {
                            //Permissions 
                            //grant permissions for specific list item
                            SPRoleDefinition roleDefinition = web.RoleDefinitions.GetByType(SPRoleType.Contributor);
                            SPRoleAssignment roleAssignment = new SPRoleAssignment(allowEditUserName);

                            roleAssignment.RoleDefinitionBindings.Add(roleDefinition);
                            currentItem.RoleAssignments.Add(roleAssignment);
                            currentItem.Update();
                        });
                    }*/
                }
                catch (Exception ex)
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Added Exception");
                        sw.WriteLine(ex);
                    }

                    throw ex;
                }                
            }
        }

        private void Log(string p, StreamWriter sw)
        {
            throw new NotImplementedException();
        }

        private void Log(Exception ex, StreamWriter sw)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);
            String path = @"D:\IDS\EventReceiver\Logs\Circular.txt";

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine("----- Item Updated Triggered -----");
            }

            using (SPWeb web = properties.OpenWeb())
            {
                try
                {
                    SPListItem currentItem = properties.ListItem;
                    String url = String.Format("{0}/{1}?ID={2}", properties.Web.Url, properties.List.Forms[PAGETYPE.PAGE_DISPLAYFORM].Url, properties.ListItem.ID);
                    System.Guid listID = properties.ListId;
                    Int32 itemID = properties.ListItem.ID;
                    String listName = properties.ListTitle;
                    String relUrl = properties.RelativeWebUrl;

                    String siteCollectionName = "circular";
                    String subSite = "";
                    String topicTitle = "";
                    String expiryDate = "";
                    String effectiveDate = "";
                    String allowRead = "";

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("URL: " + url + "; Item ID: " + itemID + "; List Name: " + listName + "; Site Collection Name: " + siteCollectionName);
                    }

                    try
                    {
                        SPFieldUserValueCollection objUserFieldValueCol =
                            new SPFieldUserValueCollection(properties.Web, properties.ListItem["AllowRead"].ToString());

                        for (int i = 0; i < objUserFieldValueCol.Count; i++)
                        {
                            SPFieldUserValue singlevalue = objUserFieldValueCol[i];

                            if (singlevalue.User == null)
                            {
                                SPGroup group = web.Groups[singlevalue.LookupValue];

                                using (StreamWriter sw = File.AppendText(path))
                                {
                                    sw.WriteLine("Allow Read Field: " + group.ToString());
                                    sw.WriteLine("Allow Read Field: " + group.ToString().Split('#')[1]);
                                }

                                if (group.ToString().Split('#')[1] == "Everyone_In_SID")
                                {
                                    allowRead = "Everyone";
                                }
                                else
                                {
                                    allowRead = "NotEveryone";
                                }
                            }
                            else
                            {
                                using (StreamWriter sw = File.AppendText(path))
                                {
                                    sw.WriteLine("Allow Read Field: " + singlevalue.ToString());
                                    sw.WriteLine("Allow Read Field: " + singlevalue.ToString().Split('#')[1]);
                                }

                                if (singlevalue.ToString().Split('#')[1] == "Everyone_In_SID")
                                {
                                    allowRead = "Everyone";
                                }
                                else
                                {
                                    allowRead = "NotEveryone";
                                }
                            }
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Updated - Permission Assignment");
                            sw.WriteLine("Item Updated - Allow Read: " + allowRead);
                        }

                    }
                    catch (Exception ex)
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Updated - Permission Assignment Exception");
                            sw.WriteLine(ex);
                        }

                        throw ex;
                    }

                    try
                    {
                        String fieldValue = properties.ListItem["Topic"].ToString();

                        if (fieldValue != null && fieldValue != "")
                        {
                            int topicIndexLeft = fieldValue.IndexOf("#");
                            int topicIndexRight = fieldValue.IndexOf("|");
                            topicTitle = fieldValue.Substring((topicIndexLeft + 1), (topicIndexRight - topicIndexLeft - 1));
                        }

                        var date = properties.ListItem["Expiry Date"].ToString();

                        if (date != null && date != "")
                        {
                            expiryDate = date.Split(' ')[0];
                            expiryDate = expiryDate.Replace("/", "-");
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Updated - Get Topic and Expiry Date");
                            sw.WriteLine("Item Updated - Topic: " + topicTitle + "; Expiry Date: " + expiryDate);
                        }

                    }
                    catch (Exception ex)
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Updated - Get Topic/Expiry Date Exception");
                            sw.WriteLine(ex);
                        }

                        throw ex;
                    }

                    var postUrl = "http://SGOMSCWFE01.sid.gov.sg:8009/api/v1/ids/create?id=" + itemID + "&listName=" + "Circular" + "&listId=" + listID + "&displayListName=" + listName + "&status=ItemUpdated" + "&siteCollectionName=" + siteCollectionName + "&subsite=" + subSite + "&topic='" + topicTitle + "'" + "&expiryDate=" + expiryDate + "&effectiveDate=" + effectiveDate;

                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Updated - Sending Post Request..");
                        sw.WriteLine("Item Updated - Post URL: " + postUrl.ToString());
                    }

                    if (allowRead == "Everyone")
                    {

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("Item Updated - Sending Post Request..");
                        }

                        try
                        {
                            var request = (HttpWebRequest)WebRequest.Create(postUrl);
                            var response = (HttpWebResponse)request.GetResponse();
                            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                        }
                        catch (Exception ex)
                        {
                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine("Item Updated - Send Post Request Exception");
                                sw.WriteLine(ex);
                            }

                            throw ex;
                        }
                    }
                }
                catch (Exception ex)
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine("Item Updated Exception");
                        sw.WriteLine(ex);
                    }

                    throw ex;
                }
            }
        }
    }
}   