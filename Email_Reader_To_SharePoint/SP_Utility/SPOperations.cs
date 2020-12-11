using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Reader_To_SharePoint
{
    class SPOperations
    {
        public void createList()
        {
            try
            {
                using (SPSite site = new SPSite(ConfigurationManager.AppSettings["siteUrl"]))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPListCollection lists = web.Lists;
                        // create new Generic list called "My List"
                        lists.Add(ConfigurationManager.AppSettings["listName"], "Task responses", SPListTemplateType.GenericList);
                        SPList newList = web.Lists["Incoming Mail List"];
                        // create Text type new column called "My Column"
                        newList.Fields.Add("Body", SPFieldType.Text, true);
                        newList.Fields.Add("Sender", SPFieldType.Text, true);
                        // make new column visible in default view
                        SPView view = newList.DefaultView;
                        view.ViewFields.Add("Body");
                        view.ViewFields.Add("Sender");
                        view.Update();
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        public void InsertMailItem(OutlookStorage.Message message)
        {
            try
            {
                using (SPSite site = new SPSite(ConfigurationManager.AppSettings["siteUrl"]))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        if (web.Lists[ConfigurationManager.AppSettings["listName"]] == null)
                        {
                            createList();
                        }
                        SPListItem newMailListItem = new web.Lists["Incoming Mail List"].AddItem();
                        newMailListItem["Title"] = message.Subject;
                        newMailListItem["Body"] = message.BodyText;
                        newMailListItem["Sender"] = new SPFieldUserValue(web, message.From);
                        newMailListItem.update();
                        foreach (OutlookStorage.Attachment attach in message.Attachments)
                        {
                            byte[] attachBytes = attach.Data;
                            FileStream attachStream = File.Create(attach.Filename);
                            attachStream.Write(attachBytes, 0, attachBytes.Length);
                            attachStream.Close();
                            newMailListItem.Attachments.Add(attach.Filename, attachBytes);
                            newMailListItem.update();
                        }
                    }
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
    }
}
