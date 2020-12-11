
using System;
using System.IO;


namespace Email_Reader_To_SharePoint
{
    class ReadMsgFile
    {
        public static void ReadOutlookMail(string filepath)
        {
            try
            {
                OutlookStorage.Message outlookMsg = new OutlookStorage.Message(filepath);
                SaveMessage(outlookMsg);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        private static void SaveMessage(OutlookStorage.Message outlookMsg)
        {
            outlookMsg.Save(outlookMsg.Subject.Replace(":", ""));

            foreach (OutlookStorage.Attachment attach in outlookMsg.Attachments)
            {
                byte[] attachBytes = attach.Data;
                FileStream attachStream = File.Create(attach.Filename);
                attachStream.Write(attachBytes, 0, attachBytes.Length);
                attachStream.Close();
            }

            foreach (OutlookStorage.Message subMessage in outlookMsg.Messages)
            {
                SaveMessage(subMessage);
            }
        }
    }
}
