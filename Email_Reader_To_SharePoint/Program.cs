using glib.Email;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Web;

namespace Email_Reader_To_SharePoint
{
    class Program
    {
        static void Main(string[] args)
        {
            readMail();


        }

        private static void readMail()
        {
            try
            {
                foreach (string file in Directory.EnumerateFiles(ConfigurationManager.AppSettings["sEmlPath"]))
                {                    
                    if (file.Contains(".msg"))
                    {
                        ReadMsgFile.ReadOutlookMail(file);
                    }
                    else
                    {
                        MimeReader mime = new MimeReader();
                        RxMailMessage mm = mime.GetEmail(file);
                        string detils = ReadEmlFile.GetPlainText(mm);
                    }
                }               

            }
            catch (Exception)
            {

                throw;
            }
        }

       
    }

}
