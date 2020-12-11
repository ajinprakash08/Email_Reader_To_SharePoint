using glib.Email;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using Limilabs.Mail.MSG;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace Email_Reader_To_SharePoint
{
   class ReadEmlFile
    {
        public static string GetPlainText(RxMailMessage mm)
        {
            // check for plain text in body
            if (!mm.IsBodyHtml && !string.IsNullOrEmpty(mm.Body))
                return mm.Body;

            string sText = string.Empty;
            foreach (AlternateView av in mm.AlternateViews)
            {
                // check for plain text
                if (string.Compare(av.ContentType.MediaType, "text/plain", true) == 0)
                    continue;// return StreamToString(av.ContentStream);

                // check for HTML text
                if (string.Compare(av.ContentType.MediaType, "text/html", true) == 0)
                    sText = StreamToString(av.ContentStream);
            }

            // HTML is our only hope
            if (sText == string.Empty && mm.IsBodyHtml && !string.IsNullOrEmpty(mm.Body))
                sText = mm.Body;

            if (sText == string.Empty)
                return string.Empty;

            // need to convert the HTML to plaintext
            return StripHtml(sText);
        }

        public static string StripHtml(string html)
        {
            // create whitespace between html elements, so that words do not run together
            html = html.Replace(">", "> ");

            // parse html
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);

            // strip html decoded text from html
            string text = HttpUtility.HtmlDecode(doc.DocumentNode.InnerText);

            // replace all whitespace with a single space and remove leading and trailing whitespace
            return Regex.Replace(text, @"\s+", " ").Trim();
        }

        public static string StreamToString(Stream stream)
        {
            string sText = string.Empty;
            using (StreamReader sr = new StreamReader(stream))
            {
                sText = sr.ReadToEnd();
                stream.Seek(0, SeekOrigin.Begin);   // leave the stream the way we found it
                stream.Close();
            }

            return sText;
        }
    }
}
