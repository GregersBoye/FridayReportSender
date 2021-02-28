using System;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;

namespace reportSender
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main()
        {
            log4net.Config.XmlConfigurator.Configure();
            int maxAge = int.Parse(ConfigurationManager.AppSettings["maxAgeInHours"]);
            DateTime deadline = DateTime.Now.AddHours(maxAge * -1);
            string folderPath = ConfigurationManager.AppSettings["folder"];
            string[] filepaths = Directory.GetFiles(folderPath);


            foreach (var path in filepaths)
            {
                DateTime editedTimestamp = File.GetLastWriteTime(path);

                if (editedTimestamp < deadline)
                {
                    log.Info($"File '{Path.GetFileName(path)}' has not been edited");
                    continue;
                }
                
                SendReport(path);
            }
        }


        public static void SendReport(string filePath)
        {
            var fileName = Path.GetFileName(filePath).Split('.')[0];
            var ol = new Outlook.Application();
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = fileName;

            // Add recipient using display name, alias, or smtp address
            mail.Recipients.Add(ConfigurationManager.AppSettings["recipient"]);
            mail.Recipients.ResolveAll();
            mail.Attachments.Add(filePath,
                Outlook.OlAttachmentType.olByValue, Type.Missing,
                Type.Missing);
            mail.Send();
            log.Info($"File '{fileName}' was sent");
        }
    }
}
