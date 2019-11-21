using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Text.RegularExpressions;

namespace OutlookAddIn1
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the Application object
            Outlook.Application application = Globals.ThisAddIn.Application;

            // Get the active Inspector object and check if is type of MailItem
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {

                string job = ThisAddIn.generateJob();
                

                switch (mailItem.BodyFormat)
                {
                    case Outlook.OlBodyFormat.olFormatHTML:

                        mailItem.HTMLBody += "\r\n\r\n" + ReadSignature("bullshit.htm").Replace("%BULLSHIT%", job);
                        ThisAddIn.lastJob = job;
                        break;
                    case Outlook.OlBodyFormat.olFormatPlain:
                        mailItem.Body += "\r\n\r\n" + ReadSignature("bullshit.txt").Replace("%BULLSHIT%", job);
                        ThisAddIn.lastJob = job;
                        break;
                    case Outlook.OlBodyFormat.olFormatRichText:
                        //mailItem.RTFBody = mailItem.RTFBody.Replace("%JOB%", job);
                        break;
                    default:
                        break;
                }
            }
        }

        private string ReadSignature(string Filter)
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string signature = "%BULLSHIT%";
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles(Filter);

                if (fiSignature.Length > 0)
                {
                    StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.UTF8);
                    signature = sr.ReadToEnd();

                    if (!string.IsNullOrEmpty(signature))
                    {
                        string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName + "_files/");
                    }
                }
            }
            return signature;
        }

        private void btnInsert_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            // Get the Application object
            Outlook.Application application = Globals.ThisAddIn.Application;

            // Get the active Inspector object and check if is type of MailItem
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                string job = ThisAddIn.generateJob();

                switch (mailItem.BodyFormat)
                {
                    case Outlook.OlBodyFormat.olFormatHTML:
                        if (mailItem.HTMLBody != null)
                        {
                            mailItem.HTMLBody = mailItem.HTMLBody.Replace("%BULLSHIT%", job);
                            if(ThisAddIn.lastJob != string.Empty)
                                mailItem.HTMLBody = mailItem.HTMLBody.Replace(ThisAddIn.lastJob, job);
                            ThisAddIn.lastJob = job;
                            //mailItem.HTMLBody = Regex.Replace(mailItem.HTMLBody, "<span class=\"bullshit\">(.*)</span>", "<span class=\"bullshit\">" + job + "</span>");
                            //mailItem.HTMLBody = mailItem.HTMLBody.Replace("%BULLSHIT%", job);
                        }
                        break;
                    case Outlook.OlBodyFormat.olFormatPlain:
                        if (mailItem.Body != null)
                            mailItem.Body = mailItem.Body.Replace("%BULLSHIT%", job);
                        if (ThisAddIn.lastJob != string.Empty)
                            mailItem.Body = mailItem.Body.Replace(ThisAddIn.lastJob, job);
                        ThisAddIn.lastJob = job;
                        break;
                    case Outlook.OlBodyFormat.olFormatRichText:
                        //mailItem.RTFBody = mailItem.RTFBody.Replace("%JOB%", job);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
