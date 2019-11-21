using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
//using Word = Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {

        Outlook.MailItem mailItem;
        public static string lastJob = string.Empty;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get the Application object
            Outlook.Application application = this.Application;

            // Get the Inspector object
            Outlook.Inspectors inspectors = application.Inspectors;

            // Get the active Inspector object
            Outlook.Inspector activeInspector = application.ActiveInspector();
            if (activeInspector != null)
            {
                // Get the title of the active item when the Outlook start.
                //MessageBox.Show("Active inspector: " + activeInspector.Caption);
            }

            // Get the Explorer objects
            Outlook.Explorers explorers = application.Explorers;


            // Get the active Explorer object
            Outlook.Explorer activeExplorer = application.ActiveExplorer();
            if (activeExplorer != null)
            {
                // Get the title of the active folder when the Outlook start.
                //MessageBox.Show("Active explorer: " + activeExplorer.Caption);
            }


            // ...
            // Add a new Inspector to the application
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_AddTextToNewMail);
            

            application.ItemLoad += Application_ItemLoad;


        }

        private void Application_ItemLoad(object Item)
        {
            if(Item.GetType() == typeof(Outlook.MailItem))
            {
                mailItem = (Outlook.MailItem)Item;
                ;
            }
        }

        void Inspectors_AddTextToNewMail(Outlook.Inspector inspector)
        {
            this.mailItem = inspector.CurrentItem as Outlook.MailItem;
            
            
            if (mailItem != null)
            {
                ;
                if (mailItem.EntryID == null)
                {
                    ;
                    if (mailItem.HTMLBody != "" )//&& mailItem.HTMLBody.Contains("<div class=WordSection1>"))
                    {
                        mailItem.Open += new Outlook.ItemEvents_10_OpenEventHandler(MailItem_Open);
                    }
                    else
                    {
                        //mailItem.HTMLBody =
                    }
                    //mailItem.PropertyChange += RecipientsPropertyChange;
                }
            }
        }

        private void App_WindowSelectionChange(Word.Selection Sel)
        {

        }

        private void MailItem_PropertyChange(string Name)
        {
            
        }

        private void MailItem_Open(ref bool Cancel)
        {
            var foo = mailItem.Body;
            var bar = mailItem.HTMLBody;

            var job = generateJob();


            switch (mailItem.BodyFormat)
            {
                case Outlook.OlBodyFormat.olFormatHTML:
                    if (mailItem.HTMLBody != null)
                    {
                        mailItem.HTMLBody = mailItem.HTMLBody.Replace("%BULLSHIT%", job);
                        ThisAddIn.lastJob = job;
                    }
                    break;
                case Outlook.OlBodyFormat.olFormatPlain:
                    if (mailItem.Body != null)
                    {
                        mailItem.Body = mailItem.Body.Replace("%BULLSHIT%", job);
                        ThisAddIn.lastJob = job;
                    }
                    break;
                case Outlook.OlBodyFormat.olFormatRichText:
                    //mailItem.RTFBody = mailItem.RTFBody.Replace("%JOB%", job);
                    break;
                default:
                    break;
            }

            mailItem.Open -= MailItem_Open;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        public static string generateJob()
        {
            
            string[] one = new string[] { "Lead", "Senior", "Direct", "Corporate", "Dynamic", "Future", "Product", "National", "Regional", "District", "Central", "Global", "Relational", "Customer", "Investor", "Dynamic", "International", "Legacy", "Forward", "Interactive", "Internal", "Human", "Chief", "Principal" };
            string[] two = new string[] { "Solutions", "Program", "Brand", "Security", "Research", "Marketing", "Directives", "Implementation", "Integration", "Functionality", "Response", "Paradigm", "Tactics", "Identity", "Markets", "Group", "Resonance", "Applications", "Optimization", "Operations", "Infrastructure", "Intranet", "Communications", "Web", "Branding", "Quality", "Assurance", "Impact", "Mobility", "Ideation", "Data", "Creative", "Configuration", "Accountability", "Interactions", "Factors", "Usability", "Metrics", "Team" };
            string[] three = new string[] { "Supervisor", "Associate", "Executive", "Liason", "Officer", "Manager", "Engineer", "Specialist", "Director", "Coordinator", "Administrator", "Architect", "Analyst", "Designer", "Planner", "Synergist", "Orchestrator", "Technician", "Developer", "Producer", "Consultant", "Assistant", "Facilitator", "Agent", "Representative", "Strategist" };
            

            string path =  Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            if(System.IO.File.Exists(path + "\\1.bullshit"))
                one = System.IO.File.ReadAllLines(path + "\\1.bullshit");
            if (System.IO.File.Exists(path + "\\2.bullshit"))
                one = System.IO.File.ReadAllLines(path + "\\2.bullshit");
            if (System.IO.File.Exists(path + "\\3.bullshit"))
                one = System.IO.File.ReadAllLines(path + "\\3.bullshit");

            ;
            Random rnd = new Random();
            return one[rnd.Next(one.Length)] + " " + two[rnd.Next(two.Length)] + " " + three[rnd.Next(three.Length)];
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
