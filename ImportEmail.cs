//headers
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using TranslateTextSample; 


namespace ConsoleApp1
{
    class ImportEmail
    {
        //Mutators/accessors
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody  { get; set; }

        //List to read in the emails
        public static List<ImportEmail> ReadEmailItem()
        {
            //create the connection
            Application outlookApplication = null;
            NameSpace outlookNameSpace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;

            //list of email items
            List<ImportEmail> ListEmailDetials = new List<ImportEmail>();
            ImportEmail emailDetails;

            //attempt to extract the emails
            try
            {
                outlookApplication = new Application();
                outlookNameSpace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                foreach (MailItem mail in mailItems)
                {
                    emailDetails = new ImportEmail();
                    emailDetails.EmailFrom = mail.SenderEmailAddress;
                    emailDetails.EmailSubject = mail.Subject;
                    emailDetails.EmailBody = mail.Body;
                    ListEmailDetials.Add(emailDetails);
                    ReleaseObject(mail);
                }
            }
            catch (System.Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            finally
            {
                ReleaseObject(mailItems);
                ReleaseObject(inboxFolder);
                ReleaseObject(outlookNameSpace);
                ReleaseObject(outlookApplication);
            }
            

            return ListEmailDetials;

        }
        //return the mail object
            private static void ReleaseObject(object obj)
            {
                if (obj != null)
                {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
                }
            }
    }
}




