using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using EmailImport;
using TranslateTextSample;
namespace Email_Translate
{
    class Driver
    {
        static async Task Main(string[] args)
        {

            var mails = ImportEmail.ReadEmailItem();
            //keeps track of the email number 
           



           
            //loop to output details about email contents
            /*
            foreach (var item in mails)
            {
                Console.WriteLine("Mail No" + i);
                Console.WriteLine("Mail Recieved From: " + item.EmailFrom);
                Console.WriteLine("Mail Subject: " + item.EmailSubject);
                Console.WriteLine("Mail Body: " + item.EmailBody);
                Console.WriteLine(" ");
                i++;

            }
            */

           


            
            //set string variable
            string path = @"C:/Users/denni/Documents/Intern-Challenge UBS 2019/Email-Translation/Email-Translate/TextOutput.txt";
            string host = "https://api.cognitive.microsofttranslator.com";
            string route = "/translate?api-version=3.0&to=en";
            string subscriptionKey = "5f9c84b31e2a4b19bd277785264c53ff";


            string statement = "            TRANSLATION OF 5 EMAILS IN DIFFERENT LANGUAGES" + "\n" +
            "-----------------------------------------------------------------------------------------" + "\n";

            // output to a txt file
            System.IO.File.WriteAllText(path, statement);


            //int for controlling mail item
            int i = 1;
            //Access mail item



            for (int x = 971; x <= 975; x++)
            {
                var item = mails[x];
                string textToTranslate = item.EmailBody;
                //string variables
                string output = "Email " + i + ": " + textToTranslate + "\n" + " -------------------------------------------------" + "\n";
                System.IO.File.AppendAllText(path, output);
                await TranslateTextSample.Program.TranslateTextRequest(subscriptionKey, host, route, textToTranslate);
                i++;
            }


            Console.WriteLine("Complete");
            Console.ReadKey();
           

            

        }
    }
}
