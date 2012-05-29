using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Xml;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

namespace PTI_Integration_Console
{
    class Globals
    {
        //Printable Data Feed Declarations
        private static Uri printableURI_packingSlip = new Uri(@"https://services.printable.com/TRANS/0.9/PackingSlip.asmx");
        private static string printableToken = "4F3515138E8BAA6A8D968BAE7F486C34";

        

        //Connection String Declarations
        private static string logicConnString = "Data Source=SQL1;Initial Catalog=pLogic;User ID=FPGwebservice;Password=kissmygrits";
        
        private static string printableConnString = "Data Source=SQL1;Initial Catalog=printable;User ID=FPGwebservice;Password=kissmygrits";
        
        //private static string logicConnString = "Data Source=PLM;Initial Catalog=devLogic;User ID=FPGwebservice;Password=kissmygrits";
        
        //private static string printableConnString = "Data Source=PLM;Initial Catalog=printable;User ID=FPGwebservice;Password=kissmygrits";
       


        //Accessor Methods
        public static string get_logicConnString
        {
            get { return Globals.logicConnString; }
        }

        public static string get_printableConnString
        {
            get { return Globals.printableConnString; }
        }

        public static Uri get_printableURI_packingSlip
        {
            get { return Globals.printableURI_packingSlip; }
        }

        public static string get_printableToken
        {
            get { return Globals.printableToken; }
        }



        //SMTP Connection Settings
        private static SmtpClient smtpClient = new SmtpClient("192.168.240.27");

        public static SmtpClient get_smtpClient
        {
            get { return Globals.smtpClient; }
        }

        public void sendEmail(string recipID, string msgBody)
        {
            ArrayList msgList = new ArrayList();

            msgList.Clear();
            msgList.Add(recipID);
                foreach (string item in msgList)
                {
                    try
                    {
                        MailMessage message = new MailMessage();
                        message.To.Add(item);
                        message.Subject = "Printable Integration Notication";
                        message.From = new MailAddress("PTI_Integration@Finelink.com");
                        message.Body = msgBody;
                        message.ReplyTo = new MailAddress("GVreeman@Finelink.com");
                        message.IsBodyHtml = true;
                        System.Net.Mail.SmtpClient smtp = Globals.get_smtpClient;
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.Write("Sending email to " + message.To.ToString());
                        smtp.Send(message);
                        Console.WriteLine(" - Success");
                        Console.ResetColor();
                        message.Dispose();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("SendEmail Error - RecipID:" + recipID + " - Receiver:" + item + " - Msg:" + msgBody);
                        Console.WriteLine(e.ToString()); 
                        Console.Beep();
                        //errorLog("Email-1", e.ToString(), "SendEmail Error - RecipID:" + recipID + " - Receiver:" + item + " - Msg:" + msgBody);
                    }
                }
            
        }

    }
}
