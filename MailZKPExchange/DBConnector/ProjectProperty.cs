using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailZKPExchange.DBConnector
{
    public static class ProjectProperty
    {
        public static string MailFuturesBoss { get; set; }
        public static string DBServer { get; set; }
        public static string DBBase { get; set; }

        public static string MailServer { get; set; }

        public static string MailUserPriceParser { get; set; }
        public static string MailUserPriceParserPassword { get; set; }

        public static string MailUserZKPParser { get; set; }
        public static string MailUserZKPParserPassword { get; set; }

 


        public static string MailCountForParser { get; set; }

        public static string MailUserAnswer { get; set; }
        public static string MailUserAnswerPassword { get; set; }

        public static string FolderForReadedMessages { get; set; }
        public static string FolderForErrorMessage { get; set; }
        public static string FolderForSimpleMessage { get; set; }
        public static string FolderForHandMessage { get; set; }

        public static string DiadocUrl { get; set; }
        public static string DiadocLogin { get; set; }
        public static string DiadocPass { get; set; }
        public static string DiadocTokenAPI { get; set; }

        public static string ProxyLogin { get; set; }
        public static string ProxyPass { get; set; }
        public static string ProxyUri { get; set; }

        public static string PathTo7Zip { get; set; }
        public static int BotUserId { get; set; }

        public static string FolderXls { get; set; }
        public static string FolderLog { get; set; }

        public static string FolderTableDesign { get; set; }

        public static string MailUserPriceParserA1 { get; set; }
        public static string MailUserPriceParserPasswordA1 { get; set; }
        public static string SmtpServer { get; set; }

        public static int SmtpPort { get; set; }
        public static string MailToAnswer { get; set; }



        static ProjectProperty()
        {
            /*DBServer = ConfigurationManager.AppSettings.Get("DBServer");
            DBBase = ConfigurationManager.AppSettings.Get("DBBase");

            MailServer = ConfigurationManager.AppSettings.Get("MailServer");

            MailUserPriceParser = ConfigurationManager.AppSettings.Get("MailUserPriceParser");
            MailUserPriceParserPassword = ConfigurationManager.AppSettings.Get("MailUserPriceParserPassword");

            MailUserAnswer = ConfigurationManager.AppSettings.Get("MailUserAnswer");
            MailUserAnswerPassword = ConfigurationManager.AppSettings.Get("MailUserAnswerPassword");

            FolderForReadedMessages = ConfigurationManager.AppSettings.Get("FolderForReadedMessages");
            FolderForErrorMessage = ConfigurationManager.AppSettings.Get("FolderForErrorMessage");
            FolderForSimpleMessage = ConfigurationManager.AppSettings.Get("FolderForSimpleMessage");
            FolderForHandMessage = ConfigurationManager.AppSettings.Get("FolderForHandMessage");

            DiadocUrl = ConfigurationManager.AppSettings.Get("DiadocUrl");
            DiadocLogin = ConfigurationManager.AppSettings.Get("DiadocLogin");
            DiadocPass = ConfigurationManager.AppSettings.Get("DiadocPass");
            DiadocTokenAPI = ConfigurationManager.AppSettings.Get("DiadocTokenAPI");

            ProxyLogin = ConfigurationManager.AppSettings.Get("ProxyLogin");
            ProxyPass = ConfigurationManager.AppSettings.Get("ProxyPass");
            ProxyUri = ConfigurationManager.AppSettings.Get("ProxyUri");

            PathTo7Zip = ConfigurationManager.AppSettings.Get("PathTo7Zip");
            BotUserId = int.Parse(ConfigurationManager.AppSettings.Get("BotUserId"));

            FolderXls = Directory.GetCurrentDirectory() + "\\" + ConfigurationManager.AppSettings.Get("FolderXls");
            FolderLog = Directory.GetCurrentDirectory() + "\\" + ConfigurationManager.AppSettings.Get("FolderLog");

            FolderTableDesign = Directory.GetCurrentDirectory() + "\\" + ConfigurationManager.AppSettings.Get("FolderTableDesign");

            if (!Directory.Exists(FolderXls))
                Directory.CreateDirectory(FolderXls);

            if (!Directory.Exists(FolderLog))
                Directory.CreateDirectory(FolderLog);

            if (!Directory.Exists(FolderTableDesign))
                Directory.CreateDirectory(FolderTableDesign);*/
        }

        public static void LoadDataAppConfig(bool flag = false)
        {

            DBServer = ConfigurationManager.AppSettings.Get("DBServer");
            DBBase = ConfigurationManager.AppSettings.Get("DBBase");

            MailServer = ConfigurationManager.AppSettings.Get("MailServer");

            MailUserZKPParser = ConfigurationManager.AppSettings.Get("MailUserZKPParser");
            MailUserZKPParserPassword = ConfigurationManager.AppSettings.Get("MailUserZKPParserPassword");

            //if(flag) //перключатель test на real
            //{
            //    DBServer = ConfigurationManager.AppSettings.Get("RealDBServer");
            //    DBBase = ConfigurationManager.AppSettings.Get("RealDBBase");

            //    MailUserZKPParser = ConfigurationManager.AppSettings.Get("RealMailUserZKPParser");
            //    MailUserZKPParserPassword = ConfigurationManager.AppSettings.Get("RealMailUserZKPParserPassword");
            //}
            //else
            //{
            //    DBServer = ConfigurationManager.AppSettings.Get("TestDBServer");
            //    DBBase = ConfigurationManager.AppSettings.Get("TestDBBase");
            //    MailUserZKPParser = ConfigurationManager.AppSettings.Get("TestMailUserZKPParser");
            //    MailUserZKPParserPassword = ConfigurationManager.AppSettings.Get("TestMailUserZKPParserPassword");
            //}


            MailCountForParser = ConfigurationManager.AppSettings.Get("MailCountForParser");

            MailUserAnswer = ConfigurationManager.AppSettings.Get("MailUserAnswer");
            MailUserAnswerPassword = ConfigurationManager.AppSettings.Get("MailUserAnswerPassword");

            FolderForReadedMessages = ConfigurationManager.AppSettings.Get("FolderForReadedMessages");
            FolderForErrorMessage = ConfigurationManager.AppSettings.Get("FolderForErrorMessage");
            FolderForSimpleMessage = ConfigurationManager.AppSettings.Get("FolderForSimpleMessage");
            FolderForHandMessage = ConfigurationManager.AppSettings.Get("FolderForHandMessage");

            ProxyLogin = ConfigurationManager.AppSettings.Get("ProxyLogin");
            ProxyPass = ConfigurationManager.AppSettings.Get("ProxyPass");
            ProxyUri = ConfigurationManager.AppSettings.Get("ProxyUri");

            PathTo7Zip = ConfigurationManager.AppSettings.Get("PathTo7Zip");
            BotUserId = int.Parse(ConfigurationManager.AppSettings.Get("BotUserId"));

            FolderXls = Directory.GetCurrentDirectory() + "\\" + ConfigurationManager.AppSettings.Get("FolderXls");
            FolderLog = Directory.GetCurrentDirectory() + "\\" + ConfigurationManager.AppSettings.Get("FolderLog");

            FolderTableDesign = Directory.GetCurrentDirectory() + "\\" + ConfigurationManager.AppSettings.Get("FolderTableDesign");

            SmtpServer = ConfigurationManager.AppSettings.Get("SmtpServer");
            SmtpPort = int.Parse(ConfigurationManager.AppSettings.Get("SmtpPort"));

            MailToAnswer = ConfigurationManager.AppSettings.Get("MailToAnswer");

            /*if (!Directory.Exists(FolderXls))
                Directory.CreateDirectory(FolderXls);

            if (!Directory.Exists(FolderLog))
                Directory.CreateDirectory(FolderLog);

            if (!Directory.Exists(FolderTableDesign))
                Directory.CreateDirectory(FolderTableDesign);*/
        }


    }
}
