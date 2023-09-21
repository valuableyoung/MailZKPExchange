using MailZKPExchange.DBConnector;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace MailZKPExchange.Helpers
{
    public static class Logger
    {
        public static int ErrorCode { get; set; }
        public static string AllMessage { get; set; }
        public static void SendMessage(string message, object idDoc = null)
        {
            message = "Время: " + DateTime.Now.ToString() + "  " + message;
            AllMessage += message + "\n";
            SendToBase(message, ErrorCode, idDoc);
        }

        public static void SendMessage(string message, LogAction typeAction, object idDoc = null)
        {
            message = "Время: " + DateTime.Now.ToString() + "  " + message;
            AllMessage += message + "\n";
            SendToBase(message, (int)typeAction, idDoc);
        }

        public static void ShowMessage(string message)
        {
            message = "Время: " + DateTime.Now.ToString() + "  " + message;
            AllMessage += message + "\n";
        }

        private static void SendToBase(string message, int idTypeAction, object idDoc)
        {
            try
            {
                SqlParameter parMessage = new SqlParameter("message", message);
                SqlParameter parIdTypeAction = new SqlParameter("idTypeAction", idTypeAction);
                SqlParameter parIdDoc = new SqlParameter("idDoc", idDoc == null ? "" : idDoc.ToString());

                StringBuilder query = new StringBuilder();
                query.Append("insert into LogeOrder ");
                query.Append("values ");
                query.Append("( ");
                query.Append("     isnull((select id_kontr_db from tov_doc (nolock) where id_doc = @idDoc), 1)  ");
                query.Append("     ,GETDATE(), @idTypeAction, @message, 0 ");
                query.Append("     ,(select id_type_doc from tov_doc (nolock) where id_doc = @idDoc) ");
                query.Append("     ,@idDoc, SYSTEM_USER ");
                query.Append(") ");

                DBExecutor.ExecuteQuery(query.ToString(), parMessage, parIdTypeAction, parIdDoc);
            }
            catch (Exception e)
            {
                Logger.WriteErrorMessage("Время: " + DateTime.Now.ToString() + " Ошибка записи в LogEOrder: " + e.Message + "\n");
            }
        }

        public static void WriteMessage()
        {
            var wayToFile = Directory.GetCurrentDirectory() + @"\log\log" + DateTime.Now.Ticks + ".txt";
            //File.WriteAllText(wayToFile, AllMessage);
            using (StreamWriter sw = new StreamWriter(wayToFile))
            {
                sw.WriteLine(AllMessage);
            }
            AllMessage = "";
        }

        public static void WriteErrorMessage(string s)
        {
            var wayToFile = Directory.GetCurrentDirectory() + @"\log\error" + DateTime.Now.Ticks + ".txt";
            using (StreamWriter sw = new StreamWriter(wayToFile))
            {
                sw.WriteLine(s);
            }
        }
        public static void WriteDebugMessage(string s)
        {
            var wayToFile = Directory.GetCurrentDirectory() + @"\log\debug" + DateTime.Now.Ticks + ".txt";
            using (StreamWriter sw = new StreamWriter(wayToFile))
            {
                sw.WriteLine(s);
            }
        }
    }
}
