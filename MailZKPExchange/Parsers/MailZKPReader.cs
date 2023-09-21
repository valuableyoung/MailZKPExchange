using MailZKPExchange.DBConnector;
using MailZKPExchange.SPR;
using MailZKPExchange.Holidays;
using SevenZip;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using MailZKPExchange.Helpers;
using MailZKPExchange.Heplers;
using MailKit.Search;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using System.Net;

namespace MailZKPExchange.Parsers
{
    public static class MailZKPReader
    {
        static string currentMail = "";
        static string currentPassword = "";

        static int idSup = 0;

        public enum MailType
        {
            Simple = 1,
            A1 = 2
        }

        public static string fields;
        public static bool checkSKU;
        public static bool AllSKUIncorrect;
        static void SetCurrentMailParams(MailType mailType)
        {
            switch (mailType)
            {
                case MailType.Simple:
                    currentMail = ProjectProperty.MailUserZKPParser;
                    currentPassword = ProjectProperty.MailUserZKPParserPassword;
                    break;
                case MailType.A1:
                    currentMail = ProjectProperty.MailUserPriceParserA1;
                    currentPassword = ProjectProperty.MailUserPriceParserPasswordA1;
                    break;
            }
        }

        public static string FindidZKP(string subject)
        {
            int startIndex = subject.IndexOf("№");
            string idZKP = subject.Substring(startIndex + 1);
            idZKP = idZKP.Substring(0, idZKP.IndexOf(" "));
            if (idZKP != null)
            {
                return idZKP;
            }
            else
            { 
                return "";
            }
        }

        public static string CheckParse(string ColumnVal, int flag , string ColumnValBrand = "")
        {
            string sql = "";
            if ((ColumnVal != null) && (ColumnVal != ""))
            {
                if(ColumnVal.Contains(',')) { ColumnVal = ColumnVal.Replace(',', '.'); }
                 
                 
                switch (flag)
                {
                    
                    case 1:
                         
                        sql = $@"select TOP(1) id_tov 
                                from spr_tov (nolock) 
                                inner join spr_tm (nolock) on spr_tov.id_tm = spr_tm.tm_id                                
                                where (spr_tov.id_tov_oem_short = dbo.f_replace_for_cross('{ColumnVal}') or spr_tov.id_tov_oem =  dbo.f_replace_for_cross('{ColumnVal}'))
                                    and spr_tm.tm_name = '{ColumnValBrand}'";
                        DataRow dr = DBExecutor.SelectRow(sql);
                        if(dr != null)
                        {
                            return ColumnVal;
                        }
                        else
                        {
                            char[] s = new char[] { '/', '\'', '*', '-', '+', '#', '№', '$', '%', '(', ')', '!', '@', '"', '[', ']', '{', '}', '&', '?', '^' };
                            foreach (char c in s) { if (ColumnVal.Length == 0) return ""; else { ColumnVal = ColumnVal.Replace(c.ToString(), ""); ColumnValBrand = ColumnValBrand.Replace(c.ToString(), "");  } }

                            sql = $@"select TOP(1) id_tov 
                                    from spr_tov (nolock) 
                                    inner join spr_tm (nolock) on spr_tov.id_tm = spr_tm.tm_id                                
                                    where (spr_tov.id_tov_oem_short = '{ColumnVal}' or spr_tov.id_tov_oem =  '{ColumnVal}')
                                        and spr_tm.tm_name = '{ColumnValBrand}'";
                            dr = DBExecutor.SelectRow(sql);
                            if (dr != null)
                            {
                                return ColumnVal;
                            }
                            else
                            {
                                return "";
                            }
                                
                                
                        }
                          // Артикул

                    case 2:
                        {
                            double count = 0;
                            int countprice = 0;
                            int i = 0;
                            List<char> nondigit = new List<char>();
                             
                          
                            if(ColumnVal.Contains("р"))  ColumnVal.Remove(ColumnVal.IndexOf("р"));
                            while (i < ColumnVal.Length)
                            {
                                if ((!char.IsDigit(ColumnVal[i])) && (Convert.ToChar(ColumnVal[i]) != '.') ) { nondigit.Add(Convert.ToChar(ColumnVal[i])); }
                                i++;
                            } 
                            
                            foreach(var c in nondigit) { if (ColumnVal.Length == 0) return ""; else { ColumnVal = ColumnVal.Replace(c.ToString(), ""); } }
                            
                            if (int.TryParse(ColumnVal, out countprice) || double.TryParse(ColumnVal, out count)) { if (Convert.ToDouble(ColumnVal) > 0) return ColumnVal; else { return ""; } }                                                                             
                            else
                            {
                                return "";
                            }
                        }
                    default: return "";
                }

            }
            else { return ""; }
        }

        public static int CheckColumn(string StringVal , string idtovoem , string brand , string fieldname , int idZKP , int idStatusCheckLog)
        {
            if (StringVal == "")  
            {
                if(idStatusCheckLog != -1)
                {
                    string sql = $@"Insert into ZKPExchangeLog (idZKP , idTovOem , TmName, idStatusCheckLog, idSup) values({idZKP}, '{idtovoem}' , '{brand}', {idStatusCheckLog} , {idSup})";
                    DBExecutor.ExecuteQuery(sql);
                } 
                fields += fieldname + ","; // для Unilogger
                return 0;
            }
            else
            {
                return 1;
            }
        }

        public static DataTable FillDataTable(DataTableCollection dtcollection , int idZKP)
        {
            DataTable dtGetZKP = new DataTable();
            dtGetZKP.Columns.Add("idTovOEM");
            //dtGetZKP.Columns.Add("needcount");
            dtGetZKP.Columns.Add("countSup");
            dtGetZKP.Columns.Add("priceSup");
            dtGetZKP.Columns.Add("Brand");


            
            int checktemp = 0;
            foreach (DataTable dt in dtcollection)
            {
                checkSKU = false;
                AllSKUIncorrect = true; 
                var a = dt.Rows.Count;
                for (int i = 1; i < dt.Rows.Count; i++)
                {

                    string[] s = new string[4];
                    s[0] = CheckParse(dt.Rows[i][1].ToString(), 1, dt.Rows[i][0].ToString());
                    //s[1] = CheckParse(dt.Rows[i][3].ToString(), 2);
                    s[1] = CheckParse(dt.Rows[i][4].ToString(), 2);
                    s[2] = CheckParse(dt.Rows[i][5].ToString(), 2);
                    s[3] = dt.Rows[i][0].ToString();




                    checktemp += CheckColumn(s[0], s[0], s[3], "Артикул", idZKP, 2);
                    //checktemp += CheckColumn(s[1], "Количество");
                    checktemp += CheckColumn(s[1], s[0], s[3], "Кол-во поставщика", idZKP, 3);
                    checktemp += CheckColumn(s[2], s[0], s[3], "Цена поставщика", idZKP, 4);
                    checktemp += CheckColumn(s[3], s[0], s[3], "Бренд", idZKP, -1);


                    if (checktemp == s.Length)
                    {
                        AllSKUIncorrect = false; 
                        DataRow row = dtGetZKP.NewRow();
                        row["idTovOEM"] = s[0];
                        //row["needcount"] = s[1];
                        row["countSup"] = s[1];
                        row["priceSup"] = s[2];
                        row["Brand"] = s[3];
                        dtGetZKP.Rows.Add(row);
                    }
                    else
                    {
                         
                        checkSKU = true;
                        fields = fields.TrimEnd(',');
                        UniLogger.WriteLog("", 0, $"Некорректное заполнение для SKU Бренд - Артикул: {s[0].ToString() + " " + s[3].ToString()}, некорректно заполнены поле(я): {fields}. SKU не была добавлена в обработку КП");
                    }
                    fields = "";
                    checktemp = 0;

                }
                if (AllSKUIncorrect)
                {
                    string sql = $@"Insert into ZKPExchangeLog (idZKP , idStatusCheckLog , idSup) values({idZKP}, 7, {idSup} )";
                    DBExecutor.ExecuteQuery(sql);
                    UniLogger.WriteLog("", 3, "Некорректная структура списка КП, изменена поставщиком, IDZKP=" + idZKP.ToString());
                }

            }
                
           

            return dtGetZKP;
        }

        public static bool CheckActualKP(int idZKP)
        {
            var d = DateTime.Now.ToString("yyyy.MM.dd");
            string sql = $@"select DateEndWaitKP from ZKP where idZKP =  {idZKP} and  DateEndWaitKP   >= cast('{d}' as DateTime)";
            DataRow DateSend = DBExecutor.SelectRow(sql);
            if(DateSend == null)
            {
                sql = $@"Insert into ZKPExchangeLog (idZKP , idStatusCheckLog , idSup) values({idZKP}, 1, {idSup} )";
                DBExecutor.ExecuteQuery(sql);
                return false;
            }
            else
            {
                UniLogger.WriteLog("", 0, "Письмо актуально по периоду ожидания КП");
                return true;
            }
             
        }


        public static void ExportData(DataTable dt, int idZKP)
        {
            var vdatatable = dt;
            //string sql = $@"EXEC [dbo].[up_InsertReceiveZKP] {vdatatable} , {idZKP}";
            //DBExecutor.ExecuteQuery(sql);
            //using (var command = new SqlCommand("InsertTable") { CommandType = CommandType.StoredProcedure })
            //{
            //    var dt = new DataTable(); //create your own data table
            //    command.Parameters.Add(new SqlParameter("@myTableType", dt));
            //    SqlHelper.Exec(command);
            //}
            try
            {
                SqlParameter par = new SqlParameter("GetZKPIn", SqlDbType.Structured);
                par.Value = dt;
                par.TypeName = "GetZKP";
                SqlParameter par2 = new SqlParameter("idZKP", SqlDbType.Int);
                par2.Value = idZKP;
                int a = dt.Rows.Count;
                DBExecutor.ExeciteProcedure("up_InsertReceiveZKP", par, par2);
                if (!checkSKU) { string sql = $@"Insert into ZKPExchangeLog (idZKP , idStatusCheckLog , idSup) values({idZKP}, 0 , {idSup})"; DBExecutor.ExecuteQuery(sql); }
            }
            catch(Exception ex)
            {
                string sql = $@"Insert into ZKPExchangeLog (idZKP , idStatusCheckLog , idSup) values({idZKP}, 7, {idSup} )";
                DBExecutor.ExecuteQuery(sql);
                UniLogger.WriteLog("", 3, "Некорректная структура списка КП, изменена поставщиком, IDZKP=" + idZKP.ToString());
            }
                
        }

        public static int CheckMailSupplier(string address)
        {
            idSup = 0;
            var sql = $"select idKontrTitle from sKontrTitle (nolock) where trim(email) = trim('{address}')";
            int id;
            var res = DBExecutor.SelectSchalar(sql);
             

            if (res != null)
            {
                if (int.TryParse(res.ToString(), out id))
                {
                    idSup = Convert.ToInt32(res.ToString());
                    return id;
                }
                else
                { return -1; }
            }
            else
            { return -1; }
        }

        public static void CheckOverDueZKP()
        {
            string sql = @" update ZNPTov
                            set idTovStatus = 60                         
                            from ZNPTov                             
                            inner join ZKP on ZKP.idZNP = ZNPTov.idZNP 
                            inner join ZKPTov on ZKPTov.idZKP = ZKP.idZKP and ZKPTov.idTov = ZNPTov.idTov
                            inner join spr_tov on ZKPTov.idTov = spr_tov.id_tov 
                            where (ZKPTov.KolSup = 0 or  ZKPTov.PriceSup = 0) and  ZNPTov.idTovStatus < 70 and ZNPTov.idTovStatus > 40 and ZKP.DateEndWaitKP <  FORMAT(GetDate(), 'yyyy-MM-dd')";

            DBExecutor.ExecuteQuery(sql);

            sql = @" update ZNP
                            set idZNPStatus = 40                         
                            from ZNP
                            inner join ZNPTov on ZNPTov.idZNP = ZNP.idZNP                            
                            inner join ZKP on ZKP.idZNP = ZNPTov.idZNP 
                            inner join ZKPTov on ZKPTov.idZKP = ZKP.idZKP and ZKPTov.idTov = ZNPTov.idTov
                            inner join spr_tov on ZKPTov.idTov = spr_tov.id_tov 
                            where (ZKPTov.KolSup = 0 or  ZKPTov.PriceSup = 0) and  ZNP.idZNPStatus  < 50 and ZNP.idZNPStatus > 20 and ZKP.DateEndWaitKP <  FORMAT(GetDate(), 'yyyy-MM-dd')";

            DBExecutor.ExecuteQuery(sql);
        }

        public static void Start(MailType mailType)
        {
            int err = 0;

            SetCurrentMailParams(mailType);

            CheckOverDueZKP(); //Обновляем на "Нет КП" если время ожидания ответа ЗКП от поставщика вышло 

            UniLogger.WriteLog("", 0, "подключаюсь к почте");

            var imap = new ImapClient();
            try
            {
                imap.Connect(ProjectProperty.MailServer, 143, MailKit.Security.SecureSocketOptions.None);

                imap.Authenticate(currentMail, currentPassword);
            }
            catch (Exception ex)
            {
                err++;
                UniLogger.WriteLog("", 1, "Неудачное подключение к почтовому серверу. Строка для программиста: " + ex.Message);
                return;
            }
            
            imap.Inbox.Open(FolderAccess.ReadWrite);

            UniLogger.WriteLog("", 0, "выбор рабочей папки inbox. загружаю сообщения...");
            
            int countForParser = Convert.ToInt32(ProjectProperty.MailCountForParser);

            var uids = imap.Inbox.Search(SearchQuery.All);

            if (uids == null || uids.Count == 0)
            {
                UniLogger.WriteLog("", 1, "Нет писем в почтовом ящике");
                imap.Disconnect(true);
                return;
            }

            
            var collectionMess = imap.Inbox.Fetch(uids, MessageSummaryItems.Envelope | MessageSummaryItems.BodyStructure).Take(countForParser);
            
            UniLogger.WriteLog("", 0, "загружено сообщений: " + collectionMess.Count());
            
            try
            {
                int cnt = 0;
                bool isActual = true;
                string sql = @"select * from ZKPMailParam (nolock)";
                DataRow row = DBExecutor.SelectRow(sql);
                foreach (var mess in collectionMess)
                {
                    int count = collectionMess.Count();
                    string address = "", subject = "";
                    string subjectTemplate = row["Subject"].ToString();
                    int idZKP = 0;
                    checkSKU = false;
                   
                    address.Trim();
                    int id = 0;

                    if (id < 0)
                    {
                        UniLogger.WriteLog("", 3, $"В справочниках КИС отсутствует поставщик с email '{address}'!");
                        continue;
                    }

                    cnt++;
                    //mess.Subject = "Re: Тема ЗКП11 №132";
                    if (mess.Envelope.Subject != null)
                    {

                        if (mess.Envelope.Subject.Contains(subjectTemplate))
                        {
                            
                            foreach (var mailbox in mess.Envelope.From.Mailboxes)
                            {
                                address = mailbox.Address.ToString();
                            }
                             
                            address = Regex.Replace(address, "(?i)['<>]", "");
                            id = CheckMailSupplier(address.Trim());
                            subject = mess.Envelope.Subject;  //ins 21.01.22 Semenkina Считываем Тему письма для идентификации по маске кода Склада поставщика
                            subject = Regex.Replace(subject, "(?i)['<>]", "");

                            idZKP = int.Parse(FindidZKP(subject));
                            isActual = CheckActualKP(idZKP);
                        }
                        else if (mess.Envelope.Subject.Contains("purchase") || mess.Envelope.Subject.Contains("КП"))
                        {

                            if(mess.Attachments.ToList().Count > 0)
                            {
                                idZKP = int.Parse(FindidZKP(mess.Attachments.ToList().First().FileName));
                                isActual = CheckActualKP(idZKP);
                                if (!isActual) { UniLogger.WriteLog("", 0, "Ответ на заявку на коммерческое предложение не актуален"); }
                            }
                            else
                            {
                                imap.Inbox.Open(FolderAccess.ReadWrite);
                                var other_fold = imap.GetFolder(ProjectProperty.FolderForSimpleMessage) ;
                                var fa = other_fold.Open(MailKit.FolderAccess.ReadWrite);                                 
                                other_fold.AddFlags(mess.UniqueId, MessageFlags.Seen, true);

                                UniLogger.WriteLog("", 0, "По данному письму не было найдено вложения, письмо перемещено в папку Прочие");
                                imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(FolderAccess.ReadWrite);
                                imap.Inbox.Open(FolderAccess.ReadWrite);
                                imap.Inbox.MoveTo(mess.UniqueId, other_fold);
                            }
                            
                        }
                        else
                        {
 
                            UniLogger.WriteLog("", 0, "По данному письму не было найдено сопоставление, письмо перемещено в папку Прочие");
                            imap.Inbox.Open(FolderAccess.ReadWrite);
                            var other_fold = imap.GetFolder(ProjectProperty.FolderForSimpleMessage);
                            var fa = other_fold.Open(MailKit.FolderAccess.ReadWrite);
                            other_fold.AddFlags(mess.UniqueId, MessageFlags.Seen, true);
                            
                            imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(FolderAccess.ReadWrite);
                            imap.Inbox.Open(FolderAccess.ReadWrite);
                            imap.Inbox.MoveTo(mess.UniqueId, other_fold);
                            continue;
                        }
                        //idZKP = FindidZKP(idZKP);
                    }

                   
                    if ((isActual) && (mess.Envelope.Subject.Contains(subjectTemplate)))
                    {
                        try
                        {
                            bool parceAnuFile = false;
                            var attachments = mess.Attachments.ToList();
                            if (attachments.Count > 0)
                            {
                                UniLogger.WriteLog("", 0, "обнаружено вложений: " + attachments.Count.ToString());
                            }
                            else
                            {
                                sql = $@"Insert into ZKPExchangeLog (idZKP , idStatusCheckLog , idSup) values({idZKP},6, {idSup} )";
                                DBExecutor.ExecuteQuery(sql);

                                imap.Inbox.Open(FolderAccess.ReadWrite);
                                var err_fold = imap.GetFolder(ProjectProperty.FolderForErrorMessage);
                                var fo = err_fold.Open(MailKit.FolderAccess.ReadWrite);
                                imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(MailKit.FolderAccess.ReadWrite);
                                imap.Inbox.Open(FolderAccess.ReadWrite);
                                imap.Inbox.MoveTo(mess.UniqueId, err_fold);

                                UniLogger.WriteLog("", 1, "нет вложений!!!, обработка письма закончена");
                                UniLogger.WriteLog("", 3, "письмо переместили в Error");
                                 
                            }

                            foreach (var attachment in attachments)
                            {
                                //setting = settings.FirstOrDefault();
                                UniLogger.WriteLog("", 0, "Чтение содержимого письма");
                                if (attachment.FileName == null)
                                {
                                    UniLogger.WriteLog("", 0, "некорректное имя файла вложения");
                                    continue;
                                }

                                //if (setting.MailFileMask != "")
                                //{
                                //    setting = settings.FirstOrDefault(p => attachment.FileName.ToLower().IndexOf(p.MailFileMask.ToLower()) != -1);
                                //    if (setting == null)
                                //    {
                                //        UniLogger.WriteLog("", 0, "некорректная маска файла");
                                //        continue;
                                //    }
                                //}

                                var mime = (MimePart)imap.Inbox.GetBodyPart(mess.UniqueId, attachment);
                                var fileName = mime.FileName;
                                string filePath = ProjectProperty.FolderXls + "\\" + attachment.FileName + (attachment.FileName.IndexOf('.') == -1 ? ".zip" : "");
                                if (string.IsNullOrEmpty(fileName)) { continue; }

                                using (var stream = File.Create(filePath))
                                    mime.ContentObject.DecodeTo(stream);


                                //string filePath = ProjectProperty.FolderXls + "\\" + attachment.FileName + (attachment.FileName.IndexOf('.') == -1 ? ".zip" : "");

                                //attachment.Save(filePath);
                                UniLogger.WriteLog("", 0, "Сохранен файл " + filePath + " Начинаю парсинг файла.");

                                var extension = filePath.Split('.').Last();

                                string arhievePath = "";
                                if (extension == "zip" || extension == "rar")
                                {
                                    SevenZipExtractor.SetLibraryPath(ProjectProperty.PathTo7Zip);
                                    SevenZip.SevenZipExtractor ex = new SevenZipExtractor(filePath);

                                    arhievePath = ProjectProperty.FolderXls + "\\" + DateTime.Now.ToString().Replace(".", "").Replace(":", "");
                                    ex.ExtractArchive(arhievePath);

                                    File.Delete(filePath);
                                    var filenames = Directory.GetFiles(arhievePath);
                                    filePath = filenames.First();
                                }

                                extension = filePath.Split('.').Last().ToLower();
                                UniLogger.WriteLog("", 0, "Расширение файла: " + extension);

                                if (extension != "xls" && extension != "xlsx" && extension != "xlsm" && extension != "csv" && extension != "txt")
                                {
                                    sql = $@"Insert into ZKPExchangeLog (idZKP , idStatusCheckLog , idSup) values({idZKP},5, {idSup} )";                                     
                                    DBExecutor.ExecuteQuery(sql);

                                    imap.Inbox.Open(FolderAccess.ReadWrite);
                                    var err_fold = imap.GetFolder(ProjectProperty.FolderForErrorMessage);
                                    var fo = err_fold.Open(MailKit.FolderAccess.ReadWrite);
                                    imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(MailKit.FolderAccess.ReadWrite);
                                    imap.Inbox.Open(FolderAccess.ReadWrite);
                                    imap.Inbox.MoveTo(mess.UniqueId, err_fold);
                                    UniLogger.WriteLog("", 3, "письмо переместили в Error");

                                    UniLogger.WriteLog("", 0, "некорректное расширение файла");
                                    continue;
                                }

                                DataTableCollection dtParceFile = ReadFileAllFormat(filePath);

                                if (dtParceFile == null)
                                {
                                    UniLogger.WriteLog("", 1, "ВНИМАНИЕ! Не удалось конвертировать файл в таблицу. IDZKP=" + idZKP.ToString());                                    
                                    File.Delete(filePath);
                                    continue;
                                }
                                else
                                {
                                    DataTable GetZKP = FillDataTable(dtParceFile, idZKP);
                                    ExportData(GetZKP, idZKP);
                                }

                                GC.Collect();

                                parceAnuFile = true;
                                UniLogger.WriteLog("", 0, "Парсинг выполнен, файл будет удален.");
                                File.Delete(filePath);
                                //if (arhievePath != "")
                                //Directory.Delete(arhievePath, true);

                                //break;

                                if (parceAnuFile)
                                {
                                    if (AllSKUIncorrect) //!checkSKU
                                    {
                                        imap.Inbox.Open(FolderAccess.ReadWrite);
                                        var err_fold = imap.GetFolder(ProjectProperty.FolderForErrorMessage);
                                        var fo = err_fold.Open(MailKit.FolderAccess.ReadWrite);
                                        imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(MailKit.FolderAccess.ReadWrite);
                                        imap.Inbox.Open(FolderAccess.ReadWrite);
                                        imap.Inbox.MoveTo(mess.UniqueId, err_fold);
                                        UniLogger.WriteLog("", 3, "письмо переместили в Error");
                                    }
                                    else
                                    {
                                        imap.Inbox.Open(FolderAccess.ReadWrite);
                                        var load_fold = imap.GetFolder(ProjectProperty.FolderForReadedMessages);
                                        var fa = load_fold.Open(FolderAccess.ReadWrite);
                                        load_fold.AddFlags(mess.UniqueId, MessageFlags.Seen, true);

                                        imap.GetFolder(ProjectProperty.FolderForReadedMessages).Open(FolderAccess.ReadWrite);
                                        imap.Inbox.Open(FolderAccess.ReadWrite);
                                        imap.Inbox.MoveTo(mess.UniqueId, load_fold);
                                    }
                                     
                                }                                     
                                else
                                {
                                    imap.Inbox.Open(FolderAccess.ReadWrite);
                                    var other_fold = imap.GetFolder(ProjectProperty.FolderForSimpleMessage);
                                    var fa = other_fold.Open(MailKit.FolderAccess.ReadWrite);
                                    imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(FolderAccess.ReadWrite);
                                    other_fold.AddFlags(mess.UniqueId, MessageFlags.Seen, true);

                                    imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(FolderAccess.ReadWrite);
                                    imap.Inbox.Open(FolderAccess.ReadWrite);
                                    imap.Inbox.MoveTo(mess.UniqueId, other_fold);
                                    UniLogger.WriteLog("", 0, "письмо переместили в обработанные");
                                }                                                                    
                            }// foreach attachments 
                        }
                        catch (Exception ex)
                        {
                            err++;
                            MessageBox.Show(ex.Message + ex.StackTrace);

                            imap.Inbox.Open(FolderAccess.ReadWrite);
                            var err_fold = imap.GetFolder(ProjectProperty.FolderForErrorMessage);
                            var fa = err_fold.Open(MailKit.FolderAccess.ReadWrite);
                            imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(MailKit.FolderAccess.ReadWrite);
                            imap.Inbox.Open(FolderAccess.ReadWrite);
                            imap.Inbox.MoveTo(mess.UniqueId, err_fold);
                            UniLogger.WriteLog("", 1, "При чтении файла возникли ошибки. От: " + address + " Строка для программиста: " + (ex.Message.Length > 500 ? ex.Message.Remove(500) : ex.Message));
                            
                            continue;
                        }

                    }
                    else
                    {
                        imap.Inbox.Open(FolderAccess.ReadWrite);
                        var other_fold = imap.GetFolder(ProjectProperty.FolderForSimpleMessage);
                        var fa = other_fold.Open(MailKit.FolderAccess.ReadWrite);
                        other_fold.AddFlags(mess.UniqueId, MessageFlags.Seen, true);

                        imap.GetFolder(ProjectProperty.FolderForSimpleMessage).Open(FolderAccess.ReadWrite);
                        imap.Inbox.Open(FolderAccess.ReadWrite);
                        imap.Inbox.MoveTo(mess.UniqueId, other_fold);                         
                        UniLogger.WriteLog("", 0, "Письмо поставщика пришло после периода ожидания , письмо переместили в обработанные");
                    }

                }
                imap.Disconnect(true);
                
                

            }
            catch (Exception exa)
            {
                err++;
                MessageBox.Show(exa.Message);
                UniLogger.WriteLog("",3, "Ошибка метода Start класс  MailPriceReader: " + exa.Message);
            }
            finally
            {
                UniLogger.WriteLog("", 0, "Работа закончена...");
                
                UniLogger.Flush();

                GC.Collect();
            }
        }

        public static DataTableCollection ReadFileAllFormat(string filePath)   //ParserSettings setting
        {
            //UniLogger.WriteLog("", 0, "Флаг ручной загрузки: " + setting.fHardLoad.ToString());
            var extension = filePath.Split('.').Last();

            var result = FileReader.Read(filePath);
            if (result == null)
            {
                UniLogger.WriteLog("", 3, "Файл имеет неизвестное расширение " + extension);
                return null;
            }
            return result;
        }
    }
}

