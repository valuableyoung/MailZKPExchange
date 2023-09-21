using MailZKPExchange.Helpers;
using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Text;

namespace MailZKPExchange.Heplers
{
    public class FileReader
    {
        public static DataTableCollection Read(string filepath, string separator = "", bool withHeaders = false)
        {

            switch (Path.GetExtension(filepath).ToLower())
            {

                case ".xls":
                    return Read_EXCEL(filepath, withHeaders);


                case ".xlsx":
                    return Read_EXCEL(filepath, withHeaders);

                case ".csv":
                    return Read_CSV(filepath, separator);

                case ".txt":
                    return Read_CSV(filepath);
            }

            return null;

        }


        public static DataTableCollection Read_EXCEL(String filepath, bool withHeaders)
        {

            DataSet result = new DataSet();

            try
            {
                using (var stream = File.Open(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = Encoding.Default }))
                    {
                        UniLogger.WriteLog("", 1, "Файл загружен. Прочитано строк: " + reader.RowCount.ToString());

                        if (withHeaders)
                        {
                            result = reader.AsDataSet(new ExcelDataSetConfiguration { ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true } });
                        }
                        else
                        {
                            result = reader.AsDataSet();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                var exm = ex.Message;
                result = ExcelXMLtoDataTable.ImportExcelXML(filepath, true, false);
            }
            return result.Tables;

        }

        public static DataTableCollection Read_CSV(String filepath, string separator = "")
        {
            DataSet result = new DataSet();


            if (separator != "")
            {
                DataSet dt = new DataSet();
                using (StreamReader sr = new StreamReader(filepath, Encoding.Default))
                {
                    string[] headers = sr.ReadLine().Split(separator[0]);
                    dt.Tables.Add();
                    foreach (string header in headers)
                    {
                        dt.Tables[0].Columns.Add(header);
                    }
                    while (!sr.EndOfStream)
                    {
                        string[] rows = sr.ReadLine().Split(separator[0]);
                        DataRow dr = dt.Tables[0].NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                        }
                        dt.Tables[0].Rows.Add(dr);
                    }

                }

                return dt.Tables;

            }
            else
            {
                using (var stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = Encoding.Default, AutodetectSeparators = new char[] { ',', ';', '\t', '|', '#' } }))
                    {
                        result = reader.AsDataSet();
                    }
                }
            }

            return result.Tables;

        }

    }
}

