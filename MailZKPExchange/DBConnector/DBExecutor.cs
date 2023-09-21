using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Runtime;
using System.Xml;

namespace MailZKPExchange.DBConnector
{
    public static class DBExecutor
    {
        public static DataSet SelectDataSet(string Query, params SqlParameter[] Parameters)
        {
            if (Connection.ConnectToDataBase())
            {
                DataSet ds = new DataSet();
                SqlDataAdapter da = new SqlDataAdapter(Query, Connection.SqlConnection);
                AddParameters(da.SelectCommand, Parameters);
                da.SelectCommand.CommandTimeout = 0;
                da.Fill(ds);
                Connection.CloseConnection();
                return ds;
            }
            return null;
        }

        private static void AddParameters(SqlCommand sqlCommand, SqlParameter[] Parameters)
        {
            if (Parameters != null && Parameters.Count() > 0)
            {
                foreach (var par in Parameters)
                    sqlCommand.Parameters.Add(par);
            }
        }


        public static DataTable SelectTable(string Query, params SqlParameter[] Parameters)
        {
            DataSet result = SelectDataSet(Query, Parameters);
            if (result != null && result.Tables.Count > 0)
                return result.Tables[0];
            return null;
        }

        public static DataRow SelectRow(string Query, params SqlParameter[] Parameters)
        {
            DataTable result = SelectTable(Query, Parameters);
            if (result != null && result.Rows.Count > 0)
                return result.Rows[0];
            return null;
        }

        public static object SelectSchalar(string Query, params SqlParameter[] Parameters)
        {
            if (Connection.ConnectToDataBase())
            {
                SqlCommand cmd = new SqlCommand(Query, Connection.SqlConnection);
                AddParameters(cmd, Parameters);
                object result = cmd.ExecuteScalar();
                Connection.CloseConnection();
                return result;
            }
            return null;
        }

        public static bool ExecuteTranzactionQuery(string Query)
        {
            return ExecuteQuery("begin tran " + Query + " commit tran");
        }

        public static bool ExecuteQuery(string Query, params SqlParameter[] Parameters)
        {
            if (Connection.ConnectToDataBase())
            {
                SqlCommand cmd = new SqlCommand(Query, Connection.SqlConnection);
                AddParameters(cmd, Parameters);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
                Connection.CloseConnection();
                return true;
            }
            return false;
        }

        public static bool ExeciteProcedure(string NameProcedure, params SqlParameter[] Parameters)
        {
            if (Connection.ConnectToDataBase())
            {
                SqlCommand cmd = new SqlCommand(NameProcedure, Connection.SqlConnection);
                cmd.Parameters.Clear();
                AddParameters(cmd, Parameters);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.ExecuteNonQuery();
                Connection.CloseConnection();
                return true;
            }
            return false;
        }

        public static DataTable ExecuteProcedureTable(string NameProcedure, params SqlParameter[] Parameters)
        {
            if (Connection.ConnectToDataBase())
            {
                SqlCommand cmd = new SqlCommand(NameProcedure, Connection.SqlConnection);
                AddParameters(cmd, Parameters);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();

                da.Fill(ds);
                Connection.CloseConnection();

                if (ds != null && ds.Tables.Count > 0)
                    return ds.Tables[0];

            }
            return null;
        }

        /// <summary>
        /// Используем класс SLTableRow для замены DataTable в целях экономии памяти
        /// </summary>
        /// <param name="NameProcedure"></param>
        /// <param name="Parameters"></param>
        /// <returns></returns>
        public static List<SLTableRow> ExecuteProcedureTableReadOnly(string NameProcedure, params SqlParameter[] Parameters)
        {
            if (Connection.ConnectToDataBase())
            {
                using (SqlCommand sqlCmd = new System.Data.SqlClient.SqlCommand(NameProcedure, Connection.SqlConnection) { CommandType = CommandType.StoredProcedure })
                {
                    //GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                    GC.Collect();

                    sqlCmd.CommandTimeout = 200;
                    sqlCmd.Parameters.Add("@datefrom", SqlDbType.DateTime).Value = Parameters[0].Value;
                    sqlCmd.Parameters.Add("@dateto", SqlDbType.DateTime).Value = Parameters[1].Value;

                    using (SqlDataReader dr = sqlCmd.ExecuteReader())
                    {

                        if (dr.HasRows)
                        {
                            SLTable slt = new SLTable();
                            slt.SLList.Capacity = 100000;

                            int i = 0;
                            while (dr.Read())
                            {
                                SLTableRow sltr = new SLTableRow();
                                sltr.sumSale = dr.IsDBNull(0) ? 0 : Convert.ToDecimal(dr[0]);
                                sltr.kolSale = dr.IsDBNull(1) ? 0 : Convert.ToDecimal(dr[1]);
                                sltr.nTerritoryShort = dr.IsDBNull(2) ? "" : dr[2].ToString();
                                sltr.Profit = dr.IsDBNull(3) ? 0 : Convert.ToDecimal(dr[3]);
                                sltr.Profittrunsport = dr.IsDBNull(4) ? 0 : Convert.ToDecimal(dr[4]);
                                sltr.MonthSale = dr.IsDBNull(5) ? -1 : Convert.ToInt32(dr[5]);
                                sltr.YearSale = dr.IsDBNull(6) ? -1 : Convert.ToInt32(dr[6]);
                                sltr.nTypeTov = dr.IsDBNull(7) ? "" : dr[7].ToString();
                                sltr.nameGroup = dr.IsDBNull(8) ? "" : dr[8].ToString();
                                sltr.nAgent = dr.IsDBNull(9) ? "" : dr[9].ToString();
                                sltr.nPodgroup = dr.IsDBNull(10) ? "" : dr[10].ToString();
                                sltr.nKontr = dr.IsDBNull(11) ? "" : dr[11].ToString();
                                sltr.tm_name = dr.IsDBNull(12) ? "" : dr[12].ToString();
                                sltr.Oblast = dr.IsDBNull(13) ? "" : dr[13].ToString();
                                sltr.Segment = dr.IsDBNull(14) ? "" : dr[14].ToString();
                                sltr.TovLine = dr.IsDBNull(15) ? "" : dr[15].ToString();
                                sltr.TOV_NAME = dr.IsDBNull(16) ? "" : dr[16].ToString();
                                sltr.in_tax = dr.IsDBNull(17) ? "" : dr[17].ToString();
                                sltr.forder = dr.IsDBNull(18) ? "" : dr[18].ToString();
                                sltr.nCategory = dr.IsDBNull(19) ? "" : dr[19].ToString();

                                slt.SLList.Add(sltr);
                                i++;
                                sltr = null;
                                if (i % 100000 == 0)
                                {
                                    sltr = null;
                                    GC.Collect();
                                }

                            }
                            GC.Collect();

                            //return ToDataSet(slt.SLList).Tables[0];
                            //return ToDataTable(slt.SLList);
                            return slt.SLList;
                        }
                    }
                }
            }

            return null;
        }

        public static DataTable ToDataTable(List<SLTableRow> sltr)
        {
            Type elementType = typeof(SLTableRow);
            DataTable t = new DataTable();

            foreach (var propInfo in elementType.GetProperties())
            {
                t.Columns.Add(propInfo.Name, propInfo.PropertyType);
            }

            foreach (SLTableRow item in sltr)
            {
                DataRow row = t.NewRow();
                foreach (var propInfo in elementType.GetProperties())
                {
                    row[propInfo.Name] = propInfo.GetValue(item, null);
                }

                t.Rows.Add(row);
                row = null;
            }
            elementType = null;
            GC.Collect();
            return t;
        }
    }

    public class SLTable
    {
        public List<SLTableRow> SLList { get; set; }
        public SLTable()
        {
            SLList = new List<SLTableRow>();
        }
    }

    public class SLTableRow
    {
        public decimal sumSale { get; set; }
        public decimal kolSale { get; set; }
        public string nTerritoryShort { get; set; }
        public decimal Profit { get; set; }
        public decimal Profittrunsport { get; set; }
        public int MonthSale { get; set; }
        public int YearSale { get; set; }
        public string nTypeTov { get; set; }
        public string nameGroup { get; set; }
        public string nAgent { get; set; }
        public string nPodgroup { get; set; }
        public string nKontr { get; set; }
        public string tm_name { get; set; }
        public string Oblast { get; set; }
        public string Segment { get; set; }
        public string TovLine { get; set; }
        public string TOV_NAME { get; set; }
        public string in_tax { get; set; }
        public string forder { get; set; }
        public string nCategory { get; set; }
    }
}
