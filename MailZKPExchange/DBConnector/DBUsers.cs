using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace MailZKPExchange.DBConnector
{
    public static class DBUsers
    {
        public static object GetCurrentWindowUserId()
        {
            string query = @"select id_kontr from spr_kontr (nolock) where sql_id = dbo.f_user_id('Arkona')";
            return DBExecutor.SelectSchalar(query);
        }

        public static object GetCurrentWindowUserName()
        {
            string query = @"select n_kontr from spr_kontr (nolock) where sql_id = dbo.f_user_id('Arkona')";
            return DBExecutor.SelectSchalar(query);
        }

        public static DataRow checkUser(string userlogin, string userpassword)
        {
            return DBExecutor.SelectRow("select u.idUser, k.n_kontr_full  from sLoginUser u (nolock) join spr_kontr k (nolock) on u.idUser = k.id_kontr where u.Login = @login and u.Pass = @pass", new System.Data.SqlClient.SqlParameter("@login", userlogin), new System.Data.SqlClient.SqlParameter("@pass", userpassword));
        }
    }
}

