using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using MailZKPExchange.DBConnector;

namespace MailZKPExchange.SPR
{
    public class User
    {

        public int IdUser { get; set; }
        public string NKontrFull { get; set; }
        public bool FDeveloper { get; set; }
        private static User _user;

        public static bool LoginUser(string userlogin = "", string userpassword = "")
        {
            DataRow row = DBUsers.checkUser(userlogin, userpassword);

            if (row != null)
            {
                _user = new User((int)row["idUser"], row["n_kontr_full"].ToString());
                return true;
            }

            return false;

        }

        public static User Current { get { return _user; } }

        public static int CurrentUserId { get { return _user == null ? 0 : _user.IdUser; } }

        public static User getInstance()
        {
            if (_user != null) return _user;
            return null;
        }

        public User(int idUser, string n_kontr_full)
        {
            this.IdUser = idUser;
            this.NKontrFull = n_kontr_full;
            this.FDeveloper = true;
        }

        public static int GetPostByUserId(int idUser)
        {
            string sql = @"
                select id_post from spr_kontr where id_kontr = @idKontr
";
            SqlParameter paridKontr = new SqlParameter("idKontr", idUser);
            return int.Parse( DBExecutor.SelectSchalar(sql, paridKontr).ToString());
        }

        public static bool InRole(int idUser, string NRole)
        {

            string sql = @"
                    select count(*) from rRoleUser
	                join sRole on rRoleUser.IdRole = sRole.IdRole
	                where IdKontr = @idKontr
	                and sRole.NRole = @NRole

                ";
            SqlParameter paridKontr = new SqlParameter("idKontr", idUser);
            SqlParameter parRoleName = new SqlParameter("NRole", NRole);
            var res = DBExecutor.SelectSchalar(sql, paridKontr, parRoleName);

            if (res.ToString() == "1") return true;

            return false;


        }
    }
}

