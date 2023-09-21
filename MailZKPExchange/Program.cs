using MailZKPExchange;
using MailZKPExchange.Base;
using MailZKPExchange.DBConnector;
using MailZKPExchange.SPR;
using System;
using System.Windows.Forms;
using System.Diagnostics;

namespace AForm.ExelForm
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            //при сборке ботов оставляем пользователя EasZakTov и раскомментируем вызов конкретного бота в args[2]
            //при сборке десктопных приложений комментим вообще все внутри директивы #Region 
            #region Закомментировать все перед релизом

            args = new string[3];

            args[0] = "EasZakTov";
            args[1] = "ddfi3)es";

            args[2] = ((int)eForm.РоботПарсингаЗКП).ToString();


            #endregion


            ProjectProperty.LoadDataAppConfig(); //  ProjectProperty.LoadDataAppConfig(true)  real
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length > 0)
            {
                if (User.LoginUser(args[0], args[1]))
                {
                    Form form = new WFMain(args[2]);
                    WindowOpener.MainForm = (WFMain)form;
                    Application.Run(form);
                    GC.Collect();
                    //Process.GetCurrentProcess().Kill();
                } 
                          
            }
        }
    }
    public enum eForm
    {
        Пусто = 0,        
        РоботПарсингаЗКП = 31
    }
}

