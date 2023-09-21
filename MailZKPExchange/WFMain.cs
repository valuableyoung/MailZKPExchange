using MailZKPExchange.Base;

using MailZKPExchange.SPR;
using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace MailZKPExchange
{
    public partial class WFMain : Form
    {

        eForm fname = eForm.Пусто;
        public WFMain(object x = null)
        {
            InitializeComponent();
            this.Text = Text + ", Сервер - " + ConfigurationManager.AppSettings["DBServer"] + ", База - " + ConfigurationManager.AppSettings["DBBase"] + ", Пользователь - " + User.Current.NKontrFull;
            fname = x == null ? eForm.Пусто : (eForm)(int.Parse(x.ToString()));
        }

        public void OpenWindow(Form wind)
        {
            if (wind.Tag == null)
            {
                MessageBox.Show("У окна не указан Тэг!!!");
                return;
            }
            var w = this.MdiChildren.FirstOrDefault(p => p.Tag.ToString() == wind.Tag.ToString());

            if (w != null)
            {
                w.Focus();
                return;
            }

            wind.MdiParent = this;
            wind.Show();
        }

        private void WFMain_Load(object sender, EventArgs e)
        {
            switch (fname)
            {
                case eForm.РоботПарсингаЗКП:
                    BotParceZKP();
                    break;
            }
        }

        private void BotParceZKP()
        {
            MailZKPExchange.Parsers.MailZKPReader.Start(Parsers.MailZKPReader.MailType.Simple);
            Close();
        }


        public enum eForm
        {
            Пусто = 0,
 
            РоботПарсингаЗКП = 31
        }
    }
}
