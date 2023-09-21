using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MailZKPExchange.Base
{
    public static class WindowOpener
    {
        public static WFMain MainForm;

        public static void OpenWindow(Form wind)
        {
            MainForm.OpenWindow(wind);
        }

    }
}
