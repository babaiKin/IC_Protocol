using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace AccessDB
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form2());
        }
    }

    static class My
    {
        internal static int poveritelListLength = new int();
        internal static string [] poveritelList = new string[500];

        internal static int oborudovanieListLength = new int();
        internal static string[] oborudovanieList = new string[500];
    }
}
