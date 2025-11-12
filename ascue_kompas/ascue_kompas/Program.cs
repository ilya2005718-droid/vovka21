using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Collections;
using ascue_kompas;
using System.Threading;

namespace ascue_kompas
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
            // Application.Run(new View());
            View v = new View();
            Model m = new Model();
            Controller cn = new Controller(v, m);
            Application.Run(v);
           
        }
      
    }
}
