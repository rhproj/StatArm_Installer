using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace StatArm_Installer01
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles(); //needed to be able to tweek parogresbar, label color and such
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
