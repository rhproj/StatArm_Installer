using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using IWshRuntimeLibrary;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace StatArm_Installer01
{
    public partial class Form2ArmLocal : Form
    {
        public Form2ArmLocal()
        {
            InitializeComponent();
        }
        private void Form2ArmLocal_Load(object sender, EventArgs e)
        {//had to specify System.IO.File bcz .File also used by IWshR..
            if (System.IO.File.Exists(@"\\fsttr02\стат. отчеты\СТАТИСТИКА\SHIF\SHIF.txt"))  //@"G:\172.16.252.4\стат. отчеты\СТАТИСТИКА\SHIF\SHIF.txt"))  //test
            {
                string[] districts = System.IO.File.ReadAllLines(@"\\fsttr02\стат. отчеты\СТАТИСТИКА\SHIF\SHIF.txt", Encoding.Default);
                foreach (string l in districts)
                {
                    comboBox1.Items.Add(l); //populating combobox with Districts
                }
            }
            else
            {
                MessageBox.Show("Не удалось найти файл SHIF на сетевом диске fsttr02 :( ");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(@"\\fsttr02\стат. отчеты\AS4.6\ARM_STAT\"))  //@"G:\172.16.252.4\стат. отчеты\ARM_STAT"))  //test
            {
                Cursor.Current = Cursors.WaitCursor;

                CopyAll(new DirectoryInfo(@"\\fsttr02\стат. отчеты\AS4.6\ARM_STAT\")  //@"G:\172.16.252.4\стат. отчеты\ARM_STAT")  //test
                      , new DirectoryInfo(@"c:\ARM_STAT\"));

                lblProgress.Text = "Копируется папка FORM...";

                CopyAll(new DirectoryInfo(@"\\fsttr02\стат. отчеты\СТАТИСТИКА\FORM\")
                      , new DirectoryInfo(@"c:\ARM_STAT\FORM\"));

                WriteShif(@"C:\ARM_STAT\STAT_WIN.INI", @"C:\ARM_STAT\SHIF\shif.txt", @"C:\ARM_STAT\SHIF\summ.txt");

                ShortCut(@"C:\ARM_STAT\", "STAT_ARM.exe", @"\Stat_Arm.lnk", @"\_EXPORT.lnk");
                //progressBar1.Value = progressBar1.Maximum;

                Cursor.Current = Cursors.Default;
                DialogResult dialog = MessageBox.Show("Готово!");
                if (dialog == DialogResult.OK)
                {
                    Application.Exit();
                }
            }
            else
            {
                MessageBox.Show("Не удалось найти необходимую папку на сетевом диске fsttr02 или нет соединения с диском! :(");
            }
        }

        private void WriteShif(string stat_win, string shif, string summ)
        {
            if (comboBox1.Text != String.Empty)
            {
                string statWinConfig = System.IO.File.ReadAllText(@"C:\ARM_STAT\STAT_WIN.INI", Encoding.Default);

                int space = comboBox1.Text.IndexOf(' ');
                string district = $"{comboBox1.Text.Substring(0, space)}-{comboBox1.Text.Substring(space + 1)}";

                statWinConfig = Regex.Replace(statWinConfig, @"summ.txt", district);
                System.IO.File.WriteAllText(@"C:\ARM_STAT\STAT_WIN.INI", statWinConfig, Encoding.Default);

                string num = $"{comboBox1.Text.Substring(0, space)}:";  //extracting only number of the district:
                System.IO.File.WriteAllText(shif, comboBox1.Text); //filling shif.txt
                System.IO.File.WriteAllText(summ, num);
            }
            else
            {
                MessageBox.Show("Не выбран район!");
            }
        }

        private void CopyAll(DirectoryInfo fromD, DirectoryInfo toD)
        {
            Directory.CreateDirectory(toD.FullName);

            //copy files
            foreach (FileInfo fI in fromD.GetFiles())
            {
                int qt = fromD.GetFiles("*", SearchOption.AllDirectories).Length;

                progressBar1.Maximum = qt;
                progressBar1.Value += 1;
                lblProgress.Text = fI.Name;
                //lblProgress.Text = toD + "/" + fI.Name;

                fI.CopyTo(Path.Combine(toD.FullName, fI.Name), true);
            }
            progressBar1.Value = 0;
            //copy sub-dirs recursively
            foreach (DirectoryInfo sourceDirs in fromD.GetDirectories())
            {
                DirectoryInfo targDirs = toD.CreateSubdirectory(sourceDirs.Name);
                CopyAll(sourceDirs, targDirs);
            }
        }

        private void ShortCut(string dir, string arm_targ, string arm_lnk, string fld_lnk)
        {
            object shCutDesk = (object)"Desktop"; //i guess that's how you create Dktp obj to reference to

            WshShell shell = new WshShell(); //Shell - interface to OS, provides work with elements like: W.Explorer, Start Menu

            string strShortcutPath = (string)shell.SpecialFolders.Item(ref shCutDesk) + arm_lnk; //Desktop Path
            string strFolderShcutPath = (string)shell.SpecialFolders.Item(ref shCutDesk) + fld_lnk;
            IWshShortcut shcut = shell.CreateShortcut(strShortcutPath); //actual act of creating ShCut
            IWshShortcut shFolcut = shell.CreateShortcut(strFolderShcutPath);

            shcut.WorkingDirectory = dir;
            shcut.TargetPath = dir + arm_targ; //"STAT_ARM - 2019-03-26.exe";//txtExePath.Text;
            shFolcut.TargetPath = dir + "_EXPORT";

            shcut.Save(); //saving changes of our shcut
            shFolcut.Save();
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Lime;
        }
        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.ForeColor = Color.White;
        }
    }
}
