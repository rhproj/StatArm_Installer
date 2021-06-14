using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IWshRuntimeLibrary;

namespace StatArm_Installer01
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            label1.BackColor = Color.Transparent; 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(@"\\fsttr02\стат. отчеты\AS4.6\ARM STAT 2019\"))
            {
                CopyAll(new DirectoryInfo(@"\\fsttr02\стат. отчеты\AS4.6\ARM STAT 2019\")
                    , new DirectoryInfo(@"c:\ARM STAT 2019\"));

                ShortCut(@"C:\ARM STAT 2019\", "STAT_ARM.exe", @"\Stat Arm.lnk", @"\EXPORT.lnk");
            }
            else
            {
                MessageBox.Show("Не удалось установить соединение fsttr02");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(@"\\fsttr02\стат. отчеты\AS4.6\ARM STAT SU\"))
            {
                CopyAll(new DirectoryInfo(@"\\fsttr02\стат. отчеты\AS4.6\ARM STAT SU\")
                           , new DirectoryInfo(@"c:\ARM STAT SU\"));

                ShortCut(@"C:\ARM STAT SU\", "STAT_ARM_SU.exe", @"\StatArm SU.lnk", @"\EXPORT SU.lnk");
            }
            else
            {
                MessageBox.Show("Не удалось установить соединение fsttr01");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Базы не определены!");
            Form2ArmLocal fL = new Form2ArmLocal();
            fL.ShowDialog();
        }

        private void CopyAll(DirectoryInfo fromD, DirectoryInfo toD)
        {
            progressBar1.ForeColor = Color.LimeGreen;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;

            Directory.CreateDirectory(toD.FullName);
            //copy files
            foreach (FileInfo fI in fromD.GetFiles())
            {
                progressBar1.PerformStep();
                lbProgress.Text = toD.FullName; //shows up what's been copied
                fI.CopyTo(Path.Combine(toD.FullName, fI.Name), true);
            }
            //copy sub-dirs recursively
            foreach (DirectoryInfo sourceDirs in fromD.GetDirectories())
            {
                DirectoryInfo targDirs = toD.CreateSubdirectory(sourceDirs.Name);
                CopyAll(sourceDirs, targDirs);
            }
            progressBar1.Value = progressBar1.Maximum;
        }

        private void ShortCut(string dir, string arm_targ, string arm_lnk, string fld_lnk)
        {
            object shCutDesk = (object)"Desktop"; 
            
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

            MessageBox.Show("Готово!");
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Lime; //FromArgb(0, 255, 42);
            
        }
        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.ForeColor = Color.White;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.ForeColor = Color.Lime;
        }
        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.ForeColor = Color.White;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.ForeColor = Color.Lime;
        }
        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.ForeColor = Color.White;
        }
    }
}
