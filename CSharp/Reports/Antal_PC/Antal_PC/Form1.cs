using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Diagnostics;
using Microsoft.PowerShell;
using System.IO;
using System.Linq;
using System.Collections;
using System.Threading;

namespace Antal_PC
{
    public partial class Form1 : Form
    {
        Excel_handler ex = new Excel_handler();
        
        public Form1()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                status_lb.Items.Clear();
                SetStatusLb("Startar rapport...");
                ex.ivantiFile = BackupFile(textBox1.Text);
                ex.ivantiSheet = "Grunddata";
                SetStatusLb($"Öppnar {ex.ivantiFile}");
                SetStatusLb($"Aktivt Sheet är '{ex.ivantiSheet}'");
                SetStatusLb($"Bearbetar Ivanti data...");

                var ivantiDepartments = ex.CheckDepartmentsInIvantiFile();

                SetStatusLb($"Antal KST i ivanti {ivantiDepartments.Count.ToString()}");
                SetStatusLb($"------------------------------");


                ex.antalPCFile = BackupFile(textBox2.Text);
                ex.antalPCSheet = "Innevarnde Månad";
                SetStatusLb($"Öppnar {ex.antalPCFile}");
                SetStatusLb($"Sheet är '{ex.antalPCSheet}'");
                SetStatusLb($"Bearbetar Antal PC data...");

                var antalPCDep = ex.CheckDepartmentsInPCFile();

                SetStatusLb($"Antal KST i Antal PC {antalPCDep.Count.ToString()}");
                SetStatusLb($"------------------------------");
                //List<string> ivantiDiff = CheckIvantiDepList(ivantiDepartments, antalPCDep);
                //List<string> pcDiff = CheckPCDepList(ivantiDepartments, antalPCDep);
                //List<string> difference = ivantiDepartments.Except(antalPCDep).Concat(antalPCDep.Except(ivantiDepartments)).ToList();
                //List<string> differenceIvanti = ivantiDepartments.Where(x => !antalPCDep.Contains(x)).ToList();

                SetStatusLb($"Jämför ivanti och antal pc KST...");

                var diff = GetIvantiDiffDict(ivantiDepartments, antalPCDep);

                SetStatusLb($"Antal KST i ivanti som saknas i Antal PC: {diff.Count.ToString()}");
                SetStatusLb($"Skriver data till Antal PC...");

                ex.WriteDataToExcel(ivantiDepartments, diff, antalPCDep);

                SetStatusLb($"");

                SetStatusLb($"Klart!");
            }
            else
            {
                status_lb.Items.Clear();
                SetStatusLb($"Måste ange sökväg till filer...");
            }
        }

        public string BackupFile(string bakFile)
        {
            
            string folderPath = @"C:\Temp\AntalPCReport_Backup";
            string filePath = bakFile;

            SetStatusLb($"Backup fil {bakFile}");

            // Check if the folder exists, and create it if it doesn't
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                SetStatusLb($"Skapar mapp: {folderPath}");
            }

            // Copy the two files into the folder
            string fileNewPath = Path.Combine(folderPath, Path.GetFileName(filePath));
            File.Copy(filePath, fileNewPath, true);
            SetStatusLb($"Kopierar {fileNewPath}");
            return fileNewPath;

        }

        private Dictionary<string, int> GetIvantiDiffDict(Dictionary<string, int> invatiDict, Dictionary<string, int> pcDict)
        {
            Dictionary<string, int> diffDict = new Dictionary<string, int>();

            foreach(var kvp in invatiDict)
            {
                if(!pcDict.ContainsKey(kvp.Key))
                {
                    diffDict.Add(kvp.Key, kvp.Value);
                }
            }

            return diffDict;
        }
        private void SetStatusLb(string status)
        {
            Thread.Sleep(500);
            status_lb.Items.Add(status);
            
            status_lb.Refresh();
        }
        private List<string> CheckIvantiDepList(List<string> ivantiList, List<string> pcList)
        {
            List<string> ivantiDiff = ivantiList.Except(pcList).ToList();
            return ivantiDiff;
        }

        private List<string> CheckPCDepList(List<string> ivantiList, List<string> pcList)
        {
            List<string> PCDiff = pcList.Except(ivantiList).ToList();
            return PCDiff;
        }

        private void btn_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            textBox1.Text = openFileDialog.FileName;
        }

        private void btn_PC_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            textBox2.Text = openFileDialog.FileName;
        }
        private void RunPowershell(string psFileUrl, string excelFileUrl)
        {
            string powerShellExe = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";
            string scriptFile = @"C:\temp\ps.ps1";
            string param1 = textBox1.Text;
            string param2 = "value2";

            // Create a new process and configure it to run the PowerShell script with parameters
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = powerShellExe;
            psi.Arguments = $"-File \"{scriptFile}\" -ivantiFile \"{param1}\" -ivantiSheet \"{param2}\"";

            // Start the process and wait for it to exit
            Process p = Process.Start(psi);
            p.WaitForExit();
        }
    }
}
