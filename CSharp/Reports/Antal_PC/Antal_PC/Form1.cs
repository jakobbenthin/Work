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
            ex.ivantiFile = textBox1.Text;
            ex.ivantiSheet = "Grunddata";
            var ivantiDepartments = ex.CheckDepartmentsInIvantiFile();

            ex.antalPCFile = textBox2.Text;
            ex.antalPCSheet = "Innevarnde Månad";
            var antalPCDep = ex.CheckDepartmentsInPCFile();

            //List<string> ivantiDiff = CheckIvantiDepList(ivantiDepartments, antalPCDep);
            //List<string> pcDiff = CheckPCDepList(ivantiDepartments, antalPCDep);
            
            //List<string> difference = ivantiDepartments.Except(antalPCDep).Concat(antalPCDep.Except(ivantiDepartments)).ToList();

            //List<string> differenceIvanti = ivantiDepartments.Where(x => !antalPCDep.Contains(x)).ToList();
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
