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

namespace Antal_PC
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
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
        private void btn_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            textBox1.Text = openFileDialog.FileName;
        }
    }
}
