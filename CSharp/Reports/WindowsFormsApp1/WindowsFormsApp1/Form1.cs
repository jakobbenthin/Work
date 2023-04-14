using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation;
using System.IO;
using System.Collections.ObjectModel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        public void Test()
        {

            // create a PowerShell instance
            PowerShell ps = PowerShell.Create();

            // specify the PowerShell script to run
            ps.AddScript("path/to/your/script.ps1");

            // add parameters to the script
            ps.AddParameter("param1", "value1");
            ps.AddParameter("param2", "value2");

            // execute the script
            ps.Invoke<Collection<PSObject>>();

        }
    }
}
