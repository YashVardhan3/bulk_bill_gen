using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using CefSharp;
using CefSharp.WinForms;
using System.Windows.Forms;
using static bulk_bill_gen.HomePage;

namespace bulk_bill_gen
{
    public partial class Billviewer : Form
    {
        public Billviewer()
        {
            InitializeComponent();
            InitBrowser();
        }

        public ChromiumWebBrowser browser;

        public void InitBrowser()
        {
            var settings = new CefSettings();
          
            settings.EnablePrintPreview();
            
            Cef.Initialize(settings);
            browser = new ChromiumWebBrowser("C:\\Users\\pc\\source\\repos\\bulk bill gen\\bulk bill gen\\invoice.html");
            this.Controls.Add(browser);
            browser.Dock = DockStyle.Fill;
        }


        private void Billviewer_Load(object sender, EventArgs e)
        {

            
    
        }

        private void button1_Click(object sender, EventArgs e)
        {

            browser.Print();
        }
    }
}
