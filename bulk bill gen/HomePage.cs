using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ganss.Excel;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Util;
using NPOI.Util;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static NPOI.HSSF.Record.UnicodeString;

namespace bulk_bill_gen
{
    public partial class HomePage : Form
    {
        public string filepath1;
        
        string filepath = @"C:\Users\pc\Desktop\Book22.xlsx";
        public HomePage()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            var bills = new ExcelMapper(filepath1).Fetch<bill>();

            if (!File.Exists("C:\\Users\\pc\\source\\repos\\bulk bill gen\\bulk bill gen" + "\\invoice.html"))
            {
                File.Delete("C:\\Users\\pc\\source\\repos\\bulk bill gen\\bulk bill gen" + "\\invoice.html");
            }
            FileStream fileStream = File.Create("C:\\Users\\pc\\source\\repos\\bulk bill gen\\bulk bill gen" + "\\invoice.html");
            string text = ("Original ");
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            {
                streamWriter.WriteLine("<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset='utf - 8' />\r\n    <style>\r\n    table,\r\n        tr,\r\n        td {\r\n            border: 2px solid black;\r\n            border-collapse: collapse;\r\n        }\r\n    }\r\n    tr.noBorder td {\r\n        border: 0;\r\n    }\r\n    tr.hide_right>td,\r\n    td.hide_right {\r\n        border-right-style: hidden;\r\n    }\r\n    tr.hide_3side>td,\r\n    td.hide_3side {\r\n        border-top-style: hidden;\r\n        border-bottom-style: hidden;\r\n        border-right-style: hidden;\r\n    }\r\n    tr.hide_ver>td,\r\n    td.hide_ver {\r\n        border-top-style: hidden;\r\n        border-bottom-style: hidden;\r\n    }\r\n    tr.hide_right_bot>td,\r\n    td.hide_right_bot {\r\n        border-right-style: hidden;\r\n        border-bottom-style: hidden;\r\n    }\r\n    tr.hide_left_bot>td,\r\n    td.hide_left_bot {\r\n        border-left-style: hidden;\r\n        border-bottom-style: hidden;\r\n    }\r\n    tr.hide_right_top>td,\r\n    td.hide_right_top {\r\n        border-right-style: hidden;\r\n        border-top-style: hidden;\r\n    }\r\n    tr.hide_left_top>td,\r\n    td.hide_left_top {\r\n        border-left-style: hidden;\r\n        border-top-style: hidden;\r\n    }\r\n    tr.hide_left>td,\r\n    td.hide_left {\r\n        border-left-style: hidden;\r\n    }\r\n    tr.hide_all>td,\r\n    td.hide_all {\r\n        border-style: hidden;\r\n    }\r\n    tr.hide_top>td,\r\n    td.hide_top {\r\n        border-top-style: hidden;\r\n    }\r\n    tr.hide_bot>td,\r\n    td.hide_bot {\r\n        border-bottom-style: hidden;\r\n    }\r\n    body {\r\n        margin: 0;\r\n        padding: 0;\r\n        background-color: white;\r\n    }\r\n    .page {\r\n        width: 29.7cm;\r\n        min-height: 21cm;\r\n        padding: 0.1cm;\r\n        margin: 1cm auto;\r\n        border: 0px #000000 solid;\r\n        border-radius: 0px;\r\n        background: white;\r\n    }\r\n    .subpage {\r\n        padding: 0.1cm;\r\n        border: 0px #000000 solid;\r\n        outline: 0cm white solid;\r\n    }\r\n    @page {\r\n        size: landscape;\r\n        margin: 0.2cm;\r\n    }\r\n    @media print {\r\n        .page {\r\n            margin: 0cm;\r\n            border: initial;\r\n            border-radius: initial;\r\n            width: initial;\r\n            min-height: initial;\r\n            box-shadow: initial;\r\n            background: initial;\r\n            page-break-after: always;\r\n        }\r\n    }\r\n    .invoice-box {\r\n        max-width: 100%;\r\n        margin: ;\r\n        padding: 0px;\r\n        border: px #000000 solid;\r\n        font-family: arial;\r\n        color: black;\r\n    }\r\n    .invoice-box table {\r\n        width: 100%;\r\n        line-height: inherit;\r\n        ;\r\n        text-align: left;\r\n        word-break: break-word;\r\n        white-space: normal;\r\n    }\r\n    .invoice-box table td {\r\n        padding: 2px;\r\n        vertical-align: center;\r\n    }\r\n    .invoice-box table tr td:nth-child(11) {}\r\n    .invoice-box table tr.top table td {\r\n        padding-bottom: 10px;\r\n    }\r\n    @media only screen and (max-width: 600px) {}\r\n    </style>\r\n</head>\r\n\r\n<body style=\"font-family:\">\r\n");
                 foreach (var item in bills)
                {

                    streamWriter.WriteLine("<div class='book'> <div class='page'> <div class='subpage'> <div class='invoice-box'> <div style='margin: 10px;'>\r\n\r\n <div> <table> <tr> <td class=\"hide_right_bot\"></td> <td class=\"hide_bot\"></td> </tr> <tr>\r\n\r\n <td class=\"hide_right_bot\" style=\"font-size: 34px; vertical-align: bottom\"><b>&thinsp;&thinsp;JAGDAMBA STONES</b> </td>\r\n\r\n <td class=\"hide_bot\" style=\"font-size: 16px; text-align: right; vertical-align: bottom\"><b>Phone - 6377693526&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp; </b></td> </tr> </table> <table>\r\n\r\n <tr>\r\n\r\n <td class=\"hide_right_top\" style=\"font-size: 15px; height: 55px;vertical-align: top \">&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;Pathano ka Mohalla, House Number-2,<br>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;Main ROAD, VILLAGE- SALEMPUR, TEHSIL- SAPOTRA,<br> &thinsp;&thinsp;&thinsp;&thinsp;&thinsp;DISTRICT-KARAULI, Karauli, Rajasthan, 322202</b> </td> <td class=\"hide_top\" style=\"font-size: 16px; text-align: right; vertical-align: top\"><b>Email - yashshekhawat2511@gmail.com&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp; </b></td>\r\n\r\n </tr> <tr> <td class=\"hide_3side\"></td> <td class=\"hide_ver\"></td> </tr> <tr> <td class=\"hide_3side\"></td> <td class=\"hide_ver\"></td> </tr> \r\n\r\n </table> ");
                    streamWriter.WriteLine("<table> <tr> <td class=\"hide_right_bot\"><b>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;GSTIN : 08AAJFJ1795Q1ZH</b></td> <td class=\"hide_right_bot\" style=\"font-size: 29px; height: 40px;color: blue\"><b>&emsp;&emsp;&emsp;TAX INVOICE</b></td> <td class=\"hide_bot\" style=\"font-size: 15px; text-align: right\"><b>ORIGINAL FOR RECIPIENT&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;</b></td> </tr> <tr> <td class=\"hide_3side\"></td> <td class=\"hide_ver\"></td> <td class=\"hide_left_bot\"></td> </tr> </table> <table> <tr> <td class=\"hide_bot\" style=\"height: 32px;font-size: 15px;text-align: center\"><b>Customer Details</b></td> <td class=\"hide_bot\" style=\"width:40%;font-size: 15px\"><b> &thinsp;&thinsp;Invoice No.</b>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&emsp;"+item.Invoice+ " </td> </tr> </table> <table> <tr> <td style=\"width: 1.5%\" class=\"hide_right\"></td> <td style=\"height: 31px ;font-size: 15px\"><b>M/S</b> &thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&emsp;" + item.Client+ " </td><td style=\"width:40% ;font-size: 15px\"><b> &thinsp;&thinsp;Invoice Date</b>&thinsp;&emsp;&emsp;"+item.Date+"</td>");
                    streamWriter.WriteLine("</tr> <tr> <td class=\"hide_right\"></td> <td style=\"height: 31px ;font-size: 15px\"><b>Address</b>&thinsp;&thinsp;&emsp;"+item.Address+ "</td> <td style=\"height: 31px ;font-size: 15px\"><b>&thinsp;&thinsp;Ravanna No.</b>&thinsp;&emsp;&thinsp;&thinsp;&thinsp; "+item.Ravanna+" </td>");
                    streamWriter.WriteLine("</tr> <tr> <td class=\"hide_right_bot\"></td> <td class=\"hide_bot\" style=\"height: 31px;font-size: 15px\"><b>GSTIN</b>&thinsp;&emsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;"+item.GSTIN+ "</td><td class=\"hide_bot\" style=\"height: 31px;font-size: 15px\"><b>&thinsp;&thinsp;Vehicle No.</b>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&emsp; "+item.Vehicle+ " </td></tr> </table>");
                    streamWriter.WriteLine("<table> <tr> <td style=\"text-align: center;font-size: 15px\" rowspan=\"2\"> <b>Name of product</b> </td> <td style=\"width:10%;text-align: center;font-size: 15px\" rowspan=\"2\"><b>Qty</b></td> <td style=\"width:7%;text-align: center;font-size: 15px\" rowspan=\"2\"><b> Rate <br> (₹)</b> </td> <td style=\"width:11%;text-align: center;font-size: 15px\" rowspan=\"2\"> <b>Taxable Value <br> (₹)</b> </td> <td style=\"height:26px;text-align: center;font-size: 15px\" colspan=\"2\"> <b>CGST</b> </td> <td style=\"text-align: center;font-size: 15px\" colspan=\"2\"> <b>SGST</b> </td> <td style=\"width:25%;text-align: center;font-size:18px\" rowspan=\"2\"><b>Total <br> (₹)</b></td> </tr> <tr> \r\n\r\n <td style=\"height:26px; width:4%;;font-size: 15px;text-align: center\"><b>%</b></td> <td style=\"width:10%;text-align: center;font-size: 15px\"> <b>Amount</b> </td> <td style=\"width:5%;text-align: center;font-size: 15px\"> <b>%</b> </td> <td style=\"width:10%;text-align: center;font-size: 15px\"><b>Amount</b></td>\r\n\r\n </tr> <tr> <td class=\"hide_bot\" style=\"height: 32px; text-align: center;font-size: 15px\"> <b>Masonary Stone</b> </td> <td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\">"+item.Ton+ "</td><td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\"><b>140</b></td> <td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\">"+item.NetPrice+"</td> <td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\"><b>2.5</b></td> <td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\">"+item.CGST+ "</td><td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\"><b>2.5</b></td> <td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\">"+item.SGST+ "</td><td class=\"hide_bot\" style=\"text-align: center;font-size: 15px\"><b>"+item.Total+"</b></td>\r\n\r\n </tr> </table>");
                    streamWriter.WriteLine("<table> <tr> <td style=\"height: 34px; text-align: center\"><b>Terms and Conditions</b></td> <td class=\"hide_right\" style=\"width:19%;font-size:17px\"><b>&thinsp;&thinsp;Total taxable amount</b></td> <td style=\"width:31%;font-size:17px\"><b>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;₹" + item.NetPrice+ "</b></td> </tr> <tr> <td style=\"vertical-align: top;font-size: 16px\" rowspan=\"6\">&thinsp;&thinsp;&thinsp;&thinsp;Subject to Gangapur City jurisdiction.<br> &thinsp;&thinsp;&thinsp;&thinsp;Our Responsibility ceases as soon as goods leaves our premises.<br>&thinsp;&thinsp;&thinsp;&thinsp;Goods once sold will not be taken back.</td> <td class=\"hide_right\" style=\"height: 33px\"><b>&thinsp;&thinsp;Total Tax </b></td> <td style=\"font-size: 17px\"><b>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;₹" + item.Tax+ "</b></td> </tr> <tr><td class=\"hide_right\" style=\"height:33px\"><b>&thinsp;&thinsp;Total amount after Tax</b></td> <td style=\"font-size: 17px\"><b>&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;&thinsp;₹" + item.Total+ "</b></td></tr> <tr> <td class=\"hide_bot\" style=\"text-align: center;height: 32px\" colspan=\"4\"><b>For Jagdamba Stone</b></td> </tr> <tr> <td class=\"hide_bot\" colspan=\"4\"><br></td> </tr> <tr> <td class=\"hide_bot\" colspan=\"4\"><br></td> </tr>\r\n\r\n <tr> <td class=\"hide_top\" style=\"text-align: center; font-size: 14px;height: 32px\" rowspan=\"2\" colspan=\"4\">Authorised Signatory</td> </tr>\r\n\r\n </table> </div> </div> </div> </div> </div> </div> \r\n");
                    
                }
                streamWriter.WriteLine("</body>\r\n\r\n</html>");
            }
            fileStream.Close();
            Billviewer billviewer =new Billviewer();
            billviewer.Show();
            

        }
        public class bill
        {
            public string Invoice{ get; set; }
            public string Ravanna { get; set; }
            public string Vehicle { get; set; }
            public string Ton { get; set; }
            public string Date { get; set; }
            public string Client { get; set; }
            public string Address { get; set; }
            public string GSTIN { get; set; }
            public string NetPrice { get; set; }
            public string CGST { get; set; }
            public string SGST { get; set; }
            public string Tax { get; set; }
            public string Total { get; set; }
        }
        string file = @"C:\Users\pc\source\repos\bulk bill gen\bulk bill gen\kcinvoice.html";
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            
            

        }

        private void filename_TextChanged(object sender, EventArgs e)
        {
            if (!filename.Text.Equals(""))
            {
                filename.Text = filepath1.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult dialogResult= openFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                string name = openFileDialog.FileName;
                try
                {
                    this.filepath1 = System.IO.Path.GetFullPath(name);
                }
                catch(System.Exception) { }
            }
            String sheetname = "Sheet1";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            filepath1 +
                            ";Extended Properties='Excel 8.0;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + sheetname + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
