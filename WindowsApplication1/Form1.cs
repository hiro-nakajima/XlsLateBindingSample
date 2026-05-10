using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                "test.xls");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (ExcelLateBindingUtility el = new ExcelLateBindingUtility())
            {
                el.Visible = true;
                el.DisplayAlerts = false;
                bool fileExists = File.Exists(textBox1.Text);
                if (!fileExists)
                {
                    el.Add(string.Empty);
                }
                else
                {
                    el.Open(textBox1.Text);
                }

                el.PrepareMap(1, 1, 60, 60, 1, 9);
                el.SetCellValue(1, 1, "九九を作った後3秒休みます。");
                for (int i = 1; i < 10; i++)
                {
                    for (int j = 1; j < 10; j++)
                    {
                        el.SetCellValue(i + 1, j, Convert.ToString((i * j)));
                        el.SetColor(i + 1, j, i + 1, j, 45);
                    }
                }
                System.Threading.Thread.Sleep(3000);
                //object rr = el.GetRange(1, 1, 60, 60);
                
                el.SetCellValue(1, 1, "3秒後に閉じます。");

                //object rt = el.GetRange(1, 1, 60, 60);
                //ExcelLateBindingUtility.CopyRange(rr, rt, ExcelLateBindingUtility.XlPasteType.xlPasteAll);
                
                //下記意味無し
                //el.SelectCell(2, 2);
                //SendKeys.SendWait("{ENTER}");

                System.Threading.Thread.Sleep(3000);
                if (checkBox1.Checked)
                {
                    if (fileExists)
                    {
                        el.Save();
                    }
                    else
                    {
                        el.SaveAs(textBox1.Text);
                    }
                }
            }
        }
    }
}
