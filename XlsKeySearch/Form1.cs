using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XlsKeySearch.ExcelHandler;

namespace XlsKeySearch
{
    public partial class Form1 : Form
    {
        public string ModuleFilePath;

        public Form1()
        {
            InitializeComponent();
            Util.InitLogInfo();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            if(fileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = fileDialog.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.textBox2.Text = path.SelectedPath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(this.textBox1.Text.Equals(string.Empty))
            {
                MessageBox.Show("Please select the module file path!");
                return;
            }

            if (this.textBox2.Text.Equals(string.Empty))
            {
                MessageBox.Show("Please select the output path!");
                return;
            }

            if (!Util.CheckStringContentToIntValid(this.textBox4.Text))
            {
                MessageBox.Show("Please input a number and which is must > 0!");
                return;
            }

            if (!Util.CheckStringContentToIntValid(this.textBox5.Text))
            {
                MessageBox.Show("Please input a number and which is must > 0!");
                return;
            }

            string fileExtension = Util.GetFileExtension(this.textBox1.Text);
            string fileName = Util.GetFileName(this.textBox1.Text);
            //this.label1.Text = fileExtension + " " + fileName + this.textBox2.Text;
            this.label1.Text = "";
            bool isRightExtension = Util.IsExcelExtension(fileExtension);
            if (isRightExtension)
            {
                Util.ExcelHandler(this.textBox1.Text, this.textBox2.Text, this.textBox4.Text, this.textBox5.Text);
                string mulKeyTitle = "\n重复的关键字如下：";
                string mulKeyTemp = "----------";
                string mulKeyStr = string.Empty;
                int multiKeysCount = Util.MultiKeys.Count;
                if (multiKeysCount > 0)
                {
                    Debug.Log("MultiKeyCount = {0}", multiKeysCount);
                    for (int i = 0; i < multiKeysCount; i++)
                    {
                        mulKeyStr = mulKeyStr + "\n" + Util.MultiKeys[i];
                    }
                    Debug.Log("{0}\n{1}\n{2}", mulKeyTitle, mulKeyTemp, mulKeyStr);
                    this.textBox3.Text = mulKeyTitle + "\n" + mulKeyTemp + "\n" + mulKeyStr;
                }
                MessageBox.Show("Generate success ! Enjoy it ! LJH 同学！");
            }
            else
            {
                MessageBox.Show("Please select file which extension is .xls or .xlsx!");
            }
        }
    }
}
