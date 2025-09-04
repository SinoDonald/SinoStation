using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SinoStation_2020
{
    public partial class Sino_Station_Form : Form
    {
        public bool trueOrFalse = false; // 確定或取消
        public double allowSpace = 0.0; // 允許間距
        public Sino_Station_Form()
        {
            InitializeComponent();
            CenterToParent(); // 置中
        }
        // 輸入Enter執行確定、輸入Esc執行取消
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                sureBtn_Click(sender, e);
            }
            if(e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
        // 確定
        private void sureBtn_Click(object sender, EventArgs e)
        {
            trueOrFalse = true;
            allowSpace = Convert.ToDouble(textBox1.Text);
            Close();
        }
        // 取消
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            trueOrFalse = false;
            Close();
        } 
        // 限制TextBox 只能輸入數字，以及限制不能使用快速鍵
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Enter)
            {
                return;
            }
            if (e.KeyChar == '.')
            {
                //判定textBox1是否有小數點
                foreach (char i in textBox1.Text)
                {
                    if (i == '.')
                    {
                        e.Handled = true;
                    }
                }
                return;
            }

            if (e.KeyChar < '0' || e.KeyChar > '9')
            {
                e.Handled = true;
            }
        }
    }
}
