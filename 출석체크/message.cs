using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class message : Form
    {
        private string[] s_name;
        public string[] GetStr
        {
            get
            {
                return this.s_name;
            }
            set
            {
                this.s_name = value;
            }
        }

        public message()
        {
            InitializeComponent();
        }
        public delegate void ChildFrom(int index);
        public event ChildFrom ChildEvent;

        private void message_Load(object sender, EventArgs e)
        {
            label.Text = s_name[s_name.Length - 1] + "가 많습니다 어느 지역인가요?";
            //  for (int i = 0; i < s_name.Length - 1; i++)
            //      comboBox1.Items.Add(s_name[i]);
            comboBox1.Items.AddRange(s_name);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.ChildEvent(comboBox1.SelectedIndex);
            this.Close();
        }
    }
}
