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
    public delegate void ChildFrom(int index); // 폼간 데이터를 전달받는 방법 2
    public partial class message : Form
    {
        public Form1 fm; // 폼간 데이터를 전달받는 방법 3
        private string[] s_name;
        public string[] GetStr // 폼간 데이터를 전달받는 방법 1
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
       
        public ChildFrom ChildEvent;

        private void message_Load(object sender, EventArgs e) // 동명이인 보여줌
        {
            label.Text =  "동명이인 중 출석한 사람을 클릭 해주세요 : " + s_name[0];
              for (int i = 1; i < s_name.Length ; i++) // 0번인덱스는 동명이인의 이름 + 설명이 저장되어 있고 1번인덱스 부터 동명이인들이 저장되어 있음
                  comboBox1.Items.Add(s_name[i]);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            fm.index = comboBox1.SelectedIndex+1; // 부모 폼에 index를 체크된 인덱스로 설정 부모 폼에 인덱스는 1부터 시작하므로 +1
            this.Close();
           // ChildEvent(comboBox1.SelectedIndex);
           // this.Close();
        }
    }
}
