using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Excel = Microsoft.Office.Interop.Excel; // 엑셀 lib 
using System.IO;
using System.Threading;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        
        Excel.Application ap = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;
        public int[] index_ar = new int[10]; // 동명이인 대비 동명이인 행 번호를 저장할 배열
        public int index; // 자식 폼에서 넘겨받을 동명이인 중 선택한 사람의 인덱스번호
        public string path = null;
        public struct Person
        {
            public string name, univ, region;
            public int row;
            public int check;
            public string state; // 장결 , 군대 구분
        }
        public Person[] stu;
        string filepath;
        DateTime dt;


        public Form1()
        {
            InitializeComponent();

        }

        private bool Ready() // 출석 설정 체크 출석 종류와 출석 주차를 확인한다.
        {
            DateTime dt = dateTimePicker1.Value;
            int week = dt.Day / 7; // 몇주차
            week += 1;
            int mon = dt.Month;


            if (radioButton1.Checked == false && radioButton2.Checked == false)
            {
                MessageBox.Show("출석 종류를 입력하세요(1,2부 , 대학부)");
                return false;
            }
            if (radioButton1.Checked == true && radioButton2.Checked == true)
            {
                MessageBox.Show("출석 타입이 두개가 선택되었습니다(1,2부 , 대학부)");
                return false;
            }
            if (MessageBox.Show(mon.ToString() + "월 " + week.ToString() + "주차 출석이 맞나요?", "출석주차 확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return false;


            return true;
        }

        private void load_name() // 엑셀파일에서 미리 이름 구조체에 저장하여 엑셀파일 접근을 최소화
        {
            System.IO.StreamReader file = new System.IO.StreamReader("path.txt");

            while((path = file.ReadLine()) == null) 
            {

            }
            file.Close();

            textBox3.Text ="엑셀파일을 불러옵니다.";
            dt = dateTimePicker1.Value;
            // int col = Convert.ToInt32(textBox4.Text);
            int week = dt.Day / 7; // 몇주차
            int col = 16 + (week * 2); // 출석체크 할 열
            int loading;
            try
            {
                // 파일 이름은 각 월을 기준으로 설정한다.
                filepath = "C:\\Users\\사용자\\Desktop\\대학부 재적정리 파일(교육국 양식)_" + dt.Year.ToString() + "_" + dt.Month.ToString() + "월.xlsx";

                if (filepath != null)
                {
                    ap = new Excel.Application(); // Excel 워크시트 가져오기 
                    wb = ap.Workbooks.Open(filepath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    int year = dt.Year % 2000;
                    string sheet = year.ToString() + "년 " + dt.Month.ToString() + "월";

                    ws = wb.Worksheets.get_Item(sheet) as Excel.Worksheet; // xx년 xx월 sheet접근
                    Excel.Range rg = ws.UsedRange; // 사용중인 엑셀 범위

                 

                    string univ; // 대학
                    string region; // 사는 지역
                    stu = new Person[rg.Rows.Count]; // 저장할 출석명단 구조체

                    for(int i = 5; i < rg.Rows.Count; i++)
                    {
                        if (ws.Cells[i, 5].Value2 == null)
                        {
                            stu[i].row = -1; // 마지막 행 구분
                            break;
                        }
                        stu[i].check = 0; // 출석체크 여부
                        stu[i].name = ws.Cells[i, 5].Value2.ToString();
                        stu[i].row = i;
                        stu[i].univ = null;
                        stu[i].region = null;
                        if (ws.Cells[i, 6].Value2 != null) stu[i].univ = ws.Cells[i, 6].Value2.ToString(); // 대학이름 혹은 지역이 기재 안되어 있으면 스킵
                        if (ws.Cells[i, 2].Value2 != null) stu[i].region = ws.Cells[i, 2].Value2.ToString();
                        loading = 100 * i / rg.Rows.Count;
                        textBox3.Text ="명단 로딩 " +loading.ToString() +"%..";

                    }
               


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("에러" + ex.Message, "에러!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            
            if(Ready()) // 날짜와 예배 종류가 선택되면 true
            {
                load_name();
                textBox3.Text = "준비중..";
                dt = dateTimePicker1.Value;
                // int col = Convert.ToInt32(textBox4.Text);
                int week = dt.Day / 7; // 몇주차
                int col = 16 + ( week * 2 ) ; // 출석체크 할 열
                if (radioButton2.Checked == true) // 출석예배 구분 
                    col += 1;
                week += 1;

                bool already_at;
                try
                {
                    
                   
                    if (filepath != null)
                    {                    
                        int year = dt.Year % 2000;
                        string sheet = year.ToString() + "년 " + dt.Month.ToString() + "월";

                        // 출석명단의 이름을 토큰 분리 
                        char[] sep = { '\n', '\t', ' ', '\r' };
                        string[] t_name = textBox2.Text.Split(sep, StringSplitOptions.RemoveEmptyEntries);

                        int loading; // 로딩 게이지
                        int same; // 동명이인 수 

                        string[] s_name = new string[10]; // 동명이인이 저장 될 배열
                        //int[] index_ar = new int[10];
                        string univ; // 대학
                        string region; // 사는 지역
                        for (int j = 0; j < t_name.Length; j++) // j는 출석명단에서 가져온 이름번호
                        {
                            univ = "미기입";
                            region = "미기입";
                            already_at = false;
                            s_name[0] = null;
                            if (t_name[j].Length > 3) // 동명이인 설명이 있는 경우
                            {
                                string temp = null;
                                s_name[0] = t_name[j]; // 설명부분을 s_name[0]인덱스에 저장하여 message폼에서 사용자에게 보여준다                           
                                for (int i = 0; i < t_name[j].Length; i++) // 설명부분을 제거 한다
                                {
                                    if (t_name[j][i] == '(' || t_name[j][i] == ':') 
                                        break;
                                    temp += t_name[j][i];
                                }
                                t_name[j] = temp;  // 이름만 저장 
                            }
                            same = 1;
                            if (t_name[j] == "") continue;
                            loading = 100 * j / t_name.Length; // 100%로 표시하기 위해 설정
                            textBox3.Text = loading.ToString() + "% 진행 중.."; // 두번째 창에 로딩 진행상황 표시

                            for (int i = 5; i < stu.Length; i++) // i는 이름이 있는 행번호
                            {

                               
                                if (stu[i].row == -1)  break; // 해당 행에 이름이 없으면 스킵

                                
                                if (t_name[j].Length == 1) // 이름이 외자일 경우에는 다른 이름에 출석이 체크될 수 있으니 2글자인 이름에만 출석을 진행한다.
                                {
                                    if (stu[i].name.Contains(t_name[j]) && stu[i].name.Length == 2)
                                    {
                                        if (stu[i].check == 1)
                                        {
                                            already_at = true;
                                            continue;
                                        } // 해당이름이 이미 체크되어 있으면 넘어간다. 

                                        if(ws.Cells[i,col].Value2 != null)
                                        stu[i].state = ws.Cells[i, col].Value2.ToString(); // 장결 군입대 같은 상태를 저장해놓습니다.
                                        ws.Cells[i, col] = 1;
                                        index_ar[same] = i; // 동명이인을 대비해 인덱스를 저장해놓습니다.
                                        if (stu[i].univ != null) univ = stu[i].univ; // 대학이름 혹은 지역이 기재 안되어 있으면 스킵
                                        if (stu[i].region != null) region = stu[i].region;
                                        stu[i].check = 1;

                                        s_name[same++] = univ + ' ' + stu[i].name + ' ' +region;

                                        //break;
                                    }

                                }

                                else if (stu[i].name.Contains(t_name[j])) // 엑셀에 있는 이름에 출석명단이름이 포함되어 있으면 출석
                                {
                                    if (stu[i].check == 1)
                                    {
                                        already_at = true;
                                        continue;
                                    } // 해당이름이 이미 체크되어 있으면 넘어간다
                                    
                                    if (ws.Cells[i, col].Value2 != null)
                                        stu[i].state = ws.Cells[i, col].Value2.ToString(); // 장결 군입대 같은 상태를 저장해놓습니다.

                                    ws.Cells[i, col] = 1;
                                    index_ar[same] = i;
                                    if (stu[i].univ != null) univ = stu[i].univ; // 대학이름 혹은 지역이 기재 안되어 있으면 스킵
                                    if (stu[i].region != null) region = stu[i].region;

                                    stu[i].check = 1;
                                    s_name[same++] = univ + ' ' + stu[i].name + ' ' + region;

                                    // break;
                                }
                            }
                            if (same > 2) // 동명이인이 두명이상 일 때
                            {
                                string[] name = new string[same]; // 실제로 할당 된 이름만 복사
                                if (s_name[0] == null)
                                    s_name[0] = t_name[j]; // 동명이인 이름을 마지막 인덱스에 저장

                                for (int i = 0; i < same; i++)
                                    name[i] = s_name[i];

                                message sameEvent = new message();
                                sameEvent.GetStr = name; // 폼간 데이터를 전달하는 방법 1
                                //sameEvent.ChildEvent += getIndex; // 방법 2
                                sameEvent.fm = this; // 방법 3
                                sameEvent.ShowDialog();


                                for (int i = 1; i < name.Length; i++) // 자식 폼에서 체크된 인덱스를 가져와 해당 이름을 제외한 나머지 이름은 출석 미처리
                                {
                                    if (i == index)
                                        ws.Cells[index_ar[i], col] = 1;

                                    else
                                    {
                                        ws.Cells[index_ar[i], col] = stu[index_ar[i]].state; // 잘못체크된 동명이인셀에 이전 상태를 저장해 놓습니다.
                                        stu[index_ar[i]].check = 0;
                                    }
                                }


                            }
                            else if (same == 1 && already_at == false) // 한번도 출석처리 되지 않은 이름
                            {
                                MessageBox.Show(t_name[j] + " 출석처리 되지 않았습니다. \n이름을 고쳐 다시 출석체크를 눌러주세요(띄어쓰기, 오타, 미기입)");
                            }

                        }

                        textBox3.Text = "출석체크 완료";

                        ap.DisplayAlerts = false;
                       // wb.Save();
                        wb.SaveCopyAs(filepath); // 본 파일 저장
                        filepath = "C:\\Users\\사용자\\Desktop\\대학부 재적정리 파일(교육국 양식)_" + dt.Year.ToString() + "_" + dt.Month.ToString() + "월_" + week.ToString() + "주차.xlsx";
                        wb.SaveCopyAs(filepath); // 백업파일로 저장

                        /*메모리 할당 해제*/
                        DeleteObject(ws);
                        DeleteObject(wb);
                        ap.Quit();
                        DeleteObject(ap);
                        /*메모리 할당 해제*/
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("에러" + ex.Message, "에러!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        
        public void getIndex(int num)
        {
            index = num;
        }
        private void DeleteObject(object obj)
        {   // 메모리 해제를 위한 사용자 정의 함수
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
            }
            finally
            {
                GC.Collect();
            }
        }


    }

}

