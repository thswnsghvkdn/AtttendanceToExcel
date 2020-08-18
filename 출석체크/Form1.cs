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
   
       

        public Form1()
        {
            InitializeComponent();

        }

        private bool Ready() // 출석 설정 체크
        {
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



            return true;
        }
        int index;

        private void button2_Click(object sender, EventArgs e)
        {

            
            if(Ready()) // 날짜와 예배 종류가 선택되면 true
            {
                DateTime dt = dateTimePicker1.Value;
                // int col = Convert.ToInt32(textBox4.Text);
                int week = dt.Day / 7; // 몇주차
                int col = 16 + ( week * 2 ) ; // 출석체크 할 열
                if (radioButton2.Checked == true) // 출석예배 구분 
                    col += 1;
                week += 1;

                try
                {
                    String filepath = "C:\\Users\\사용자\\Desktop\\대학부 재적정리 파일(교육국 양식)_" + dt.Year.ToString() + "_" + dt.Month.ToString() + "월.xlsx";

                    if (filepath != null)
                    {
                        ap = new Excel.Application(); // Excel 워크시트 가져오기 
                        wb = ap.Workbooks.Open(filepath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        int year = dt.Year % 2000;
                        string sheet = year.ToString() + "년 " + dt.Month.ToString() + "월";

                        ws = wb.Worksheets.get_Item(sheet) as Excel.Worksheet; // xx년 xx월 sheet접근


                        Excel.Range rg = ws.UsedRange; // 사용중인 엑셀 범위

                        // 출석명단의 이름을 토큰 분리 
                        char[] sep = { '\n', '\t', ' ' };
                        string[] t_name = textBox2.Text.Split(sep, StringSplitOptions.RemoveEmptyEntries);

                        int loading; // 로딩 %
                        int same;
                        string[] s_name = new string[10];

                        for (int j = 0; j < t_name.Length; j++) // j는 출석명단에서 가져온 이름번호
                        {
                            same = 0;
                            for (int i = 5; i < rg.Rows.Count; i++) // i는 이름이 있는 행번호
                            {
                                if (i > 150) break;
                                loading = 100 * i / rg.Rows.Count; // 100%로 표시하기 위해 설정
                                textBox3.Text = loading.ToString() + "% 진행 중.."; // 두번째 창에 로딩 진행상황 표시
                                if (ws.Cells[i, 5].Value.ToString() == "") continue; // 해당 행에 이름이 없으면 스킵


                                if (t_name[j].Length == 1) // 이름이 외자일 경우에는 다른 이름에 출석이 체크될 수 있으니 2글자인 이름에만 출석을 진행한다.
                                {
                                    if (ws.Cells[i, 5].Value.ToString().Contains(t_name[j]) && ws.Cells[i, 1].Value.ToString().Length == 2)
                                    {
                                        ws.Cells[i, col] = 1;
                                        s_name[same++] = t_name[j] + ' ' + ws.Cells[i, 2] ;
                                        break;
                                    }

                                }
                                else if (ws.Cells[i, 5].Value.ToString().Contains(t_name[j])) // 엑셀에 있는 이름에 출석명단이름이 포함되어 있으면 출석
                                {
                                    ws.Cells[i, col] = 1;
                                    s_name[same++] = t_name[j]+ ' ' + ws.Cells[i, 2];
                                    break;
                                }
                            }
                            if(same > 0)
                            {
                                message sameEvent = new message();
                                s_name[same] = t_name[j];
                                sameEvent.GetStr = s_name;
                                sameEvent.ChildEvent += getIndex;
                                sameEvent.ShowDialog();

                                for(int i = 0; i < s_name.Length - 1; i++)
                                {
                                    if (i == index)
                                    {
                                        ws.Cells[i, col] = 1;
                                    }
                                    else
                                        ws.Cells[i, col] = "";
                                }

                            }

                        }

                        textBox3.Text = "출석체크 완료!!!";

                        wb.SaveCopyAs(filepath); // 본 파일 저장
                        filepath = "C:\\Users\\사용자\\Desktop\\대학부 재적정리 파일(교육국 양식)_" + dt.Year.ToString() + "_" + dt.Month.ToString() + "월_" + week.ToString() +"주차.xlsx";
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

