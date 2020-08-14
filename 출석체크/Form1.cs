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


        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                String filepath = "C:\\Users\\사용자\\Desktop\\aa.xlsx";
                if (filepath != null)
                {
                    ap = new Excel.Application(); // Excel 워크시트 가져오기 
                    wb = ap.Workbooks.Open(filepath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    ws = wb.Worksheets.get_Item("sheet1") as Excel.Worksheet; // 1번째 워크시트 
              

                    Excel.Range rg = ws.UsedRange; // 사용중인 엑셀 범위
                    
                    // 출석명단의 이름을 토큰 분리 
                    char[] sep = { '\n', '\t', ' ' };
                    string[] t_name = textBox2.Text.Split(sep, StringSplitOptions.RemoveEmptyEntries);

                    int loading; // 로딩 %
                    for (int i = 2; i < rg.Rows.Count; i++) // i는 이름이 있는 행번호
                    {
                        loading = 100 * i / rg.Rows.Count; // 100%로 표시하기 위해 설정
                        textBox3.Text = loading.ToString() + "% 진행 중.."; // 두번째 창에 로딩 진행상황 표시

                        if (ws.Cells[i] == null) continue; // 해당 행에 이름이 없으면 스킵
                    
                        for (int j = 0; j < t_name.Length; j++) // j는 출석명단에서 가져온 이름번호
                        {
                            if(t_name[j].Length == 1) // 이름이 외자일 경우에는 다른 이름에 출석이 체크될 수 있으니 2글자인 이름에만 출석을 진행한다.
                            {
                                if(ws.Cells[i, 1].Value.ToString().Contains(t_name[j]) && ws.Cells[i, 1].Value.ToString().Length == 2)
                                {
                                    ws.Cells[i, 3] = 1;
                                    break;
                                }

                            }
                            else if (ws.Cells[i,1].Value.ToString().Contains(t_name[j])) // 엑셀에 있는 이름에 출석명단이름이 포함되어 있으면 출석
                            {
                                ws.Cells[i , 3] = 1;
                                break;
                            }
                        }
                    }
                    textBox3.Text = "출석체크 완료!";

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

