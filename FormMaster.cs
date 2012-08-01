using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;


namespace KIWI
{
    public partial class FormMaster : Form
    {
        public FormMaster()
        {
            InitializeComponent();
            FormUserInput frm = new FormUserInput();
            panelSet(frm);
        }

        /// <summary>
        /// 파일열기
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (this.panel1.Controls.Count > 0)
            {
                if (this.panel1.Controls[0] is Form)
                {
                    if ((this.panel1.Controls[0] as Form).Name == "FormAdmin")
                    {
                        (this.panel1.Controls[0] as FormAdmin).openFileDialog(sender, e);
                        return;
                    }
                }
            }
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel File|*.xlsx";
            openFileDialog1.Title = "Select a Excel File";
            openFileDialog1.RestoreDirectory = true;

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    // Create a PictureBox.
                    try
                    {
                        CommonUtil.ReadExcelFileToData(file);

                        // CDataControl의 파일정보변수(g_File*)에 담겨있는 데이터를 일반정보변수에 딥카피
                        // 엑셀내용중 시트 1의 내용만 옮겨짐
                        CommonUtil.deepCopyBasicInput(CDataControl.g_FileBasicInput, CDataControl.g_BasicInput);
                        CommonUtil.deepCopyBusinessData(CDataControl.g_FileDetailInput, CDataControl.g_DetailInput);
                    }
                    catch (Exception ex)
                    {
                        // Could not load the image - probably related to Windows file system permissions.
                        MessageBox.Show("파일을 열 수 없습니다.\n\nReported error: " + ex.Message);
                    }
                }
                FormUserInput frm = new FormUserInput();
                panelSet(frm);
            }

        }

        /// <summary>
        /// 입력
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            FormUserInput frm = new FormUserInput();
            panelSet(frm);
        }

        /// <summary>
        /// 결과
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            FormUserOutput frm = new FormUserOutput();
            panelSet(frm);
        }

        /// <summary>
        /// 분석
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            FormUserAnalysis frm = new FormUserAnalysis();
            panelSet(frm);
        }

        /// <summary>
        /// 시뮬레이트입력
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            FormUserSimulateOutput frm = new FormUserSimulateOutput();
            panelSet(frm);
        }

        /// <summary>
        /// 업데이트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            // 업데이트
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Update File|manager.csv";

            openFileDialog1.Title = "업데이트 파일을 선택하세요.";
            openFileDialog1.RestoreDirectory = true;

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    // Create a PictureBox.
                    try
                    {
                        String path = Application.StartupPath;
                        System.IO.File.Copy(file, path+"\\"+"manager.csv", true);
                    }
                    catch (Exception ex)
                    {
                        // Could not load the image - probably related to Windows file system permissions.
                        MessageBox.Show("파일을 열 수 없습니다.\n\nReported error: " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            //if (CommonUtil.openAsName == null)
            //{
            //    MessageBox.Show("먼저 파일을 열어주세요.");
            //    return;
            //}
            if (this.panel1.Controls.Count > 0)
            {
                if (this.panel1.Controls[0] is Form)
                {
                    if ((this.panel1.Controls[0] as Form).Name == "FormAdmin")
                    {

                    }
                    else if ((this.panel1.Controls[0] as Form).Name == "FormReport")
                    {

                    }
                    else if ((this.panel1.Controls[0] as Form).Name == "FormUserAnalysis")
                    {
                        //파일경로명
                        (this.panel1.Controls[0] as FormUserAnalysis).saveComments();
                    }
                    else if ((this.panel1.Controls[0] as Form).Name == "FormUserInput")
                    {
                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                        saveFileDialog1.Filter = "Excel File|*.xlsx";
                        saveFileDialog1.Title = "Select a Excel File";
                        saveFileDialog1.ShowDialog();

                        // If the file name is not an empty string open it for saving.
                        if (saveFileDialog1.FileName != "")
                        {
                            string filename = CommonUtil.defaultName;


                            FileInfo fi2 = new FileInfo(filename);
                            fi2.CopyTo(saveFileDialog1.FileName, true);

                            CommonUtil.saveAsName = saveFileDialog1.FileName;

                            //excel.Workbook _Workbook = CommonUtil.GetExcel_WorkBook(saveFileDialog1.FileName);
                            //excel.Worksheet _WorkSheet1 = _Workbook.Sheets[1] as excel.Worksheet;
                            //excel.Worksheet _WorkSheet2 = _Workbook.Sheets[2] as excel.Worksheet;
                            (this.panel1.Controls[0] as FormUserInput).saveAsInput();
                            CommonUtil.WriteDataToExcelFile(CommonUtil.saveAsName, CDataControl.g_BasicInput, CDataControl.g_DetailInput);
                        }
                    }
                    else if ((this.panel1.Controls[0] as Form).Name == "FormUserOutput")
                    {

                    }
                    else if ((this.panel1.Controls[0] as Form).Name == "FormUserSimulateInput")
                    {

                    }
                }
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            new Printer();
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            FormAdmin frm = new FormAdmin();
            panelSet(frm);
        }

        private void outOfFormUserInput_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.panel1.Controls.Count; i++)
            {
                if (this.panel1.Controls[i] is Form)
                {
                    if (this.panel1.Controls[i].Name == "FormUserInput")
                    {
                        (this.panel1.Controls[i] as FormUserInput).saveAsInput();
                    }
                }
            }
        }

        private void panelSet(Form form)
        {
            form.TopLevel = false;
            if (panel1.Controls.Count > 0)
            {
                if (panel1.Controls[0] is Form)
                {
                    if ((panel1.Controls[0] as Form).Name == "FormUserInput" ||
                        (panel1.Controls[0] as Form).Name == "FormUserAnalysis" ||
                        (panel1.Controls[0] as Form).Name == "FormUserSimulateInput")
                    {


                    }
                    (panel1.Controls[0] as Form).Close();
                }
            }

            panel1.Controls.Add(form);
            form.Show();
        }

        private void FormMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            CommonUtil.GetExcel_WorkBook_CLOSE();
        }



    }
}
