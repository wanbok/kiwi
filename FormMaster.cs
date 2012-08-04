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
            openFileDialog1.Filter = "LGE File|*.lge|Excel File|*.xlsx|All File|*.*";
            openFileDialog1.Title = "Select a File";
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LGE Data";
            openFileDialog1.DefaultExt = "lge";
            openFileDialog1.AutoUpgradeEnabled = true;
            openFileDialog1.AddExtension = true;
            openFileDialog1.RestoreDirectory = true;

            // If the directory doesn't exist, create it.
            if (!Directory.Exists(openFileDialog1.InitialDirectory))
            {
                Directory.CreateDirectory(openFileDialog1.InitialDirectory);
            }

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    // Create a PictureBox.
                    try
                    {
                        if (file.EndsWith("lge")) 
                        {
                            CommonUtil.readLGEFile(file, "|");
                        }
                        else if (file.EndsWith("xlsx"))
                        {
                            CommonUtil.ReadExcelFileToData(file);
                        }
                        else
                        {
                            throw new Exception("지원하지 않는 확장자");
                        }

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
        /// 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (!outOfFormUserInput_Click(sender, e)) return;

            if ((this.panel1.Controls[0] as Form).Name == "FormUserSimulateOutput")
            {
                (this.panel1.Controls[0] as FormUserSimulateOutput).saveSimulateFile();
                return;
            }

            if (this.panel1.Controls.Count > 0)
            {
                if (this.panel1.Controls[0] is Form)
                {
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "LGE File|*.lge|Excel File|*.xlsx";
                    saveFileDialog1.Title = "Select a File";
                    saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LGE Data";
                    saveFileDialog1.DefaultExt = "lge";
                    saveFileDialog1.AutoUpgradeEnabled = true;
                    saveFileDialog1.AddExtension = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = CDataControl.g_ReportData.get지역() + "_" + CDataControl.g_ReportData.get대리점() + "_" + CDataControl.g_ReportData.get판매자() + "_" + DateTime.Now.ToString("yyyyMMddHHmm");

                    // If the directory doesn't exist, create it.
                    if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LGE Data"))
                    {
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LGE Data");
                    }

                    if ((this.panel1.Controls[0] as Form).Name == "FormUserAnalysis")
                    {
                        (this.panel1.Controls[0] as FormUserAnalysis).saveComments();
                    }
                    if ((this.panel1.Controls[0] as Form).Name == "FormUserInput" ||
                        (this.panel1.Controls[0] as Form).Name == "FormUserOutput" ||
                        (this.panel1.Controls[0] as Form).Name == "FormUserAnalysis" ||
                        (this.panel1.Controls[0] as Form).Name == "FormReport" ||
                        (this.panel1.Controls[0] as Form).Name == "FormAdmin" ||
                        (this.panel1.Controls[0] as Form).Name == "FormUserSimulateInput")
                    {
                        saveFileDialog1.ShowDialog();

                        if (saveFileDialog1.FileName.EndsWith("lge"))
                        {
                            CommonUtil.writeLGEFile(saveFileDialog1.FileName, "|");
                        }
                        else if (saveFileDialog1.FileName.EndsWith("xlsx"))
                        {
                            FileInfo fi2 = new FileInfo(CommonUtil.defaultName);
                            fi2.CopyTo(saveFileDialog1.FileName, true);

                            CommonUtil.WriteDataToExcelFile(saveFileDialog1.FileName, false); 
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 업데이트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            if (!outOfFormUserInput_Click(sender, e)) return;
            // 업데이트
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Update File|*.lgm";
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
                        System.IO.File.Copy(file, CommonUtil.defaultManagerFileName, true);
                        toolStripButton4_Click(sender, e);  // UserInput으로 보내서 Refresh역할을 함
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
            if (!outOfFormUserInput_Click(sender, e)) return;
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
            if (!outOfFormUserInput_Click(sender, e)) return;
            FormUserAnalysis frm = new FormUserAnalysis();
            panelSet(frm);
        }

        /// <summary>
        /// 시뮬레이션입력
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            if (!outOfFormUserInput_Click(sender, e)) return;
            FormUserSimulateOutput frm = new FormUserSimulateOutput();
            panelSet(frm);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (!outOfFormUserInput_Click(sender, e)) return;
            new Printer();
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            if (!outOfFormUserInput_Click(sender, e)) return;
            FormAdmin frm = new FormAdmin();
            panelSet(frm);
        }

        private Boolean outOfFormUserInput_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.panel1.Controls.Count; i++)
            {
                if (this.panel1.Controls[i] is Form)
                {
                    if (this.panel1.Controls[i].Name == "FormUserInput")
                    {
                        if (!(this.panel1.Controls[i] as FormUserInput).validateData())
                        {
                            MessageBox.Show("지역, 대리점명, 마케터를 반드시 적어야 합니다.");
                            return false;
                        }
                        (this.panel1.Controls[i] as FormUserInput).saveAsInput();
                    }
                }
            }
            return true;
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
