using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        string extension;

        private void btn_load_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Data.filePath = ofd.FileName;
                    extension = Path.GetExtension(Data.filePath);
                    long size = new FileInfo(Data.filePath).Length;

                    MessageBoxButtons button = MessageBoxButtons.OK;

                    using (Stream str = ofd.OpenFile())
                    {
                        bool excheck = false;
                        bool sicheck = false;

                        excheck = CheckFileExtension(extension, button);
                        sicheck = CheckFileSize(size, button);

                        if (excheck && sicheck)
                        {
                            MessageBox.Show("불러오기 완료.", "Success", button);


                            DirectoryInfo parentDir = Directory.GetParent(Data.filePath);
                            Data.currentPath = parentDir.FullName;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBoxButtons button = MessageBoxButtons.OK;
                    MessageBox.Show("파일을 닫은 후 실행해주세요.", "Fail", button);
                }
            }
        }

        private bool CheckFileExtension(string extension, MessageBoxButtons button)
        {
            if (extension == ".hwp" || extension == ".odt")
            {
                return true;
            }

            MessageBox.Show("파일 확장자를 확인해주세요.", "Fail", button);
            return false;
        }

        private bool CheckFileSize(long size, MessageBoxButtons button)
        {
            if (size < 209715000)
            {
                return true;
            }

            MessageBox.Show("20mb 이하의 파일만 불러올 수 있습니다.", "Fail", button);
            return false;
        }

        private void bnt_convert_Click(object sender, EventArgs e)
        {
            if (extension == ".hwp")
            {
                HwpToOdt hto = new HwpToOdt();
                hto.Convert();
                SaveOdtFile();
            }
            else if (extension == ".odt")
            {
                OdtToHwp oth = new OdtToHwp();
                oth.Convert();
                MessageBoxButtons button = MessageBoxButtons.OK;
                MessageBox.Show("변환 완료.", "Success", button);
            }
        }

        private void SaveOdtFile()
        {
            MessageBoxButtons button = MessageBoxButtons.OK;
            MessageBox.Show("변환 완료.", "Success", button);
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string savefile = sfd.FileName;
                if (Path.GetExtension(savefile) != ".odt")
                    savefile += ".odt";
                try
                {
                    ZipFile.CreateFromDirectory(Application.StartupPath + @"\New File", savefile);
                }
                catch (IOException)
                {
                    File.Delete(savefile);
                    ZipFile.CreateFromDirectory(Application.StartupPath + @"\New File", savefile);
                }
                MessageBox.Show("저장 완료.", "Success", button);
            }
            sfd.FileName = "NewFile";
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }
    }
}
