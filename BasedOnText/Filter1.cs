using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasedOnText
{
    class Filter1
    {
        Common common = new Common();
        Constant constant = new Constant();
        public List<string> Filter1WordList = new List<string>();

        public void ReadTextFromExcelFile(Excel.Application xlApp, string path)
        {
            Filter1WordList = common.ReadTextFromExcelFile(xlApp, path);
        }

        public void FileUpload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool uploadCheck = */common.FileOneUpload(relativePath, constant.excelExtension, constant.filter1Filename);
                //if (uploadCheck)
                //{
                //    MessageBox.Show(constant.filter1Filename + "파일 업로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter1Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter1Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
            }
        }

        public string ChangeWord(string text)
        {
            try
            {
                //짝수는 왼쪽 홀수는 오른쪽
                for (int i = 0; i < Filter1WordList.Count; i = i + 2)
                {
                    text = text.Replace(Filter1WordList[i], Filter1WordList[i + 1]);
                }
                return text;
            }
            catch (Exception ex)
            {
                common.makeLogFile("Filter1에서 단어 바꾸기 오류!");
                return text;
            }
        }

        public void throughtFilter1FileUpload(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                //fileName = Path.GetFileNameWithoutExtension(fileName);
                string relativePath = constant.rootPath + constant.troughtFilter1Path;
                string fullPath = common.MakeFullpath(relativePath);
                if(common.makeFolder(fullPath) && common.FileOneUpload(relativePath, constant.docxExtension, fileName))
                {
                    //MessageBox.Show(fileName + "파일 업로드에 성공하엇습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("필터1 파일업로드 오류:\n" + ex.ToString());
            }
        }

        public void throughtFilter1FileDownload(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.troughtFilter1Path;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + fileName))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, fileName);
                    //if (downstatus)
                    //    MessageBox.Show(fileName + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(fileName + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile("파일 다운로드시 오류.\n" + ex.ToString());
            }
        }

        public void fileDownload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + constant.filter1Filename))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, constant.filter1Filename);
                    //if (downstatus)
                    //    MessageBox.Show(constant.filter1Filename + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(constant.filter1Filename + "파일이 존재하지 않습니다. 파일을 업로드하세요.");
                    common.makeLogFile(constant.filter1Filename + "파일이 존재하지 않습니다. 파일을 업로드하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter1Filename + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter1Filename + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
    }
}
