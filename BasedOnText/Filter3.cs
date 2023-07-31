﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasedOnText
{
    class Filter3
    {
        Common common = new Common();
        Constant constant = new Constant();
        public List<string> Filter3List = new List<string>();
        public List<string> FilterWordList = new List<string>();
        public void ReadTextFromExcelFile(Excel.Application xlApp, string path)
        {
            try
            {
                FilterWordList = common.ReadTextFromExcelFile(xlApp, path);

            }
            catch (Exception ex)
            {
                common.makeLogFile("Filter3에서 익셀파일 읽기 오류!\n" + ex.ToString());
            }
        }

        public void FileUpload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool uploadCheck = */common.FileOneUpload(relativePath, constant.excelExtension, constant.filter3Filename);
                //if (uploadCheck)
                //{
                //    MessageBox.Show(constant.filter3Filename + "파일 업로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter3Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter3Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
            }
        }

        public string ChangeWord(string text)
        {
            try
            {
                //짝수는 왼쪽 홀수는 오른쪽
                for (int i = 0; i < FilterWordList.Count; i = i + 2)
                {
                    text = text.Replace(FilterWordList[i], FilterWordList[i + 1]);
                }
                return text;
            }
            catch(Exception ex)
            {
                common.makeLogFile("Filter3에서 단어 바꾸기 오류!");
                return text;
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

                if (File.Exists(fullPath + "\\" + constant.filter3Filename))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, constant.filter3Filename);
                    //if (downstatus)
                    //    MessageBox.Show(constant.filter3Filename + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(constant.filter3Filename + "파일이 존재하지 않습니다. 파일을 업로드하세요.");
                    common.makeLogFile(constant.filter3Filename + "파일이 존재하지 않습니다. 파일을 업로드하세요.");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(constant.filter3Filename + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter3Filename + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
    }
}
