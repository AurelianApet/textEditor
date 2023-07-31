using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace BasedOnText
{
    class Filter2
    {
        Common common = new Common();
        public List<string> Filter2WordList = new List<string>();
        Constant constant = new Constant();

        BasedOnText Form;
        public Filter2(BasedOnText MainForm)
        {
            this.Form = MainForm;
        }
        public bool ReadTextFromExcelFile(Excel.Application xlApp, string path)
        {
            try
            {
                Filter2WordList = common.ReadTextFromExcelFile(xlApp, path);
                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(path + "파일 읽기시 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("파일 읽기시 오류.\n" + ex.ToString());
                return false;
            }
        }

        public bool ReadLeftTextFromExcelFile(Excel.Application xlApp, string path)
        {
            try
            {
                Filter2WordList = common.ReadLeftTextFromExcelFile(xlApp, path);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(path + "파일 읽기시 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("파일 읽기시 오류.\n" + ex.ToString());
                return false;
            }
        }

        public void MakeDanoFile(List<wordStruct> wordList)
        {
            List<string> DanoWordList = new List<string>();
            List<string> DuplicateList = new List<string>();

            List<wordStruct> wList = new List<wordStruct>();

            for (int i = 0; i < wordList.Count; i++)
            {
                for (int j = 0; j < wordList.Count; j++)
                {
                    if (i == wordList[j].wNo)
                        wList.Add(wordList[j]);
                }
            }

            for(int i = 0; i < wordList.Count; i++)
            { 
                for(int j = 0; j < wordList[i].wList.Count; j ++)
                {
                    if (!Filter2WordList.Contains(wordList[i].wList[j]))
                    {
                        if (!DanoWordList.Contains(wordList[i].wList[j]))
                            DanoWordList.Add(wordList[i].wList[j]);
                        //else
                        //{
                        //    if (!DuplicateList.Contains(wordList[i]))
                        //        DuplicateList.Add(wordList[i]);
                        //}
                    }
                    else
                    {
                        if (!DuplicateList.Contains(wordList[i].wList[j]))
                            DuplicateList.Add(wordList[i].wList[j]);
                    }
                }
            }

            string relativePath = constant.rootPath + constant.wordFile;
            string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
            common.makeFolder(fullPath);

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false; // 파일 저장시 같은 파일을 선택하게 되면 sheet내용을 모두 삭제하고 새로 만든다.

            common.MakeExcelFile(xlApp, DanoWordList, fullPath + "\\" + constant.danoFilename);
            common.MakeExcelFile(xlApp, DuplicateList, fullPath + "\\" + constant.duplicateFilename);

            common.ReleaseExcelComObjects(xlApp, null, null);
            Thread.Sleep(100);

            this.Form.DanoDuplicateWord(DanoWordList.Count, DanoWordList.Count + DuplicateList.Count, DuplicateList.Count);
            MessageBox.Show("단어파일 생성 완료엇습니다.");
        }

        public void FileUpload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool uploadCheck = */common.FileOneUpload(relativePath, constant.excelExtension, constant.filter2Filename);
                //if (uploadCheck)
                //{
                //    MessageBox.Show(constant.filter2Filename + "파일 업로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter2Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter2Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
            }
        }

        public void ChangeWord(word.Application createword, string text, bool Rulecheck, string filePath)
        {          
            //짝수는 왼쪽 홀수는 오른쪽
            for (int i = 0; i < Filter2WordList.Count; i = i + 2)
            {
                try
                {
                    if (Filter2WordList[i] != null && Filter2WordList[i] != "")
                    {
                        if (Filter2WordList[i + 1] == null)
                            Filter2WordList[i + 1] = "";
                        string[] similarWord = Filter2WordList[i + 1].Split(',');
                        word.Document createdoc = new word.Document();
                        object fileName = filePath;
                        // Define an object to pass to the API for missing parameters
                        object missing = System.Reflection.Missing.Value;
                        createdoc = createword.Documents.Open(ref fileName, ref missing, ReadOnly:false, missing, missing,
                            missing, missing, missing, missing, missing, missing, Visible: false, missing, missing, missing, missing);
                        createword.Visible = false;
                        text = ChangeFilter2Word(createword, createdoc, text, Filter2WordList[i], similarWord, filePath, Rulecheck);
                        createdoc.Save();
                        Thread.Sleep(100);
                        common.ReleaseWordComObjects(null, createdoc);
                        Thread.Sleep(100);
                        text = common.ReadTextFromWordFile(createword, filePath);
                    }
                }
                catch (Exception ex)
                {
                    common.makeLogFile("룰1/필터2 실행시 단어바꾸기에서 오류.\n" + ex.ToString());
                }
            }
        }

        int wordNumber = 0;
        public string ChangeFilter2Word(word.Application createword, word.Document createdoc, string text, string word, string[] similarWord, string filePath, bool rulecheck)
        {
            if (rulecheck && similarWord.Length > 0)
            {
                wordNumber = common.getRandomIndex(similarWord.Length);
            }
            else
            {
                if (wordNumber >= similarWord.Length) wordNumber = 0;
            }
            text = text.Replace("’", "'");
            text = text.Replace("”", "\"");
            word = word.Replace("’", "'");
            word = word.Replace("”", "\"");

            int index = text.LastIndexOf(word);
            if (index != -1)
            {
                try
                {
                    text = text.Substring(0, index);
                    
                    Thread.Sleep(100);
                    Object start = index;
                    Object end = index + word.Length;
                    word.Range range = createdoc.Range(ref start, ref end);
                    if (range.Font.Bold == 0)
                    {
                        range.Text = similarWord[wordNumber];
                        range.Font.Bold = -1;
                    }
                    wordNumber++;
                    ChangeFilter2Word(createword, createdoc, text, word, similarWord, filePath, rulecheck);
                }
                catch (Exception ex)
                {
                    common.makeLogFile("Filter2에서 오류!\n" + ex.ToString());
                }
            }
            return text;
        }

        public void fileDownload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + constant.filter2Filename))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, constant.filter2Filename);
                    //if (downstatus)
                    //    MessageBox.Show(constant.filter2Filename + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(constant.filter1Filename + "파일이 존재하지 않습니다. 파일을 업로드하세요.");
                    common.makeLogFile(constant.filter1Filename + "파일이 존재하지 않습니다. 파일을 업로드하세요.");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(constant.filter1Filename + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter1Filename + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
    }
}
