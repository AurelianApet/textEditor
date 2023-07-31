using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasedOnText
{
    class MainFile
    {
        Common common = new Common();
        Filter1 filter1 = new Filter1();
        Filter2 filter2;
        BasedOnText Form;
        Constant constant = new Constant();
        public MainFile(BasedOnText MainForm)
        {
            this.Form = MainForm;
            filter2 = new Filter2(this.Form);
        }

        public List<string> MainFileList = new List<string>();  //원본파일리스트
        public List<string> TroughFilter1List = new List<string>(); //필터1 통과한 파일 리스트 
        public List<string> TroughFilter2List = new List<string>(); //필터2 통과한 파일 리스트
        public List<string> TroughRule2List = new List<string>(); //룰2 통과한 파일 리스트
        public List<string> TroughRule3List = new List<string>(); //룰2 통과한 파일 리스트

        //원본파일 업로드
        public void FileUpload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.mainFilePath;
                string fileExtention = constant.docxExtension;
                MainFileList = common.FileUpload(relativePath, fileExtention, MainFileList);

                for (int i = 0; i < MainFileList.Count; i++)
                {
                    string FName = Path.GetFileName(MainFileList[i]);
                    this.Form.mainFileListShow(FName, i + 1); //업로드된 원본파일 패널에 노출
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일 업로드시 오류가 발생하었습니다.");
            }
        }

        public void FileDelete(object sender, EventArgs e)
        {
            try
            {
                Button btn = sender as Button;
                int btnFileDeleteNumber = Convert.ToInt32(btn.Name);
                string havetoDeletePath = MainFileList[btnFileDeleteNumber - 1];
                File.Delete(havetoDeletePath);
                Thread.Sleep(100);
                MainFileList.Remove(MainFileList[btnFileDeleteNumber - 1]);

                for (int i = 0; i < MainFileList.Count; i++)
                {
                    string FName = Path.GetFileName(MainFileList[i]);
                    this.Form.mainFileListShow(FName, i + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("파일 삭제시 오류!\n" + ex.ToString());
            }
        }

        public void Rule2FileDelete(object sender, EventArgs e)
        {
            try
            {
                Button btn = sender as Button;
                int btnFileDeleteNumber = Convert.ToInt32(btn.Name);
                string havetoDeletePath = TroughRule2List[btnFileDeleteNumber - 1];
                File.Delete(havetoDeletePath);
                Thread.Sleep(100);
                TroughRule2List.Remove(TroughRule2List[btnFileDeleteNumber - 1]);
                for (int i = 0; i < TroughRule2List.Count; i++)
                {
                    string FName = Path.GetFileName(TroughRule2List[i]);
                    this.Form.mainFileTroughRule2Show(FName, i + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("파일 삭제시 오류가 발생하엇습니다.");
            }
        }

        public void FileAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(MainFileList);
                MainFileList.RemoveRange(0, MainFileList.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    MainFileList.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.mainFileListShow(filename, i + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("원본파일 일괄삭제시 오류가 발생하엇습니다.");
            }
        }


        //파일 이어서보기 파일 생성
        public bool FileConnectView(string type)
        {
            try
            {
                string filename = constant.mainFileConnectViewName;
                bool status = common.FileConnectView(MainFileList, filename, type);
                return status;
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일 이어서 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("원본파일 이어서 보기 실행오류.\n" + ex.ToString());
                return false;
            }
        }

        //파일 이어서 보기 파일 다운로드
        public void FileConnectViewDownload()
        {
            string relativePath = constant.rootPath + constant.mainFilePath;
            string fullPath = common.MakeFullpath(relativePath);
            common.makeFolder(fullPath);
            string filename = constant.mainFileConnectViewName;
            if (File.Exists(fullPath + "\\" + constant.mainFileConnectViewName))
            {
                try
                {
                    common.FileConnectViewDownload(fullPath, filename);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(constant.mainFileConnectViewName + "파일 생성후 다운로드시 오류가 발생하엇습니다.");
                    common.makeLogFile(constant.mainFileConnectViewName + "파일 생성후 다운로드시 오류가 발생하엇습니다.");
                }
            }
            else
            {
                MessageBox.Show(constant.mainFileConnectViewName + "파일이 생성되지 않앗습니다. 다시 실행시켜주십시오.");
                return;
            }
        }

        //필터1통과한 파일 이어서보기 파일 생성
        public bool Filter1ConnectView(string type)
        {
            try
            {
                string filename = constant.TroughFilter1FileConnectViewName;
                bool status = common.FileConnectView(TroughFilter1List, filename, type);
                return status;
            }
            catch (Exception ex)
            {
                common.makeLogFile("필터1통과한 파일 이어서 보기 실행중 오류가 발생하엇습니다.");
                MessageBox.Show("필터1통과한 파일 이어서 보기 실행중 오류가 발생하엇습니다.");
                return false;
            }
        }

        //필터1 통과한 파일 이어서 보기 파일 다운로드
        public void Filter1ConnectViewDownload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.troughtFilter1Path;
                string fullPath = common.MakeFullpath(relativePath);
                common.makeFolder(fullPath);
                string filename = constant.TroughFilter1FileConnectViewName;
                if (File.Exists(fullPath + "\\" + constant.TroughFilter1FileConnectViewName))
                {
                    common.FileConnectViewDownload(fullPath, filename);
                }
                else
                {
                    MessageBox.Show(constant.TroughFilter1FileConnectViewName + "파일이 생성되지 않앗습니다. 다시 실행시켜주십시오.");
                    return;
                }
                //MessageBox.Show("필터1 통과한 파일 이어서보기 다운로드에 성공하엇습니다.");
            }
            catch (Exception ex)
            {
                common.makeLogFile("필터1통과한 파일 이어서 보기 파일 생성후 다운로드시 오류가 발생하었습니다.");
                MessageBox.Show("필터1통과한 파일 이어서 보기 파일 생성후 다운로드시 오류가 발생하었습니다.");
            }
        }

        public void GetMainFileFilter1Text(bool check)
        {
            try
            {
                TroughFilter1List.Clear();
                string relativePath1 = constant.rootPath + constant.troughtFilter1Path;
                string fullPath1 = common.MakeFullpath(relativePath1);
                if (!common.makeFolder(fullPath1)) return;

                this.Form.mainFileTrougphFilter1PanelInitialize();

                if (MainFileList.Count == 0)
                {
                    MessageBox.Show("원본파일을 업로드하세요!");
                    return;
                }

                if (check)
                {
                    string relativePath = constant.rootPath + constant.filterFilePath;
                    string fullPath = common.MakeFullpath(relativePath);
                    bool checkstatus = common.makeFolder(fullPath);
                    if (!checkstatus) return;

                    if (File.Exists(fullPath + "\\" + constant.filter1Filename))
                    {
                        Excel.Application xlApp = new Excel.Application();
                        filter1.ReadTextFromExcelFile(xlApp, fullPath + "\\" + constant.filter1Filename); // Filter1단어파일 생성
                        common.ReleaseExcelComObjects(xlApp, null, null);
                        Thread.Sleep(100);
                    }
                }

                curThreadId_filter1 = 0;
                Thread[] threadArray = new Thread[MainFileList.Count];

                //word.Application createword = new word.Application();
                for (int i = 0; i < MainFileList.Count; i++)
                {
                    //exchangeFilter1Word(createword, MainFileList[i], status, fullPath1/*, i*/);
                    try
                    {
                        threadArray[i] = new Thread(new ThreadStart(() => this.exchangeFilter1Word(MainFileList[i], check, fullPath1, i)));
                        threadArray[i].Start();
                        Thread.Sleep(100);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(Path.GetFileName(MainFileList[i]) + "파일 필터1 통과시 오류가 발생하엇습니다.");
                        common.makeLogFile(ex.ToString());
                    }
                }

                for (int i = 0; i < MainFileList.Count; i++)
                {
                    threadArray[i].Join();
                }
                //common.ReleaseWordComObjects(createword, null);
                //Thread.Sleep(100);
                for (int i = 0; i < TroughFilter1List.Count; i++)
                {
                    this.Form.mainFileTroughFilter1Show(Path.GetFileName(TroughFilter1List[i]), i + 1);
                }
                MessageBox.Show("필터1 통과시키기가 완료되엇습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("필터1 통과시키기 실행중 오류가 발생하엇습니다.");
                common.makeLogFile("필터1 통과시키기 실행중 오류.\n" + ex.ToString());
            }
        }

        int curThreadId_filter1 = 0;
        public void exchangeFilter1Word(string file, bool check, string fullPath1, int threadId)
        {
            string FileName = Path.GetFileNameWithoutExtension(file);
            word.Application createword = new word.Application();
            try
            {
                var text = common.ReadTextFromWordFile(createword, file);
                if (check)
                {
                    text = filter1.ChangeWord(text);
                }
                FileName += "_Filter1";
                common.MakeWordFile(createword, text, fullPath1 + "\\" + FileName + ".docx");
                while (curThreadId_filter1 != threadId)
                    Thread.Sleep(100);
                TroughFilter1List.Add(fullPath1 + "\\" + FileName + ".docx");
                if (curThreadId_filter1 < MainFileList.Count - 1)
                    curThreadId_filter1++;
                else
                    curThreadId_filter1 = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("필터1 통과시키기 실행중 " + FileName + "파일생성에서 워드파일의 단어치환시 오류가 발생하엇습니다.");
                common.makeLogFile(ex.ToString());
            }
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
        }

        public void TroughFilter1ListAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(TroughFilter1List);
                TroughFilter1List.RemoveRange(0, TroughFilter1List.Count);
                for (int i = 0; i < errorList.Count; i++)
                {
                    this.Form.mainFileTroughFilter1Show(errorList[i], i + 1);
                    string filename = Path.GetFileName(errorList[i]);
                    TroughFilter1List.Add(filename);
                }
            }
            catch (Exception ex)
            {
                common.makeLogFile("필터1 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
                MessageBox.Show("필터1 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
            }
        }


        List<wordStruct> words = new List<wordStruct> ();
        void makeWordFromFilter1List(string fname, int threadId)
        {
            word.Application createword = new word.Application();
            string filename = Path.GetFileName(fname);
            var paragraph = common.ReadTextFromWordFileToParagraph(createword, fname);
            List<string> wList = new List<string>();
            for (int k = 0; k < paragraph.Count; k++)
            {
                string[] word = paragraph[k].Split(' ');
                for (int j = 0; j < word.Length; j++)
                {
                    try
                    {
                        if (word[j] == "")
                            continue;
                        if (word[j].Contains('.'))
                        {
                            string[] dot = word[j].Split('.');
                            //if (!WordList.Contains(dot[0] + "."))
                            wList.Add(dot[0] + ".");
                            if (dot[1] != ""/* && !WordList.Contains(dot[1])*/)
                                wList.Add(dot[1]);
                        }
                        if (word[j].Contains('!'))
                        {
                            string[] gantam = word[j].Split('!');
                            //if (!WordList.Contains(gantam[0] + "."))
                            wList.Add(gantam[0] + ".");
                            if (gantam[1] != ""/* && !WordList.Contains(gantam[1])*/)
                                wList.Add(gantam[1]);
                        }
                        if (word[j].Contains('?'))
                        {
                            string[] question = word[j].Split('?');
                            //if (!WordList.Contains(question[0] + "."))
                            wList.Add(question[0] + ".");
                            if (question[1] != ""/* && !WordList.Contains(question[1])*/)
                                wList.Add(question[1]);
                        }
                        if (!word[j].Contains('.') && !word[j].Contains('!') && !word[j].Contains('?'))
                        {
                            //if(!WordList.Contains(word[j]))
                            wList.Add(word[j]);
                        }
                    }
                    catch (Exception ex)
                    {
                        common.makeLogFile("단어파일 생성시 오류.\n" + ex.ToString());
                        MessageBox.Show("단어파일 생성시 오류가 발생하엇습니다.");
                    }
                }
                if (wList.Contains("=>"))
                {
                    wList.Remove("=>");
                }
            }
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
            wordStruct wSt = new wordStruct();
            wSt.wList = wList;
            wSt.wNo = threadId;
            words.Add(wSt);
        }
        public void MakeWordFileStart()
        {
            if (TroughFilter1List.Count == 0)
            {
                MessageBox.Show("필터1 통과한 파일리스트 없습니다!");
                return;
            }
            Thread[] threadArray = new Thread[TroughFilter1List.Count];

            for (int i = 0; i < TroughFilter1List.Count; i++)
            {
                try
                {
                    threadArray[i] = new Thread(new ThreadStart(() => makeWordFromFilter1List(TroughFilter1List[i], i)));
                    threadArray[i].Start();
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                }
            }

            for (int i = 0; i < MainFileList.Count; i++)
            {
                threadArray[i].Join();
            }

            string relativePath = constant.rootPath + constant.filterFilePath;
            string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
            bool status = common.makeFolder(fullPath);
            if (!status) return;

            if (File.Exists(fullPath + "\\" + constant.filter2Filename))
            {
                Excel.Application xlApp = new Excel.Application();
                bool filter2status = filter2.ReadLeftTextFromExcelFile(xlApp, fullPath + "\\" + constant.filter2Filename); //Filter2단어파일 생성
                common.ReleaseExcelComObjects(xlApp, null, null);
                Thread.Sleep(100);
                if (!filter2status) return;
            }
            filter2.MakeDanoFile(words);
        }

        public void GetMainFileFilter2Text(bool Filtercheck, bool Rulecheck)
        {
            try
            {
                TroughFilter2List.RemoveRange(0, TroughFilter2List.Count);
                string relativePath1 = constant.rootPath + constant.troughtFilter2Path;
                string fullPath1 = common.MakeFullpath(relativePath1);
                if (!common.makeFolder(fullPath1)) return;

                this.Form.mainFileTrougphFilter2PanelInitialize();
                if (MainFileList.Count == 0)
                {
                    MessageBox.Show("원본파일을 업로드하세요!");
                    return;
                }

                if (Filtercheck)
                {
                    string relativePath = constant.rootPath + constant.filterFilePath;
                    string fullPath = common.MakeFullpath(relativePath);
                    bool filterstatus = common.makeFolder(fullPath);
                    if (!filterstatus) return;

                    if (File.Exists(fullPath + "\\" + constant.filter2Filename))
                    {
                        Excel.Application xlApp = new Excel.Application();
                        filter2.ReadTextFromExcelFile(xlApp, fullPath + "\\" + constant.filter2Filename); // Filter2 단어파일 생성
                        common.ReleaseExcelComObjects(xlApp, null, null);
                        Thread.Sleep(100);
                    }
                }

                curThreadId_filter2 = 0;
                Thread[] threadArray = new Thread[MainFileList.Count];
                //word.Application createword = new word.Application();

                for (int i = 0; i < MainFileList.Count; i++)
                {
                    string filename = Path.GetFileName(MainFileList[i]);
                    //exchangeFilter2Word(createword, MainFileList[i], fullPath1, filterCheck, ruleCheck/*, i*/);
                    try
                    {
                        threadArray[i] = new Thread(new ThreadStart(() => exchangeFilter2Word(MainFileList[i], fullPath1, Filtercheck, Rulecheck, i)));
                        threadArray[i].Start();
                        Thread.Sleep(100);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(filename + "파일 필터2/룰1 통과시키기 실행중 오류가 발생하엇습니다.");
                    }
                }

                for (int i = 0; i < MainFileList.Count; i++)
                {
                    threadArray[i].Join();
                }
                //common.ReleaseWordComObjects(createword, null);
                //Thread.Sleep(100);


                for (int i = 0; i < TroughFilter2List.Count; i++)
                {
                    string FileName = Path.GetFileName(TroughFilter2List[i]);
                    this.Form.mainFileTroughFilter2Show(FileName, i + 1);
                }
                MessageBox.Show("룰1/필터2 통과하기 완료되엇습니다.");
            }
            catch (Exception ex)
            {
                common.makeLogFile("룰1/필터2 통과시키기 실행중 오류가 발생하엇습니다.");
                MessageBox.Show("룰1/필터2 통과시키기 실행중 오류가 발생하엇습니다.");
            }
        }
        int curThreadId_filter2 = 0;
        public void exchangeFilter2Word(string file, string fullPath1, bool filterCheck, bool ruleCheck, int threadId)
        {
            word.Application createword = new word.Application();
            var text = common.ReadTextFromWordFile(createword, file);
            string FileName = Path.GetFileNameWithoutExtension(file);
            FileName += "_Ruled1";
            try
            {
                common.copyWordFile(file, fullPath1 + "\\" + FileName + constant.docxExtension);
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show("룰2 실행중 프로세스가 작동중이므로 " + FileName + constant.docxExtension + "복사시 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜주십시오.");
                common.makeLogFile("룰2 실행 - 복사시 오류\n" + ex.ToString());
                return;
            }

            if (filterCheck)
            {
                filter2.ChangeWord(createword, text, ruleCheck, fullPath1 + "\\" + FileName + constant.docxExtension);
            }
            while (curThreadId_filter2 != threadId)
                Thread.Sleep(100);
            TroughFilter2List.Add(fullPath1 + "\\" + FileName + constant.docxExtension);
            if (curThreadId_filter2 < MainFileList.Count - 1)
                curThreadId_filter2++;
            else
                curThreadId_filter2 = 0;
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
        }

        public void TroughFilter2ListAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(TroughFilter2List);
                TroughFilter2List.RemoveRange(0, TroughFilter2List.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    TroughFilter2List.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.mainFileTroughFilter2Show(filename, i + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("필터2/룰1 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("필터2/룰1 통과한 파일리스트 일괄 삭제시 오류.\n" + ex.ToString());
            }
        }

        public void TroughRule2ListAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(TroughRule2List);
                TroughRule2List.RemoveRange(0, TroughRule2List.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    TroughRule2List.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.mainFileTroughRule2Show(filename, i + 1);
                }
            }
            catch (Exception ex)
            {
                common.makeLogFile("룰2 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
                MessageBox.Show("룰2 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
            }
        }

        public void TroughRule3ListAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(TroughRule3List);
                TroughRule3List.RemoveRange(0, TroughRule3List.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    TroughRule3List.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.mainFileTroughRule3Show(filename, i + 1);
                }
            }
            catch (Exception ex)
            {
                common.makeLogFile("룰3 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
                MessageBox.Show("룰3 통과한 파일리스트 일괄 삭제시 오류가 발생하엇습니다.");
            }
        }

        List<int> Rule2IntFileList = new List<int>();
        int curRule2ThreadId = 0;
        void runRule2(string frule, string Rule2Space, string fPath, int threadId)
        {
            word.Application createword = new word.Application();
            word.Document readdoc = new word.Document();

            int filenumber = 0;
            try
            {
                filenumber = Convert.ToInt32(frule);
            }
            catch (Exception ex)
            {
                MessageBox.Show("룰2박스 파일 리스트 오더링을 올바로 입력하세요.");
                return;
            }

            int duplicate = 0;
            for (int p = 0; p < Rule2IntFileList.Count; p++)
            {
                if (Rule2IntFileList[p] == filenumber)
                    duplicate++;
            }

            Rule2IntFileList.Add(filenumber);

            if (filenumber <= TroughFilter2List.Count)
            {
                string FileName = Path.GetFileNameWithoutExtension(MainFileList[filenumber - 1]);
                if (duplicate == 0)
                {
                    FileName += "_Ruled2.docx";
                }
                else
                {
                    FileName += "_Ruled2_Dupli" + duplicate + ".docx";
                }
                try
                {
                    common.copyWordFile(TroughFilter2List[filenumber - 1], fPath + "\\" + FileName);
                    Thread.Sleep(100);
                }
                catch (Exception)
                {
                    return;
                }
                object fileName = fPath + "\\" + FileName;
                string[] array = common.SeperateComma(Rule2Space);
                for (int k = 0; k < array.Length; k++)
                {
                    try
                    {
                        var text = common.ReadTextFromWordFile(createword, fPath + "\\" + FileName);
                        object missing = System.Reflection.Missing.Value;
                        readdoc = createword.Documents.Open(ref fileName, ref missing, ReadOnly: false, missing, missing,
                            missing, missing, missing, missing, missing, missing, Visible: false, missing, missing, missing, missing);
                        createword.Visible = false;
                        List<string> sentenses = new List<string>();
                        sentenses = common.FindSentence(text);

                        string[] sentenceNumber = array[k].Split('-');
                        int start = 0;
                        int Length = 0;
                        start = text.IndexOf(sentenses[Convert.ToInt32(sentenceNumber[0]) - 1]);
                        Length = sentenses[Convert.ToInt32(sentenceNumber[0]) - 1].Length;
                        int spaceNumber = common.GetSpaceNumber(sentenses[Convert.ToInt32(sentenceNumber[0]) - 1], Convert.ToInt32(sentenceNumber[1]));
                        if (spaceNumber == -1) continue;

                        Object first = start;
                        Object middle = start + spaceNumber;
                        Object last = start + Length - 1;
                        if (spaceNumber != -1)
                        {
                            word.Range range = readdoc.Range(ref middle, ref last);
                            range.Underline = word.WdUnderline.wdUnderlineSingle;
                            range.InsertAfter("\t");
                            range.Select();
                            Thread.Sleep(100);
                            range.Cut();
                            Thread.Sleep(100);
                            readdoc.Range(ref first, ref first).PasteSpecial(word.WdPasteOptions.wdKeepSourceFormatting);
                        }
                        readdoc.Save();
                        common.ReleaseWordComObjects(null, readdoc);
                        Thread.Sleep(100);
                    }
                    catch (Exception ex)
                    {
                        common.ReleaseWordComObjects(null, readdoc);
                        Thread.Sleep(100);
                    }
                }
                while(curRule2ThreadId != threadId)
                {
                    Thread.Sleep(100);
                }
                TroughRule2List.Add(fPath + "\\" + FileName);
                curRule2ThreadId++;
            }
            else
            {
                while (curRule2ThreadId != threadId)
                {
                    Thread.Sleep(100);
                }
                curRule2ThreadId++;
            }
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
        }
        public void Rule2Start(bool check, List<string> Rule2SpaceList, List<string> Rule2FileList)
        {
            TroughRule2List.Clear();
            if (TroughFilter2List.Count == 0)
            {
                MessageBox.Show("필터2 통과한 파일 리스트 없습니다!");
                return;
            }

            string relativePath1 = constant.rootPath + constant.troughtRule2Path;
            string fullPath1 = common.MakeFullpath(relativePath1);

            if (!common.makeFolder(fullPath1)) return;

            this.Form.mainFileTrougphRule2PanelInitialize();
            if (check)
            {
                int rule2Cnt = 0;
                for (int i = 0; i < Rule2FileList.Count; i++)
                {
                    if(Rule2FileList[i] != "")
                    {
                        rule2Cnt++;
                    }
                }
                Thread[] threadArray = new Thread[rule2Cnt];
                rule2Cnt = 0;
                curRule2ThreadId = 0;
                for (int i = 0; i < Rule2FileList.Count; i++)
                {
                    try
                    {
                        if (Rule2FileList[i] == "")
                        {
                            continue;
                        }
                        bool isFind = false;
                        for(int j = 0; j < TroughFilter2List.Count; j++)
                        {
                            if ((j + 1).ToString() == Rule2FileList[i])
                            {
                                isFind = true; break;
                            }
                        }
                        if (isFind)
                        {
                            threadArray[rule2Cnt] = new Thread(new ThreadStart(() => runRule2(Rule2FileList[i], Rule2SpaceList[i], fullPath1, rule2Cnt)));
                            threadArray[rule2Cnt].Start();
                            Thread.Sleep(100);
                            rule2Cnt++;
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                for (int i = 0; i < rule2Cnt; i++)
                {
                    threadArray[i].Join();
                }

                for (int i = 0; i < TroughFilter2List.Count; i++)
                {
                    try
                    {
                        if (!Rule2IntFileList.Contains(i + 1))
                        {
                            string FileName = Path.GetFileNameWithoutExtension(MainFileList[i]);
                            FileName += "_Ruled2.docx";
                            common.copyWordFile(TroughFilter2List[i], fullPath1 + "\\" + FileName);
                            Thread.Sleep(100);
                            TroughRule2List.Add(fullPath1 + "\\" + FileName);
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            else
            {
                for (int i = 0; i < TroughFilter2List.Count; i++)
                {
                    try
                    {
                        string FileName = Path.GetFileNameWithoutExtension(MainFileList[i]);
                        FileName += "_Ruled2.docx";
                        common.copyWordFile(TroughFilter2List[i], fullPath1 + "\\" + FileName);
                        Thread.Sleep(100);
                        TroughRule2List.Add(fullPath1 + "\\" + FileName);
                    }
                    catch (Exception) { }
                }
            }

            for (int i = 0; i < TroughRule2List.Count; i++)
            {
                string filename = Path.GetFileName(TroughRule2List[i]);
                this.Form.mainFileTroughRule2Show(filename, i + 1);
            }
            MessageBox.Show("룰2 실행이 완료되엇습니다.");
        }

        public int SpaceNumber(string text, int number)
        {
            if (text.IndexOf(" ") == 0 || text.IndexOf("\r") == 0 || text.IndexOf("\n") == 0)
            {
                number++;
                text = text.Substring(2, text.Length - 2);
                SpaceNumber(text, number);
            }
            return number;
        }

        //int thread3Cnt = 0;
        public void Rule3Start(bool check, List<string> Rule3ListOrdering, List<string> Rule3ParagraphList, List<string> Rule3SentenceList, List<string> Rule3FileName)
        {
            TroughRule3List.RemoveRange(0, TroughRule3List.Count);
            if (TroughRule2List.Count == 0)
            {
                MessageBox.Show("룰2 통과한 파일 리스트가 없습니다.");
                return;
            }

            string relativePath1 = constant.rootPath + constant.troughtRule3Path;
            string fullPath1 = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath1)));
            bool status = common.makeFolder(fullPath1);
            if (!status) return;

            this.Form.mainFileTrougphRule3PanelInitialize();
            //Thread[] threadArray = new Thread[Rule3ListOrdering.Count];
            //thread3Cnt = 0;
            if (check)
            {
                //curRule3ThreadId = 0;
                for (int i = 0; i < Rule3ListOrdering.Count; i++)
                {
                    if (Rule3ListOrdering[i] != "")
                    {
                        string filename = Rule3FileName[i] + ".docx";
                        try
                        {
                            //threadArray[thread3Cnt] = new Thread(new ThreadStart(() => rule3Thread(Rule3ListOrdering[i], Rule3ParagraphList[i], Rule3SentenceList[i], Rule3FileName[i], fullPath1, i)));
                            //threadArray[thread3Cnt].Start();
                            //Thread.Sleep(100);
                            //thread3Cnt++;
                            rule3Thread(Rule3ListOrdering[i], Rule3ParagraphList[i], Rule3SentenceList[i], Rule3FileName[i], fullPath1, i);
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(filename + "파일 생성시 프로세스가 작동중이므로 오류가 발생하엇습니다. 우선 파일을 종료하고 다시 실행시켜주십시오.");
                            common.makeLogFile("룰3파일 생성시 오류!\n" + ex.ToString());
                            continue;
                        }
                    }
                }

                //for (int i = 0; i < thread3Cnt; i++)
                //{
                //    threadArray[i].Join();
                //}

                for (int i = 0; i < TroughRule3List.Count; i++)
                {
                    string filename = Path.GetFileName(Rule3FileName[i]) + ".docx";
                    this.Form.mainFileTroughRule3Show(filename, i + 1);
                }
            }
            else
            {
                MessageBox.Show("우선 룰3 체크 박스를 체크해 주세요.");
            }
            MessageBox.Show("룰3 실행이 완료되엇습니다.");
        }

        //int curRule3ThreadId = 0;
        public void rule3Thread(string Rule3ListOrderingFile, string Rule3ParagraphListFile, string Rule3SentenceListFile, string Rule3FileNameFile, string fullPath1, int threadId)
        {
            string[] fileNo = Rule3ListOrderingFile.Split(',');
            string[] Paragraph = new string[] { };
            string[] Sentence = new string[] { };
            if (Rule3ParagraphListFile != "")
            {
                Paragraph = Rule3ParagraphListFile.Split(',');
            }

            if (Rule3SentenceListFile != "")
            {
                Sentence = Rule3SentenceListFile.Split(',');
            }

            word.Application readword = new word.Application();
            word.Document readdoc = new word.Document();
            word.Document createdoc = new word.Document();
            object readOnly = true;
            object missing = System.Reflection.Missing.Value;
            for (int j = 0; j < fileNo.Length; j++)
            {
                int fno = 0;
                try
                {
                    fno = Convert.ToInt32(fileNo[j]);
                }
                catch(Exception ex)
                {
                    MessageBox.Show("룰3 문장/문단 수집 입력값들이 올바르지 않습니다!");
                    common.makeLogFile("룰3 문장/문단 수집 입력값 오류!\n" + ex.ToString());
                    continue;
                }
                if(Convert.ToInt32(fileNo[j]) > TroughRule2List.Count)
                {
                    continue;
                }
                string fname = Path.GetFileName(TroughRule2List[Convert.ToInt32(fileNo[j]) - 1]);
                object fileName = TroughRule2List[Convert.ToInt32(fileNo[j]) - 1];
                // Define an object to pass to the API for missing parameters
                try
                {
                    readdoc = readword.Documents.Open(ref fileName, ref missing, readOnly, missing, missing,
                        missing, missing, missing, missing, missing, missing, Visible: false, missing, missing, missing, missing);
                    readword.Visible = false;
                }
                catch(Exception ex)
                {
                    MessageBox.Show(fname + "파일 읽기시 오류가 발생하었습니다!");
                    common.makeLogFile("룰3 실행-파일 읽기시 오류!\n" + ex.ToString());
                    continue;
                }
                List<word.Range> para = new List<word.Range>();
                for(int k = 1; k <= readdoc.Paragraphs.Count; k++)
                {
                    try
                    {
                        if (readdoc.Paragraphs[k].Range.Text != "\r")
                            para.Add(readdoc.Paragraphs[k].Range);
                    }catch(Exception ex)
                    {

                    }
                }
                Object obj = new Object();

                for (int k = 0; k < Paragraph.Length; k++)
                {
                    Monitor.Enter(obj);
                    try
                    {
                        if (Convert.ToInt32(Paragraph[k]) <= para.Count)
                        {
                            readdoc.Activate();
                            para[Convert.ToInt32(Paragraph[k]) - 1].Select();
                            //Thread.Sleep(100);
                            para[Convert.ToInt32(Paragraph[k]) - 1].Copy();
                            //Thread.Sleep(100);
                            createdoc.Activate();
                            createdoc.Select();
                            //Thread.Sleep(100);
                            createdoc.Content.Paragraphs.Last.Range.Paste();
                            //Thread.Sleep(100);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("룰3 박스 문단 입력란이 올바르지 않습니다!");
                    }
                    finally
                    {
                        Monitor.Exit(obj);
                    }
                }
                string text = readdoc.Content.Text;
                List<string> sentenses = new List<string>();
                sentenses = common.FindSentence(text);
                if (sentenses.Count == 0) continue;

                for (int l = 0; l < Sentence.Length; l++)
                {
                    try
                    {
                        if (sentenses.Count > Convert.ToInt32(Sentence[l]))
                        {
                            int start = text.IndexOf(sentenses[Convert.ToInt32(Sentence[l]) - 1]);
                            int Length = sentenses[Convert.ToInt32(Sentence[l]) - 1].Length;
                            Object startpos = start;
                            Object endpos = start + Length;
                            readdoc.Activate();
                            readdoc.Range(ref startpos, ref endpos).Select();
                            //Thread.Sleep(100);
                            if(readdoc.Range(ref startpos, ref endpos) != null)
                            {
                                readdoc.Range(ref startpos, ref endpos).Copy();
                                //Thread.Sleep(100);
                                createdoc.Activate();
                                createdoc.Select();
                                //Thread.Sleep(100);
                                
                                createdoc.Content.Paragraphs.Last.Range.Paste();
                                //Thread.Sleep(100);
                                createdoc.Paragraphs.Add(ref missing);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        common.makeLogFile("룰3 문장 박스 입력값이 올바르지 않습니다.\n" + ex.ToString());
                        //MessageBox.Show("룰3 문장 박스 입력값이 올바르지 않습니다.");
                    }
                }
            }
            common.ReleaseWordComObjects(null, readdoc);
            Thread.Sleep(100);
            //while (curRule3ThreadId != threadId)
            //{
            //    Thread.Sleep(100);
            //}
            //curRule3ThreadId ++;
            TroughRule3List.Add(fullPath1 + "\\" + Rule3FileNameFile + ".docx");
            if(common.makeFolder(fullPath1))
                createdoc.SaveAs(fullPath1 + "\\" + Rule3FileNameFile + ".docx");
            Thread.Sleep(100);

            //while(curRule3ThreadId != thread3Cnt - 1)
            //{
            //    Thread.Sleep(100);
            //}
            common.ReleaseWordComObjects(readword, createdoc);
            Thread.Sleep(100);
        }

        public void throughtFilter1AllDownload()
        {
            try
            {
                /*bool status = */common.AllDownload(TroughFilter1List);
                //if(status)
                //{
                //    MessageBox.Show("필터1 통과한 파일 일괄 다운로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                common.makeLogFile("필터1통과한 파일 일괄 다운로드시 오류가 발생하엇습니다.");
                MessageBox.Show("필터1통과한 파일 일괄 다운로드시 오류.\n" + ex.ToString());
            }
        }

        //이어서 보기 파일 다운로드
        public void ConnectViewDownload(string relativePath, string filename)
        {
            try
            {
                string fullPath = common.MakeFullpath(relativePath);
                common.makeFolder(fullPath);

                if (File.Exists(fullPath + "\\" + filename))
                {
                    common.FileConnectViewDownload(fullPath, filename);
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                common.makeLogFile(filename + "파일 생성후 다운로드시 오류.\n" + ex.ToString());
            }
        }

        //이어서보기 파일 생성
        public bool ConnectView(string type, string filename, string state)
        {
            try
            {
                bool status = false;
                if (state == "Rule1")
                {
                    status = common.FileConnectView(TroughFilter2List, filename, type);

                }
                else if (state == "Rule2")
                {
                    status = common.FileConnectView(TroughRule2List, filename, type);

                }
                else if (state == "Rule3")
                {
                    status = common.FileConnectView(TroughRule3List, filename, type);
                }
                return status;
            }
            catch (Exception ex)
            {
                MessageBox.Show(state + "통과한 파일 이어서 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile(state + "통과한 파일 이어서 보기 실행시 오류.\n" + ex.ToString());
                return false;
            }
        }

        //일괄 다운로드
        public void AllDownload(string state)
        {
            try
            {
                //bool status = false;
                if (state == "Rule1")
                {
                    /*status = */common.AllDownload(TroughFilter2List);
                }
                else if (state == "Rule2")
                {
                    /*status = */common.AllDownload(TroughRule2List);
                }
                else if (state == "Rule3")
                {
                    /*status = */common.AllDownload(TroughRule3List);
                }
                //if (status)
                //{
                //    MessageBox.Show(state + "파일 일괄 다운로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(state + "통과한 파일리스트 일괄 다운로드 실행중 오류가 발생하엇습니다.");
                common.makeLogFile(state + "통과한 파일리스트 일괄 다운로드 실행중 오류.\n" + ex.ToString());
            }
        }
    }
}
