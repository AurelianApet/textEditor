using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasedOnText
{
    class TurchFile
    {
        List<string> MainFileList2 = new List<string>();
        List<string> LetterConnectList = new List<string>();
        List<string> CodeAddList = new List<string>();
        List<string> ChangeWordList = new List<string>();
        List<string> DeleteWordList = new List<string>();
        List<string> DeleteWordConnectList = new List<string>();
        Common common = new Common();
        Filter3 filter3 = new Filter3();
        Constant constant = new Constant();

        BasedOnText Form;
        public TurchFile(BasedOnText MainForm)
        {
            this.Form = MainForm;
        }
        public void MainFileUpload()
        {
            try
            {
                string relativePath = constant.rootPath + constant.mainFile2;
                string fileExtention = constant.docxExtension;
                MainFileList2 = common.FileUpload(relativePath, fileExtention, MainFileList2);

                for (int i = 0; i < MainFileList2.Count; i++)
                {
                    string FName = Path.GetFileName(MainFileList2[i]);
                    this.Form.turchMainFileListShow(FName, i + 1); //업로드된 원본파일 패널에 노출
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일(#2) 업로드시 오류가 발생하었습니다.");
                common.makeLogFile("원본파일(#2) 업로드시 오류가 발생하었습니다.");
            }
        }

        public void MainFileAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(MainFileList2);
                MainFileList2.RemoveRange(0, MainFileList2.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    MainFileList2.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.turchMainFileListShow(filename, i + 1);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("원본파일(#2) 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("원본파일(#2) 일괄삭제시 오류가 발생하엇습니다.");
            }
        }

        public void LetterConnect()
        {
            LetterConnectList.RemoveRange(0, LetterConnectList.Count);
            if (MainFileList2.Count == 0)
            {
                MessageBox.Show("업로드된 원본파일(#2)이 없습니다. 원본파일(#2)을 업로드 하십시오.");
                return;
            }
            this.Form.turchConnectLetterInitialize();
            curThreadId_connect = 0;
            Thread[] threadArray = new Thread[MainFileList2.Count];
            //word.Application createword = new word.Application();

            for (var i = 0; i < MainFileList2.Count; i++)
            {
                try
                {
                    string filename = Path.GetFileName(MainFileList2[i]);
                    //LetterConnectThread(MainFileList2[i], i + 1, i);
                    threadArray[i] = new Thread(new ThreadStart(() => LetterConnectThread(MainFileList2[i], i + 1, i)));
                    threadArray[i].Start();
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {

                }
            }

            for (var i = 0; i < MainFileList2.Count; i++)
            {
                threadArray[i].Join();
            }
            //common.ReleaseWordComObjects(createword, null);
            //Thread.Sleep(100);

            for (int i = 0; i < LetterConnectList.Count; i ++)
            {
                string filename = Path.GetFileName(LetterConnectList[i]);
                this.Form.turchLetterConnectShow(filename, i + 1);
            }
            MessageBox.Show("글자 잇기 실행에 완료되엇습니다.");
        }

        int curThreadId_connect = 0;
        public void LetterConnectThread(string file, int count, int threadId)
        {
            word.Application createword = new word.Application();
            var text = common.ReadTextFromWordFile(createword, file);
            string FileName = Path.GetFileNameWithoutExtension(file);
            FileName += "_Nospace";
            string extention = "docx";
            EndTagDelete(createword, text, FileName, extention, count, threadId);
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
        }

        public void EndTagDelete(word.Application createword, string text, string FileName, string extention, int count, int threadId)
        {
            try
            {
                int dotSpace = text.IndexOf(". ");
                if (dotSpace != -1) text = text.Remove(dotSpace + 1, 1);
                int dotEnter = text.IndexOf(".\r");
                if (dotEnter != -1) text = text.Remove(dotEnter + 1, 1);
                int questionSpace = text.IndexOf("? ");
                if (questionSpace != -1) text = text.Remove(questionSpace + 1, 1);
                int questionEnter = text.IndexOf("?\r");
                if (questionEnter != -1) text = text.Remove(questionEnter + 1, 1);
                int comSpace = text.IndexOf("! ");
                if (comSpace != -1) text = text.Remove(comSpace + 1, 1);
                int comEnter = text.IndexOf("!\r");
                if (comEnter != -1) text = text.Remove(comEnter + 1, 1);
                if (text.Contains(". ") || text.Contains(".\r") || text.Contains("! ") || text.Contains("!\r") || text.Contains("? ") || text.Contains("?\r"))
                {
                    EndTagDelete(createword, text, FileName, extention, count, threadId);
                    return;
                }
                else
                {
                    string relativePath = constant.rootPath + constant.connectLetterFile;
                    string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                    common.makeFolder(fullPath);

                    while (threadId != curThreadId_connect)
                        Thread.Sleep(100);
                    LetterConnectList.Add(fullPath + "\\" + FileName + "." + extention);
                    if (curThreadId_connect < MainFileList2.Count - 1)
                        curThreadId_connect++;
                    else
                        curThreadId_connect = 0;
                    /*bool status = */common.MakeWordFile(createword, text, fullPath + "\\" + FileName + "." + extention);
                }
            }
            catch(Exception ex)
            {
                common.makeLogFile("파일터치의 태그 지우기 실행시 오류가 발생하엇습니다.");
            }
        }

        public void LetterConnectAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(LetterConnectList);
                LetterConnectList.RemoveRange(0, LetterConnectList.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    LetterConnectList.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.turchLetterConnectShow(filename, i + 1);
                }              
            }
            catch(Exception ex)
            {
                MessageBox.Show("글자 잇기 파일 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("글자 잇기 파일 일괄삭제시 오류가 발생하엇습니다.");
            }
        }

        public void CodeAdd()
        {
            CodeAddList.RemoveRange(0, CodeAddList.Count);
            if (MainFileList2.Count == 0)
            {
                MessageBox.Show("업로드된 원본파일(#2)이 없습니다. 원본파일(#2)을 업로드하십시오!");
                return;
            }
            this.Form.turchCodeAddInitialize();
            curThreadId_addcode = 0;
            Thread[] threadArray = new Thread[MainFileList2.Count];

            for (int i = 0; i < MainFileList2.Count; i++)
            {
                try
                {
                    string filename = Path.GetFileName(MainFileList2[i]);
                    //addCodeThread(MainFileList2[i], i);
                    threadArray[i] = new Thread(new ThreadStart(() => addCodeThread(MainFileList2[i], i)));
                    threadArray[i].Start();
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {

                }
            }

            for (int i = 0; i < MainFileList2.Count; i++)
            {
                threadArray[i].Join();
            }

            for (int i = 0; i < CodeAddList.Count; i++)
            {
                string filename = Path.GetFileName(CodeAddList[i]); 
                this.Form.turchCodeAddShow(filename, i + 1);
            }
            MessageBox.Show("코드 추가 실행이 완료되엇습니다.");
        }
        int curThreadId_addcode = 0;

        public void addCodeThread(string file, int threadId)
        {
            List<string> text = new List<string>();
            word.Application createword = new word.Application();
            text = common.ReadTextFromWordFileToParagraph(createword, file);
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);

            for (int j = 0; j < text.Count; j++)
            {
                if(text[j].Trim() != "")
                {
                    text[j] = text[j].Trim() + "<br><br>";
                }
            }

            string relativePath = constant.rootPath + constant.addCode;
            string fullPath = common.MakeFullpath(relativePath);
            common.makeFolder(fullPath);

            string FileName = Path.GetFileNameWithoutExtension(file) + "_Coded.docx";
            while (threadId != curThreadId_addcode)
                Thread.Sleep(100);
            CodeAddList.Add(fullPath + "\\" + FileName);
            if (curThreadId_addcode < MainFileList2.Count - 1)
                curThreadId_addcode++;
            else
                curThreadId_addcode = 0;
            common.MakeWordFileToParagraph(text, fullPath + "\\" + FileName);
        }

        public void CodeAddAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(CodeAddList);
                CodeAddList.RemoveRange(0, CodeAddList.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    CodeAddList.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.turchCodeAddShow(filename, i + 1);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("코드 추가 파일 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("코드 추가 파일 일괄삭제시 오류가 발생하엇습니다.");
            }
        }

        public void WordChangeStart(bool check)
        {
            ChangeWordList.Clear();
            string relativePath1 = constant.rootPath + constant.changeWordFile;
            string fullPath1 = common.MakeFullpath(relativePath1);
            if (!common.makeFolder(fullPath1)) return; ;
            this.Form.turchChangeWordPanelInitialize();
                
            if (MainFileList2.Count == 0)
            {
                MessageBox.Show("업로드된 원본파일이 없습니다. 원본파일(#2) 업로드하십시오!");
                return;
            }

            if (check)
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                common.makeFolder(fullPath);

                FileInfo fileInfo = new FileInfo(fullPath + "\\Filter3.xlsx");

                if (File.Exists(fullPath + "\\" + constant.filter3Filename))
                {
                    Excel.Application xlApp = new Excel.Application();
                    filter3.ReadTextFromExcelFile(xlApp, fullPath + "\\" + constant.filter3Filename);//Filter3단어파일 생성
                    common.ReleaseExcelComObjects(xlApp, null, null);
                    Thread.Sleep(100);
                }
            }
            Thread[] threadArray = new Thread[MainFileList2.Count];

            //word.Application createword = new word.Application();
            curThreadId_changeword = 0;
            for (int i = 0; i < MainFileList2.Count; i++)
            {
                try
                {
                    string filename = Path.GetFileName(MainFileList2[i]);
                    //changeWordThread(MainFileList2[i], fullPath1, check, i);
                    threadArray[i] = new Thread(new ThreadStart(() => changeWordThread(MainFileList2[i], fullPath1, check, i)));
                    threadArray[i].Start();
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {

                }
            }
            for (int i = 0; i < MainFileList2.Count; i++)
            {
                threadArray[i].Join();
            }
            //common.ReleaseWordComObjects(createword, null);
            //Thread.Sleep(100);

            for (int i = 0; i < ChangeWordList.Count; i++)
            {
                string filename = Path.GetFileName(ChangeWordList[i]);
                this.Form.turchChangeWordShow(filename, i + 1);
            }
            MessageBox.Show("단어바꾸기 실행이 완료되엇습니다.");
        }

        int curThreadId_changeword = 0;

        public void changeWordThread(/*word.Application createword, */string file, string fullPath1, bool check, int threadId)
        {
            word.Application createword = new word.Application();
            var text = common.ReadTextFromWordFile(createword, file);

            if(check)
            {
                text = filter3.ChangeWord(text);
            }

            string FileName = Path.GetFileNameWithoutExtension(file) + "_CFN";
            common.MakeWordFile(createword, text, fullPath1 + "\\" + FileName + ".docx");
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
            while (threadId != curThreadId_changeword)
                Thread.Sleep(100);
            ChangeWordList.Add(fullPath1 + "\\" + FileName + ".docx");
            if (curThreadId_changeword < MainFileList2.Count - 1)
                curThreadId_changeword++;
            else
                curThreadId_changeword = 0;
        }

        public void ChangeWordAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(ChangeWordList);
                ChangeWordList.RemoveRange(0, ChangeWordList.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    ChangeWordList.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.turchChangeWordShow(filename, i + 1);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("단어 바꾸기 파일 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("단어 바꾸기 파일 일괄삭제시 오류가 발생하엇습니다.");
            }
        }

        public void DeleteLetterAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(DeleteWordList);
                DeleteWordList.RemoveRange(0, DeleteWordList.Count);

                for (int i = 0; i < errorList.Count(); i++)
                {
                    DeleteWordList.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.turchRule4DeleteWordShow(filename, i + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("글자 제거하기 파일 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("글자 제거하기 파일 일괄삭제시 오류.\n" + ex.ToString());
            }
        }

        public void DeleteConnectLetterAllDelete()
        {
            try
            {
                List<string> errorList = common.FileAllDelete(DeleteWordConnectList);
                DeleteWordConnectList.RemoveRange(0, DeleteWordConnectList.Count);

                for(int i = 0; i < errorList.Count(); i ++)
                {
                    DeleteWordConnectList.Add(errorList[i]);
                    string filename = Path.GetFileName(errorList[i]);
                    this.Form.turchDeleteLetterConnectShow(filename, i + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("글자 제가하기/글자 잇기 파일 일괄삭제시 오류가 발생하엇습니다.");
                common.makeLogFile("글자 제가하기/글자 잇기 파일 일괄삭제시 오류가 발생하엇습니다.");
            }
        }

        public void Rule4Start(bool check, List<string> Rule4SpaceList, List<string> Rule4FileList)
        {
            DeleteWordList.Clear();
            if (MainFileList2.Count == 0)
            {
                MessageBox.Show("업로드된 원본파일(#2) 존재하지 않습니다. 파일을 업로드 하세요!");
                return;
            }
            string relativePath1 = constant.rootPath + constant.deleteFile;
            string fullPath1 = common.MakeFullpath(relativePath1);
            if (!common.makeFolder(fullPath1)) return;
            this.Form.TurchFileTrougphRule4PanelInitialize();
            if (check)
            {
                word.Application createword = new word.Application();
                List<int> Rule4IntFileList = new List<int>();
                for (int j = 0; j < Rule4FileList.Count; j++)
                {
                    if (Rule4FileList[j] == "")
                    {
                        continue;
                    }

                    int filenumber = 0;
                    try
                    {
                        filenumber = Convert.ToInt32(Rule4FileList[j]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("룰4 박스의 입력란을 올바로 입력하세요.");
                    }
                    if (filenumber <= MainFileList2.Count)
                    {
                        var text = common.ReadTextFromWordFile(createword, MainFileList2[filenumber - 1]);

                        List<string> sentenses = new List<string>();
                        string[] array = common.SeperateComma(Rule4SpaceList[j]);
                        for (int k = 0; k < array.Length; k++)
                        {
                            try
                            {
                                sentenses = common.FindSentence(text);
                                if (sentenses.Count > 0)
                                {
                                    string[] sentenceNumber = array[k].Split('-');
                                    int start = text.IndexOf(sentenses[Convert.ToInt32(sentenceNumber[0]) - 1]);
                                    int Length = sentenses[Convert.ToInt32(sentenceNumber[0]) - 1].Length;

                                    string sentence = text.Substring(start, Length);
                                    int startNumber = common.GetSpaceNumber(sentence.Trim(), Convert.ToInt32(sentenceNumber[1]));
                                    int endNumber = common.GetSpaceNumber(sentence.Trim(), Convert.ToInt32(sentenceNumber[1]) + 1);

                                    if (endNumber == -1 || endNumber >= Length) endNumber = start + Length - 1;

                                    if (startNumber != -1 && endNumber != -1)
                                    {
                                        string changesentence = sentence.Remove(startNumber, endNumber - startNumber);
                                        text = text.Replace(sentence, changesentence);
                                    }
                                }
                            }
                            catch(Exception ex)
                            {
                                MessageBox.Show("룰4 박스 입력란을 올바로 입력하세요.");
                                continue;
                            }
                        }

                        string FileName = Path.GetFileNameWithoutExtension(MainFileList2[filenumber - 1]);
                        int duplicate = 0;
                        for (int i = 0; i < Rule4IntFileList.Count; i++)
                        {
                            if (Rule4IntFileList[i] == filenumber)
                                duplicate++;
                        }
                        Rule4IntFileList.Add(filenumber);

                        if (duplicate == 0)
                        {
                            FileName += "_Deleted.docx";

                        }
                        else
                        {
                            FileName += "_Deleted_Dupli" + duplicate + ".docx";
                        }
                        common.MakeWordFile(createword, text, fullPath1 + "\\" + FileName);
                        DeleteWordList.Add(fullPath1 + "\\" + FileName);
                    }
                }
                for(int i = 0; i < MainFileList2.Count; i++)
                {
                    try
                    {
                        if (!Rule4IntFileList.Contains(i + 1))
                        {
                            string FileName = Path.GetFileNameWithoutExtension(MainFileList2[i]);
                            FileName += "_Deleted.docx";
                            var text = common.ReadTextFromWordFile(createword, MainFileList2[i]);
                            common.MakeWordFile(createword, text, fullPath1 + "\\" + FileName);
                            DeleteWordList.Add(fullPath1 + "\\" + FileName);
                        }
                    }
                    catch(Exception)
                    {
                        continue;
                    }
                }
                common.ReleaseWordComObjects(createword, null);
                Thread.Sleep(100);
            }
            else
            {
                curThreadId_delword = 0;
                Thread[] threadArray = new Thread[MainFileList2.Count];
                //word.Application createword = new word.Application();
                for (int i = 0; i < MainFileList2.Count; i++)
                {
                    DeleteWordThreadNeverRule(MainFileList2[i], fullPath1, i);
                    threadArray[i] = new Thread(new ThreadStart(() => DeleteWordThreadNeverRule(MainFileList2[i], fullPath1, i)));
                    threadArray[i].Start();
                    Thread.Sleep(100);
                }
                for (int i = 0; i < MainFileList2.Count; i++)
                {
                    threadArray[i].Join();
                }
                //common.ReleaseWordComObjects(createword, null);
                //Thread.Sleep(100);
            }
            for (int i = 0; i < DeleteWordList.Count; i++)
            {
                string filename = Path.GetFileName(DeleteWordList[i]);
                this.Form.turchRule4DeleteWordShow(filename, i + 1);
            }
            MessageBox.Show("글자 제거하기 실행이 완료되엇습니다.");
        }

        int curThreadId_delword = 0;
        public void DeleteWordThreadNeverRule(string file, string fullPath1, int threadId)
        {
            string FileName = Path.GetFileNameWithoutExtension(file) + "_Deleted.docx";
            word.Application createword = new word.Application();
            var text = "";
            try
            {
                text = common.ReadTextFromWordFile(createword, file);

            }
            catch(Exception ex)
            {
                MessageBox.Show(file + "파일읽기중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile(file + "파일읽기중 프로세스가 작동중이므로 오류.\n" + ex.ToString());
                return;
            }
            common.MakeWordFile(createword, text, fullPath1 + "\\" + FileName);
            while (threadId != curThreadId_delword)
                Thread.Sleep(100);
            DeleteWordList.Add(fullPath1 + "\\" + FileName);
            if (curThreadId_delword < MainFileList2.Count - 1)
                curThreadId_delword++;
            else
                curThreadId_delword = 0;
            common.ReleaseWordComObjects(createword, null);
            Thread.Sleep(100);
        }

        public void Rule4StartAndConnectLetter(bool check, List<string> Rule4SpaceList, List<string> Rule4FileList)
        {
            DeleteWordConnectList.Clear();
            if (MainFileList2.Count == 0)
            {
                MessageBox.Show("업로드된 원본파일(#2) 존재하지 않습니다. 파일을 업로드 하세요!");
                return;
            }

            string relativePath1 = constant.rootPath + constant.deleteConnectFile;
            string fullPath1 = common.MakeFullpath(relativePath1);
            if (!common.makeFolder(fullPath1)) return;
            this.Form.TurchFileTrougphRule4ConnectPanelInitialize();
            if (check)
            {
                int count = 0;
                word.Application createword = new word.Application();
                List<int> Rule4IntFileList = new List<int>();
                for (int j = 0; j < Rule4FileList.Count; j++)
                {
                    if (Rule4FileList[j] == "")
                    {
                        continue;
                    }
                    int filenumber = 0;
                    try
                    {
                        filenumber = Convert.ToInt32(Rule4FileList[j]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("룰4 박스 입력란을 올바로 입력하세요.");
                    }
                    if (Convert.ToInt32(Rule4FileList[j]) <= MainFileList2.Count)
                    {
                        var text = common.ReadTextFromWordFile(createword, MainFileList2[filenumber - 1]);
                        string[] array = common.SeperateComma(Rule4SpaceList[j]);
                        for (int k = 0; k < array.Length; k++)
                        {
                            try
                            {
                                List<string> sentenses = new List<string>();

                                sentenses = common.FindSentence(text);

                                if (sentenses.Count > 0)
                                {
                                    string[] sentenceNumber = array[k].Split('-');
                                    int start = text.IndexOf(sentenses[Convert.ToInt32(sentenceNumber[0]) - 1]);
                                    int Length = sentenses[Convert.ToInt32(sentenceNumber[0]) - 1].Length;

                                    string sentence = text.Substring(start, Length);
                                    int startNumber = common.GetSpaceNumber(sentence.Trim(), Convert.ToInt32(sentenceNumber[1]));
                                    int endNumber = common.GetSpaceNumber(sentence.Trim(), Convert.ToInt32(sentenceNumber[1]) + 1);

                                    if (endNumber == -1 || endNumber >= Length) endNumber = start + Length - 1;
                                    if (startNumber != -1 && endNumber != -1)
                                    {
                                        string changesentence = sentence.Remove(startNumber, endNumber - startNumber);
                                        text = text.Replace(sentence, changesentence);
                                    }
                                }
                            }
                            catch(Exception ex)
                            {
                                MessageBox.Show("룰4 박스 입력란을 올바로 입력하세요.");
                                continue;
                            }
                        }

                        string FileName = Path.GetFileNameWithoutExtension(MainFileList2[filenumber -1]);

                        int duplicate = 0;
                        for(int i = 0; i < Rule4IntFileList.Count; i++)
                        {
                            if (Rule4IntFileList[i] == filenumber)
                                duplicate++;
                        }
                        Rule4IntFileList.Add(filenumber);
                        if (duplicate == 0)
                        {
                            FileName += "_Deleted_Nospace";
                        }
                        else
                        {
                            FileName += "_Deleted_Nospace_Dupli" + duplicate;
                        }

                        count++;
                        ConnectLetter(createword, text, FileName, "docx", count);
                    }
                }

                for (int i = 0; i < MainFileList2.Count; i++)
                {
                    if(!Rule4IntFileList.Contains(i + 1))
                    {
                        var text = common.ReadTextFromWordFile(createword, MainFileList2[i]);
                        string FileName = Path.GetFileNameWithoutExtension(MainFileList2[i]);
                        FileName += "_Deleted_Nospace";
                        count++;
                        ConnectLetter(createword, text, FileName, "docx", count);
                    }
                }
                common.ReleaseWordComObjects(createword, null);
                Thread.Sleep(100);
            }
            else
            {
                word.Application createword = new word.Application();
                for (int i = 0; i < MainFileList2.Count; i++)
                {
                    var text = common.ReadTextFromWordFile(createword, MainFileList2[i]);
                    string FileName = Path.GetFileNameWithoutExtension(MainFileList2[i]);
                    FileName += "_Deleted_Nospace";
                    ConnectLetter(createword, text, FileName, "docx", i + 1);
                }
                common.ReleaseWordComObjects(createword, null);
                Thread.Sleep(100);
            }
            MessageBox.Show("글자 제거하기/잇기 동시에 하기 실행이 완료되엇습니다.");
        }

        public void ConnectLetter(word.Application createword, string text, string FileName, string extention, int count)
        {
            int dotSpace = text.IndexOf(". ");
            if (dotSpace != -1) text = text.Remove(dotSpace + 1, 1);
            int dotEnter = text.IndexOf(".\r");
            if (dotEnter != -1) text = text.Remove(dotEnter + 1, 1);
            int questionSpace = text.IndexOf("? ");
            if (questionSpace != -1) text = text.Remove(questionSpace + 1, 1);
            int questionEnter = text.IndexOf("?\r");
            if (questionEnter != -1) text = text.Remove(questionEnter + 1, 1);
            int comSpace = text.IndexOf("! ");
            if (comSpace != -1) text = text.Remove(comSpace + 1, 1);
            int comEnter = text.IndexOf("!\r");
            if (comEnter != -1) text = text.Remove(comEnter + 1, 1);
            if (text.Contains(". ") || text.Contains(".\r") || text.Contains("! ") || text.Contains("!\r") || text.Contains("? ") || text.Contains("?\r"))
            {
                ConnectLetter(createword, text, FileName, extention, count);
                return;
            }
            else
            {
                string relativePath = constant.rootPath + constant.deleteConnectFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                common.makeFolder(fullPath);

                DeleteWordConnectList.Add(fullPath + "\\" + FileName + "." + extention);
                if (common.MakeWordFile(createword, text, fullPath + "\\" + FileName + "." + extention))
                {
                    this.Form.turchDeleteLetterConnectShow(FileName + "." + extention, count);
                }
            }
        }

        //원본파일(#2) 이어서보기 파일 생성
        public bool MainFile2ConnectView(string type)
        {
            try
            {
                string filename = constant.mainFile2ConnectViewName;
                bool status = common.FileConnectView(MainFileList2, filename, type);
                return status;
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일(#2) 이어서 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("원본파일(#2) 이어서 보기 실행시 오류가 발생하엇습니다.");
                return false;
            }
        }

        //원본파일(#2) 이어서 보기 파일 다운로드
        public bool MainFile2ConnectViewDownload()
        {
            string relativePath = constant.rootPath + constant.mainFile2;
            string fullPath = common.MakeFullpath(relativePath);
            bool status = common.makeFolder(fullPath);
            if (!status) return false;

            string filename = constant.mainFile2ConnectViewName;

            if (File.Exists(fullPath + "\\" + constant.mainFile2ConnectViewName))
            {
                try
                {
                    bool viewStatus = common.FileConnectViewDownload(fullPath, filename);
                    if (!viewStatus) return false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(constant.mainFile2ConnectViewName + "파일 생성후 다운로드시 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜 주십시오.");
                    common.makeLogFile(constant.mainFile2ConnectViewName + "파일 생성후 다운로드시 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜 주십시오.");
                    return false;
                }
                return true;
            }
            else
            {
                MessageBox.Show(constant.mainFile2ConnectViewName + "파일이 생성되지 않앗습니다. 다시 실행시켜주십시오.");
                return false; ;
            }
        }

        //이어서 보기 파일 다운로드
        public bool ConnectViewDownload(string relativePath , string filename)
        {
            try
            {
                string fullPath = common.MakeFullpath(relativePath);
                if (!common.makeFolder(fullPath)) return false;

                if (File.Exists(fullPath + "\\" + filename))
                {
                    if (!common.FileConnectViewDownload(fullPath, filename)) return false;
                    return true;
                }
                else
                {
                    MessageBox.Show(filename + "파일이 생성되지 않앗습니다. 다시 실행시켜주십시오.");
                    common.makeLogFile(filename + "파일이 생성되지 않앗습니다. 다시 실행시켜주십시오.");
                    return false; ;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(filename + "파일 이어서 보기 파일 생성후 다운로드시 오류가 발생하었습니다.");
                common.makeLogFile(filename + "파일 이어서 보기 파일 생성후 다운로드시 오류가 발생하었습니다.");
                return false;
            }
        }

        //이어서보기 파일 생성
        public bool ConnectView(string type, string filename, string state)
        {
            try
            {
                bool status = false;
                if (state == "connectLetter")
                {
                     status = common.FileConnectView(LetterConnectList, filename, type);

                }
                else if (state == "deleteLetter")
                {
                    status = common.FileConnectView(DeleteWordList, filename, type);

                }
                else if (state == "deleteConnectLetter")
                {
                    status = common.FileConnectView(DeleteWordConnectList, filename, type);

                }
                else if (state == "addCode")
                {
                    status = common.FileConnectView(CodeAddList, filename, type);

                }
                else if (state == "changeWord")
                {
                    status = common.FileConnectView(ChangeWordList, filename, type);

                }

                return status;
            }
            catch (Exception ex)
            {
                common.makeLogFile(state + "파일 이어서 보기 실행중 오류가 발생하엇습니다.");
                MessageBox.Show(state + "파일 이어서 보기 실행중 오류가 발생하엇습니다.");
                return false;
            }
        }

        //일괄 다운로드
        public void AllDownload(string state)
        {
            try
            {
                //bool status = false;
                if (state == "connectLetter")
                {
                    /*status = */common.AllDownload(LetterConnectList);
                }
                else if (state == "deleteLetter")
                {
                    /*status = */common.AllDownload(DeleteWordList);
                }
                else if (state == "deleteConnectLetter")
                {
                    /*status = */common.AllDownload(DeleteWordConnectList);
                }
                else if (state == "addCode")
                {
                    /*status = */common.AllDownload(CodeAddList);
                }
                else if (state == "changeWord")
                {
                    /*status = */common.AllDownload(ChangeWordList);
                }

                //if (status)
                //{
                //    MessageBox.Show(state + "파일 일괄 다운로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                common.makeLogFile(state + "파일 일괄 다운로드시 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜주십시오.");
                MessageBox.Show(state + "파일 일괄 다운로드시 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜주십시오.");
            }
        }

        //개객 다운로드
        public void OneFileDownload(object sender, EventArgs e, string relativePath)
        {
            try
            {
                Button btn = sender as Button;
                string fileName = btn.Name.Split('>')[0];
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                common.makeFolder(fullPath);
                common.FileConnectViewDownload(fullPath, fileName);
            }
            catch(Exception ex)
            {
                common.makeLogFile("파일 다운로드 오류!\n" + ex.ToString());
            }
        }
    }
}
