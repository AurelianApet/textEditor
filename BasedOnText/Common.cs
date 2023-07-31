using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text.RegularExpressions;

namespace BasedOnText
{
    struct wordStruct
    {
        public List<string> wList;
        public int wNo;
    }

    public class Common
    {
        Constant constant = new Constant();
        List<string> processList = new List<String>();
        ForceExit forceExitDialog = new ForceExit();
        public static bool over = false;

        public void ReleaseWordComObjects(word.Application wrd, word.Document doc, bool isQuitting = true)
        {
            try
            {
                if (isQuitting)
                {
                    if (doc != null)
                    {
                        doc.Close(false, System.Type.Missing, System.Type.Missing);
                        Thread.Sleep(100);
                    }
                    if (wrd != null)
                    {
                        wrd.Quit();
                        Thread.Sleep(100);
                    }
                }
                if (doc != null) { Marshal.ReleaseComObject(doc); doc = null; Thread.Sleep(20);}
                if (wrd != null) {Marshal.ReleaseComObject(wrd); wrd = null; Thread.Sleep(20); }
            }
            catch { }
            finally { GC.Collect(); }
        }
        public void ReleaseExcelComObjects(Excel.Application xlApp, Excel.Workbook xlWorkBook, Excel.Worksheet xlWorkSheet, bool isQuitting = true)
        {
            try
            {
                if (isQuitting)
                {
                    if (xlWorkBook != null)
                    {
                        xlWorkBook.Close(false, System.Type.Missing, System.Type.Missing);
                        Thread.Sleep(100);
                    }
                    if (xlApp != null)
                    {
                        xlApp.Quit(); Thread.Sleep(100);
                    }
                }
                if (xlWorkSheet != null) { Marshal.ReleaseComObject(xlWorkSheet); xlWorkSheet = null; Thread.Sleep(10); }
                if (xlWorkBook != null) {Marshal.ReleaseComObject(xlWorkBook); xlWorkBook = null; Thread.Sleep(10); }
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); xlApp = null; Thread.Sleep(10); }              
            }
            catch { }
            finally { GC.Collect(); }
        }

        public List<string> ReadTextFromExcelFile(Excel.Application xlApp, string FilePath)
        {
            List<string> ExcelWordList = new List<string>();
            try
            {
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(FilePath, 0, FileAccess.Read);
                Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);
                xlApp.Visible = false;
                for (int i = 1; i <= xlWorkSheet.Rows.Count; i++)
                {
                    if (xlWorkSheet.Cells[i, 1].value == null || xlWorkSheet.Cells[i, 1].value == "")
                    {
                        break;
                    }
                    else
                    {
                        ExcelWordList.Add(Convert.ToString(xlWorkSheet.Cells[i, 1].value));
                        ExcelWordList.Add(Convert.ToString(xlWorkSheet.Cells[i, 2].value));
                    }
                }
                ReleaseExcelComObjects(null, xlWorkBook, xlWorkSheet);
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                makeLogFile(ex.ToString());
            }
            return ExcelWordList;
        }

        public List<string> ReadLeftTextFromExcelFile(Excel.Application xlApp, string FilePath)
        {
            List<string> ExcelWordList = new List<string>();
            try
            {
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(FilePath, 0, FileAccess.Read);
                Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);
                xlApp.Visible = false;
                for (int i = 1; i <= xlWorkSheet.Rows.Count; i++)
                {
                    if (xlWorkSheet.Cells[i, 1].value == null || xlWorkSheet.Cells[i, 1].value == "")
                    {
                        break;
                    }
                    else
                    {
                        ExcelWordList.Add(Convert.ToString(xlWorkSheet.Cells[i, 1].value));
                    }
                }
                ReleaseExcelComObjects(null, xlWorkBook, xlWorkSheet);
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                makeLogFile(ex.ToString());
            }
            return ExcelWordList;
        }

        //파일 업로드
        public List<string> FileUpload(string relativePath, string fileExtention, List<string> UploadFiles)
        {
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                openDialog.Multiselect = true;
                openDialog.ShowDialog();
                foreach (string s in openDialog.FileNames)
                {
                    var extension = Path.GetExtension(s);
                    string filename = Path.GetFileName(s);

                    if (extension == fileExtention)
                    {
                        string fullPath = MakeFullpath(relativePath);
                        bool status = this.makeFolder(fullPath);
                        if (!status) return UploadFiles;

                        if (File.Exists(fullPath + "\\" + filename))
                        {
                            File.Delete(fullPath + "\\" + filename);

                            //오버로드된 파일이 원본파일리스트에 존재하는지 체크
                            bool check = false;
                            for (var k = 0; k < UploadFiles.Count; k++)
                            {
                                if (UploadFiles[k] == fullPath + "\\" + filename)
                                    check = true;
                            }
                            if (!check)
                                UploadFiles.Add(fullPath + "\\" + filename); // 파일 추가
                        }
                        else
                        {
                            UploadFiles.Add(fullPath + "\\" + filename); // 파일 추가
                        }
                        File.Copy(s, fullPath + "\\" + filename);
                        Thread.Sleep(100);
                    }
                    else
                    {
                        MessageBox.Show(extension + "파일입니다 " + filename + "파일 입력해주세요!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 업로드시 오류가 발생하엇습니다.");
                makeLogFile(ex.ToString());
            }
            return UploadFiles;
        }

        //파일 개개 업로드
        public bool FileOneUpload(string relativePath, string FileExtention, string filename)
        {
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                openDialog.ShowDialog();
                string fname = openDialog.FileName;
                string extension = Path.GetExtension(fname);

                if (extension == FileExtention)
                {
                    string fullPath = this.MakeFullpath(relativePath);
                    bool status = this.makeFolder(fullPath);
                    if (!status) return false;

                    if (File.Exists(fullPath + "\\" + filename))
                    {
                        File.Delete(fullPath + "\\" + filename);
                    }
                    File.Copy(fname, fullPath + "\\" + filename);
                    Thread.Sleep(100);
                    return true;
                }
                else
                {
                    MessageBox.Show(filename + "은 확장자가 " + FileExtention + "파일입니다 " + FileExtention + "파일 입력해주세요!");
                    return false;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(filename + "파일 업로드시 오류가 발생하엇습니다.");
                this.makeLogFile(ex.ToString());
                return false;
            }
        }

        public void MakeExcelFile(Excel.Application xlApp, List<string> WordList, string path)
        {
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {
                if (xlApp != null)
                {
                    for (int i = 0; i < WordList.Count; i++)
                    {
                        xlWorkSheet.Cells[i + 1, 1] = WordList[i];
                    }
                    if(File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    xlWorkBook.SaveAs(path);
                    ReleaseExcelComObjects(null, xlWorkBook, xlWorkSheet);
                    Thread.Sleep(100);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(path + "파일이 실행중이므로 종료하고 다시 작동시켜주십시오.");
                this.makeLogFile(ex.ToString());
                ReleaseExcelComObjects(null, xlWorkBook, xlWorkSheet);
            }
            
        }

        public void MakeSeperateExcelFile(Excel.Application xlApp, List<string> WordList, string path)
        {
            xlApp.DisplayAlerts = false; // 파일 저장시 같은 파일을 선택하게 되면 sheet내용을 모두 삭제하고 새로 만든다.

            if (xlApp != null)
            {
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                int count = 1;
                for (int i = 0; i < WordList.Count; i = i + 2)
                {
                    xlWorkSheet.Cells[count, 1] = WordList[i];
                    xlWorkSheet.Cells[count, 2] = WordList[i + 1];
                    count++;
                }
                xlWorkBook.SaveAs(path);
                ReleaseExcelComObjects(null, xlWorkBook, xlWorkSheet);
                Thread.Sleep(100);
            }
        }

        //이어서 보기 
        public bool FileConnectView(List<string> FileList, string filename, string type)
        {
            bool ret = false;
            if (FileList.Count == 0)
            {
                MessageBox.Show("파일이 없습니다!");
                return ret;
            }

            word.Application createword = new word.Application();
            word.Document readdoc = new word.Document();

            word.Document createdoc = new word.Document();
            for (var i = 0; i < FileList.Count; i++)
            {
                object fileName = FileList[i];
                string fname = Path.GetFileName(FileList[i]);
                try
                {
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    readdoc = createword.Documents.Open(ref fileName, ref missing, readOnly, missing, missing,
                        missing, missing, missing, missing, missing, missing, Visible:false, missing, missing, missing, missing);
                    createword.Visible = false;
                    word.Range fnameRange = createdoc.Paragraphs.Last.Range;
                    fnameRange.Text = fname;

                    fnameRange.Select();
                    Thread.Sleep(100);
                    fnameRange.Copy();
                    Thread.Sleep(100);
                    createdoc.Paragraphs.Last.Range.Paste();
                    Thread.Sleep(100);

                    readdoc.Content.Select();
                    Thread.Sleep(100);
                    readdoc.Content.Copy();
                    Thread.Sleep(100);
                    createdoc.Paragraphs.Last.Range.Paste();
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                    ReleaseWordComObjects(createword, readdoc);
                    Thread.Sleep(100);
                    this.makeLogFile(ex.ToString());
                    MessageBox.Show("파일 이어서 보기시 " + fname + "읽기중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                    return ret;
                }
                ReleaseWordComObjects(null, readdoc);
                Thread.Sleep(100);
            }

            string fpath = Path.GetFullPath(FileList[0]);
            string saveUrl =  fpath.Substring(0, fpath.IndexOf(Path.GetFileName(fpath))) + filename;
            if (File.Exists(saveUrl))
            {
                try
                {
                    File.Delete(saveUrl);
                }
                catch (Exception ex)
                {
                    ReleaseWordComObjects(createword, createdoc);
                    Thread.Sleep(100);
                    MessageBox.Show(saveUrl + "오버로드시 프로세스가 작동중이므로 오류가 발생하엇습니다. 프로그램을 종료하고 다시 실행시켜주십시오.");
                    this.makeLogFile(ex.ToString());
                    return ret;
                }
            }
            try
            {
                createword.Visible = false;
                createdoc.SaveAs(saveUrl);
                if (type == "download")
                {
                    ReleaseWordComObjects(createword, createdoc);
                    Thread.Sleep(100);
                }
                else
                {
                    createdoc.Close();
                    Thread.Sleep(100);

                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    createword.Documents.Open(saveUrl, ref missing, readOnly, missing, missing,
                        missing, missing, missing, missing, missing, missing, Visible: true, missing, missing, missing, missing);
                    createword.Visible = true;
                    ReleaseWordComObjects(createword, createdoc, false);
                    Thread.Sleep(100);
                }
                ret = true;
            }
            catch (Exception ex)
            {
                ReleaseWordComObjects(createword, createdoc);
                Thread.Sleep(100);
                MessageBox.Show("파일 이어서 보기 파일 창조시 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                this.makeLogFile(ex.ToString());
            }
            return ret;
        }

        //다운로드
        public bool FileConnectViewDownload(string fullpath, string filename)
        {
            SaveFileDialog connectfile = new SaveFileDialog();
            connectfile.Title = "다운로드";
            connectfile.FileName = filename;
            connectfile.DefaultExt = "docx";
            connectfile.Filter = "Docx files (*.docx)|*.doc|All files (*.*)|*.*";
            if (connectfile.ShowDialog() == DialogResult.OK)
            {

                if (File.Exists(connectfile.FileName))
                {
                    File.Delete(connectfile.FileName);
                    Thread.Sleep(100);
                }
                File.Copy(fullpath + "\\" + filename, connectfile.FileName);
                Thread.Sleep(100);
                return true;
            }
            return false;
        }

        //파일 일괄 삭제
        public List<string> FileAllDelete(List<string> FileList)
        {
            List<string> errorList = new List<string>();
            for (int i = 0; i < FileList.Count; i++)
            {
                try
                {
                    File.Delete(FileList[i]);
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                    errorList.Add(FileList[i]);
                    MessageBox.Show(Path.GetFileName(FileList[i]) + "파일 삭제시 프로세스가 작동중이므로 삭제할수 없습니다.");
                    makeLogFile(ex.ToString());
                    continue;
                }
            }
            return errorList;
        }

        public bool MakeWordFile(word.Application createword, string text, string path)
        {
            try
            {
                word.Document createdoc = new word.Document();
                createdoc.Content.Text = text;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)));
                this.makeFolder(fullPath);
                createdoc.SaveAs(path);
                ReleaseWordComObjects(null, createdoc);
                Thread.Sleep(100);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(path + "파일 생성중 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실해시켜주십시오.");
                this.makeLogFile(ex.ToString());
            }
            return false;
        }

        public bool MakeWordFileToParagraph(List<string> text, string path)
        {
            word.Application createword = new word.Application();
            word.Document createdoc = new word.Document();
            bool ret = false;
            try
            {
                for (int i = 0; i < text.Count; i++)
                {
                    createdoc.Content.Text += text[i]/* + "\r\n"*/;
                }

                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)));
                this.makeFolder(fullPath);
                createdoc.SaveAs(path);
                ret = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            ReleaseWordComObjects(createword, createdoc);
            Thread.Sleep(100);
            return ret;
        }

        public string ReadTextFromWordFile(word.Application createword, string FilePath)
        {
            string text = "";
            word.Document createdoc = new word.Document();
            try
            {
                object fileName = FilePath;
                // Define an object to pass to the API for missing parameters
                object readOnly = true;
                object missing = System.Reflection.Missing.Value;
                createdoc = createword.Documents.Open(ref fileName, ref missing, readOnly, missing, missing,
                        missing, missing, missing, missing, missing, missing, Visible: false, missing, missing, missing, missing);
                createword.Visible = false;
                text = createdoc.Content.Text;
                ReleaseWordComObjects(null, createdoc);
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                ReleaseWordComObjects(null, createdoc);
                Thread.Sleep(200);
                //MessageBox.Show(FilePath + "파일 읽기시 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜주십시오.");
                ReadTextFromWordFile(createword, FilePath);
            }
            return text;
        }

        public List<string> ReadTextFromWordFileToParagraph(word.Application createword, string FilePath)
        {
            word.Document createdoc = new word.Document();
            List<string> data = new List<string>();
            object fileName = FilePath;
            // Define an object to pass to the API for missing parameters

            object readOnly = true;
            object missing = System.Reflection.Missing.Value;
            createdoc = createword.Documents.Open(ref fileName, ref missing, readOnly, missing, missing,
                        missing, missing, missing, missing, missing, missing, Visible: false, missing, missing, missing, missing);
            createword.Visible = false;
            for (int j = 0; j <= createdoc.Paragraphs.Count; j++)
            {
                try
                {
                    data.Add(createdoc.Paragraphs[j + 1].Range.Text.Trim());
                }
                catch (Exception ex)
                {
                    this.makeLogFile(ex.ToString());
                }
            }
            ReleaseWordComObjects(null, createdoc);
            Thread.Sleep(100);
            return data;
        }

        public int getRandomIndex(int totalCount)    //랜덤수 생성
        {
            int index = 0;
            Random rand = new Random();
            index = rand.Next(totalCount);
            return index;
        }

        public List<string> FindSentence(string text)
        {
            List<string> sentenses = new List<string>();
            try
            {
                string temp = text.Replace(Environment.NewLine, " ");
                char[] arrSplitChars = { '.', '?', '!', '\r', '\v' };
                string[] splitSentences = temp.Split(arrSplitChars, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < splitSentences.Length; i++)
                {
                    if (splitSentences[i].Trim() == "" || splitSentences[i].Trim() == "\r")
                        continue;
                    int pos = temp.IndexOf(splitSentences[i].ToString());
                    char[] arrChars = temp.Trim().ToCharArray();
                    char c = arrChars[pos + splitSentences[i].Length];
                    if(splitSentences[i].ToString().Trim() + c.ToString() != "\r")
                    {
                        sentenses.Add(splitSentences[i].ToString() + c.ToString());
                    }
                }
                return sentenses;
            }
            catch (Exception ex)
            {
                this.makeLogFile(ex.ToString());
                return sentenses;
            }
        }

        public List<int> SortingList(List<int> endPointList)
        {
            List<int> SortingNumberList = new List<int>();
            try
            {
                int smallestNumber = 0;
                for (int j = 0; j < endPointList.Count; j++)
                {
                    int getNumber = 10000000;
                    for (int i = 0; i < endPointList.Count; i++)
                    {
                        if (getNumber > endPointList[i] && endPointList[i] > smallestNumber)
                        {
                            getNumber = endPointList[i];
                        }
                    }
                    SortingNumberList.Add(getNumber);
                    smallestNumber = getNumber;
                }
                return SortingNumberList;
            }
            catch (Exception ex)
            {
                this.makeLogFile("sorting 오류입니다\n" + ex.ToString());
                return SortingNumberList;
            }
        }

        public int GetSpaceNumber(string text, int number)
        {
            List<int> SpaceList = new List<int>();
            try
            {
                int previous = 0;
                string[] space = text.Split(' ');
                int spacenumber = 0;
                if (number > space.Length)
                {
                    return -1;
                }
                for (int k = 0; k < space.Length; k++)  //덜기 1은 마지막 제거
                {
                    previous = spacenumber;
                    spacenumber += space[k].Length + 1; //1 은 공백 포함
                    if(spacenumber - previous != 1)
                        SpaceList.Add(spacenumber);
                }
                return SpaceList[number - 1];
            }
            catch (Exception ex)
            {
                this.makeLogFile("공백 얻기 오류!\n" + ex.ToString());
                return -1;
            }
        }

        public string[] SeperateComma(string Space)
        {
            string[] Array = Space.Split(',');
            return Array;
        }

        public bool AllDownload(List<string> filelist)
        {
            if (filelist.Count == 0)
            {
                MessageBox.Show("파일이 없습니다!");
                return false;
            }

            FolderBrowserDialog connectfile = new FolderBrowserDialog();
            if (connectfile.ShowDialog() == DialogResult.OK)
            {
                string folderName = connectfile.SelectedPath;
                for (int i = 0; i < filelist.Count; i++)
                {
                    try
                    {
                        string filename = folderName + "\\" + Path.GetFileName(filelist[i]);
                        if (File.Exists(filename))
                        {
                            forceExitDialog.ShowText(filename);
                            forceExitDialog.ShowDialog();
                            if (over)
                            {
                                File.Delete(filename);
                                Thread.Sleep(100);
                                File.Copy(filelist[i], filename);
                                Thread.Sleep(100);
                            }
                        }
                        else
                        {
                            File.Copy(filelist[i], filename);
                            Thread.Sleep(100);
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
                return true;
            }
            return false;
        }
        
        public bool FileView(string filePath)
        {
            string filename = Path.GetFileName(filePath);
            if (!File.Exists(filePath))
            {
                MessageBox.Show(filename + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                this.makeLogFile(filename + "파일이 존재하지 않습니다.");
                return false;
            }

            bool ret = false;
            var excelApp = new Excel.Application();
            try
            {
                Excel.Workbook wb = excelApp.Workbooks.Open(filePath, 0, FileAccess.Read);
                excelApp.Visible = true;
                ret = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(filename + "파일 보기 실행중 파일이 작동중이므로 오류가 발생하엇습니다. 파일을 끄고 다시 실행시켜주십시오.");
                this.makeLogFile(ex.ToString());
            }
            return ret;
        }

        public void killExcelProcess()
        {
            System.Diagnostics.Process[] AfterExcelPsrocess;
            AfterExcelPsrocess = System.Diagnostics.Process.GetProcessesByName("Excel");
            for(int i = 0; i < AfterExcelPsrocess.Length; i++)
            {
                AfterExcelPsrocess[i].Kill();
            }
        }

        public bool makeFolder(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path); //폴더 창조
                }
                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(path + "폴더 생성하기 오류입니다.");
                this.makeLogFile(ex.ToString());
                return false;
            }
        }

        public void makeLogFile(string errorLog)
        {
            try
            {
                string relativePath = constant.rootPath + constant.logFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                this.makeFolder(fullPath);

                DateTime d = DateTime.Now;
                string text = "";
                if (File.Exists(fullPath + "\\log.log"))
                {
                    text = ReadTxtFile(fullPath + "\\log.log");
                    text += "\r\n" + d + " : " + errorLog;
                }
                else
                {
                    text = "\r\n" + d + " : " + errorLog;
                }
                this.MakeTxtFile(text, fullPath + "\\log.log");
            }
            catch(Exception)
            {
            }
        }

        public void filterSave(bool filter1, bool filter2, bool filter3)
        {
            try
            {
                string relativePath = constant.rootPath + constant.logFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = this.makeFolder(fullPath);
                if (!status) return;
                int one = 0;
                int two = 0;
                int three = 0;
                if (filter1) one = 1;
                if (filter2) two = 1;
                if (filter3) three = 1;
                string text = one + "," + two+ "," + three + ",";
                this.MakeTxtFile(text, fullPath + "\\filter.log");
            }
            catch(Exception ex)
            {
                MessageBox.Show("필터 체크 저장시 오류가 발생하엇습니다.");
                this.makeLogFile("필터 체크 저장시 오류가 발생하엇습니다.\n" + ex.ToString());
            }
        }

        public void MakeTxtFile(string text, string path)
        {
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine(text);
            }
        }
        public string ReadTxtFile(string path)
        {
            string line = "";
            string text = "";
            using (StreamReader sr = new StreamReader(path))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    text += line + "\r\n";
                }
            }
            return text;
        }

        public string MakeFullpath(string relativePath)
        {
            string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
            return fullPath;
        }

        public void copyWordFile(string path, string newPath)
        {
            if(File.Exists(newPath))
            {
                File.Delete(newPath);
                Thread.Sleep(100);
            }
            File.Copy(path, newPath);
            Thread.Sleep(100);
        }

        public void exchangeWordFile(word.Application createword, int index, int length, string similarword, string filepath)
        {
            word.Document createdoc = new word.Document();
            try
            {
                object fileName = filepath;
                object missing = System.Reflection.Missing.Value;
                // Define an object to pass to the API for missing parameters
                createdoc = createword.Documents.Open(ref fileName, ref missing, ReadOnly:false, missing, missing,
                    missing, missing, missing, missing, missing, missing, Visible: false, missing, missing, missing, missing);
                createword.Visible = false;
                Object start = index;
                Object end = index + length;
                word.Range range = createdoc.Range(ref start, ref end);
                if(range.Font.Bold == 0)
                {
                    range.Text = similarword;
                    range.Font.Bold = -1;
                }
                
            }
            catch(Exception ex)
            {
                //MessageBox.Show(filepath + "파일 보관시 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜주십시오.");
            }
            createdoc.Save();
            Thread.Sleep(100);
            ReleaseWordComObjects(null, createdoc);
            Thread.Sleep(100);
        }

        public bool wordFileView(string fullPath, string filename)
        {
            if (!File.Exists(fullPath + "\\" + filename))
            {
                MessageBox.Show(filename + "파일이 없습니다!");
                return false;
            }
            word.Application createword = new word.Application();
            word.Document createdoc = new word.Document();
            bool ret = false;
            try
            {
                object readOnly = true;
                object missing = System.Reflection.Missing.Value;
                createword.Documents.Open(fullPath + "\\" + filename, ref missing, readOnly, missing, missing,
                    missing, missing, missing, missing, missing, missing, Visible: true, missing, missing, missing, missing);
                createword.Visible = true;
                ret = true;
                ReleaseWordComObjects(createword, createdoc, false);
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                ReleaseWordComObjects(createword, createdoc);
                Thread.Sleep(100);
                MessageBox.Show(filename + "파일보기 실행중 프로세스가 작동중이므로 오류가 발생하엇습니다. 파일을 종료하고 다시 실행시켜주십시오.");
                this.makeLogFile("워드파일보기 오류!\n" + ex.ToString());
            }
            return ret;
        }

        private string RemoveSpace(string str)
        {
            if (string.IsNullOrEmpty(str))
                return str;

            str = str.Trim().Replace("\t", " ").Replace("\r", " ").Replace("\n", " ");
            str = (new System.Text.RegularExpressions.Regex(" +")).Replace(str, " ");

            return str;
        }
    }
}
