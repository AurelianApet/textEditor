using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace BasedOnText
{
    public partial class BasedOnText : Form
    {
        //패널마다 노출 위한 값
        public static int y_space = 20;
        public static int y_length = 25;
        public static int Rule_Ylength = 60;
        public static int Rule_Yspace = 50;
        public static int x_length1 = 10;
        public static int x_length2 = 30;
        public static int x_length3 = 230;
        public static int ruleBoxCount = 25;
        
        MainFile mainFile;
        TurchFile turchFile;
        Filter1 filter1 = new Filter1();
        Filter3 filter3 = new Filter3();
        Filter2 filter2;
        Common common = new Common();
        Constant constant = new Constant();

        //룰 세이브
        bool rule1check = false;
        bool rule2check = false;
        bool rule3check = false;
        bool rule4check = false;

        List<string> Rule2SpaceList = new List<string>();
        List<string> Rule2FileList = new List<string>();

        List<string> Rule3ListOrderingList = new List<string>();
        List<string> Rule3ParagraphList = new List<string>();
        List<string> Rule3SentenceList = new List<string>();
        List<string> Rule3FileName = new List<string>();

        List<string> Rule4SpaceList = new List<string>();
        List<string> Rule4FileList = new List<string>();

        public BasedOnText()
        {
            readRuldBoxSave();
            InitializeComponent();
            SetElementInRulePanel();
            mainFile = new MainFile(this);
            turchFile = new TurchFile(this);
            filter2 = new Filter2(this);
            readFilterCheckSave();
            readRuleCheckSave();
        }

        //원본파일 개개 삭제
        public void btnMainFileDelete_click(object sender, EventArgs e)
        {
            this.mainfilePanel.Controls.Clear();
            mainFile.FileDelete(sender, e);
        }

        //Rule2 통과시키기 이어서 보기
        private void btnRule2Connect_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "view";
            string filename = constant.TroughRule2FileConnectViewName;
            string state = "Rule2";
            bool status = mainFile.ConnectView(type, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("룰2 통과한 파일 이어서 보기 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("룰2 통과한 파일 이어서 보기 실패하엇습니다.");
            }
        }
        //Rule2 통과시키기 일괄 다운로드
        private void btnRule2AllDownload_Click(object sender, EventArgs e)
        {
            string state = "Rule2";
            this.mainFile.AllDownload(state);
        }
        //룰2 통과시키기 실행
        private void btnRule2Start_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            this.mainFile.Rule2Start(rule2check, Rule2SpaceList, Rule2FileList);
            loadingPanel.Visible = false;
        }

        //Filter1/Rule1 통과시키기 이어서 다운로드
        private void btnFilter2Rule1ConnectDownload_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "download";
            string filename = constant.TroughFilter2FileConnectViewName;
            string state = "Rule1";
            bool status = mainFile.ConnectView(type, filename, state);

            if (status == false)
            {
                loadingPanel.Visible = false;
                return;
            }

            string relativePath = constant.rootPath + constant.troughtFilter2Path;
            mainFile.ConnectViewDownload(relativePath, filename);

            loadingPanel.Visible = false;
            //MessageBox.Show("필터2/룰1파일 이어서 보기 다운로드에 성공하엇습니다.");
        }

        //룰1 필터2 통과하기 실행
        private void btnFilter2Start_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            mainFile.GetMainFileFilter2Text(this.CheckBoxFilter2.Checked, rule1check);
            loadingPanel.Visible = false;
        }
        //단어파일 생성하기
        private void btnMakeWordFile_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            mainFile.MakeWordFileStart();
            loadingPanel.Visible = false;
        }
        // Dano파일 Duplicate파일 수량
        public void DanoDuplicateWord(float danoCount, float TotalWord, float duplicateWord)
        {
            this.DanoNumberText.Text = Convert.ToString(danoCount);
            this.DuplicateWord.Text = Convert.ToString(duplicateWord);
            this.Rate.Text = Convert.ToString(duplicateWord / TotalWord);
        }

        //Filter0 다운로드
        private void btnFilter0FileDownload_Click(object sender, EventArgs e)
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + constant.filter0Filename))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, constant.filter0Filename);
                    //if (downstatus)
                    //    MessageBox.Show(constant.filter0Filename + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(constant.filter0Filename + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                    common.makeLogFile(constant.filter0Filename + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter0Filename + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter0Filename + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
        //유의어 찾기
        private void btnFindSimilarWord_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            try
            {
                string relativePath = constant.rootPath + constant.wordFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status)
                {
                    loadingPanel.Visible = false;
                    return;
                }

                if (File.Exists(fullPath + "\\" + constant.danoFilename))
                {
                    List<string> findSimilar = new List<string>();
                    List<string> ResultWord = new List<string>();
                    List<string> data = dictionaryLogin();
                    if (data.Count == 0) return;

                    Excel.Application xlApp = new Excel.Application();
                    findSimilar = common.ReadTextFromExcelFile(xlApp, fullPath + "\\" + constant.danoFilename);

                    string userId = data[0];
                    string key = data[1];

                    for (int i = 0; i < findSimilar.Count; i = i + 2)
                    {
                        if (findSimilar[i] == null || findSimilar[i] == "" ||  ResultWord.Contains(findSimilar[i]))
                            continue;
                        ResultWord.Add(findSimilar[i]);

                        var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://m.wordnet.co.kr/api/dic/search/pyojaeItems");
                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Method = "POST";

                        using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            string json = "{\"params\": {\"pageNum\":0," +
                                            "\"pageSize\":10," +
                                            "\"searchWord\":\"" + findSimilar[i] + "\"}}";

                            streamWriter.Write(json);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();
                            string[] array = result.Split('}');
                            //var check = false;
                            string Similarwords = "";
                            for (int j = 0; j < array.Length; j++)
                            {
                                int startword = array[j].IndexOf("WORD");
                                int endword = array[j].IndexOf("HANJA");
                                if (startword != -1 && endword != -1)
                                {
                                    string word = array[j].Substring(startword + 7, endword - startword - 10);
                                    if (word == findSimilar[i])
                                    {
                                        //check = true;
                                        int start = array[j].IndexOf("\"ID\"");
                                        int end = array[j].IndexOf("\"WORD\"");
                                        string id = array[j].Substring(start + 5, end - start - 6);
                                        if(Similarwords == "")
                                        {
                                            Similarwords = findSimilarWord(id, userId, key);
                                        }
                                        else
                                        {
                                            Similarwords = Similarwords + "," + findSimilarWord(id, userId, key);
                                        }
                                    }
                                }
                            }
                            if (Similarwords.Length > 0 && Similarwords[Similarwords.Length - 1] == ',')
                            {
                                Similarwords = Similarwords.Substring(0, Similarwords.Length - 1);
                            }
                            ResultWord.Add(Similarwords);
                            //if (!check)
                            //    ResultWord.Add("");
                        }
                    }
                    string relativePath1 = constant.rootPath + constant.filterFilePath;
                    string fullPath1 = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath1)));
                    if (!common.makeFolder(fullPath1))
                    {
                        common.ReleaseExcelComObjects(xlApp, null, null);
                        Thread.Sleep(100);
                        loadingPanel.Visible = false;
                        return;
                    }
                    common.MakeSeperateExcelFile(xlApp, ResultWord, fullPath1 + "\\" + constant.filter0Filename);
                    common.ReleaseExcelComObjects(xlApp, null, null);
                    Thread.Sleep(100);
                    loadingPanel.Visible = false;
                    MessageBox.Show("유의어 찾기가 완료되엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                    MessageBox.Show(constant.danoFilename + "파일이 없습니다!");
                }
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show("유의어 찾기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("유의어 찾기 실행시 오류.\n" + ex.ToString());
            }
        }

        private List<string> dictionaryLogin()
        {
            List<string> array = new List<string>();

            try
            {
                var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://m.wordnet.co.kr/api/user/login");
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    string json = "{\"email\":\"" + constant.dictionaryId + "\"," +
                                  "\"password\":\"" + constant.dictionaryPass + "\"}";

                    streamWriter.Write(json);
                    streamWriter.Flush();
                    streamWriter.Close();
                }

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    int start = result.IndexOf("id");
                    int end = result.IndexOf("userName");
                    if (start != -1 && end != -1)
                    {
                        string id = result.Substring(start + 4, end - start - 6);

                        int keystart = result.IndexOf("curKey");
                        int keyend = result.IndexOf("curIP");
                        string key = result.Substring(keystart + 9, keyend - keystart - 12);
                        array.Add(id);
                        array.Add(key);
                    }
                    return array;
                }
            }
            catch (Exception ex)
            {
                common.makeLogFile("낱말사전 로그인에 실패하엇습니다.");
                MessageBox.Show("낱말사전 로그인에 실패하엇습니다.");
                //array.Add("9502");
                //array.Add("b2duoVvtxDWHqE4");
                return array;
            }
        }

        private string findSimilarWord(string id, string userId, string key)
        {
            string similarWords = "";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://m.wordnet.co.kr/api/dic/search/wordSynItems");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                string json = "{\"params\":{\"userId\":" + userId + "," +
                              "\"pyojaeId\":\"" + id + "\"," +
                              "\"userAuthKey\":\"" + key + "\"}}";


                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();
            }

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
                string[] array = result.Split('}');
                for (int k = 0; k < array.Length; k++)
                {
                    if (array[k].Contains("비슷한말"))
                    {
                        int start = array[k].IndexOf("\"WORD\"");
                        int end = array[k].IndexOf('(');
                        if (start != -1 && end != -1)
                        {
                            string word = array[k].Substring(start + 8, end - start - 8);
                            similarWords += word + ",";
                        }
                    }
                }
                if (similarWords.Length > 0)
                {
                    similarWords = similarWords.Substring(0, similarWords.Length - 1);
                }
            }
            return similarWords;
        }

        //Duplicate 파일 보기
        private void btnDuplicateFileView_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string relativePath = constant.rootPath + constant.wordFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                if(common.makeFolder(fullPath) && common.FileView(fullPath + "\\" + constant.duplicateFilename))
                {
                    loadingPanel.Visible = false;
                    //MessageBox.Show(constant.duplicateFilename + "파일 보기에 성공하엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                    MessageBox.Show(constant.duplicateFilename + "파일 보기에 실패하엇습니다.");
                }
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(constant.duplicateFilename + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile(constant.duplicateFilename + "파일 보기 실행시 오류가 발생하엇습니다.");
            }
        }
        //Duplicate 다운로드
        private void btnDuplicateFileDownload_Click(object sender, EventArgs e)
        {
            try
            {
                string relativePath = constant.rootPath + constant.wordFile;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + constant.duplicateFilename))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, constant.duplicateFilename);
                    //if(downstatus)
                    //    MessageBox.Show(constant.duplicateFilename + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(constant.duplicateFilename + "파일이 존재하지 않습니다. 파일을 업로드하십시오.");
                    common.makeLogFile(constant.duplicateFilename + "파일이 존재하지 않습니다. 파일을 업로드하십시오.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.duplicateFilename + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(constant.duplicateFilename + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
        //원본파일 터치
        //원본파일 업로드

        private void btnFileUpload_Click(object sender, EventArgs e)
        {
            mainFile.FileUpload();
        }

        //원본 모든파일 일괄 삭제 
        private void btnMainFileAllDelete_Click(object sender, EventArgs e)
        {
            this.mainfilePanel.Controls.Clear(); //패널 초기화
            this.mainFile.FileAllDelete(); //업로드된 원본파일 초기화

        }

        //원본 모든 파일 이어서 보기 
        private void btnMainFileConnect_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            /*bool status =  */mainFile.FileConnectView(view);
            loadingPanel.Visible = false;
            //if (status)
            //{
            //    MessageBox.Show("원본파일 이어서 보기에 성공하엇습니다.");
            //}
        }

        //원본파일 이어서 보기 다운로드
        private void btnMainFileConnectDownload_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "download";
            if (mainFile.FileConnectView(type))
            {
                mainFile.FileConnectViewDownload();
                loadingPanel.Visible = false;
                //MessageBox.Show("원본파일 이어서보기 다운로드에 성공하엇습니다.");
            }
            else
            {
                loadingPanel.Visible = false;
                MessageBox.Show("원본파일 이어서보기 다운로드에 실패하엇습니다.");
            }
        }

        //원본파일리스트 노출
        public void mainFileListShow(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                
            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length2, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDelete = new Button();
            btnDelete.Text = "삭제";
            btnDelete.Name = Convert.ToString(Count);
            btnDelete.Location = new System.Drawing.Point(x_length3, y_length * Count - y_space);
            btnDelete.Click += new EventHandler(btnMainFileDelete_click);
            btnDelete.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));


            this.mainfilePanel.Controls.Add(numberLabel);
            this.mainfilePanel.Controls.Add(filenameLabel);
            this.mainfilePanel.Controls.Add(btnDelete);
        }
        //필터1 통과한 파일 일괄 삭제
        private void btnTroghtFilter1AllDelete_Click(object sender, EventArgs e)
        {
            this.mainFileFilter1Panel.Controls.Clear();
            this.mainFile.TroughFilter1ListAllDelete();

        }
        //필터1 통과한 파일 이어서 다운로드
        private void btnFilter1ConnectDownload_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "download";
            if (mainFile.Filter1ConnectView(type))
            {
                mainFile.Filter1ConnectViewDownload();
                loadingPanel.Visible = false;
            }
            else
            {
                loadingPanel.Visible = false;
                MessageBox.Show("필터1 통과한 파일 이어서 다운로드에 실패하엇습니다.");
            }
        }
        //필터1 통과한 파일 이어서 보기
        private void btnFilter1Connect_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            bool status = mainFile.Filter1ConnectView(view);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("필터1 통과한 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("필터1 통과한 파일 이어서 보기에 실패하엇습니다.");
            }
        }
        //필터1 통과한 파일 일괄 다운로드
        private void btnFilter1AllDownload_Click(object sender, EventArgs e)
        {
            this.mainFile.throughtFilter1AllDownload();
        }

        //Rule2 통과시키기 이어서 다운로드
        private void btnRule2ConnectDownload_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "download";
            string filename = constant.TroughRule2FileConnectViewName;
            string state = "Rule2";

            if (mainFile.ConnectView(type, filename, state))
            {
                string relativePath = constant.rootPath + constant.troughtRule2Path;
                mainFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
            }
            else
            {
                loadingPanel.Visible = false;
                MessageBox.Show("Rule2 통과시키기 이어서 다운로드에 실패하었습니다.");
            }
        }
        //룰2 통과시키기 리스트 일괄 삭제
        private void btnTroghtRule2AllDelete_Click(object sender, EventArgs e)
        {
            this.mainFileRule2Panel.Controls.Clear();
            this.mainFile.TroughRule2ListAllDelete();
        }
        //룰3 통과시키기 실행
        private void btnRule3Start_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            this.mainFile.Rule3Start(rule3check, Rule3ListOrderingList, Rule3ParagraphList, Rule3SentenceList, Rule3FileName);
            loadingPanel.Visible = false;
        }
        //Rule3 통과시키기 일괄 다운로드
        private void btnRule3AllDownload_Click(object sender, EventArgs e)
        {
            string state = "Rule3";
            this.mainFile.AllDownload(state);
            
        }
        //Rule3 통과시키기 이어서 보기
        private void btnRule3Connect_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "view";
            string filename = constant.TroughRule3FileConnectViewName;
            string state = "Rule3";
            bool status = mainFile.ConnectView(type, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("룰3 통과한 파일 이어서 보기 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("룰3 통과한 파일 이어서 보기 실패하엇습니다.");
            }
        }
        //Rule3 통과시키기 이어서 다운로드
        private void btnRule3ConnectDownload_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "download";
            string filename = constant.TroughRule3FileConnectViewName;
            string state = "Rule3";
            if (mainFile.ConnectView(type, filename, state))
            {
                string relativePath = constant.rootPath + constant.troughtRule3Path;
                mainFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
            }
            else
            {
                loadingPanel.Visible = false;
                MessageBox.Show("Rule3 통과시키기 이어서 다운로드에 실패하었습니다.");
            }
        }
        //룰3 통과시키기 리스트 일괄 삭제
        private void btnTroghtRule3AllDelete_Click(object sender, EventArgs e)
        {
            this.mainFileRule3Panel.Controls.Clear();
            this.mainFile.TroughRule3ListAllDelete();
            
        }

        //Filter1/Rule1 통과시키기 일괄 다운로드
        private void btnFilter2Rule1AllDownload_Click(object sender, EventArgs e)
        {
            string state = "Rule1";
            this.mainFile.AllDownload(state);

        }

        //Filter2/Rule1 통과시키기 이어서 보기
        private void btnFilter2Rule1Connect_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string type = "view";
            string filename = constant.TroughFilter2FileConnectViewName;
            string state = "Rule1";
            bool status = mainFile.ConnectView(type, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("룰1/필터2 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("룰1/필터2 파일 이어서 보기에 실패하엇습니다.");
            }
        }

        //룰1 필터2 통과하기 일괄 삭제
        private void btnTroghtFilter2AllDelete_Click(object sender, EventArgs e)
        {
            this.mainFileFilter2Panel.Controls.Clear();
            this.mainFile.TroughFilter2ListAllDelete();

        }
        //필터1 실행
        private void btnFilter1Start_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            mainFile.GetMainFileFilter1Text(CheckBoxFilter1.Checked);
            loadingPanel.Visible = false;
        }
        //필터 1 실행 패널 초기화.
        public void mainFileTrougphFilter1PanelInitialize()
        {
            try
            {
                this.mainFileFilter1Panel.Controls.Clear();
            }
            catch(Exception ex)
            {
                MessageBox.Show("필터1 패널 초기화시 오류가 발생하엇습니다.");
                common.makeLogFile("필터1 패널 초기화시 오류가 발생하엇습니다.\n" + ex.ToString());
            }
        }
        //필터1파일 업로드
        private void btnFilter1FileUpload_Click(object sender, EventArgs e)
        {
            filter1.FileUpload();
        }

        //Dano 파일 업로드
        private void btnDanoFileUpload_Click(object sender, EventArgs e)
        {
            try
            {
                string relativePath = constant.rootPath + constant.wordFile;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool uploadCheck = */common.FileOneUpload(relativePath, constant.excelExtension, constant.danoFilename);
                //if(uploadCheck)
                //{
                //    MessageBox.Show(constant.danoFilename + "파일 업로드에 성공하엇습니다.");
                //}
            }
            catch(Exception ex)
            {
                MessageBox.Show(constant.danoFilename +"파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("단어파일 업로드 오류.\n" + ex.ToString());
            }
        }
        //Dano 파일 다운로드
        private void btnDanoFileDownload_Click(object sender, EventArgs e)
        {
            try
            {
                string relativePath = constant.rootPath + constant.wordFile;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + constant.danoFilename))
                {
                    /*bool downstatus = */common.FileConnectViewDownload(fullPath, constant.danoFilename);
                    //if (downstatus)
                    //    MessageBox.Show(constant.danoFilename + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(constant.danoFilename + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.danoFilename + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile("단처파일 다운로드시  오류!\n" + ex.ToString());
            }
        }
        //Dano파일 보기
        private void btnDanoFileView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            try
            {
                string relativePath = constant.rootPath + constant.wordFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool viewStatus = */common.FileView(fullPath + "\\" + constant.danoFilename);
                loadingPanel.Visible = false;
                //if (viewStatus)
                //{
                //    MessageBox.Show(constant.danoFilename + "파일 보기에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(constant.danoFilename + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("단어파일 보기 오류!\n" + ex.ToString());
            }
        }

        //Filter0 업로드
        private void btnFilter0FileUpload_Click(object sender, EventArgs e)
        {
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool uploadCheck = */common.FileOneUpload(relativePath, constant.excelExtension, constant.filter0Filename);
                //if (uploadCheck)
                //{
                //    MessageBox.Show(constant.filter0Filename + "파일 업로드에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter0Filename + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("단어파일 업로드중 오류.\n" + ex.ToString());
            }
        }

        //Filter0 보기
        private void btnFilter0FileView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                if(common.makeFolder(fullPath) && common.FileView(fullPath + "\\" + constant.filter0Filename))
                {
                    loadingPanel.Visible = false;
                    //MessageBox.Show(constant.filter0Filename + "파일 보기에 성공하엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(constant.filter0Filename + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter0Filename + "파일 보기 실행시 오류가 발생하엇습니다.");
            }
        }
        //Filter0을 Filter2에 합치기
        private void btnAddFilter0ToFilter2_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (!File.Exists(fullPath + "\\" + constant.filter0Filename))
                {
                    loadingPanel.Visible = false;
                    MessageBox.Show(constant.filter0Filename + "파일이 없습니다! 파일을 업로드 하세요.");
                    return;
                }
                if (!File.Exists(fullPath + "\\" + constant.filter2Filename))
                {
                    loadingPanel.Visible = false;
                    MessageBox.Show(constant.filter2Filename + "파일이 없습니다! 파일을 업로드 하세요.");
                    return;
                }

                List<string> filter0WordList = new List<string>();
                List<string> filter2WordList = new List<string>();
                Excel.Application xlApp = new Excel.Application();

                filter0WordList = common.ReadTextFromExcelFile(xlApp, fullPath + "\\" + constant.filter0Filename);
                filter2WordList = common.ReadTextFromExcelFile(xlApp, fullPath + "\\" + constant.filter2Filename);
                for (int i = 0; i < filter0WordList.Count; i = i + 2)
                {
                    bool check = false;
                    for (int j = 0; j < filter2WordList.Count; j = j + 2)
                    {
                        if (filter0WordList[i] == filter2WordList[j]) check = true;
                    }
                    if (!check)
                    {
                        filter2WordList.Add(filter0WordList[i]);
                        filter2WordList.Add(filter0WordList[i + 1]);
                    }
                }
                common.MakeSeperateExcelFile(xlApp, filter2WordList, fullPath + "\\" + constant.filter2Filename);
                common.ReleaseExcelComObjects(xlApp, null, null);
                Thread.Sleep(100);
                MessageBox.Show(constant.filter0Filename + "파일을 " + constant.filter2Filename + "파일에 합치기 완료되엇습니다.");
            }
            catch (Exception ex)
            {
                common.makeLogFile("필터0파일을 필터 1파일에 합치기시 오류.\n" + ex.ToString());
                loadingPanel.Visible = false;
                MessageBox.Show(constant.filter0Filename + "파일을 " + constant.filter2Filename + "파일에 합치기 실패하엇습니다.");
            }
            loadingPanel.Visible = false;
        }
        //필터1 다운로드
        private void btnFilter1FileDownload_Click(object sender, EventArgs e)
        {
            this.filter1.fileDownload();
        }
        //필터1 파일 보기
        private void btnFilter1FileView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            try
            {
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool viewStatus = */common.FileView(fullPath + "\\" + constant.filter1Filename);
                //if (viewStatus)
                //{
                //    MessageBox.Show(constant.filter1Filename + "파일 보기에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(constant.filter1Filename + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile(constant.filter1Filename + "파일 보기 실행시 오류가 발생하엇습니다.");
            }
            loadingPanel.Visible = false;
        }
        //필터2 다운로드
        private void btnFilter2FileDownload_Click(object sender, EventArgs e)
        {
            this.filter2.fileDownload();

        }
        //필터2 파일 보기
        private void btnFilter2FileView_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                if (common.makeFolder(fullPath) && common.FileView(fullPath + "\\" + constant.filter2Filename))
                {
                    loadingPanel.Visible = false;
                    //MessageBox.Show(constant.filter2Filename + "파일 보기에 성공하엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(constant.filter2Filename + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("파일 보기 실행시 오류.\n" + ex.ToString());
            }
        }
        //필터2파일 업로드
        private void btnFilter2FileUpload_Click(object sender, EventArgs e)
        {
            filter2.FileUpload();
        }
        //필터3 다운로드
        private void btnFilter3FileDownload_Click(object sender, EventArgs e)
        {
            this.filter3.fileDownload();

        }
        // 필터3 파일 보기
        private void btnFilter3FileView_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string relativePath = constant.rootPath + constant.filterFilePath;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                if (common.makeFolder(fullPath) && common.FileView(fullPath + "\\" + constant.filter3Filename))
                {
                    loadingPanel.Visible = false;
                    //MessageBox.Show(constant.filter3Filename + "파일 보기에 성공하엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(constant.filter3Filename + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("파일 보기 실행시 오류.\n" + ex.ToString());
            }
        }
        //필터3 업로드
        private void btnFilter3Upload_Click(object sender, EventArgs e)
        {
            filter3.FileUpload();
        }
        //파일터치의 단어바꾸기 노출
        public void turchChangeWordShow(string filename, int Count)
        {
            try
            {
                Label numberLabel = new Label();
                numberLabel.AutoSize = true;
                numberLabel.Text = Convert.ToString(Count);
                numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
                numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                Label filenameLabel = new Label();
                filenameLabel.AutoSize = true;
                filenameLabel.Text = filename;
                filenameLabel.Location = new System.Drawing.Point(x_length2, y_length * Count - y_space);
                filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                this.TurchChangeWordPanel.Controls.Add(numberLabel);
                this.TurchChangeWordPanel.Controls.Add(filenameLabel);
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일(#2)코드추가 노출 오류!");
            }
        }
        //파일터치단어바꾸기패널 초기화.
        public void turchChangeWordPanelInitialize()
        {
            this.TurchChangeWordPanel.Controls.Clear();
        }

        //파일터치 글자 잇기 패널 초기화.
        public void turchConnectLetterInitialize()
        {
            this.TurchLetterConnectPanel.Controls.Clear();
        }

        //파일터치 코드 추가 패널 초기화.
        public void turchCodeAddInitialize()
        {
            this.TurchCodeAddPanel.Controls.Clear();
        }
        //원본파일 필터 2 실행 패널 초기화.
        public void mainFileTrougphFilter2PanelInitialize()
        {
            this.mainFileFilter2Panel.Controls.Clear();
        }

        //원본파일 룰2 실행 패널 초기화.
        public void mainFileTrougphRule2PanelInitialize()
        {
            this.mainFileRule2Panel.Controls.Clear();
        }

        //원본파일(#2) 룰4/ 글자 제거하기 실행 패널 초기화.
        public void TurchFileTrougphRule4PanelInitialize()
        {
            this.TurchDeleteLetterPanel.Controls.Clear();
        }

        //원본파일(#2) 룰4/ 글자 제거/잇기동시에 하기 실행 패널 초기화.
        public void TurchFileTrougphRule4ConnectPanelInitialize()
        {
            this.TurchDeleteLetterConnectPanel.Controls.Clear();
        }

        //원본파일 룰3 실행 패널 초기화.
        public void mainFileTrougphRule3PanelInitialize()
        {
            this.mainFileRule3Panel.Controls.Clear();
        }

        //파일터치의 원본파일 업로드
        private void btnTurchMainFileUpload_Click(object sender, EventArgs e)
        {
            this.turchFile.MainFileUpload();
        }
        //파일터치의 원본파일리스트 노출
        public void turchMainFileListShow(string filename, int Count)
        {
            try
            {
                Label numberLabel = new Label();
                numberLabel.AutoSize = true;
                numberLabel.Text = Convert.ToString(Count);
                numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
                numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                Label filenameLabel = new Label();
                filenameLabel.AutoSize = true;
                filenameLabel.Text = filename;
                filenameLabel.Location = new System.Drawing.Point(x_length2 + 70, y_length * Count - y_space);
                filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                this.TurchMainFilePanel.Controls.Add(numberLabel);
                this.TurchMainFilePanel.Controls.Add(filenameLabel);
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일(#2)리스트 노출 오류!");
            }
        }
        //원본파일(#2)의 글자 제거하기 및 잇기 실행
        private void bntTurchDeleteLetterConnectStart_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            this.turchFile.Rule4StartAndConnectLetter(rule4check, Rule4SpaceList, Rule4FileList);
            loadingPanel.Visible = false;
        }
        //파일터치의 원본파일(#2)이어서 보기
        private void btnTurchMainFileAllView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            bool status = turchFile.MainFile2ConnectView(view);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("원본파일(#2) 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("원본파일(#2) 이어서 보기에 실패하엇습니다.");
            }
        }
        //파일터치의 원본파일(#2)이어서 보기 다운로드
        private void btnTurchMainFileAllDownload_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string type = "download";
                bool status = turchFile.MainFile2ConnectView(type);
                if (status == false)
                {
                    loadingPanel.Visible = false;
                    return;
                }
                /*bool downStatus = */turchFile.MainFile2ConnectViewDownload();
                loadingPanel.Visible = false;
                //if (downStatus)
                //    MessageBox.Show("원본파일(#2)이어서 보기 다운로드에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show("원본파일(#2)의 이어서 보기 다운로드에 실패하엇습니다.");
                common.makeLogFile("원본파일(#2)의 이어서 보기 다운로드 오류.\n" + ex.ToString());
            }
        }
        //파일터치의 원본파일리스트 일괄삭제

        private void btnTurchMainFileAllDelete_Click(object sender, EventArgs e)
        {
            this.TurchMainFilePanel.Controls.Clear();
            this.turchFile.MainFileAllDelete();

        }
        //글자 제거/잇기 동시에 하기 일괄 다운로드
        private void btnTurchDeleteConnectLetterADownload_Click(object sender, EventArgs e)
        {
            string state = "deleteConnectLetter";
            this.turchFile.AllDownload(state);
        }
        //글자 제거/잇기 동시에 하기 이어서 보기
        private void btnTurchDeleteConnectLetterAllView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            string filename = constant.deleteLetterConnectLetterConnectViewName;
            string state = "deleteConnectLetter";
            bool status = turchFile.ConnectView(view, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("글자 제거하기/잇기 동시에하기 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("글자 제거하기/잇기 동시에하기 파일 이어서 보기에 실패하엇습니다.");
            }
        }
        //파일터치의 글자 제기/잇기 동시에 하기 이어서 다운로드
        private void btnTurchDeleteConnectLetterAllDownload_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string type = "download";
                string filename = constant.deleteLetterConnectLetterConnectViewName;
                string state = "deleteConnectLetter";
                bool status = turchFile.ConnectView(type, filename, state);

                if (status == false)
                {
                    loadingPanel.Visible = false;
                    return;
                }

                string relativePath = constant.rootPath + constant.deleteConnectFile;
                /*bool downStatus = */turchFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
                //if (downStatus)
                //    MessageBox.Show("글자 제거하기/잇기 파일 이어서 보기 다운로드에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show("글자 제거하기/잇기 파일 이어서 보기 다운로드에 실패하엇습니다.");
                common.makeLogFile("글자 제거하기/잇기 파일 이어서 보기 다운로드에 실패하엇습니다.");
            }
        }
        //원본파일(#2)의 글자 제거하기 및 잇기 일괄삭제
        private void btnTurchDeleteConnectLetterFileAllDelete_Click(object sender, EventArgs e)
        {
            this.TurchDeleteLetterConnectPanel.Controls.Clear();
            this.turchFile.DeleteConnectLetterAllDelete();

        }
        //파일터치의 코드 추가하기
        private void btnTurchCodeAdd_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            turchFile.CodeAdd();
            loadingPanel.Visible = false;
        }
        //파일터치의 코드추가 노출
        public void turchCodeAddShow(string filename, int Count)
        {
            try
            {
                Label numberLabel = new Label();
                numberLabel.AutoSize = true;
                numberLabel.Text = Convert.ToString(Count);
                numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
                numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                Label filenameLabel = new Label();
                filenameLabel.AutoSize = true;
                filenameLabel.Text = filename;
                filenameLabel.Location = new System.Drawing.Point(x_length2 + 70, y_length * Count - y_space);
                filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                this.TurchCodeAddPanel.Controls.Add(numberLabel);
                this.TurchCodeAddPanel.Controls.Add(filenameLabel);
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일(#2)코드추가 노출 오류!");
            }
        }

        //파일터치의 글자잇기
        private void btnTurchLetterConnect_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            turchFile.LetterConnect();
            loadingPanel.Visible = false;
        }
        //파일터치의 글자잇기 노출
        public void turchLetterConnectShow(string filename, int Count)
        {
            try
            {
                Label numberLabel = new Label();
                numberLabel.AutoSize = true;
                numberLabel.Text = Convert.ToString(Count);
                numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
                numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));


                Label filenameLabel = new Label();
                filenameLabel.AutoSize = true;
                filenameLabel.Text = filename;
                filenameLabel.Location = new System.Drawing.Point(x_length2 + 70, y_length * Count - y_space);
                filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));


                this.TurchLetterConnectPanel.Controls.Add(numberLabel);
                this.TurchLetterConnectPanel.Controls.Add(filenameLabel);
            }
            catch (Exception ex)
            {
                MessageBox.Show("원본파일(#2)글자잇기 노출 오류!");
            }
        }
        //파일터치의 글자제거하기 글자잇기  노출
        public void turchDeleteLetterConnectShow(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length2 + 10, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDownload = new Button();
            btnDownload.Text = "다운로드";
            btnDownload.Name = filename + ">" + Convert.ToString(Count);
            btnDownload.Location = new System.Drawing.Point(x_length1 + 250, y_length * Count - y_space);
            btnDownload.Click += new EventHandler(btnTurchFileDeleteConnectLetterDownload_click);
            btnDownload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnDownload.Size = new System.Drawing.Size(70, 20);

            this.TurchDeleteLetterConnectPanel.Controls.Add(numberLabel);
            this.TurchDeleteLetterConnectPanel.Controls.Add(filenameLabel);
            this.TurchDeleteLetterConnectPanel.Controls.Add(btnDownload);
        }
        //파일 터치의 글자 제거/잇기 동시에 하기 개개 다운로드
        public void btnTurchFileDeleteConnectLetterDownload_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name.Split('>')[0];
            try
            {
                string relativePath = constant.rootPath + constant.deleteConnectFile;
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
                    common.makeLogFile(fileName + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
        //글자 잇기 일괄 다운로드
        private void btnTurchConnectLetterADownload_Click(object sender, EventArgs e)
        {
            string state = "connectLetter";
            this.turchFile.AllDownload(state);
        }
        //글자 잇기 이어서 보기
        private void btnTurchConnectLetterAllView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            string filename = constant.connectLetterConnectViewName;
            string state = "connectLetter";
            bool status = turchFile.ConnectView(view, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("글자 잇기 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("글자 잇기 파일 이어서 보기에 실패하엇습니다.");
            }
        }
        //파일터치의 글자 잇기 이어서 다운로드
        private void btnTurchConnectLetterAllDownload_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string type = "download";
                string filename = constant.connectLetterConnectViewName;
                string state = "connectLetter";
                bool status = turchFile.ConnectView(type, filename, state);
                if (!status)
                {
                    loadingPanel.Visible = false;
                    return;
                }

                string relativePath = constant.rootPath + constant.connectLetterFile;
                /*bool downStatus = */turchFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
                //if (downStatus)
                //    MessageBox.Show("글자 잇기 파일 이어서 보기 다운로드에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                MessageBox.Show("글자 잇기 파일 이이서 보기 다운로드에 실패하엇습니다.");
                common.makeLogFile("글자 잇기 파일 이이서 보기 다운로드에 실패하엇습니다.");
            }
        }
        //파일터치의 글자잇기 일괄삭제

        private void btnTurchLetterConnectAllDelete_Click(object sender, EventArgs e)
        {
            this.TurchLetterConnectPanel.Controls.Clear();
            this.turchFile.LetterConnectAllDelete();
        }
        //코드 추가 일괄 다운로드
        private void btnTurchAddCodeADownload_Click(object sender, EventArgs e)
        {
            string state = "addCode";
            this.turchFile.AllDownload(state);
        }
        //코드 추가 이어서 보기
        private void btnTurchAddCodeAllView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            string filename = constant.addCodeConnectViewName;
            string state = "addCode";
            bool status = turchFile.ConnectView(view, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("코드 추가 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("코드 추가 파일 이어서 보기에 실패하엇습니다.");
            }
        }
        //파일터치의 코드 추가 이어서 다운로드
        private void btnTurchAddCodeAllDownload_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string type = "download";
                string filename = constant.addCodeConnectViewName;
                string state = "addCode";
                bool status = turchFile.ConnectView(type, filename, state);

                if (status == false)
                {
                    loadingPanel.Visible = false;
                    return;
                }

                string relativePath = constant.rootPath + constant.addCode;
                /*bool downStatus = */turchFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
                //if (downStatus)
                //    MessageBox.Show("코드 추가 파일 이어서 보기 다운로드에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show("코드 추가 파일 이어서 보기 다운로드에 실패하엇습니다.");
                common.makeLogFile("코드 추가 파일 이어서 보기 다운로드에 실패하엇습니다.");
            }
        }
        //파일터치에서 코드추가파일 일괄삭제
        private void btnCodeAddFileAllDelete_Click(object sender, EventArgs e)
        {
            this.TurchCodeAddPanel.Controls.Clear();
            this.turchFile.CodeAddAllDelete();
        }
        //파일터치의 글자 제거하기 실행
        private void bntTurchDeleteLetterStart_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            this.turchFile.Rule4Start(rule4check, Rule4SpaceList, Rule4FileList);
            loadingPanel.Visible = false;
        }
        //파일터치에서 단어바꾸기 실행
        private void btnTurchChangeWord_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            this.turchFile.WordChangeStart(this.CheckBoxFilter3.Checked);
            loadingPanel.Visible = false;
        }

        //글자 제거하기 일괄 다운로드
        private void btnTurchDeleteLetterADownload_Click(object sender, EventArgs e)
        {
            string state = "deleteLetter";
            this.turchFile.AllDownload(state);
        }
        //글자 제거하기 이어서 보기
        private void btnTurchDeleteLetterAllView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            string filename = constant.deleteLetterConnectViewName;
            string state = "deleteLetter";
            bool status = turchFile.ConnectView(view, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("글자 제거하기 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("글자 제거하기 파일 이어서 보기에 실패하엇습니다.");
            }
        }
        //파일터치의 글자 제거하기 이어서 다운로드
        private void btnTurchDeleteLetterAllDownload_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string type = "download";
                string filename = constant.deleteLetterConnectViewName;
                string state = "deleteLetter";
                bool status = turchFile.ConnectView(type, filename, state);

                if (status == false)
                {
                    loadingPanel.Visible = false;
                    return;
                }

                string relativePath = constant.rootPath + constant.deleteFile;
                /*bool downCheck = */turchFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
                //if (downCheck)
                //    MessageBox.Show("글자 제거하기 파일 이어서 보기 다운로드에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show("글자 제거하기 파일 이어서 보기 다운로드에 실패하엇습니다.");
                common.makeLogFile("글자 제거하기 파일 이어서 보기 다운로드에 실패하엇습니다.");
            }
        }
        //파일터치의 글자 제거하기 일괄 삭제
        private void btnTurchDeleteLetterFileAllDelete_Click(object sender, EventArgs e)
        {
            this.TurchDeleteLetterPanel.Controls.Clear();
            this.turchFile.DeleteLetterAllDelete();

        }
        //룰2 룰3 룰4에 25개 박스 넣기
        private void SetElementInRulePanel()
        {
            for (int i = 1; i <= 25; i++)
            {
                //룰2 패널 추가
                TextBox r2orderLabel = new TextBox();
                r2orderLabel.ReadOnly = true;
                r2orderLabel.BorderStyle = default;
                r2orderLabel.AutoSize = true;
                r2orderLabel.Text = Convert.ToString(i);
                r2orderLabel.Location = new System.Drawing.Point(x_length1 + 30, Rule_Ylength * i - Rule_Yspace);
                r2orderLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                TextBox Rule2Space = new TextBox();
                Rule2Space.Name = "Rule2Space" + Convert.ToString(i);
                Rule2Space.Size = new System.Drawing.Size(200, 55);
                Rule2Space.Location = new System.Drawing.Point(x_length1 + 180, Rule_Ylength * i - Rule_Yspace);
                Rule2Space.Multiline = true;
                Rule2Space.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule2Space.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule2Space.Text = Rule2SpaceList[i - 1];

                TextBox Rule2File = new TextBox();
                Rule2File.Name = "Rule2File" + Convert.ToString(i);
                Rule2File.Size = new System.Drawing.Size(55, 55);
                Rule2File.Location = new System.Drawing.Point(x_length1 + 530, Rule_Ylength * i - Rule_Yspace);
                Rule2File.Multiline = true;
                Rule2File.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule2File.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule2File.Text = Rule2FileList[i - 1];

                this.Rule2Panel.Controls.Add(r2orderLabel);
                this.Rule2Panel.Controls.Add(Rule2Space);
                this.Rule2Panel.Controls.Add(Rule2File);

                //룰3 패널 추가
                TextBox r3orderLabel = new TextBox();
                r3orderLabel.ReadOnly = true;
                r3orderLabel.BorderStyle = default;
                r3orderLabel.Size = new System.Drawing.Size(20, 100);
                r3orderLabel.Text = Convert.ToString(i);
                r3orderLabel.Location = new System.Drawing.Point(x_length1, Rule_Ylength * i - Rule_Yspace);
                r3orderLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                TextBox Rule3ListOrdering = new TextBox();
                Rule3ListOrdering.Name = "Rule3ListOrdering" + Convert.ToString(i);
                Rule3ListOrdering.Size = new System.Drawing.Size(155, 55);
                Rule3ListOrdering.Location = new System.Drawing.Point(x_length1 + 50, Rule_Ylength * i - Rule_Yspace);
                Rule3ListOrdering.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule3ListOrdering.Multiline = true;
                Rule3ListOrdering.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule3ListOrdering.Text = Rule3ListOrderingList[i - 1];

                TextBox Rule3SeleteParagraph = new TextBox();
                Rule3SeleteParagraph.Name = "Rule3SeleteParagraph" + Convert.ToString(i);
                Rule3SeleteParagraph.Size = new System.Drawing.Size(155, 55);
                Rule3SeleteParagraph.Location = new System.Drawing.Point(x_length1 + 210, Rule_Ylength * i - Rule_Yspace);
                Rule3SeleteParagraph.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule3SeleteParagraph.Multiline = true;
                Rule3SeleteParagraph.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule3SeleteParagraph.Text = Rule3ParagraphList[i - 1];

                TextBox Rule3ListSentence = new TextBox();
                Rule3ListSentence.Name = "Rule3ListSentence" + Convert.ToString(i);
                Rule3ListSentence.Size = new System.Drawing.Size(155, 55);
                Rule3ListSentence.Location = new System.Drawing.Point(x_length1 + 370, Rule_Ylength * i - Rule_Yspace);
                Rule3ListSentence.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule3ListSentence.Multiline = true;
                Rule3ListSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule3ListSentence.Text = Rule3SentenceList[i - 1];

                var nowDate = DateTime.Now;
                string filename = "FN" + nowDate.Year + nowDate.Month + nowDate.Day + "_" + Convert.ToString(i);
                TextBox Rule3FileName = new TextBox();
                Rule3FileName.Name = "Rule3Text" + Convert.ToString(i);
                Rule3FileName.ReadOnly = true;
                Rule3FileName.BorderStyle = default;
                Rule3FileName.AutoSize = true;
                Rule3FileName.Text = filename;
                Rule3FileName.Location = new System.Drawing.Point(x_length1 + 540, Rule_Ylength * i - Rule_Yspace);
                Rule3FileName.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                this.Rule3Panel.Controls.Add(r3orderLabel);
                this.Rule3Panel.Controls.Add(Rule3ListOrdering);
                this.Rule3Panel.Controls.Add(Rule3FileName);
                this.Rule3Panel.Controls.Add(Rule3SeleteParagraph);
                this.Rule3Panel.Controls.Add(Rule3ListSentence);

                //룰4 패널 추가
                TextBox r4orderLabel = new TextBox();
                r4orderLabel.ReadOnly = true;
                r4orderLabel.BorderStyle = default;
                r4orderLabel.Size = new System.Drawing.Size(20, 100);
                r4orderLabel.Text = Convert.ToString(i);
                r4orderLabel.Location = new System.Drawing.Point(x_length1 + 30, Rule_Ylength * i - Rule_Yspace);
                r4orderLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                TextBox Rule4Space = new TextBox();
                Rule4Space.Name = "Rule4Space" + Convert.ToString(i);
                Rule4Space.Size = new System.Drawing.Size(200, 55);
                Rule4Space.Location = new System.Drawing.Point(x_length1 + 180, Rule_Ylength * i - Rule_Yspace);
                Rule4Space.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule4Space.Multiline = true;
                Rule4Space.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule4Space.Text = Rule4SpaceList[i - 1];

                TextBox Rule4File = new TextBox();
                Rule4File.Name = "Rule4File" + Convert.ToString(i);
                Rule4File.Size = new System.Drawing.Size(55, 55);
                Rule4File.Location = new System.Drawing.Point(x_length1 + 530, Rule_Ylength * i - Rule_Yspace);
                Rule4File.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Rule4File.Multiline = true;
                Rule4File.ScrollBars = System.Windows.Forms.ScrollBars.Both;
                Rule4File.Text = Rule4FileList[i - 1];
                
                this.Rule4Panel.Controls.Add(r4orderLabel);
                this.Rule4Panel.Controls.Add(Rule4Space);
                this.Rule4Panel.Controls.Add(Rule4File);
            }
        }
        //단어 바꾸기 일괄 다운로드
        private void btnTurchChangeWordADownload_Click(object sender, EventArgs e)
        {
            string state = "changeWord";
            this.turchFile.AllDownload(state);
        }
        //단어 바꾸기 이어서 보기
        private void btnTurchChangeWordAllView_Click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            string view = "view";
            string filename = constant.changeWordConnectViewName;
            string state = "changeWord";
            bool status = turchFile.ConnectView(view, filename, state);
            loadingPanel.Visible = false;
            if (status)
            {
                //MessageBox.Show("단어 바꾸기 파일 이어서 보기에 성공하엇습니다.");
            }
            else
            {
                MessageBox.Show("단어 바꾸기 파일 이어서 보기에 실패하엇습니다.");
            }
        }
        //룰 세이브클릭
        private void btnRuleSave_Click(object sender, EventArgs e)
        {
            foreach (TextBox component in this.Rule2Panel.Controls)
            {

            }
        }
        //원본파일의 필터1 실행노출
        public void mainFileTroughFilter1Show(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length1 + 20, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDownload = new Button();
            btnDownload.Text = "다운로드";
            btnDownload.Name = filename + ">" + Convert.ToString(Count);
            btnDownload.Size = new System.Drawing.Size(70, 20);
            btnDownload.Location = new System.Drawing.Point(x_length1 + 180, y_length * Count - y_space);
            btnDownload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnDownload.Click += new EventHandler(btnMainFileFilter1Download_click);

            Button btnUpload = new Button();
            btnUpload.Text = "업로드";
            btnUpload.Name = filename + ">" + Convert.ToString(Count);
            btnUpload.Location = new System.Drawing.Point(x_length1 + 250, y_length * Count - y_space);
            btnUpload.Click += new EventHandler(btnMainFileFilter1Upload_click);
            btnUpload.Size = new System.Drawing.Size(70, 20);
            btnUpload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            this.mainFileFilter1Panel.Controls.Add(numberLabel);
            this.mainFileFilter1Panel.Controls.Add(filenameLabel);
            this.mainFileFilter1Panel.Controls.Add(btnDownload);
            this.mainFileFilter1Panel.Controls.Add(btnUpload);
        }

        //원본파일의 필터2 실행노출
        public void mainFileTroughFilter2Show(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length1 + 20, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDownload = new Button();
            btnDownload.Text = "다운로드";
            btnDownload.Name = filename + ">" + Convert.ToString(Count);
            btnDownload.Location = new System.Drawing.Point(x_length1 + 220, y_length * Count - y_space);
            btnDownload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnDownload.Click += new EventHandler(btnRule1Download_click);
            btnDownload.Size = new System.Drawing.Size(70, 20);

            Button btnView = new Button();
            btnView.Text = "보기";
            btnView.Name = filename + ">" + Convert.ToString(Count);
            btnView.Location = new System.Drawing.Point(x_length1 + 290, y_length * Count - y_space);
            btnView.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnView.Click += new EventHandler(btnRule1View_click);
            btnView.Size = new System.Drawing.Size(70, 20);

            Button btnUpload = new Button();
            btnUpload.Text = "업로드";
            btnUpload.Name = filename + ">" + Convert.ToString(Count);
            btnUpload.Location = new System.Drawing.Point(x_length1 + 360, y_length * Count - y_space);
            btnUpload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnUpload.Click += new EventHandler(btnRule1Upload_click);
            btnUpload.Size = new System.Drawing.Size(70, 20);

            this.mainFileFilter2Panel.Controls.Add(numberLabel);
            this.mainFileFilter2Panel.Controls.Add(filenameLabel);
            this.mainFileFilter2Panel.Controls.Add(btnDownload);
            this.mainFileFilter2Panel.Controls.Add(btnView);
            this.mainFileFilter2Panel.Controls.Add(btnUpload);
        }
        public void btnRule1Upload_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                //fileName = Path.GetFileNameWithoutExtension(fileName);
                string relativePath = constant.rootPath + constant.troughtFilter2Path;
                string fullPath = common.MakeFullpath(relativePath);
                if(common.makeFolder(fullPath) && common.FileOneUpload(relativePath, constant.docxExtension, fileName))
                {
                    //MessageBox.Show(fileName + "파일 업로드에 성공하엇습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("파일 업로드중 오류!\n" + ex.ToString());
            }
        }
        //원본파일의 룰2 실행노출
        public void mainFileTroughRule2Show(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length1 + 20, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDownload = new Button();
            btnDownload.Text = "다운로드";
            btnDownload.Name = filename + ">" + Convert.ToString(Count);
            btnDownload.Location = new System.Drawing.Point(x_length1 + 180, y_length * Count - y_space);
            btnDownload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnDownload.Click += new EventHandler(btnRule2Download_click);
            btnDownload.Size = new System.Drawing.Size(70, 20);

            Button btnView = new Button();
            btnView.Text = "보기";
            btnView.Name = filename + ">" + Convert.ToString(Count);
            btnView.Location = new System.Drawing.Point(x_length1 + 250, y_length * Count - y_space);
            btnView.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnView.Click += new EventHandler(btnRule2View_click);
            btnView.Size = new System.Drawing.Size(70, 20);

            Button btnUpload = new Button();
            btnUpload.Text = "업로드";
            btnUpload.Name = filename + ">" + Convert.ToString(Count);
            btnUpload.Location = new System.Drawing.Point(x_length1 + 320, y_length * Count - y_space);
            btnUpload.Click += new EventHandler(btnRule2Upload_click);
            btnUpload.Size = new System.Drawing.Size(70, 20);
            btnUpload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDelete = new Button();
            btnDelete.Text = "삭제";
            btnDelete.Name = Convert.ToString(Count);
            btnDelete.Location = new System.Drawing.Point(x_length1 + 390, y_length * Count - y_space);
            btnDelete.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnDelete.Click += new EventHandler(btnRule2Delete_click);
            btnDelete.Size = new System.Drawing.Size(70, 20);

            this.mainFileRule2Panel.Controls.Add(numberLabel);
            this.mainFileRule2Panel.Controls.Add(filenameLabel);
            this.mainFileRule2Panel.Controls.Add(btnDownload);
            this.mainFileRule2Panel.Controls.Add(btnView);
            this.mainFileRule2Panel.Controls.Add(btnUpload);
            this.mainFileRule2Panel.Controls.Add(btnDelete);
        }
        //파일터치의 단어 바꾸기 이어서 다운로드
        private void btnTurchChangeWordAllDownload_Click(object sender, EventArgs e)
        {
            try
            {
                loadingPanel.Visible = true;
                string type = "download";
                string filename = constant.changeWordConnectViewName;
                string state = "changeWord";
                bool status = turchFile.ConnectView(type, filename, state);

                if (!status)
                {
                    loadingPanel.Visible = false;
                    return;
                }

                string relativePath = constant.rootPath + constant.changeWordFile;
                /*bool downStatus = */turchFile.ConnectViewDownload(relativePath, filename);
                loadingPanel.Visible = false;
                //if (downStatus)
                //    MessageBox.Show("단어바꾸기 파일 이어서 보기 다운로드에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show("단어바꾸기 파일 이어서 보기 다운로드에 실패하엇습니다.");
                common.makeLogFile("단어바꾸기 파일 이어서 보기 다운로드 오류.\n" + ex.ToString());
            }
        }
        //룰2 통과한 파일 다운로드
        private void btnRule2Download_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.troughtRule2Path;
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
                    common.makeLogFile(fileName + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile("파일 다운로드시  오류!\n" + ex.ToString());
            }
        }
        //룰2 통과한 파일 보기
        private void btnRule2View_click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.troughtRule2Path;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                /*bool viewStatus = */common.wordFileView(fullPath, fileName);
                loadingPanel.Visible = false;
                //if (viewStatus)
                //{
                //    MessageBox.Show(fileName + "파일 보기에 성공하엇습니다.");
                //}
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(fileName + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("파일 보기 실행시 오류.\n" + ex.ToString());
            }
        }
        //룰2 통과한 파일 업로드
        private void btnRule2Upload_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                //fileName = Path.GetFileNameWithoutExtension(fileName);
                string relativePath = constant.rootPath + constant.troughtRule2Path;
                string fullPath = common.MakeFullpath(relativePath);
                if(common.makeFolder(fullPath) && common.FileOneUpload(relativePath, constant.docxExtension, fileName))
                {
                    //MessageBox.Show(fileName + "파일 업로드에 성공하엇습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 업로드중 프로세스가 작동중이므로 오류가 발생하엇습니다.");
                common.makeLogFile("파일 업로드중 오류\n" + ex.ToString());
            }
        }
        //룰2 통과한 파일 삭제
        private void btnRule2Delete_click(object sender, EventArgs e)
        {
            try
            {
                this.mainFileRule2Panel.Controls.Clear();
                mainFile.Rule2FileDelete(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 삭제 오류!");
                common.makeLogFile("파일 삭제 오류\n" + ex.ToString());
            }
        }
        //원본파일(#2)의 룰4/글자 지우기 실행노출
        public void turchRule4DeleteWordShow(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length1 + 50, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDownload = new Button();
            btnDownload.Text = "다운로드";
            btnDownload.Name = filename + ">" + Convert.ToString(Count);
            btnDownload.Location = new System.Drawing.Point(x_length1 + 230, y_length * Count - y_space);
            btnDownload.Click += new EventHandler(btnTurchFileDeleteLetterDownload_click);
            btnDownload.Size = new System.Drawing.Size(70, 20);
            btnDownload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            this.TurchDeleteLetterPanel.Controls.Add(numberLabel);
            this.TurchDeleteLetterPanel.Controls.Add(filenameLabel);
            this.TurchDeleteLetterPanel.Controls.Add(btnDownload);
        }
        //파일 터치의 글자 제거하기 개개 다운로드
        public void btnTurchFileDeleteLetterDownload_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.deleteFile;
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
                    common.makeLogFile(fileName + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }

        //원본파일의 룰3 실행노출
        public void mainFileTroughRule3Show(string filename, int Count)
        {
            Label numberLabel = new Label();
            numberLabel.AutoSize = true;
            numberLabel.Text = Convert.ToString(Count);
            numberLabel.Location = new System.Drawing.Point(x_length1, y_length * Count - y_space);
            numberLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Label filenameLabel = new Label();
            filenameLabel.AutoSize = true;
            filenameLabel.Text = filename;
            filenameLabel.Location = new System.Drawing.Point(x_length1 + 20, y_length * Count - y_space);
            filenameLabel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            Button btnDownload = new Button();
            btnDownload.Text = "다운로드";
            btnDownload.Name = filename + ">" + Convert.ToString(Count);
            btnDownload.Location = new System.Drawing.Point(x_length1 + 180, y_length * Count - y_space);
            btnDownload.Size = new System.Drawing.Size(70, 20);
            btnDownload.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnDownload.Click += new EventHandler(btnRule3Download_click);

            Button btnView = new Button();
            btnView.Text = "보기";
            btnView.Name = filename + ">" + Convert.ToString(Count);
            btnView.Location = new System.Drawing.Point(x_length1 + 250, y_length * Count - y_space);
            btnView.Size = new System.Drawing.Size(70, 20);
            btnView.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnView.Click += new EventHandler(btnRule3View_click);

            this.mainFileRule3Panel.Controls.Add(numberLabel);
            this.mainFileRule3Panel.Controls.Add(filenameLabel);
            this.mainFileRule3Panel.Controls.Add(btnDownload);
            this.mainFileRule3Panel.Controls.Add(btnView);
        }
        // 룰3 통과한 파일 다운로드
        private void btnRule3Download_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.troughtRule3Path;
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
                    common.makeLogFile(fileName + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
            }
        }
        //룰 3 통과한 파일 보기
        private void btnRule3View_click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.troughtRule3Path;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                if (common.makeFolder(fullPath) && common.wordFileView(fullPath, fileName))
                {
                    loadingPanel.Visible = false;
                    //MessageBox.Show(fileName + "파일 보기에 성공하엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                loadingPanel.Visible = false;
                MessageBox.Show(fileName + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile(fileName + "파일 보기 실행시 오류가 발생하엇습니다.");
            }
        }

        //필터1 통과한 파일 업로드
        public void btnMainFileFilter1Upload_click(object sender, EventArgs e)
        {
            this.filter1.throughtFilter1FileUpload(sender, e);
        }

        //필터1 통과한 파일 다운로드
        public void btnMainFileFilter1Download_click(object sender, EventArgs e)
        {
            this.filter1.throughtFilter1FileDownload(sender, e);
        }
        //필터2 통과한 파일 업로드
        public void btnRule1Download_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileName = btn.Name;
            fileName = fileName.Split('>')[0];
            try
            {
                string relativePath = constant.rootPath + constant.troughtFilter2Path;
                string fullPath = common.MakeFullpath(relativePath);
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                if (File.Exists(fullPath + "\\" + fileName))
                {
                    common.FileConnectViewDownload(fullPath, fileName);
                        //MessageBox.Show(fileName + "파일 다운로드에 성공하엇습니다.");
                }
                else
                {
                    MessageBox.Show(fileName + "파일이 존재하지 않습니다. 파일을 업로드 하세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 다운로드시  오류가 발생하엇습니다.");
                common.makeLogFile("파일 다운로드시 오류\n" + ex.ToString());
            }

        }
        //필터2 통과한 파일 보기
        public void btnRule1View_click(object sender, EventArgs e)
        {
            loadingPanel.Visible = true;
            Button btn = sender as Button;
            string fileName = btn.Name;
            try
            {
                fileName = fileName.Split('>')[0];
                string relativePath = constant.rootPath + constant.troughtFilter2Path;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                if (common.makeFolder(fullPath) && common.wordFileView(fullPath, fileName))
                {
                    loadingPanel.Visible = false;
                    //MessageBox.Show(fileName + "파일 보기에 성공하엇습니다.");
                }
                else
                {
                    loadingPanel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(fileName + "파일 보기 실행시 오류가 발생하엇습니다.");
                common.makeLogFile("파일 보기 실행시 오류.\n" + ex.ToString());
            }
        }

        //파일터치 단어바꾸기 일괄 삭제
        private void btnTurchChangeWordAllDelete_Click(object sender, EventArgs e)
        {
            this.TurchChangeWordPanel.Controls.Clear();
            this.turchFile.ChangeWordAllDelete();
        }

        //룰 세이브
        private void btnSaveRule_Click(object sender, EventArgs e)
        {
            Rule2SpaceList.RemoveRange(0, Rule2SpaceList.Count);
            Rule2FileList.RemoveRange(0, Rule2FileList.Count);

            Rule3ListOrderingList.RemoveRange(0, Rule3ListOrderingList.Count);
            Rule3ParagraphList.RemoveRange(0, Rule3ParagraphList.Count);
            Rule3SentenceList.RemoveRange(0, Rule3SentenceList.Count);
            Rule3FileName.RemoveRange(0, Rule3FileName.Count);

            Rule4SpaceList.RemoveRange(0, Rule4SpaceList.Count);
            Rule4FileList.RemoveRange(0, Rule4FileList.Count);
            
            try
            {
                if (Rule1CheckBox.Checked)
                {
                    rule1check = true;
                }
                else rule1check = false;

                if (Rule2CheckBox.Checked)
                {
                    rule2check = true;
                }
                else rule2check = false;

                foreach (TextBox component in this.Rule2Panel.Controls)
                {
                    if (component.Name.ToString().Contains("Rule2Space"))
                    {
                        Rule2SpaceList.Add(component.Text);
                    }
                    else if (component.Name.ToString().Contains("Rule2File"))
                    {
                        Rule2FileList.Add(component.Text);
                    }
                }

                if (Rule3CheckBox.Checked)
                {
                    rule3check = true;
                }
                else rule3check = false;

                foreach (TextBox component in this.Rule3Panel.Controls)
                {
                    if (component.Name.ToString().Contains("Rule3ListOrdering"))
                    {
                        Rule3ListOrderingList.Add(component.Text);
                    }
                    else if (component.Name.ToString().Contains("Rule3SeleteParagraph"))
                    {
                        Rule3ParagraphList.Add(component.Text);
                    }
                    else if (component.Name.ToString().Contains("Rule3ListSentence"))
                    {
                        Rule3SentenceList.Add(component.Text);
                    }
                    else if (component.Name.ToString().Contains("Rule3Text"))
                    {
                        Rule3FileName.Add(component.Text);
                    }
                }

                if (Rule4CheckBox.Checked)
                {
                    rule4check = true;
                }
                else rule4check = false;

                foreach (TextBox component in this.Rule4Panel.Controls)
                {
                    if (component.Name.ToString().Contains("Rule4Space"))
                    {
                        Rule4SpaceList.Add(component.Text);
                    }
                    else if (component.Name.ToString().Contains("Rule4File"))
                    {
                        Rule4FileList.Add(component.Text);
                    }
                }

                string relativePath = constant.rootPath + constant.logFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                bool status = common.makeFolder(fullPath);
                if (!status) return;

                string text = rule1check + ";" + rule2check + ";" + rule3check + ";" + rule4check + ";\r";
                for(int i = 0; i < Rule2FileList.Count; i ++)
                {
                    text += Rule2FileList[i] + ";" + Rule2SpaceList[i] + ";\n";
                }
                for (int i = 0; i < Rule3ListOrderingList.Count; i++)
                {
                    text += Rule3ListOrderingList[i] + ";" + Rule3ParagraphList[i] + ";" + Rule3SentenceList[i] + ";" + Rule3FileName[i] + ";\n";
                }
                for (int i = 0; i < Rule4FileList.Count; i++)
                {
                    text += Rule4FileList[i] + ";" + Rule4SpaceList[i] + ";\n";
                }
                common.MakeTxtFile(text, fullPath + "\\rule.log");
                MessageBox.Show("룰 세브에 성공하엇습니다.");
            }
            catch(Exception ex)
            {
                MessageBox.Show("룰 세브시 오류가 발생하엇습니다.");
                common.MakeFullpath("룰 세브시 오류\n" + ex.ToString());
            }
        }

        private void CheckBoxFilter1_CheckedChanged(object sender, EventArgs e)
        {
            common.filterSave(CheckBoxFilter1.Checked, CheckBoxFilter2.Checked, CheckBoxFilter3.Checked);
        }

        private void CheckBoxFilter2_CheckedChanged(object sender, EventArgs e)
        {
            common.filterSave(CheckBoxFilter1.Checked, CheckBoxFilter2.Checked, CheckBoxFilter3.Checked);
        }

        private void CheckBoxFilter3_CheckedChanged(object sender, EventArgs e)
        {
            common.filterSave(CheckBoxFilter1.Checked, CheckBoxFilter2.Checked, CheckBoxFilter3.Checked);
        }

        public void readFilterCheckSave()
        {
            try
            {
                string relativePath = constant.rootPath + constant.logFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                common.makeFolder(fullPath);

                if (File.Exists(fullPath + "\\filter.log"))
                {
                    string text = common.ReadTxtFile(fullPath + "\\filter.log");
                    string[] textArray = text.Split(',');
                    if (textArray[0] == "1") this.CheckBoxFilter1.Checked = true;
                    if (textArray[1] == "1") this.CheckBoxFilter2.Checked = true;
                    if (textArray[2] == "1") this.CheckBoxFilter3.Checked = true;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("저장된 필터 체크 읽기시 오류가 발생하엇습니다.");
                common.makeLogFile("저장된 필터 체크 읽기시 오류!\n" + ex.ToString());
            }
        }

        public void readRuldBoxSave()
        {
            try
            {
                string relativePath = constant.rootPath + constant.logFile;
                string fullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(relativePath)));
                common.makeFolder(fullPath);

                if (File.Exists(fullPath + "\\rule.log"))
                {
                    string text = common.ReadTxtFile(fullPath + "\\rule.log");
                    string[] textArray = text.Split('\n');
                    string[] ruleCheck = textArray[0].Split(';');
                    if (ruleCheck[0] == "True") rule1check = true;
                    if (ruleCheck[1] == "True") rule2check = true;
                    if (ruleCheck[2] == "True") rule3check = true;
                    if (ruleCheck[3] == "True") rule4check = true;
                    for (int i = 1; i <= ruleBoxCount; i ++)
                    {
                        string[] box = textArray[i].Split(';');
                        Rule2FileList.Add(box[0]);
                        Rule2SpaceList.Add(box[1]);
                    }
                    for (int i = 1; i <= ruleBoxCount; i++)
                    {
                        string[] box = textArray[ruleBoxCount + i].Split(';');
                        Rule3ListOrderingList.Add(box[0]);
                        Rule3ParagraphList.Add(box[1]);
                        Rule3SentenceList.Add(box[2]);
                        Rule3FileName.Add(box[3]);
                    }
                    for (int i = 1; i <= ruleBoxCount; i++)
                    {
                        string[] box = textArray[ruleBoxCount * 2 + i].Split(';');
                        Rule4FileList.Add(box[0]);
                        Rule4SpaceList.Add(box[1]);
                    }
                }
                else
                {
                    for (int i = 1; i <= ruleBoxCount; i++)
                    {
                        Rule2FileList.Add("");
                        Rule2SpaceList.Add("");
                        Rule3ListOrderingList.Add("");
                        Rule3ParagraphList.Add("");
                        Rule3SentenceList.Add(""); 
                        Rule3FileName.Add(""); 
                        Rule4FileList.Add("");
                        Rule4SpaceList.Add("");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("저장된 룰 체크 읽기시 오류가 발생하엇습니다.");
                common.makeLogFile("저장된 룰 체크 읽기시 오류가 발생하엇습니다.");
            }
        }

        public void readRuleCheckSave()
        {
            this.Rule1CheckBox.Checked = this.rule1check;
            this.Rule2CheckBox.Checked = this.rule2check;
            this.Rule3CheckBox.Checked = this.rule3check;
            this.Rule4CheckBox.Checked = this.rule4check;
        }
    }
}
