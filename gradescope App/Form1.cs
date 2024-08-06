using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using Microsoft.VisualBasic.FileIO;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;
using System.Net;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace gradescope_App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Load += new System.EventHandler(this.Form_Resize);
            this.Resize += new System.EventHandler(this.Form_Resize);
            this.AutoSize = true;

            progressBar1.Visible = false;
        }

        private async void WebView_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess && uploadInProgress == true)
            {
                await webView.ExecuteScriptAsync("alert(\"An error navigating to a page has occured. The upload may have failed or be partially complete.\")");
            }

        }

        private void Form_Resize(object sender, EventArgs e)
        {
            webView.Size = this.ClientSize - new System.Drawing.Size(webView.Location);
            button4.Left = this.ClientSize.Width - button4.Width;
            buttonSelectFile.Left = button4.Left - buttonSelectFile.Width;
            label1.Left = buttonSelectFile.Left - label1.Width;
            button1.Left = 2;
            label2.Left = this.ClientSize.Width/2;
            progressBar1.Left = label2.Left - progressBar1.Width;
        }

        private List<string> filePath = new List<string>();
        private void buttonSelectFile_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            label2.Text = string.Empty;
            var fileContentLine = string.Empty;
            filePath.Clear();
            List<string> fileContent = new List<string>();

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.Filter = "csv files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in openFileDialog.FileNames)
                    {
                        filePath.Add(file);
                    }
                    string label = "";
                    foreach (string fileName in openFileDialog.SafeFileNames)
                    {
                        label += ", " + fileName;
                    }
                    label1.Text = label.Remove(0, 1);

                }

            }

            this.Resize += new System.EventHandler(this.Form_Resize);
        }



        private async Task checkErrorsUploadButton()
        {
            string uri = webView.Source.ToString();
            string gradePageURL = await webView.ExecuteScriptAsync("document.querySelectorAll('[href$=\"grade\"]')[0].getAttribute(\"href\")");
            if (filePath.Count <= 0)
            {
                MessageBox.Show("Please select a file");
            }
            else if (uri == "https://www.gradescope.com.au/login" || uri == "https://www.gradescope.com.au/saml" || !uri.Contains("gradescope"))
            {
                MessageBox.Show("Please login to upload files");
            }
            else if (gradePageURL == "null")
            {
                MessageBox.Show("Please choose a course and assignment before uploading files");
            }
            else { uploadInProgress = true; }
        }

        public string ShowMyDialogBox(string fileName)
        {
            Form2 questionNoForm = new Form2();
            this.AcceptButton = questionNoForm.button1;
            questionNoForm.label1.Text = fileName;
            string input;

            if (questionNoForm.ShowDialog(this) == DialogResult.OK)
            {
                input = questionNoForm.textBox1.Text;
            }
            else
            {
                input = "Cancelled";
            }
            questionNoForm.Dispose();
            return input;
        }

        private bool uploadInProgress = false;
        private int navigateDelay = 5000;
        private int actionDelay = 3000;

        private async void button4_Click(object sender, EventArgs e)
        {
            if (uploadInProgress) {return; }
            await checkErrorsUploadButton();
            if (!uploadInProgress) { return; }
            string baseURL = "https://www.gradescope.com.au";
            string gradePageURL = await webView.ExecuteScriptAsync("document.querySelectorAll('[href$=\"grade\"]')[0].getAttribute(\"href\")");
            gradePageURL = gradePageURL.Substring(1, gradePageURL.Length - 2);
            string URL = baseURL + gradePageURL;
            float fileProgress = 100 / filePath.Count;
            float rowProgress;
            float totalProgress = 0;
            int fileNo = 0;
            progressBar1.Visible = true;

            foreach (string file in filePath)
            {
                fileNo++;
                updateProgressBar(totalProgress, fileNo);
                string[,] fileContent = getFileArray(file);
                rowProgress = fileProgress/(fileContent.GetLength(0) - 1);

                int startIndex = 7;
                int endIndex;

                for (endIndex = 7; endIndex < fileContent.GetLength(1); endIndex++)
                {
                    if (fileContent[0, endIndex] == "Adjustment")
                    {
                        endIndex--;
                        break;
                    }
                }
                int commentsIndex = endIndex + 2;

                float time = ((endIndex - startIndex + 2) * actionDelay/1000 + navigateDelay/1000) * (fileContent.GetLength(0) - 1) + 2 * navigateDelay/1000;
                float urlNaviagateProgress = rowProgress/time * navigateDelay/1000 * (fileContent.GetLength(0) - 1);
                float fileRowProgress = rowProgress/time * actionDelay/1000 * (fileContent.GetLength(0) - 1);

                webView.CoreWebView2.Navigate(URL);
                await Task.Delay(navigateDelay);
                totalProgress += urlNaviagateProgress;
                updateProgressBar(totalProgress, fileNo);
                string questionNo = file.Substring(file.LastIndexOf('\\') + 1)[0].ToString();
                while (int.TryParse(questionNo, out int result) == false || questionNo == "Cancelled")
                {
                    questionNo = ShowMyDialogBox(file.Substring(file.LastIndexOf('\\') + 1));
                }
                string question = "question" + questionNo;
                string questionURL = await webView.ExecuteScriptAsync($"document.querySelector(\"[data-px={question}]\").getAttribute(\"href\");");
                if (questionURL == "null")
                {
                    MessageBox.Show("The question number is invalid. Upload failed");
                    uploadInProgress = false;
                    return; 
                }
                questionURL = (baseURL + questionURL.Substring(1, questionURL.Length - 2)).Replace("/grade", "/submissions");
                webView.CoreWebView2.Navigate(questionURL);
                await Task.Delay(navigateDelay);
                totalProgress += urlNaviagateProgress;
                updateProgressBar(totalProgress, fileNo);


                for (int i = 1; i < fileContent.GetLength(0); i++)
                {
                    string newURL = questionURL + "/" + fileContent[i, 1] + "/grade";
                    Console.WriteLine(newURL);
                    webView.CoreWebView2.Navigate(newURL);
                    await Task.Delay(navigateDelay);
                    totalProgress += urlNaviagateProgress;
                    updateProgressBar(totalProgress, fileNo);
                    //update grades
                    for (int j = startIndex; j <= endIndex; j++)
                    {
                        string wantedValue = fileContent[i, j].ToLower();
                        string javascriptButton = $"document.querySelectorAll('li.rubricEntryDragContainer')[{j - 7}].querySelector('button')";
                        string currentValue = await webView.ExecuteScriptAsync(javascriptButton + ".getAttribute(\"aria-pressed\");");
                        if (currentValue == "\"true\"" && wantedValue == "false")
                        {
                            await webView.ExecuteScriptAsync(javascriptButton + ".click();");
                        }
                        if (currentValue == "\"false\"" && wantedValue == "true")
                        {
                            await webView.ExecuteScriptAsync(javascriptButton + ".click();");
                        }
                        await Task.Delay(actionDelay);
                        totalProgress += fileRowProgress;
                        updateProgressBar(totalProgress, fileNo);
                    }
                    //update comment
                    string comment = fileContent[i, commentsIndex];
                    string javascriptCommentBox = $"document.getElementById(\"adjustment-comment\")";
                    await webView.ExecuteScriptAsync(
                        $"let textArea = {javascriptCommentBox};" +
                        $"\r\nconst nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;" +
                        $"\r\nnativeInputValueSetter.call(textArea, \"{comment}\");" +
                        $"\r\nconst event = new Event('change', {{ bubbles: true }});" +
                        $"\r\ntextArea.dispatchEvent(event);"
                        );
                    Console.WriteLine(await webView.ExecuteScriptAsync(javascriptCommentBox + ".value"));
                    totalProgress += fileRowProgress;
                    updateProgressBar(totalProgress, fileNo);
                    await Task.Delay(actionDelay);
                    
                }
            }
            uploadInProgress = false;
        }

        private void updateProgressBar(float totalProgress, int fileNo)
        {
            progressBar1.Value = (int)Math.Round(totalProgress, 0);
            if (progressBar1.Value == 100)
            {
                label2.Text = "Upload Complete";
            }
            else 
            {
                label2.Text = $"{progressBar1.Value}%    uploading file {fileNo} of {filePath.Count}";
            }
        }

        private void resetProgressBar()
        {
            progressBar1.Value = 0;
            progressBar1.Visible = false;
        }

        private string[,] getFileArray(string filePath)
        {
            string[] fileContentLine;
            List<string> fileContent = new List<string>();
            int numberOfColumns = 0;
            int numberOfRows;
            //read csv file to get row and column size
            using (TextFieldParser csvParser = new TextFieldParser(filePath))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                while (!csvParser.EndOfData)
                {
                    fileContentLine = csvParser.ReadFields();
                    if (fileContentLine[0] == "" || fileContentLine[0] == "Point Values")
                    {
                        break;
                    }
                    else
                    {
                        fileContent.Add(fileContentLine[0]);
                    }
                    if (csvParser.LineNumber == 2)
                    {
                        numberOfColumns = fileContentLine.Length;
                    }
                }
                numberOfRows = fileContent.Count;
            }
            string[,] content = new string[numberOfRows, numberOfColumns];
            //read csv file to get data into two-dimensional array
            using (TextFieldParser csvParser = new TextFieldParser(filePath))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                while (!csvParser.EndOfData && csvParser.LineNumber - 1 < numberOfRows)
                {
                    fileContentLine = csvParser.ReadFields();
                    for (int i = 0; i < fileContentLine.Length; i++)
                    {
                        content[csvParser.LineNumber - 2, i] = fileContentLine[i];
                    }

                }
            }
            return content;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            webView.CoreWebView2.Navigate("https://www.gradescope.com.au");
        }
    }
}
