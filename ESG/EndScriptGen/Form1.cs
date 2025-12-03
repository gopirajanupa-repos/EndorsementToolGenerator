using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using newfileword = Microsoft.Office.Interop.Word;
using System.Globalization;
using Range = Microsoft.Office.Interop.Word.Range;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace EndScriptGen
{
    Range paragraphRange = paragraph.Range;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"d:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = ".docx",
                Filter = @"All Files|*.txt;*.docx;*.doc;*.pdf*.xls;*.xlsx;*.pptx;*.ppt|Text File (.txt)|*.txt|Word File (.docx ,.doc)|*.docx;*.doc|PDF (.pdf)|*.pdf|Spreadsheet (.xls ,.xlsx)|  *.xls ;*.xlsx|Presentation (.pptx ,.ppt)|*.pptx;*.ppt",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }

        }
        string MainHeading;
        string MainHeadingCaptilizeeachword = "";
        string Mainfootertext;
        StringBuilder builder = new StringBuilder();
        int programid = 99;
        string edtionid = "";

        public StringBuilder GenerateEndorsmentsTwoLines()
        {

            builder.AppendLine("delete from table_v400 where programid=" + programid + " and id='" + edtionid);
            builder.AppendLine("insert into table_v400 where programid=" + programid + " and id='" + edtionid);

            return builder;
        }


        static void ApplyBoldTags(Range range, int startWordIndex, int endWordIndex, StringBuilder modifiedContent)
        {
            Range startRange = range.Words[startWordIndex];
            Range endRange = range.Words[endWordIndex];

            string boldText = range.Range(startRange.Start, endRange.End).Text;
            modifiedContent.Append("<b>").Append(boldText).Append("</b> ");
        }
        private void button2_Click(object sender, EventArgs e)
        {

            //read main heading 
            string filePath = textBox1.Text.ToString();

            // Create an instance of the Word application
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            //GenerateEndorsmentsTwoLines();

            // Open the Word document
            Document document = wordApp.Documents.Open(filePath);

            //spelling and grammer check

            // Enable proofing for the document
            document.ShowSpellingErrors = true;
            //document.ShowGrammaticalErrors = true;
            // Get the spelling errors

            // Display the spelling errors   
            listBox1.Items.Add("Spelling checks in progress.....!");
            foreach (Range error in document.SpellingErrors)
            {
                int lineNumber = error.Information[WdInformation.wdFirstCharacterLineNumber];
                listBox1.Items.Add(error.Text + "  in  Line number  " + lineNumber);
            }
            listBox1.Items.Add("Spelling checks Completed.....!");
            listBox1.Items.Add("Feteching header and footer texts.....!");
             
            // Get the main header from the first section
            Section section = document.Sections[1];
            HeaderFooter header = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
            HeaderFooter footer = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
            // Read the content of the main header
            //TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;

            //// Capitalize each word in the sentence
            //MainHeadingCaptilizeeachword = textInfo.ToTitleCase(MainHeading.ToLower());
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            // Changes a string to titlecase.
            MainHeadingCaptilizeeachword = textInfo.ToTitleCase(header.Range.Text);
            string headerText = header.Range.Text;
            //builder.Append("'" + textInfo.ToTitleCase(MainHeadingCaptilizeeachword) + "',");
            //builder.Append("<b>" + headerText + "</b>,");
            string footerText = footer.Range.Text;
            listBox1.Items.Add("HEADER =>" + headerText + " FOOTER=>" + footerText);
            // Close the Word document and application
            listBox1.Items.Add("Applying leftintendation tabs...!");
            // Get the number of paragraphs in the document
            int paragraphCount = document.Paragraphs.Count;
            for (int i = 1; i <= paragraphCount; i++)
            {
                // Get the paragraph based on its index
                Paragraph paragraph = document.Paragraphs[i];

                float leftIndent = paragraph.Format.LeftIndent;

                // Calculate the intendent space
                int tabSpaces = (int)(leftIndent / 28.35); // Assuming 1 tab space is equal to 28.35 points (default Word tab width)

                //MessageBox.Show($"Paragraph Left Indentation: {leftIndent} points");
                //MessageBox.Show($"Number of Tab Spaces: {tabSpaces}");
                if (tabSpaces >= 1)
                {
                    for (int j = 0; j <= tabSpaces; j++)
                    {
                        builder.Append("<t><t>");
                    }
                }
                int numberOfSpaces = 0;
                int numberOfTabs = 0;
                // Get the prefix (leading spaces/tabs) of the paragraph
                string prefix = paragraph.Range.Text.Substring(0, paragraph.Range.Text.Length - paragraph.Range.Text.TrimStart().Length);

                // Count the number of spaces and tabs in the prefix
                foreach (char c in prefix)
                {
                    if (c == ' ')
                        numberOfSpaces++;
                    else if (c == '\t')
                        numberOfTabs++;
                }
                if (numberOfTabs >= 1)
                {
                    for (int j = 0; j <= tabSpaces; j++)
                    {
                        builder.Append("<t><t>");
                    }
                }

                //MessageBox.Show("number of spaces" + numberOfSpaces);
                //MessageBox.Show("number of spaces" + numberOfTabs);


                // Iterate over each word in the paragraph
                bool valueapplied = false;
                if (paragraph.Range.ListFormat.ListType == WdListType.wdListBullet
                        || paragraph.Range.ListFormat.ListType == WdListType.wdListListNumOnly
                        || paragraph.Range.ListFormat.ListType == WdListType.wdListMixedNumbering

                         || paragraph.Range.ListFormat.ListType == WdListType.wdListOutlineNumbering
                          || paragraph.Range.ListFormat.ListType == WdListType.wdListPictureBullet
                         || paragraph.Range.ListFormat.ListType == WdListType.wdListSimpleNumbering
                         )
                {
                    // Check if the paragraph is a numbered list item
                    if (paragraph.Range.ListFormat.ListType == WdListType.wdListBullet
                        || paragraph.Range.ListFormat.ListType == WdListType.wdListListNumOnly
                        || paragraph.Range.ListFormat.ListType == WdListType.wdListMixedNumbering
                        || paragraph.Range.ListFormat.ListType == WdListType.wdListNoNumbering
                         || paragraph.Range.ListFormat.ListType == WdListType.wdListOutlineNumbering
                          || paragraph.Range.ListFormat.ListType == WdListType.wdListPictureBullet
                         || paragraph.Range.ListFormat.ListType == WdListType.wdListSimpleNumbering)

                    {

                        if (paragraph.Range.ListFormat.ListType == WdListType.wdListBullet)
                        {
                            string unicodeString = "\u2022";
                            builder.Append(unicodeString);
                            builder.Append("<t>");
                        }
                        else
                        {
                            string listNumber = paragraph.Range.ListFormat.ListString;
                            builder.Append(listNumber);
                            builder.Append("<t>");
                        }

                    }
                    Range range = document.Content;
                    range.Find.ClearFormatting();

                    bool inBold = false;
                    int startPosition = 0;

                    StringBuilder modifiedContent = new StringBuilder();
                    foreach (Range wordRange in paragraphRange.Words)
                    {
                        //Range wordRanges = range.Words[i];
                        bool isBold = wordRange.Bold == 1;

                        if (isBold && !inBold)
                        {
                            inBold = true;
                            startPosition = i;
                        }
                        else if (!isBold && inBold)
                        {
                            inBold = false;
                            ApplyBoldTags(range, startPosition, i - 1, modifiedContent);
                        }

                        if (i == range.Words.Count - 1 && inBold)
                        {
                            ApplyBoldTags(range, startPosition, i, modifiedContent);
                        }
                        // Check if the word is bold
                        //int isBold = wordRange.Font.Bold;

                        // Check if the word is italic
                        //string word = wordRange.Text;

                        //if (isBold == -1 && wordRange.Font.Underline != WdUnderline.wdUnderlineNone)
                        //{
                        //    builder.Append("<u><b>" + word.Trim() + "</b></u> ");
                        //    valueapplied = true;
                        //}

                        //else if (isBold == -1 && valueapplied == false)
                        //{

                        //    builder.Append("<b>" + word.Trim() + "</b> ");
                        //    valueapplied = true;

                        //}
                        //else if (wordRange.Font.Underline != WdUnderline.wdUnderlineNone && valueapplied == false)
                        //{
                        //    builder.Append("<u>" + word.Trim() + "</u> ");
                        //    valueapplied = true;
                        //}
                        //else
                        //{
                        //    if (word.Contains("$"))
                        //    {
                        //        string replacing_word = word;
                        //        word = "[$________]";
                        //        builder.Append(word);
                        //        listBox1.Items.Add("REPLACED TEXT  WITH **" + replacing_word + "**with New " + word);
                        //    }
                        //    else if (word.StartsWith("_") && word.EndsWith("_"))
                        //    {
                        //        string replacing_word = word;
                        //        word = "[________]";
                        //        builder.Append(word);
                        //        listBox1.Items.Add("REPLACED TEXT  WITH **" + replacing_word + "**with New " + word);
                        //    }
                        //    else
                        //    {
                        //        builder.Append(word);
                        //    }
                        //}

                        valueapplied = false;
                    }
                    string modifiedText = modifiedContent.ToString();

                }
                else
                {
                    foreach (Range wordRange in paragraphRange.Words)
                    {

                        // Check if the word is bold
                        int isBold = wordRange.Font.Bold;

                        // Check if the word is italic
                        string word = wordRange.Text;

                        if (isBold == -1 && wordRange.Font.Underline != WdUnderline.wdUnderlineNone)
                        {
                            builder.Append("<u><b>" + word.Trim() + "</b></u> ");
                            valueapplied = true;
                        }

                        else if (isBold == -1 && valueapplied == false)
                        {
                            builder.Append("<b>" + word.Trim() + "</b> ");
                            valueapplied = true;
                        }
                        else if (wordRange.Font.Underline != WdUnderline.wdUnderlineNone && valueapplied == false)
                        {
                            builder.Append("<u>" + word.Trim() + "</u> ");
                            valueapplied = true;
                        }
                        else
                        {
                            if (word.Contains("$"))
                            {
                                string replacing_word = word;
                                word = "[$________]";
                                builder.Append(word);
                                listBox1.Items.Add("REPLACED TEXT WITH**" + replacing_word + " **with New " + word);
                            }
                            else if (word.StartsWith("_") && word.EndsWith("_"))
                            {
                                string replacing_word = word;
                                word = "[________]";
                                builder.Append(word);
                                listBox1.Items.Add("REPLACED TEXT WITH **" + replacing_word + "**with New " + word);
                            }
                            else
                            {
                                builder.Append(word);
                            }
                        }
                        valueapplied = false;

                    }
                    //builder.AppendLine();
                }
            }
            listBox1.Items.Add("Applying bold and underline tags!");
            // Create a new instance of Word Application

            if (textBox3.Text.Length > 1)
            {
                string Suggesttag = "','<Available><MiscText2>=";
                builder.Append(Suggesttag);
                builder.Append(textBox3.Text);
                string SuggestClosetag = "/<MiscText2>";
                builder.Append(SuggestClosetag);

            }
            if (textBox4.Text.Length > 1)
            {
                string Formttag = "<Form>=";
                builder.Append(Formttag);
                builder.Append(textBox4.Text);
                string SuggestFormtag = "/<Form><Available1>,'null,null,GetDate(),null,0)";
                builder.Append(SuggestFormtag);
            }
            newfileword.Application NewwordApp = new newfileword.Application();
            // Create a new document
            newfileword.Document doc = NewwordApp.Documents.Add();
            doc.Content.Text = builder.ToString();
            Random random = new Random();
            int value = random.Next(1, 300);
            string newdocpath = @"D:\test" + value + ".docx";
            // Check if the file already exists
            doc.SaveAs2(newdocpath);
            //Specify the file path to save the text
            if (MainHeading == "")
            {
                MainHeading = "test";
            }
            string filePaths = @"D:\" + MainHeading + ".sql";
            listBox1.Items.Add("GENERATED FILE SAVED PATH " + newdocpath);
            // Write the StringBuilder content to the file without combining the lines
            using (StreamWriter writer = new StreamWriter(filePaths))
            {
                foreach (string line in builder.ToString().Split('\n'))
                {
                    line.Replace(" . ", ". ");
                    line.Replace(" , ", ", ");
                    writer.WriteLine(line);
                }
            }

            MessageBox.Show(builder.ToString());
            builder.Clear();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //loadDomicile();
            loadProgram();
            //loadCostCenter();
            //loadForms();
        }

        public void loadexcelvalues(string Params, string programselected)
        {
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            Dictionary<int, string> programdetails = new Dictionary<int, string>();
            programdetails.Add(1, "GPL");
            programdetails.Add(2, "MPL");
            programdetails.Add(3, "HH");
            programdetails.Add(4, "ALPHA GREEN");
            programdetails.Add(5, "PUBLIC D&O");
            programdetails.Add(6, "PRIVATE D&O");
            programdetails.Add(7, "FI");

            int key = programdetails.FirstOrDefault(pair => pair.Value == programselected).Key;
            //load forms
            System.Data.DataTable dt = new System.Data.DataTable();
            Excel.Application excelApp = new Excel.Application();
            // Open the workbook
            Excel.Workbook workbook = excelApp.Workbooks.Open(Params);
            // Get the first worksheet from the workbook
            Excel.Worksheet worksheet = workbook.Worksheets[key]; // Assuming data is on the first sheet (index 1)
            // Read data from the worksheet
            int rowCount = worksheet.UsedRange.Rows.Count;
            int columnCount = worksheet.UsedRange.Columns.Count;
            for (int col = 1; col <= 1; col++)
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    // Read the cell value
                    Excel.Range range = worksheet.Cells[row, col];

                    string cellValue = range.Value?.ToString();
                    // Do something with the cell value
                    if (cellValue != null)
                        comboBox4.Items.Add(cellValue);

                }
            }
            for (int col = 2; col <= 2; col++)
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    // Read the cell value
                    Excel.Range range = worksheet.Cells[row, col];
                    string cellValue = range.Value?.ToString();
                    // Do something with the cell value
                    if (cellValue != null)
                        comboBox3.Items.Add(cellValue);

                }
            }
            // Close the workbook and release resources
            workbook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            comboBox3.SelectedIndex = 1;
            comboBox4.SelectedIndex = 1;

        }


        public void loadDomicile()
        {
            comboBox1.Items.Add("ALL");
            comboBox1.Items.Add("CA");
            comboBox1.Items.Add("TX");
            comboBox1.Items.Add("AK");
            comboBox1.Items.Add("WA");
        }
        public void loadProgram()
        {
            comboBox1.Items.Add("GPL");
            comboBox1.Items.Add("MPL");
            comboBox1.Items.Add("HH");
            comboBox1.Items.Add("ALPHA GREEN");
            comboBox1.Items.Add("FI");
            comboBox1.Items.Add("PUBLIC D&O");
            comboBox1.Items.Add("PRIVATE D&O");

        }
        public void loadCostCenter()
        {
            List<String> ListCostcenter = new List<string>();
            ListCostcenter.Add("abc");


        }
        public void loadForms()
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
                loadexcelvalues(textBox2.Text, comboBox1.SelectedItem.ToString());
            else
                MessageBox.Show("please select any program");

        }

        private void button4_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the initial directory and filter for the file types
            openFileDialog.InitialDirectory = "d:\\";
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            // Display the dialog and check if the user clicked the "OK" button
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file path
                string fileiNPUTPath = openFileDialog.FileName;
                textBox2.Text = fileiNPUTPath;


            }

        }
        string Copeningtag = "<MiscText>";
        string CClosingtag = "</MiscText>";
        StringBuilder SBCostcenter = new StringBuilder();
        string fopeningtag = "<FORM>";
        string FClosingtag = "</FORM>";
        StringBuilder SBfORM = new StringBuilder();
        private void button5_Click(object sender, EventArgs e)
        {

            if (textBox3.Text != "")
            {
                //textBox3.Text = "/" + comboBox3.SelectedItem.ToString();
                //string Stext = textBox3.Text.ToString() + comboBox3.SelectedItem.ToString();
                //textBox3.Text = Stext;
                SBCostcenter.Append("/" + comboBox3.SelectedItem);
                textBox3.Text = SBCostcenter.ToString();

            }
            else
            {
                SBCostcenter.Append("/" + comboBox3.SelectedItem);
                textBox3.Text = SBCostcenter.ToString();
            }
            //SBCostcenter.Insert(0, Copeningtag);
            //SBCostcenter.Append("/"+CClosingtag);
        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (textBox4.Text != "")
            {
                //textBox3.Text = "/" + comboBox3.SelectedItem.ToString();
                //string Stext = textBox3.Text.ToString() + comboBox3.SelectedItem.ToString();
                //textBox3.Text = Stext;
                SBfORM.Append("/" + comboBox4.SelectedItem);
                textBox4.Text = SBfORM.ToString();

            }
            else
            {
                SBfORM.Append("/" + comboBox4.SelectedItem);
                textBox4.Text = SBfORM.ToString();
            }
        }
    }
}
