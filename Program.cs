using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WordToSaras
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the Document to store data.
            // NOTE: this will take a param that is the saveTo filename
            ExcelDoc xlDoc = new ExcelDoc();
            ParserHelper p = new ParserHelper();
            // NOTE: args should be replaced by the command line parameter that is the test filename
            p.parseFile(args, xlDoc);
            // This is just to help me see when the program is finished
            Console.Write("Press any key: ");
            Console.ReadKey();
        }
    }
    class ParserHelper
    {
        Excel.Application excel;
        Word.Application word;
        public void parseFile(string[] args, ExcelDoc xlDoc)
        {
            // Set up excel
            excel = xlDoc.xlApp;

            // This file will later be passed in trough an arg
            FileInfo file = new FileInfo(@"C:\Users\Daniel\Desktop\Bio461Orig\Bio461Exam1.docx");

            // Open the document as readOnly
            word = new Word.Application();
            object miss = Missing.Value;
            object path = file.FullName;
            object readOnly = true;
            Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            int MAX = docs.Paragraphs.Count;

            // paragraph we are currently on
            int p = 1;

            // To set up a question
            string qTxt = "";
            List<string> rTxt = new List<string>();

            // extra string for checking what I have
            string tmp = "";
            // Chars to trim so it will upload to Saras correctly
            char[] charsToTrim = { '\r', ' ', '\n', '\t' };

            // Get the title
            //string title = "";
            //int index = 0;

            // To bypass the Title of the document
            for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue == 0; p++)
            {
                /*
                tmp = docs.Paragraphs[p].Range.Text.ToString();
                if (tmp != null && tmp != "\r" && tmp != "" && tmp != "\n")
                {
                    title += docs.Paragraphs[p].Range.Text.ToString() + "\n";
                }
                */
            }
            //title = title.Trim(charsToTrim);
            //rng.MoveEnd(Word.WdUnits.wdParagraph, index);
            //xlDoc.ws.Range["A1", "A1"].Value2 = title;

            // Grabbing the rest of the document
            while (p < MAX)
            {
                // Get the questionHeader
                for (tmp = ""; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue == 0; p++)
                {
                    tmp += docs.Paragraphs[p].Range.Text.ToString() + '\n';
                }
                tmp = tmp.Trim(charsToTrim);
                if (tmp != null && tmp != "")
                    xlDoc.writePassage(tmp);

                // Get the question
                if (p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue != 0)
                {
                    // question main line
                    qTxt = docs.Paragraphs[p].Range.Text.ToString().Trim(charsToTrim) + '\n';
                    p++;
                    // question responses NOTE: THIS WILL BE USED TO DO MULTIPLE CHOICE LATER
                    // ListValue == 0 tells me we have not reached the answers yet
                    for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue == 0; p++)
                    {
                        // This checks if it is a blank line (do not need to add them into questionText)
                        tmp = docs.Paragraphs[p].Range.Text.ToString();
                        if (tmp != null && tmp != "\r")
                            qTxt += docs.Paragraphs[p].Range.Text.ToString().Trim(charsToTrim) + '\n';
                    }
                }
                // Get the answers
                for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue != 0; p++)
                {
                    char letter = 'A';
                    tmp = docs.Paragraphs[p].Range.Text.ToString();
                    if (tmp != "\r")
                        rTxt.Add(docs.Paragraphs[p].Range.Text.ToString().Trim(charsToTrim));
                    else
                    {
                        rTxt.Add(letter.ToString());
                        ++letter;
                    }
                }
                // put question into excel document
                xlDoc.writeQuestion(qTxt.Trim(charsToTrim), new List<string>(rTxt));

                // clear up the responses
                rTxt.Clear();
            }

            // Make like a tree
            docs.Close();
            word.Quit();
        }
    }
    class ExcelDoc
    {
        public Excel.Application xlApp;
        public Excel.Workbook wb;
        public Excel.Worksheet ws;
        private int row;
        private int questionNumber;

        public void writeQuestion(string questionText, List<string> responses)
        {
            // CODE
            ws.Cells[row, 1].Value2 = "MultiChoiceQ Code #" + (questionNumber);
            // LABEL
            ws.Cells[row, 2].Value2 = "MiltiChoiceQ Label #" + (questionNumber);
            // PARENT REF (QUESTION HEADER) NOTE: WE MAY NEED TO DO THIS LATER
            // NOTE: ALL headers are at least going to be related to question following
            if (ws.Cells[row - 1, 4].Value2 == "P")
                ws.Cells[row, 3].Value2 = ws.Cells[row - 1, 1];
            // ITEM TYPE
            ws.Cells[row, 4].Value2 = "MC";
            // ITEM CORRECT MARKS
            ws.Cells[row, 6].Value2 = 1;
            // ITEM WRONG MARKS
            ws.Cells[row, 7].Value2 = 0;
            // LANGUAGE
            ws.Cells[row, 11].Value2 = "English";
            // SHUFFLE
            ws.Cells[row, 12].Value2 = "True";
            // NUM OF OPTIONS
            ws.Cells[row, 13].Value2 = responses.Count;
            // CORRECT OPTION
            ws.Cells[row, 14].Value2 = 1;
            // ITEM TEXT
            // NOTE: WE NEED RICH TEXT, THIS IS CURRENTLY PLAIN TEXT
            ws.Cells[row, 15].Value2 = questionText;
            // OPTIONS
            int col = 18;
            foreach (string option in responses)
            {
                ws.Cells[row, col].Value2 = option;
                col += 3;
            }

            // Get variables ready for next question
            ++row;
            ++questionNumber;
        }

        public void writePassage(string passageText)
        {
            // Write each question
            {
                // CODE
                ws.Cells[row, 1].Value2 = "Passage Code #" + (questionNumber);
                // LABEL
                ws.Cells[row, 2].Value2 = "Passage Label #" + (questionNumber);
                // ITEM TYPE
                ws.Cells[row, 4].Value2 = "P";
                // LANGUAGE
                ws.Cells[row, 11].Value2 = "English";
                // ITEM TEXT
                // NOTE: WE NEED RICH TEXT, THIS IS CURRENTLY PLAIN TEXT
                ws.Cells[row, 15].Value2 = passageText;

                // Get variables ready for next entry
                ++row;
            }
        }

        public ExcelDoc()
        {
            row = 2;
            questionNumber = 1;
            xlApp = new Excel.Application();
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = false; // Turn this to true if you want to see the program in the foreground

            wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            ws = (Excel.Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            ws.Cells[1, 1].EntireRow.Font.Bold = true;

            // Setup the labels
            ws.Cells[1, 1].Value2 = "Code";
            ws.Cells[1, 2].Value2 = "Label";
            ws.Cells[1, 3].Value2 = "ParentItemRef";
            ws.Cells[1, 4].Value2 = "ItemType";
            ws.Cells[1, 5].Value2 = "ItemLevelScore";
            ws.Cells[1, 6].Value2 = "ItemCorrectMarks";
            ws.Cells[1, 7].Value2 = "ItemWrongMarks";
            ws.Cells[1, 8].Value2 = "Difficulty";
            ws.Cells[1, 9].Value2 = "Classification";
            ws.Cells[1, 10].Value2 = "Experience";
            ws.Cells[1, 11].Value2 = "Language";
            ws.Cells[1, 12].Value2 = "Shuffle";
            ws.Cells[1, 13].Value2 = "NoOfOptions";
            ws.Cells[1, 14].Value2 = "CorrectOption";
            ws.Cells[1, 15].Value2 = "ItemText";
            ws.Cells[1, 16].Value2 = "ItemImage";
            ws.Cells[1, 17].Value2 = "ItemRationale";
            int col = 18;
            for (int o = 1; o <= 10; o++, col += 3)
            {
                ws.Cells[1, col].Value2 = "Option" + o;
                ws.Cells[1, col + 1].Value2 = "Option" + o + "_Image";
                ws.Cells[1, col + 2].Value2 = "Option" + o + "_Rationale";
            }
        }

        ~ExcelDoc()
        {
            // This should make the worksheet look pretty 
            ws.Columns.AutoFit();
            ws.Rows.AutoFit();

            // Close the document
            xlApp.DisplayAlerts = false; // This prevents the overwrite message popping up every time
            wb.SaveAs(@"C:\Users\Daniel\Desktop\Bio461Orig\Bio461Exam1Excel.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            xlApp.Quit();
        }
    }
}
