using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace WordToSaras
{
    class Passage
    {
        private string text;
        public Passage(string txt)
        {
            text = txt;
        }
        public string getText() { return text; }
    }
    class Question
    {
        static int numQuestions = 0;
        public string text;
        private List<string> responses;
        
        public Question(string txt, List<string> r)
        {
            numQuestions++;
            text = txt;
            responses = r;
        }
        public List<string> getResponses() { return responses; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            // Note: questions do not include headers because headers pertain to multiple questions
            List<Question> questionList = new List<Question>();
            List<Passage> questionHeaderList = new List<Passage>();
            ParserHelper p = new ParserHelper();
            p.parseFile(args, questionList, questionHeaderList);
            ExcelCreator e = new ExcelCreator();
            e.writeExcelDocument(args, questionList, questionHeaderList);
            Console.ReadKey();
        }
    }
    class ParserHelper
    { 
        public void parseFile(string[] args, List<Question> questionList, List<Passage> questionHeaderList)
        {
            
            // This file will later be passed in trough an arg
            FileInfo file = new FileInfo(@"C:\Users\Daniel\Desktop\Bio461Orig\Bio461Exam1.docx");

            // Open the document as readOnly
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = Missing.Value;
            object path = file.FullName;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            int MAX = docs.Paragraphs.Count;

            // Items I want to parse
            string title = "";
            
            // line/paragraph
            int p = 1;
            
            // To set up a question
            string qTxt = "";
            List<string> rTxt = new List<string>();

            // To set up questionHeader
            string tmp = "";

            // Get the title
            for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue == 0; p++)
            {
                tmp = docs.Paragraphs[p].Range.Text.ToString();
                if (tmp != null && tmp != "\r")
                    title += " \r\n " + tmp;
            }
            while (p < MAX)
            {
                // Get the questionHeader
                tmp = "";
                for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue == 0; p++)
                {
                    tmp += docs.Paragraphs[p].Range.Text.ToString();
                }
                if (tmp != null && tmp != "" && tmp != "\r")
                    questionHeaderList.Add(new Passage(" \r\n " + tmp));
                // Get the question
                if (p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue != 0)
                {
                    // question main line
                    qTxt = docs.Paragraphs[p].Range.Text.ToString();
                    p++;
                    // question responses NOTE: THIS WILL BE USED TO DO MULTIPLE CHOICE LATER
                    for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue == 0; p++)
                    {
                        tmp = docs.Paragraphs[p].Range.Text.ToString();
                        if (tmp != null && tmp != "\r")
                            qTxt += " \r\n " + docs.Paragraphs[p].Range.Text.ToString();
                    }
                }
                // Get the answers
                for (; p < MAX && docs.Paragraphs[p].Range.ListFormat.ListValue != 0; p++)
                {
                     rTxt.Add(docs.Paragraphs[p].Range.Text.ToString());
                }
                // put question into questionList
                questionList.Add(new Question(qTxt, new List<string>(rTxt)));
                // clear up the responses
                rTxt.Clear();
            }

            // Make like a tree
            docs.Close();
            word.Quit();
        }
    }
    class ExcelCreator
    {
        public void writeExcelDocument(string[] args, List<Question> questionList, List<Passage> questionHeaderList)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = false; // Turn this to true if you want to see the program in the foreground

            Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

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
            
            // Write each question
            int row = 2;
            foreach (Question question in questionList)
            {
                // CODE
                ws.Cells[row, 1].Value2 = "Sample MultiChoiceQ Code #" + (row - 1);
                // LABEL
                ws.Cells[row, 2].Value2 = "Sample MiltiChoiceQ Label #" + (row - 1);
                // PARENT REF (QUESTION HEADER)
                ws.Cells[row, 3].Value2 = "";
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
                ws.Cells[row, 13].Value2 = question.getResponses().Count;
                // CORRECT OPTION
                ws.Cells[row, 14].Value2 = 1;
                // ITEM TEXT
                // NOTE: WE NEED RICH TEXT, THIS IS CURRENTLY PLAIN TEXT
                ws.Cells[row, 15].Value2 = question.text;
                // OPTIONS
                col = 18;
                char letter = 'A';
                foreach (string option in question.getResponses())
                {
                    if (option != "\r")
                        ws.Cells[row, col].Value2 = option;
                    else
                    {
                        ws.Cells[row, col].Value2 = letter.ToString();
                        ++letter;
                    }
                    col += 3;
                }
                ++row;
            }

            // Close the document
            xlApp.DisplayAlerts = false; // This prevents the overwrite message popping up every time
            wb.SaveAs(@"C:\Users\Daniel\Desktop\Bio461Orig\Bio461Exam1Excel.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            xlApp.Quit();
        }
    }
}
