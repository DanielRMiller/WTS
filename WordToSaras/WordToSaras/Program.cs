using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace WordToSaras
{
    class Junk
    {
        /* THIS CODE WORKS FINE
            DirectoryInfo dir = new DirectoryInfo(@"C:\Users\Daniel\Desktop\Bio461");
            Console.WriteLine("Full Name is : {0}", dir.FullName);
            Console.WriteLine("Attributes are : {0}",
                               dir.Attributes.ToString());
            FileInfo[] docxfiles = dir.GetFiles("*.docx");
            Console.WriteLine("Total number of bmp files", docxfiles.Length);
            foreach(FileInfo f in docxfiles)
                {
                Console.WriteLine("Name is : {0}", f.Name);
                Console.WriteLine("Length of the file is : {0}", f.Length);
                Console.WriteLine("Creation time is : {0}", f.CreationTime);
                Console.WriteLine("Attributes of the file are : {0}",
                f.Attributes.ToString());
            }
            */
        //Console.WriteLine("Full Name is : {0}", file.FullName);
        //Console.WriteLine("Length of the file is : {0}", file.Length);
        //Console.WriteLine("Creation time is : {0}", file.CreationTime);
        //Console.WriteLine("Attributes of the file are : {0}", file.Attributes.ToString());
        // For Testing
        //string totaltext = "";
        // temporary question number
        // int tqNum = 0;
        // question number ** May need for location for Q headers... will see
        // int qNum = 1;

        // questionHeaders[qNum] += " \r\n " + docs.Paragraphs[p + 1].Range.Text.ToString();
        /*
                    // Get the rest of the document
                    for (; p < 25; p++) // docs.Paragraphs.Count; i++)
            {
                totaltext += " \r\n " + (p + 1) + " LINE: " + docs.Paragraphs[p + 1].Range.Text.ToString();
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.ListString;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.List;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.ListLevelNumber;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.ListPictureBullet;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.ListTemplate;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.ListType;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.ListValue;
                totaltext += " \r\n " + (p + 1) + " Info: " + docs.Paragraphs[p + 1].Range.ListFormat.SingleList;


            }
            */
        /* This algorithm need revised
        for (; i < docs.Paragraphs.Count; i++) //<-- For full document
        {
            Console.WriteLine(docs.Paragraphs[i + 1].Range.ListFormat.ListString);

            if (docs.Paragraphs[i + 1].Range.ListFormat.ListString == null)
            {
                questionHeaders[qNum + 1] = "\r\n" + docs.Paragraphs[i + 1].Range.Text.ToString();
                i++;
                while (i < docs.Paragraphs.Count && docs.Paragraphs[i + 1].Range.ListFormat.ListString == null)
                {
                    questionHeaders[qNum + 1] += " \r\n" + docs.Paragraphs[i + 1].Range.Text.ToString();
                    i++;
                }
                i--;
            }
            else if (Int32.TryParse(docs.Paragraphs[i + 1].Range.ListFormat.ListString, out tqNum))
            {
                qNum = tqNum;
                questions[qNum] = "\r\n" + docs.Paragraphs[i + 1].Range.Text.ToString();
                i++;
                while (i < docs.Paragraphs.Count && docs.Paragraphs[i + 1].Range.ListFormat.ListString == null)
                {
                    questions[qNum] += " \r\n" + docs.Paragraphs[i + 1].Range.Text.ToString();
                    i++;
                }
                i--;
            }
            else
            {
                int j = 2;
                responses[qNum, 1] = docs.Paragraphs[i + 1].Range.Text.ToString();
                i++;
                while (i < docs.Paragraphs.Count && 
                    docs.Paragraphs[i + 1].Range.ListFormat.ListString != null && 
                    !Int32.TryParse(docs.Paragraphs[i + 1].Range.ListFormat.ListString, out tqNum))
                {
                    //responses[qNum, j] += "\r\n" + docs.Paragraphs[i + 1].Range.Text.ToString();
                    i++;
                    j++;
                }
                i--;
            }
             totaltext += " \r\n" + docs.Paragraphs[i + 1].Range.ListFormat.ListString + docs.Paragraphs[i + 1].Range.Text.ToString();
            // totaltext += " \r\n " + (i + 1) + " Info: " + docs.Paragraphs[i + 1].Range.ListFormat.ListString;
        }
        /*


        /* Check the beginning output of the program
        totaltext += " \r\n 1: " + docs.Paragraphs[1].Range.Text.ToString();
        totaltext += " \r\n 2: " + docs.Paragraphs[2].Range.Text.ToString();
        totaltext += " \r\n 3: " + docs.Paragraphs[3].Range.Text.ToString();
        totaltext += " \r\n 4: " + docs.Paragraphs[4].Range.Text.ToString();
        totaltext += " \r\n 5: " + docs.Paragraphs[5].Range.Text.ToString();
        totaltext += " \r\n 6: " + docs.Paragraphs[6].Range.Text.ToString();
        totaltext += " \r\n 7: " + docs.Paragraphs[7].Range.Text.ToString();
        totaltext += " \r\n 8: " + docs.Paragraphs[8].Range.Text.ToString();
        */

        /* Secondary test for levels
        totaltext += " \r\n 4: " + docs.Paragraphs[4].Range.Text.ToString();
        totaltext += " \r\n 4 Info: " + docs.Paragraphs[4].Range.ListFormat.ListString;
        totaltext += " \r\n 5: " + docs.Paragraphs[5].Range.Text.ToString();
        totaltext += " \r\n 5 Info: " + docs.Paragraphs[5].Range.ListFormat.ListString;
        */

        /* LIST TEST
        List<Question> qs = new List<Question>();
        qs.Add(new Question("blah"));
        Console.WriteLine(qs.Last().getText());
        Console.ReadKey();
        */
    }
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
        /*
        * 0: 
        * 1: 
        * 2:
        */
        //private int type;
        //private string codename;
        public string text;
        private List<string> responses;
        //private List<int> score;

        
        public Question(string txt, List<string> r)
        {
            numQuestions++;
            //type = newType;
            //codename = numQuestions + "-" + type;
            text = txt;
            responses = r;
            //score = s;
        }
        //public int getType() { return type; }
        //public string getCodename() { return codename; }
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
            
            // This file will later be passed in
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
                    // question pos responses NOTE: THIS WILL BE USED TO DO MULTIPLE CHOICE LATER
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

            

            //Console.WriteLine(totaltext);
            Console.WriteLine(title);
            foreach (Question q in questionList)
            {
                Console.WriteLine("Q: " + q.text);
                foreach (string response in q.getResponses())
                {
                    Console.WriteLine("R: " + response);
                }
            }
            foreach  (Passage qh in questionHeaderList)
            {
                //Console.WriteLine("H: " + qh.getText());
            }
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
            ws.Cells[1, 18].Value2 = "Option1";
            ws.Cells[1, 19].Value2 = "Option1_Image";
            ws.Cells[1, 20].Value2 = "Option1_Rationale";
            ws.Cells[1, 21].Value2 = "Option2";
            ws.Cells[1, 22].Value2 = "Option2_Image";
            ws.Cells[1, 23].Value2 = "Option2_Rationale";
            ws.Cells[1, 24].Value2 = "Option3";
            ws.Cells[1, 25].Value2 = "Option3_Image";
            ws.Cells[1, 26].Value2 = "Option3_Rationale";
            ws.Cells[1, 27].Value2 = "Option4";
            ws.Cells[1, 28].Value2 = "Option4_Image";
            ws.Cells[1, 29].Value2 = "Option4_Rationale";
            ws.Cells[1, 30].Value2 = "Option5";
            ws.Cells[1, 31].Value2 = "Option5_Image";
            ws.Cells[1, 32].Value2 = "Option5_Rationale";
            ws.Cells[1, 33].Value2 = "Option6";
            ws.Cells[1, 34].Value2 = "Option6_Image";
            ws.Cells[1, 35].Value2 = "Option6_Rationale";
            ws.Cells[1, 36].Value2 = "Option7";
            ws.Cells[1, 37].Value2 = "Option7_Image";
            ws.Cells[1, 38].Value2 = "Option7_Rationale";

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
                ws.Cells[row, 15].Value2 = question.text;
                // OPTIONS
                int col = 18;
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
