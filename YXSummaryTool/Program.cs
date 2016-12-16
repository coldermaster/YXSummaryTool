using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Linq;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace YXSummaryTool
{
    class Program
    {
        static void Main(string[] args)
        {

            //input file
            string XLSFileName = @"D:\Document\XY\週六晚4A-201611.xls";
            string SummaryDocPath = @"D:\Document\XY\New";
            
            //output file
            string OutputFileName = @"D:\Document\XY\New\opt\Sample.docx";

            DataSet ds = NPOIHelp.GetDataTableFromExcelFile(XLSFileName, 1);
            Dictionary<int, Atendee> AtendeeList = Reformat(ds.Tables[0]);
            string docPath = SummaryDocPath;
            string[] AllFiles = Directory.GetFiles(docPath);
            List<string> DocFiles = new List<string>();
            foreach (string filename in AllFiles)
            {
                if (Path.GetExtension(filename) == ".doc" || Path.GetExtension(filename) == ".docx")
                {
                    DocFiles.Add(filename);
                }
            }
            //Construct all Summary
            List<SummaryDoc> AllSummaryDoc = new List<SummaryDoc>();
            foreach (string DocFile in DocFiles)
            {
                SummaryDoc Summary = new SummaryDoc(DocFile);
                AllSummaryDoc.Add(Summary);
            }
            
            //parse all doc
            foreach (SummaryDoc Summary in AllSummaryDoc)
            {
                foreach (KeyValuePair<int, Atendee> AtendeeRec in AtendeeList)
                {
                    if (Summary.GetAuthor(AtendeeRec.Value.Name))
                    {
                        CheckAndTryInsert(AtendeeRec.Value, Summary, AtendeeList.Count());
                        break;                        
                    }
                   
                }
            }
            //Warning for those docs do not match to Atendee
            WhileLoopCheckSummaryDocList(AllSummaryDoc, AtendeeList);

            //End Warning for those doc do not match to Atendee

            //Warning for those Atendee do not match to Atendee
            Console.WriteLine("Name list for lack of summary:");
            foreach (KeyValuePair<int,Atendee> ad in AtendeeList)
            {
                if (!(ad.Value.HasSummary))
                {
                    Console.WriteLine( ad.Value.ID.ToString()+" " + ad.Value.Name);
                }
            }
            //End arning for those Atendee do not match to Atendee


            //Do output file

            Document document = new Document();
            //Create a new secition
            
            //Create a new paragraph
            
            //Append Text


            foreach (KeyValuePair<int, Atendee> ATD in AtendeeList)
            {
                Section section = document.AddSection();
                if (!ATD.Value.HasSummary)
                {
                    Paragraph Tital = section.AddParagraph();
                    Tital.AppendText(ATD.Key.ToString() + ATD.Value.Name);
                    //Tital.Format.OutlineLevel = OutlineLevel.Level1;
                    //Tital.ApplyStyle(BuiltinStyle.Heading1);
                    //Paragraph Paragraph = section.AddParagraph();                    
                    //Paragraph.Format.OutlineLevel = OutlineLevel.Body;
                    //Paragraph.AppendBreak(0);
                }
                else
                {
                    //Paragraph Tital = section.AddParagraph();
                    //Tital.AppendText(ATD.Value.Summary.Titel);
                    //Tital.ApplyStyle(BuiltinStyle.Title);
                    Paragraph Paragraph2 = section.AddParagraph();
                    Paragraph2.AppendText(ATD.Value.Summary.FullText);
                    //Paragraph2.Format.OutlineLevel = OutlineLevel.Level9;
                    //Paragraph.AppendBreak(0); 
                }             

            }
            //Save doc file.



            document.SaveToFile(OutputFileName, FileFormat.Docx);


            Console.WriteLine("End");
            Console.Read();

        }
        private static int GetNameRow(DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    var value = dt.Rows[i][j];
                    if (value.ToString() == "姓名")
                    {
                        return i ;
                    }
                }
            }
            return -1; 
        }
        static Dictionary<int, Atendee> Reformat(DataTable dt)
        {
            int NameRow = GetNameRow(dt);
            Dictionary<int, Atendee> NameList = new Dictionary<int, Atendee>();
            for (int i = 0; i < dt.Columns.Count-1; i++)
            {
                var tmp = dt.Rows[NameRow][i];
                var tmp2 = dt.Rows[NameRow][i+1];
                if ((tmp.ToString() == "No.") && (tmp2.ToString() == "姓名"))
                {
                    for (int j = NameRow+1; j < dt.Rows.Count; j++)
                    {
                        int id;
                        if (int.TryParse(dt.Rows[j][i].ToString(), out id) && !(dt.Rows[j][i + 1].ToString().Trim() == "") )
                        {
                            Atendee AD = new Atendee(id, (string)dt.Rows[j][i + 1]);
                            NameList.Add(id, AD);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                
                
            }
            return NameList;
            
        }
        private static int LoadAndValidateNumber(string str, int range)
        {
            int rt = -1;
            Console.WriteLine(str);
            string UserKeyin = Console.ReadLine();
            if (!(int.TryParse(UserKeyin, out rt)))
            {
                Console.BackgroundColor = ConsoleColor.Blue;
                Console.WriteLine("Error: Please enter a number.");
                Console.ResetColor();
                rt = LoadAndValidateNumber(str,range);
            }
            else if ((rt < 0) || (rt > range))
            {
                Console.BackgroundColor = ConsoleColor.Blue;
                Console.WriteLine("Error: Out of range.");
                Console.ResetColor();
                rt = LoadAndValidateNumber(str, range);
            }
            return rt;
        }
        private static void CheckAndTryInsert(Atendee AR, SummaryDoc SD,int MaxNumber)
        {
                         if (AR.HasSummary)
                        {
                            //two docs match to one name
                            //do some warning
                            Warning();
                            string Warn1 = string.Format("{0} has been matched to {1} who has already has summary named {2}.\nPlease Enter:\n  1 to assign {0} to {1}.\n  0 to ignore {0}.", Path.GetFileName(SD.FileName),AR.Name,Path.GetFileName(AR.Summary.FileName));
                            int NewLocationOrIgnore = LoadAndValidateNumber(Warn1,1);
                            if (NewLocationOrIgnore==0)
                            {
                                SD.ignored =true;
                            }
                            else
                            {      
                                AR.CleanSummary();
                                AR.SetSummary(SD);
                            }
                            //互動可以作在這邊，及時修正, 可能可以列出選項1. replace/2. ignore
                        }
                        else
                        {
                            AR.SetSummary(SD);                             
                        }
        }
        private static void WhileLoopCheckSummaryDocList(List<SummaryDoc> SDL, Dictionary<int, Atendee> ADD)
        {
            var WarningDocs = SDL.Where(c => ((c.HasAuther == false) &&(!(c.ignored))));
            if (WarningDocs.Count() > 0)
            {
                foreach (SummaryDoc SD in (IEnumerable<SummaryDoc>)WarningDocs)
                {
                    Warning();
                    int newKey = LoadAndValidateNumber("No mathced name found for {" + Path.GetFileName(SD.FileName) + "}.\nPlease enter the sequence numner of the author or 0 to ignore it.", ADD.Count());
                    if (newKey == 0)
                    {
                        SD.ignored = true;
                    }
                    else
                    {
                        CheckAndTryInsert(ADD[newKey], SD, ADD.Count());
                    }
                }
                WhileLoopCheckSummaryDocList(SDL, ADD);
            }  

            
        }
        private static void Warning()
        {
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.ForegroundColor = ConsoleColor.Black;
            Console.WriteLine("Warning!");
            Console.ResetColor();
        }
        
    }
}
