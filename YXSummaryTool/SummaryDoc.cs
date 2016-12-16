using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using System.IO;

namespace YXSummaryTool
{
    class SummaryDoc
    {
        public string FileName;
        public string Titel;
        public string Date;
        private bool m_HasTitel;
        public string Author;
        public string FullText;
        private string[] Content;
        private bool m_HasAuther;
        public bool ignored;
        public bool HasAuther
        {
            set {  m_HasAuther = value; }
            get { return m_HasAuther; }
        }
        public bool HasTitel
        {
            //set { m_HasTitel = value; }
            get { return m_HasTitel; }
        }
        public SummaryDoc(string filename)
        {
            this.FileName = filename;
            Document document = new Document();
            document.LoadFromFile(filename);
            this.FullText = document.GetText();
            this.Content = this.FullText.Split('\n');
            this.m_HasAuther = false;
            this.m_HasTitel = false;
            this.GetTitel();
            this.ignored = false;
        }
        public void GetTitel()
        {
            for (int i = 0; i < Content.Length; i++)
            {
                if (!(Content[i]==String.Empty))
                {
                    m_HasTitel = true;
                    Titel = Content[i];
                    break;
                }
            }
        }
        public bool GetAuthor(string name)
        {
            if (this.FileName.Contains(name))
            {
                this.Author = name;
                return true;
            }
            else if (SearchForAuthor(name))
            {
                this.Author = name;                
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool SearchForAuthor(string name)
        {
            bool rt = false;
            List<string> LinesWeShouldWatch = new List<string>();
            int HowManyLinesWeCheck = 4;
            if (Content.Length > HowManyLinesWeCheck)
            {
                for (int i = 0; i < HowManyLinesWeCheck; i++)
                {
                    LinesWeShouldWatch.Add(Content[i]);
                }
                LinesWeShouldWatch.Add(Content[Content.Length - 1]);
            }
            foreach (string line in LinesWeShouldWatch)
            {
                if (line.Contains(name))
                {
                    return true;
                }
            }
            return rt;
        }
    }
}
