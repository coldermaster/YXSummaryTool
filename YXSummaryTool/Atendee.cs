using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YXSummaryTool
{
    class Atendee
    {
        public string Name;
        public int ID;        
        private bool m_HasSummary;
        public bool HasSummary 
        {
            get{ return m_HasSummary;}
        }
        public SummaryDoc Summary;
        public Atendee(int dD, string name)
        {
            this.ID = dD;
            this.Name = name;
            this.m_HasSummary = false;
        }
        public void SetSummary(SummaryDoc summary)
        {
            this.Summary = summary;
            this.m_HasSummary = true;
            summary.HasAuther = true;
        }
        public void CleanSummary()
        {
            this.Summary.HasAuther = false;
            this.Summary = null;
        }
    }
}
