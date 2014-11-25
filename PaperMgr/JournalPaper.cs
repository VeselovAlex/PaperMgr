using System.Collections.Generic;
namespace PaperMgr.Entity
{
    abstract class JournalPaper : Paper
    {
        protected JournalPaper()
        {
            JournalTitle = "";
            JournalNumber = 0;
            Volume = 0;
            FirstPage = "";
            LastPage = "";
            IsInWACList = false;
            IsInScopus = false;
            IsInWoS = false;
        }
        public string JournalTitle{ get; set; }
        public int JournalNumber { get; set; }
        public int Volume { get; set; }
        public string FirstPage { get; set; }
        public string LastPage { get; set; }
        public bool IsInWACList { get; set; }
        public bool IsInScopus { get; set; }
        public bool IsInWoS { get; set; }

    }
}
