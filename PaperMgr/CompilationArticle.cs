namespace PaperMgr.Entity
{
    class CompilationArticle : Paper
    {
        public CompilationArticle()
        {
            Volume = 0;
            City = "";
            Publisher = "";
            CompilationTitle = "";
            FirstPage = "";
            LastPage = "";
        }
        public int Volume { get; set; }
        public string City { get; set; }
        public string Publisher { get; set; }
        public string CompilationTitle { get; set; }
        public string FirstPage { get; set; }
        public string LastPage { get; set; }
    }
}
