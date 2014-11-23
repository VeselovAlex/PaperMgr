namespace PaperMgr.Entity
{
    class CompilationArticle : Paper
    {
        public int Volume { get; set; }
        public string City { get; set; }
        public string Publisher { get; set; }
        public string CompilationTitle { get; set; }
    }
}
