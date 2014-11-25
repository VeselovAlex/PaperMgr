namespace PaperMgr.Entity
{
    class Dissertation : Paper
    {
        public Dissertation()
        {
            City = "";
            Publisher = "";
            Branch = "";
            Degree = "";
        }
        public string City { get; set; }
        public string Publisher { get; set; }
        public string Branch { get; set; }
        public string Degree { get; set; }
    }
}
