using System.Collections.Generic;
namespace PaperMgr.Entity
{
    abstract class JournalPaper : Paper
    {
        public string JournalTitle{ get; set; }
        public int JournalNumber { get; set; }
        public int Volume { get; set; }
        protected List<Person> m_VACList = new List<Person>();

        public List<Person> VAC
        {
            get
            {
                List<Person> copy = new List<Person>(m_VACList);
                return copy;
            }
        }

        public void addToVAC(Person person)
        {
            m_VACList.Add(person);
        }
    }
}
