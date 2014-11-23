using System;
using System.Collections.Generic;
using PaperMgr.Exceptions;

namespace PaperMgr.Entity
{
    /// <summary>
    /// Base class for papers
    /// </summary>
    abstract class Paper
    {
        const int MIN_PUBLICATION_YEAR = 1920;

        protected List<Person> m_authors = new List<Person>();
        protected int m_year;
        protected int m_pageCount;

        public string Title { get; set; }
        public Labels.LabelNames Label { get; set; }
        public string Reference { get; set; }
        public string Additional { get; set; }

        /// <summary>
        /// Copy of list of authors
        /// </summary>
        public List<Person> Authors
        {
            get
            {
                List<Person> copy = new List<Person>(m_authors);
                return copy;
            }
        }
        /// <summary>
        /// Year between MIN_PUBLICATION_YEAR and current year
        /// </summary>
        public int Year
        {
            get
            {
                return m_year;
            }
            set
            {
                int maxYear = DateTime.Now.Year;
                if (value > maxYear || value < MIN_PUBLICATION_YEAR)
                    throw new YearOutOfRangeException("Year must be in range between " + MIN_PUBLICATION_YEAR + " and " + maxYear + ".");
                else
                    m_year = value;
            }
        }
        /// <summary>
        /// Number of pages
        /// </summary>
        public int PageCount
        {
            get
            {
                return m_pageCount;
            }
            set
            {
                if (m_pageCount < 0)
                    throw new PageCountNumberException("Page count must be positive.");
                else
                    m_pageCount = value;
            }
        }

        public void addAuthor(Person author)
        {
            m_authors.Add(author);
        }
    }
}
