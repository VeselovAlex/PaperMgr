using PaperMgr.Entity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PaperMgr.FileWriters
{
    class TexBibliographyWriter : BibWriter
    {
        public TexBibliographyWriter(string fileName)
            : base(fileName)
        {
            FileReadyMsg = "TeX-файл готов";
        }

        protected StreamWriter output;

        protected bool OutputIsOpen
        {
            get
            {
                return output != null && output.BaseStream != null;
            }
        }

        protected override void PrepareBibliographyTask(ICollection<Paper> papers)
        {
            string curFileName = "";
            curFileName = FileName;
            using (output = OutputIsOpen ? output : new StreamWriter(curFileName))
            {
                output.WriteLine("\\begin{thebibliography}[" + papers.Count + ".]");
                foreach (Paper paper in papers)
                {
                    Write(paper, false);
                }
                output.WriteLine("\n\\end{thebibliography}");
                output.Close();
            }
        }

        protected void Write(Paper paper, bool closeOnExit)
        {
            try
            {
                output = OutputIsOpen ? output : new StreamWriter(FileName);
                List<Person> authors = paper.Authors;
                output.WriteLine("\n\\bibitem{" + getKeyFromAuthors(authors) + paper.Year + "}");//BibItem Header
                output.Write("\\emph{ " + concatAuthorsNames(authors) + "\\/} ");
                output.Write(paper.Title);
                if (paper is Dissertation)
                {
                    Dissertation ppr = (Dissertation)paper;
                    output.WriteLine(". " + ppr.Publisher + ", " + ppr.Year + ". " + ppr.PageCount + "~с.");
                }
                else if (paper is CompilationArticle)
                {
                    CompilationArticle ppr = (CompilationArticle)paper;
                    output.WriteLine(" // " + ppr.CompilationTitle + ". " + ppr.City + ", " + ppr.Year +
                        ". С.~" + ppr.FirstPage + (ppr.LastPage.Equals("") ? "--" : "") + ppr.LastPage + ".");
                }
                else if (paper is JournalArticle)
                {
                    JournalArticle ppr = (JournalArticle)paper;
                    output.WriteLine(" // " + ppr.JournalTitle + ". " + ppr.Year + ". \\textnumero~" + ppr.JournalNumber +
                        ". С.~" + ppr.FirstPage + (ppr.LastPage.Equals("") ? "" : "--") + ppr.LastPage + ".");
                }
            }
            finally
            {
                if (closeOnExit)
                    output.Close();
            }
        }

        private string getKeyFromAuthors(ICollection<Person> authors)
        {
            string res = "";
            if (authors.Count > 1)
            {
                foreach (Person person in authors)
                    res += person.Surname.Substring(0, 1).ToUpper();
            }
            else if (authors.Count == 1)
            {
                string surname = authors.First<Person>().Surname;
                res = surname.Substring(0, Math.Min(surname.Length, 5)).ToUpper(); ;
            }

            //TODO Убрать костыль
            string[] latinUpper = { "A", "B", "V", "G", "D", "E", "YO", "ZH", "Z", "I", "Y", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TS", "CH", "SH", "SHCH", "E", "YU", "YA" };
            string[] cyrUpper = { "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Э", "Ю", "Я" };
            for (int i = 0; i < 30; i++)
                res = res.Replace(cyrUpper[i], latinUpper[i]);

            return res;
        }

        private string concatAuthorsNames(List<Person> authors)
        {
            if (authors == null || authors.Count == 0)
                return "";
            Person last = authors.Last<Person>();
            string result = "";
            foreach (Person author in authors)
                result += author.GetName().Replace(' ', '~') + (last == author ? "" : ", ");
            return result;
        }

    }

}
