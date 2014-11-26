using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using PaperMgr.Entity;
using System.Threading;

namespace PaperMgr.FileWriters
{
    class TexBibliographyWriter
    {
        protected TexBibliographyWriter()
        {
            FileName = "file.tex";
        }

        private sealed class TexBibliographyWriterFactory
        {
            private static readonly TexBibliographyWriter instance = new TexBibliographyWriter();
            public static TexBibliographyWriter Instance { get { return instance; } }
        }

        protected static object lockOn = new Object();

        public static TexBibliographyWriter Instance
        {
            get
            {
                return TexBibliographyWriterFactory.Instance;
            }
        }

        private StreamWriter output;
        public string FileName { get; set; }

        protected bool OutputIsOpen { get { return output != null && output.BaseStream != null; } }

        protected void PrepareBibliographyTask(ICollection<Paper> papers)
        {
            lock (lockOn)
            {
                string curFileName = FileName;
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
                OnPrepComplete(curFileName);
                Monitor.PulseAll(lockOn);
                Monitor.Wait(lockOn);
            }

        }

        protected static void OnPrepComplete(string filename)
        {
            DialogResult result = MessageBox.Show("Файл " + filename + " готов. Открыть?", "TeX-файл готов",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (result == DialogResult.Yes)
                System.Diagnostics.Process.Start(filename);
        }

        public static void PrepareBibliography(string filename, ICollection<Paper> papers)
        {
            TexBibliographyWriter writer = TexBibliographyWriter.Instance;
            writer.FileName = filename;
            writer.PrepareBibliographyTask(papers);
        }

        public static Task PrepareBibliographyAsync(string filename, ICollection<Paper> papers)
        {
            TexBibliographyWriter writer = TexBibliographyWriter.Instance;
            writer.FileName = filename;
            return Task.Run(() => writer.PrepareBibliographyTask(papers));
        }


        public void Write(Paper paper)
        {
            Write(paper, true);
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
                    output.WriteLine(" // " + ppr.JournalTitle + ". " + ppr.Year + ". \\No~" + ppr.JournalNumber +
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
            string result = "";
            foreach (Person author in authors)
                result += author.GetName().Replace(' ', '~') + ", ";
            return result;
        }

    }

}
