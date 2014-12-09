using PaperMgr.Entity;
using PaperMgr.FileWriters;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PaperMgr
{
    class MainClass
    {
        private static async Task runAsync(string filename, ICollection<Paper> papers)
        {
            IBibWriter writer = new TexBibliographyWriter(filename + ".tex");
            IBibWriter writer2 = new WordListBibliographyWriter(filename + ".docx");
            var t =  writer.PrepareBibliographyAsync(papers);
            var t2 = writer2.PrepareBibliographyAsync(papers);
            await t;
            await t2;
        }

        private static void runSync(string filename, ICollection<Paper> papers)
        {
            IBibWriter writer = new TexBibliographyWriter(filename + ".tex");
            IBibWriter writer2 = new WordListBibliographyWriter(filename + ".docx");
            writer.PrepareBibliography(papers);
            writer2.PrepareBibliography(papers);
        }

        public static void Main()
        {

            JournalPaper paper = new JournalArticle();
            paper.addAuthor(new Person("Ivan", "Ivanovich","Ivanov"));
            paper.Title = "Introduction to C#";
            paper.PageCount = 90;
            paper.Label = Labels.LabelNames.Electronic;
            paper.JournalNumber = 15;
            paper.JournalTitle = "Applied programming";
            paper.Year = 2014;
            paper.FirstPage = "51";
            paper.LastPage = (51 + paper.PageCount - 1).ToString();
                       
            Dissertation disser = new Dissertation();
            disser.Label = Labels.LabelNames.Printed;
            disser.Title = "First disser";
            disser.addAuthor(new Person("Сергей", "Васильевич", "Петров"));
            disser.addAuthor(new Person("Сергей", "Васильевич", "Смирнов"));
            disser.addAuthor(new Person("Ivan", "Ivanovich", "Ivanov"));
            disser.Year = 2010;
            disser.PageCount = 100;
            disser.Degree = "Проф.";
            disser.Publisher = "СПбГУ";

            List<Paper> testSet = new List<Paper>();
            for (int i = 0; i < 10; i++)
            {
                testSet.Add(paper);
                testSet.Add(disser);
            }

            try
            { /*Task t = runAsync("fullTest", testSet);
                t.Wait();*/
                (new WordTableBibliographyWriter("testfiles/testTable.docx", new Person("Ivan", "Ivanovich", "Ivanov"))).PrepareBibliography(testSet);
                (new WordListBibliographyWriter("testfiles/testList.docx")).PrepareBibliography(testSet);
            }
            catch (System.Exception exc)
            {
                System.Console.WriteLine("Can not write file " + exc.ToString());
                System.Console.ReadKey(true);
            }
            
            
        }
    }
}
