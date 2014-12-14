using PaperMgr.Entity;
using PaperMgr.FileWriters;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PaperMgr
{
    class MainClass
    {
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
            disser.Publisher = "«Вега-Инфо";

            List<Paper> testSet = new List<Paper>();
            for (int i = 0; i < 10; i++)
            {
                testSet.Add(paper);
                testSet.Add(disser);
            }

            try
            { 
                var tex = new TexBibliographyWriter("C:/testfiles/testBib.tex").PrepareBibliographyAsync(testSet);
                var word = new WordListBibliographyWriter("C:/testfiles/testList.docx").PrepareBibliographyAsync(testSet);
                var excel = new ExcelBibliographyWriter("C:/testfiles/testSheet.xlsx").PrepareBibliographyAsync(testSet);
                tex.Wait();
                word.Wait();
                excel.Wait();
            }
            catch (System.Exception exc)
            {
                System.Console.WriteLine("Can not write file " + exc.ToString());
                System.Console.ReadKey(true);
            }
            
            
        }
    }
}
