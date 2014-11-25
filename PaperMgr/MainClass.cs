using System;
using PaperMgr.Entity;
using PaperMgr.FileWriters;

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
                       
            Dissertation disser = new Dissertation();
            disser.Title = "First disser";
            disser.addAuthor(new Person("Сергей", "Васильевич", "Петров"));
            disser.Year = 2010;
            disser.PageCount = 100;
            disser.Degree = "Проф.";
            disser.Publisher = "СПбГУ";

            TexBibliographyWriter writer = TexBibliographyWriter.Instance;
            writer.FileName = "test1.tex";
            writer.PrepareBibliography(new Paper[] { paper, disser });

            Console.WriteLine(disser.Title);
            //Console.ReadKey(false);

        }
    }
}
