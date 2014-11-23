using System;
using PaperMgr.Entity;

namespace PaperMgr
{
    class MainClass
    {
        public static void Main()
        {
            JournalPaper paper = new JournalArticle();
            paper.addAuthor(new Person("Ivanov", "Ivan", "Ivanovich"));
            paper.Title = "Introduction to C#";
            paper.PageCount = 90;
            paper.Label = Labels.LabelNames.Electronic;
            paper.JournalNumber = 15;
            paper.JournalTitle = "Applied programming";
            paper.Year = 2014;
            Console.WriteLine(paper.ToString());
            try
            {
                paper.Year = 2015;
            }
            catch (Exceptions.YearOutOfRangeException)
            {
                Console.WriteLine("Year fault test successful!");
                Console.WriteLine(paper.Year);
            }
            Paper disser = new Dissertation();
            disser.Title = "First disser";
            Console.WriteLine(disser.Title);
            Console.ReadKey(false);
            
        }
    }
}
