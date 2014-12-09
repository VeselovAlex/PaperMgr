using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperMgr.Entity;
using System.Collections.Generic;
using System.Linq;

//Interop R.I.P.
namespace PaperMgr.FileWriters
{
    /// <summary>
    /// Class for creation bibliography as Word document
    /// </summary>
    class WordListBibliographyWriter : BibWriter
    {
        /// <summary>
        /// Constructor for Word Bibliography Writer (List style)
        /// </summary>
        /// <param name="filename">Path to new file</param>
        public WordListBibliographyWriter(string filename)
            : base(filename)
        {
            FileReadyMsg = "Документ Word готов.";
        }       
        
        /// <summary>
        /// Creates Bibliography file with Elenemts from <paramref name="papers"/>
        /// </summary>
        /// <param name="papers">Collection of papers to add to bibliography</param>
        protected override void PrepareBibliographyTask(ICollection<Paper> papers)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(FileName, WordprocessingDocumentType.Document))
            {
                //Create document
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                Document doc = new Document();
                mainPart.Document = doc;
                Body body = new Body();
                doc.AppendChild(body);

                //Create page title
                Paragraph pageTitle = new Paragraph();
                ParagraphProperties titleParProp = new ParagraphProperties();
                titleParProp.Justification = new Justification() { Val = JustificationValues.Center };
                titleParProp.SpacingBetweenLines = new SpacingBetweenLines { After = "0" };
                pageTitle.ParagraphProperties = titleParProp;
                Run titleRun = new Run(new Text("Список"));
                RunProperties titleProp = new RunProperties();
                titleProp.FontSize = new FontSize() { Val = "32" };
                titleProp.Bold = new Bold();
                titleRun.RunProperties = titleProp;
                pageTitle.AppendChild(titleRun);
                body.AppendChild(pageTitle);

                //Create subtitle
                Paragraph pageSubTitle = new Paragraph(new Run(new Text("опубликованных работ")));
                ParagraphProperties subTitleProps = new ParagraphProperties();
                subTitleProps.Justification = new Justification() { Val = JustificationValues.Center };
                pageSubTitle.ParagraphProperties = subTitleProps;
                body.AppendChild(pageSubTitle);

                //Create list by Year
                var groupByYear = from p in papers
                                  group p by p.Year;

                int i = 1;
                foreach (var year in groupByYear)
                {
                    Paragraph title = new Paragraph();
                    Run run = new Run(new Text(year.Key.ToString()));
                    RunProperties prop = new RunProperties();
                    prop.Bold = new Bold();
                    prop.FontSize = new FontSize() { Val = "28" };
                    run.RunProperties = prop;
                    title.AppendChild(run);
                    body.AppendChild(title);
                    foreach (Paper paper in year)
                    {
                        body.AppendChild(PaperParagraph(paper, i));
                    }
                    i++;
                }
                
                wordDoc.Close();
            }
        }

        /// <summary>
        /// Method to create paper paragraph with/without numeration
        /// </summary>
        /// <param name="paper">Source paper</param>
        /// <param name="listNumber">-1, for creation w/o numeration, else number of list</param>
        /// <returns>Paper paragraph</returns>
        protected Paragraph PaperParagraph(Paper paper, int listNumber = -1)
        {
            string res = "";
            res += concatAuthorsNames(paper.Authors) + " ";
            res += paper.Title;

            if (paper is JournalPaper)
            {
                JournalPaper journal = (JournalPaper)paper;
                res += " // ";
                res += journal.JournalTitle + ". ";
                res += journal.Year + ". ";
                res += (journal.JournalNumber > 0) ? "№ " + journal.JournalNumber + " " : "";
                res += (journal.Volume > 0) ? "Т. " + journal.Volume + ". " : "";
                res += "С. " + journal.FirstPage;
                res += (journal.LastPage != "") ? "–" + journal.LastPage + ". " : ". ";
                res += (journal.Reference != "") ? "URL: " + journal.Reference + ". " : "";
                res += journal.Additional;
            }
            else if (paper is Dissertation)
            {
                Dissertation diss = (Dissertation)paper;
                res += ". ";
                res += "Дисс. " + diss.Degree + " по специальности " + diss.Branch + ". ";
                res += diss.City + ": " + diss.Publisher + ", " + diss.Year + ". ";
                res += diss.PageCount + " с.";
                res += (diss.Reference != "") ? "URL: " + diss.Reference + ". " : "";
                res += diss.Additional;
            }
            else if (paper is CompilationArticle)
            {
                CompilationArticle art = (CompilationArticle)paper;
                res += " // ";
                res += art.CompilationTitle + ". ";
                res += (art.Volume > 0) ? "Т. " + art.Volume + ". " : "";
                res += art.City + ": " + art.Publisher + ", " + art.Year + ". ";
                res += "С. " + art.FirstPage;
                res += (art.LastPage != "") ? "–" + art.LastPage + ". " : ". ";
                res += (art.Reference != "") ? "URL: " + art.Reference + ". " : "";
                res += art.Additional;
            }
            //Create Numerated List Element
            Paragraph par = new Paragraph();
            par.ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines { After = "0" });
            if (listNumber != -1)
            {
                ParagraphProperties prop = new ParagraphProperties();
                prop.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };
                prop.NumberingProperties = new NumberingProperties();
                prop.NumberingProperties.NumberingLevelReference = new NumberingLevelReference() { Val = 0 };
                prop.NumberingProperties.NumberingId = new NumberingId() { Val = listNumber };
                par.AppendChild(prop);
            }
            par.AppendChild(new Run(new Text(res)));
            return par;
        }

        private string concatAuthorsNames(List<Person> authors)
        {
            if (authors == null || authors.Count == 0)
                return "";
            Person last = authors.Last<Person>();
            string result = "";
            foreach (Person author in authors)
                result += author.GetName() + (last == author ? "" : ", ");
            return result;
        }
    }
}
