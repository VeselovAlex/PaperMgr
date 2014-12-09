using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperMgr.Entity;
using System.Collections.Generic;
using System.Linq;

namespace PaperMgr.FileWriters
{
    /// <summary>
    /// Class for creation bibliography as Word document
    /// </summary>
    class WordTableBibliographyWriter : BibWriter
    {
        /// <summary>
        /// Constructor for Word Bibliography Writer (List style)
        /// </summary>
        /// <param name="filename">Path to new file</param>
        /// <param name="author">Current author</param>
        public WordTableBibliographyWriter(string filename, Person author)
            : base(filename)
        {
            FileReadyMsg = "Документ Word готов.";
            m_author = author;
        }

        protected Person m_author;
        
        /// <summary>
        /// Class to generate cell
        /// </summary>
        protected static class WordTableCell
        {
            /// <summary>
            /// Generate table cell
            /// </summary>
            /// <param name="text">Cell text</param>
            /// <param name="widthPercent">Width of cell in percent of table width</param>
            /// <param name="justification">Justification of text (Left by default)</param>
            /// <param name="isBold">Bold text flag (false by default)</param>
            /// <returns>Table cell with current text and format</returns>
            public static TableCell CreateWordTableCell(string text, int widthPercent, JustificationValues justification = JustificationValues.Left, bool isBold = false)
            {
                Run r = new Run(new Text(text));
                if (isBold)
                    r.RunProperties = new RunProperties(new Bold());
                Paragraph p = new Paragraph(r);
                p.ParagraphProperties = new ParagraphProperties(new Justification { Val = justification });
                TableCell tc = new TableCell(p);
                TableCellProperties prop = new TableCellProperties();
                prop.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = (5 * widthPercent).ToString() };
                tc.TableCellProperties = prop;
                return tc;
            }
        }
       
        /// <summary>
        /// Creates Bibliography file with Elements of <paramref name="papers"/>
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
                pageTitle.ParagraphProperties = titleParProp;
                Run titleRun = new Run(new Text("СПИСОК"));
                RunProperties titleProp = new RunProperties();
                titleProp.FontSize = new FontSize() { Val = "32" };
                titleProp.Bold = new Bold();
                titleRun.RunProperties = titleProp;
                pageTitle.AppendChild(titleRun);
                body.AppendChild(pageTitle);

                //Create subtitle
                Paragraph pageSubTitle = new Paragraph(new Run(new Text("научных работ")));
                ParagraphProperties subTitleProps = new ParagraphProperties();
                subTitleProps.SpacingBetweenLines = new SpacingBetweenLines { After = "0" };
                subTitleProps.Justification = new Justification() { Val = JustificationValues.Center };
                pageSubTitle.ParagraphProperties = subTitleProps;
                body.AppendChild(pageSubTitle);

                //Create author name paragraph
                Run aRun = new Run(new Text(m_author.GetName()));
                aRun.RunProperties = new RunProperties(new Color() { Val = "red" });
                Paragraph aName = new Paragraph(aRun);
                ParagraphProperties aNameProps = new ParagraphProperties();
                aNameProps.Justification = new Justification() { Val = JustificationValues.Center };
                aName.ParagraphProperties = aNameProps;
                body.AppendChild(aName);

                Table table = new Table();
                AppendTableHeader(table);
                
                int i = 1;
                foreach (Paper paper in papers)
                {
                    table.AppendChild(PaperRow(paper, i++));
                }
                body.AppendChild(table);
                wordDoc.Close();
            }
        }
        
        /// <summary>
        /// Create header for bib table
        /// </summary>
        /// <param name="table">Table to add header</param>
        protected void AppendTableHeader(Table table)
        {
            TableProperties tProp = new TableProperties(
                    new TableBorders(
                        new TopBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 4
                        },
                        new BottomBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 4
                        },
                        new LeftBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 4
                        },
                        new RightBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 4
                        },
                        new InsideHorizontalBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 4
                        },
                        new InsideVerticalBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 4
                        }
                        )
                    );
            table.AppendChild(tProp);

            TableRow numRow = new TableRow();
            TableRow numRow2 = new TableRow();
            int[] sizes = {5, 25, 10, 25, 10, 25};
            for (int j = 1; j <= 6; j++)
            {
                numRow.AppendChild(WordTableCell.CreateWordTableCell(j.ToString(), sizes[j - 1], JustificationValues.Center, true));
                numRow2.AppendChild(WordTableCell.CreateWordTableCell(j.ToString(), sizes[j - 1], JustificationValues.Center, true));
            }
            table.AppendChild(numRow);

            TableRow headRow = new TableRow(WordTableCell.CreateWordTableCell("№ п/п", 5, JustificationValues.Center, true),
                                            WordTableCell.CreateWordTableCell("Наименование работы, ее вид", 25, JustificationValues.Center, true),
                                            WordTableCell.CreateWordTableCell("Форма работы", 10, JustificationValues.Center, true),
                                            WordTableCell.CreateWordTableCell("Выходные данные", 25, JustificationValues.Center, true),
                                            WordTableCell.CreateWordTableCell("Объем, стр.", 10, JustificationValues.Center, true),
                                            WordTableCell.CreateWordTableCell("Соавторы", 25, JustificationValues.Center, true));
            table.AppendChild(headRow);
            table.AppendChild(numRow2);
        }
        
        /// <summary>
        /// Method to create paper row
        /// </summary>
        /// <param name="paper">Source paper</param>
        /// <param name="number">Number of paper row</param>
        /// <exception cref="Exceptions.AuthorMismatchException">Thrown, if current author is not current paper's author</exception>
        /// <returns>Paper row</returns>
        protected TableRow PaperRow(Paper paper, int number)
        {
            var authors = paper.Authors;
            string match = m_author.GetFullName();
            if (authors.RemoveAll(a => Equals(a.GetFullName(), match)) == 0)
                //Current author is not current paper's author
                throw new Exceptions.AuthorMismatchException();
            string coAuthors = concatAuthorsNames(authors);

            string type = "";
            string refEnd = "";

            if (paper is JournalPaper)
            {
                JournalPaper journal = (JournalPaper)paper;
                if (journal is JournalArticle)
                    type = "статья";
                else if (journal is JournalMessage)
                    type = "сообщение";
                refEnd += journal.JournalTitle + ". ";
                refEnd += journal.Year + ". ";
                refEnd += (journal.JournalNumber > 0) ? "№ " + journal.JournalNumber + " " : "";
                refEnd += (journal.Volume > 0) ? "Т. " + journal.Volume + ". " : "";
                refEnd += (journal.Reference != "") ? "URL: " + journal.Reference + ". " : "";
                refEnd += journal.Additional;
            }
            else if (paper is Dissertation)
            {
                Dissertation diss = (Dissertation)paper;
                type = "диссертация";
                refEnd += "Дисс. " + diss.Degree + " по специальности " + diss.Branch + ". ";
                refEnd += diss.City + ": " + diss.Publisher + ", " + diss.Year + ". ";
                refEnd += (diss.Reference != "") ? "URL: " + diss.Reference + ". " : "";
                refEnd += diss.Additional;
            }
            else if (paper is CompilationArticle)
            {
                CompilationArticle art = (CompilationArticle)paper;
                type = "статья";
                refEnd += art.CompilationTitle + ". ";
                refEnd += (art.Volume > 0) ? "Т. " + art.Volume + ". " : "";
                refEnd += art.City + ": " + art.Publisher + ", " + art.Year + ". ";
                refEnd += (art.Reference != "") ? "URL: " + art.Reference + ". " : "";
                refEnd += art.Additional;
            }

            //Row preparation
            TableRow row = new TableRow(WordTableCell.CreateWordTableCell(number.ToString(), 5, JustificationValues.Center),
                                        WordTableCell.CreateWordTableCell(paper.Title + '(' + type + ')', 25),
                                        WordTableCell.CreateWordTableCell(Labels.ToString(paper.Label), 10, JustificationValues.Center),
                                        WordTableCell.CreateWordTableCell(refEnd, 25),
                                        WordTableCell.CreateWordTableCell(paper.PageCount.ToString(), 10, JustificationValues.Center),
                                        WordTableCell.CreateWordTableCell(coAuthors, 25));
            return row;
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
