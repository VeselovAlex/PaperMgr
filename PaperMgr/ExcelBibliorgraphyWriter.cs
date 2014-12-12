using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PaperMgr.Entity;
using System.Collections.Generic;
using System.Linq;

namespace PaperMgr.FileWriters
{
    /// <summary>
    /// Class for creation bibliography as Word document
    /// </summary>
    class ExcelBibliographyWriter : BibWriter
    {
        /// <summary>
        /// Constructor for Word Bibliography Writer (List style)
        /// </summary>
        /// <param name="filename">Path to new file</param>
        /// <param name="author">Current author</param>
        public ExcelBibliographyWriter(string filename)
            : base(filename)
        {
            FileReadyMsg = "Документ Excel готов.";
        }

        protected override void PrepareBibliographyTask(ICollection<Paper> papers)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart part = doc.AddWorkbookPart();
                part.Workbook = new Workbook();

                Sheets sheets = part.Workbook.AppendChild(new Sheets());

                //Group papers by paper type
                var papersByType = from paper in papers
                                   group paper by paper.GetType();

                uint sheetId = 1;
                foreach (var group in papersByType)
                {
                    //Create sheet for category
                    var keyType = group.Key;
                    string sheetName = "";
                    if (keyType == typeof(JournalArticle))
                        sheetName = "Статьи в журналах";
                    else if (keyType == typeof(JournalMessage))
                        sheetName = "Сообщения";
                    else if (keyType == typeof(CompilationArticle))
                        sheetName = "Статьи в сборниках";
                    else if (keyType == typeof(Dissertation))
                        sheetName = "Диссертации";

                    WorksheetPart sheetPart = part.AddNewPart<WorksheetPart>();
                    SheetData shData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(shData);

                    Sheet sheet = new Sheet() {Id = part.GetIdOfPart(sheetPart), SheetId = sheetId++, Name = sheetName};
                    sheets.Append(sheet);
                    shData.Append(HeaderRow(keyType));

                    uint j = 2;
                    foreach (Paper p in group)
                    {
                        shData.Append(PaperRow(p, j++));
                    }
                    part.Workbook.Save();
                }
                doc.Close();
            }
        }

        protected Row HeaderRow(System.Type type)
        {
            Row row = new Row { RowIndex = 1 };
            RefBuilder refs = new RefBuilder(1, 'A'); 
            if (type.IsSubclassOf(typeof(Paper)))
            {
                appendCell(row, "Авторы", refs.Ref);
                appendCell(row, "Название", refs.Ref);
                appendCell(row, "Год", refs.Ref);
                appendCell(row, "Число страниц", refs.Ref);
                appendCell(row, "Метка", refs.Ref);
                appendCell(row, "Ссылка", refs.Ref);
                appendCell(row, "Дополнительно", refs.Ref);
            }

            if (type == typeof(Dissertation))
            {
                appendCell(row, "Специальность", refs.Ref);
                appendCell(row, "Степень", refs.Ref);
                appendCell(row, "Город", refs.Ref);
                appendCell(row, "Издательство", refs.Ref);
            }
            else if (type.IsSubclassOf(typeof(JournalPaper)))
            {
                appendCell(row, "Первая страница", refs.Ref);
                appendCell(row, "Последяя страница", refs.Ref);
                appendCell(row, "Название журнала", refs.Ref);
                appendCell(row, "Номер журнала", refs.Ref);
                appendCell(row, "Том", refs.Ref);
                appendCell(row, "В ВАК", refs.Ref);
                appendCell(row, "В SCOPUS", refs.Ref);
                appendCell(row, "В Web of Science", refs.Ref);
            }
            else if (type == typeof(CompilationArticle))
            {
                appendCell(row, "Первая страница", refs.Ref);
                appendCell(row, "Последяя страница", refs.Ref);
                appendCell(row, "Том", refs.Ref);
                appendCell(row, "Город", refs.Ref);
                appendCell(row, "Издательство", refs.Ref);
                appendCell(row, "Название сборника", refs.Ref);
            }

            return row;
        }

        protected Row PaperRow(Paper paper, uint idx)
        {
            Row row = new Row { RowIndex = idx};
            string authors = concatAuthorsNames(paper.Authors);
            RefBuilder refs = new RefBuilder(idx, 'A'); 
            appendCell(row, authors, refs.Ref);
            appendCell(row, paper.Title, refs.Ref);
            appendCell(row, paper.Year, refs.Ref);
            appendCell(row, paper.PageCount, refs.Ref);
            appendCell(row, Labels.ToString(paper.Label), refs.Ref);
            appendCell(row, paper.Reference, refs.Ref);
            appendCell(row, paper.Additional, refs.Ref);

            if (paper is Dissertation)
            {
                Dissertation disser = (Dissertation)paper;
                appendCell(row, disser.Branch, refs.Ref);
                appendCell(row, disser.Degree, refs.Ref);
                appendCell(row, disser.City, refs.Ref);
                appendCell(row, disser.Publisher, refs.Ref);
            }
            else if (paper is JournalPaper)
            {
                JournalPaper journal = (JournalPaper)paper;
                appendCell(row, journal.FirstPage, refs.Ref);
                appendCell(row, journal.LastPage, refs.Ref);
                appendCell(row, journal.JournalTitle, refs.Ref);
                appendCell(row, journal.JournalNumber, refs.Ref);
                appendCell(row, journal.Volume, refs.Ref);
                appendCell(row, intValue(journal.IsInWACList), refs.Ref);
                appendCell(row, intValue(journal.IsInScopus), refs.Ref);
                appendCell(row, intValue(journal.IsInWoS), refs.Ref);
            }
            else if (paper is CompilationArticle)
            {
                CompilationArticle article = (CompilationArticle)paper;
                appendCell(row, article.FirstPage, refs.Ref);
                appendCell(row, article.LastPage, refs.Ref);
                appendCell(row, article.Volume, refs.Ref);
                appendCell(row, article.City, refs.Ref);
                appendCell(row, article.Publisher, refs.Ref);
                appendCell(row, article.CompilationTitle, refs.Ref);
            }
            return row;
        }

        private int intValue(bool val) { return val ? 1 : 0; }

        private class RefBuilder
        {
            public RefBuilder(uint row, char col)
            {
                this.row = row;
                this.col = col;
            }

            private uint row = 1;
            private char col = 'A';

            public string Ref
            {
                get
                {
                    return (col++).ToString() + row;
                }
            }
        }

        private void appendCell(Row row, object obj, string cellRef)
        {
            Cell cell = new Cell(new CellValue(obj.ToString()));
            cell.CellReference = cellRef;
            cell.DataType = obj is int ? CellValues.Number : CellValues.String;
            row.AppendChild(cell);
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
