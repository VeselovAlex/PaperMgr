using PaperMgr.Entity;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PaperMgr.FileWriters
{
    abstract class BibWriter : IBibWriter
    {
        /// <summary>
        /// Constructor for bibliography file writer
        /// </summary>
        /// <param name="fileName">Path to new bibliography file</param>
        /// <exception cref="System.IO.FileLoadException">Throw if new file path is null or file already exists</exception>
        public BibWriter(string fileName)
        {
            FileName = fileName;
        }
        
        private string mFileName;
        /// <summary>
        /// Represents current file path
        /// </summary>
        /// <exception cref="System.IO.FileLoadException">Throw if new file path is null or file already exists</exception>
        public string FileName
        {
            get
            {
                return mFileName;
            }
            protected set
            {
                if (value == null)
                    throw new FileLoadException("Invalid file path");
                else if (File.Exists(value))
                    throw new FileLoadException("File " + value + " already exists", value);
                else
                    mFileName = value;
            }
        }
        protected string FileReadyMsg { get; set; }
        /// <summary>
        /// Method to prepare biblography file
        /// </summary>
        /// <param name="papers"></param>
        protected abstract void PrepareBibliographyTask(ICollection<Paper> papers);
        /// <summary>
        /// Method to call after file preparation
        /// </summary>
        protected virtual void OnPrepComplete()
        {
            DialogResult result = MessageBox.Show("Файл " + this.FileName + " готов. Открыть?", FileReadyMsg,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (result == DialogResult.Yes)
            {
                FileInfo info = new FileInfo(this.FileName);
                System.Diagnostics.Process.Start(info.FullName);
            }
        }
        public void PrepareBibliography(ICollection<Paper> papers)
        {
            this.PrepareBibliographyTask(papers);
            OnPrepComplete();
        }
        public Task PrepareBibliographyAsync(ICollection<Paper> papers)
        {
            return Task.Run(() => PrepareBibliography(papers));
        }
    }
}
