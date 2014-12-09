using PaperMgr.Entity;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PaperMgr
{
    /// <summary>
    /// Interface for PaperMgr Bibliography writers
    /// </summary>
    interface IBibWriter
    {
        /// <summary>
        /// Write records from <paramref name="papers"/> to bibliography file
        /// </summary>
        /// <param name="papers">Collection of <typeparamref name="Paper"/> should be written to bibliography</param>
        void PrepareBibliography(ICollection<Paper> papers);
        /// <summary>
        /// Write records from <paramref name="papers"/> to bibliography file
        /// </summary>
        /// <param name="papers">Collection of <typeparamref name="Paper"/> should be written to bibliography</param>
        /// <returns>Biblography Preparation <typeparamref name="Task"/></returns>
        Task PrepareBibliographyAsync(ICollection<Paper> papers);
    }
}
