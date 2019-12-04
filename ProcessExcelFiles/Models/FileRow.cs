using System.Collections.Generic;
namespace ProcessExcelFiles
{
    public class FileRow
    {
        public IEnumerable<FileCell> Cells { get; set; }
    }
}