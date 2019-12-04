using System.Collections.Generic;
namespace ProcessExcelFiles
{
    public class FileJob
    {
        public IEnumerable<FileRow> Rows { get; set; }

        public Dictionary<int, string> Dictionary { get; set; }
    }
}