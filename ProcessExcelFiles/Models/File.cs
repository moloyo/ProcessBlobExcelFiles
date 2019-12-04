using System.Collections.Generic;
using System.Linq;
namespace ProcessExcelFiles
{
    public class File
    {
        public File(IEnumerable<FileRow> rows, Dictionary<int, string> dictionary)
        {
            Dictionary = dictionary;
            Rows = rows;
        }

        public IEnumerable<IEnumerable<FileRow>> BreakRowsIntoStacks(int size = 1000)
        {
            return Rows.Select((row, index) => new { row, group = index / size })
                .GroupBy(obj => obj.group)
                .Select(grouped => grouped.Select(obj => obj.row));
        }

        public IEnumerable<FileRow> Rows { get; set; }

        public Dictionary<int, string> Dictionary { get; set; }
    }
}