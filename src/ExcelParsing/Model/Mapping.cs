using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParsing.Model
{
    public class ColumnMapping
    {
        public Dictionary<string, string> Map { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public void Add(string excelColumnName, string entityPropertyName)
        {
            Map[excelColumnName] = entityPropertyName;
        }
    }
}
