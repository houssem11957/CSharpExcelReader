using ExcelParsing.Core;
using ExcelParsing.Entities;
using ExcelParsing.Model;

namespace ExcelParsing
{
    internal class Program
    {
        static string filePath = "people.xlsx";
        static void Main(string[] args)
        {
            var fileInfo = new System.IO.FileInfo(filePath);

            var mapping = new ColumnMapping();
            // excel column name -> Attribute name
            mapping.Add("myId", "Id");
            mapping.Add("Name of the Person", "Name");
            mapping.Add("Date of birth", "DateOfbirth");
            mapping.Add("The Job Title", "JobTitle");

            var items = ExcelReader.ReadToEntities<Person>(filePath, mapping: mapping);
        }
    }
}
