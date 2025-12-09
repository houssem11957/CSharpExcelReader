using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParsing.Entities
{
    internal class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime DateOfbirth { get; set; }
        public string JobTitle { get; set; }
    }
}
