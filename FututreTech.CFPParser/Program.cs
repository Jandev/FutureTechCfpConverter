using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using FututreTech.CFPParser.Model.Excel;

namespace FututreTech.CFPParser
{
    class Program
    {
        static void Main(string[] args)
        {
            Wufoo.ExtractCfps();
            Papercall.ExtractCfps();
        }
    }
}
