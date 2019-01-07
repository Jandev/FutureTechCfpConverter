using System.Linq;
using System.Text;
using TinyCsvParser;
using TinyCsvParser.Mapping;

namespace FututreTech.CFPParser
{
    class Papercall
    {
        internal static void ExtractCfps()
        {
            var filePath = @"D:\Temp\futuretech.csv";
            CsvParserOptions csvParserOptions = new CsvParserOptions(false, ',');
            var csvParser = new CsvParser<Model.Csv.Cfp>(csvParserOptions, new CsvPersonMapping());

            var result = csvParser
                .ReadFromFile(filePath, Encoding.UTF8)
                .ToList();
        }

        private class CsvPersonMapping : CsvMapping<Model.Csv.Cfp>
        {
            public CsvPersonMapping()
                : base()
            {
                MapProperty(0, x => x.name);
                MapProperty(1, x => x.email);
                MapProperty(2, x => x.avatar);
                MapProperty(3, x => x.location);
                MapProperty(4, x => x.bio);
                MapProperty(5, x => x.talk_format);
                MapProperty(6, x => x.twitter);
                MapProperty(7, x => x.url);
                MapProperty(8, x => x.organization);
                MapProperty(9, x => x.shirt_size);
                MapProperty(10, x => x.title);
                MapProperty(11, x => x.@abstract);
                MapProperty(12, x => x.description);
                MapProperty(13, x => x.notes);
                MapProperty(14, x => x.audience_level);
                MapProperty(15, x => x.tags);
                MapProperty(16, x => x.rating);
                MapProperty(17, x => x.state);
                MapProperty(18, x => x.confirmed);
                MapProperty(19, x => x.created_at);
                MapProperty(20, x => x.additional_info);
            }
        }
    }
}