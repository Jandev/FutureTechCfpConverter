using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FututreTech.CFPParser.Model.Csv;
using TinyCsvParser;
using TinyCsvParser.Mapping;

namespace FututreTech.CFPParser
{
    class Papercall
    {
        internal static void ExtractCfps()
        {
            var filePath = @"D:\Temp\futuretech.csv";
            var cfpCollection = GetCfpCollection(filePath);

            WriteCfpCollectionToTextFile(cfpCollection.Select(c => c.Result), @"D:\Temp\FutureTech - Papercall CFPs.txt");
        }

        private static List<CsvMappingResult<Cfp>> GetCfpCollection(string filePath)
        {
            CsvParserOptions csvParserOptions = new CsvParserOptions(false, ',');
            var csvParser = new CsvParser<Model.Csv.Cfp>(csvParserOptions, new CsvPersonMapping());

            var cfpCollection = csvParser
                .ReadFromFile(filePath, Encoding.UTF8)
                .ToList();
            return cfpCollection;
        }

        private static void WriteCfpCollectionToTextFile(IEnumerable<Cfp> cfpCollection, string fullPath)
        {
            File.Delete(fullPath);
            using (var stream = File.AppendText(fullPath))
            {
                foreach (var cfp in cfpCollection.OrderBy(c => c?.name))
                {
                    if (cfp == null)
                    {
                        continue;
                    }
                    Console.WriteLine($"Processing CFP {cfp.title}.");
                    stream.WriteLine($"Speaker:\t\t\t\t{cfp.name}");
                    stream.WriteLine($"Speaker e-mail\t\t\t{cfp.email}");
                    if(!string.IsNullOrEmpty(cfp.bio)) stream.WriteLine($"Speaker bio:\t\t\t{cfp.bio.Replace("<br>", "\n")}");
                    if(!string.IsNullOrEmpty(cfp.location)) stream.WriteLine($"Speaker location:\t\t{cfp.location}");
                    if (!string.IsNullOrEmpty(cfp.twitter)) stream.WriteLine($"Twitter:\t\t\t\t{cfp.twitter}");
                    if (!string.IsNullOrEmpty(cfp.url))stream.WriteLine($"Url:\t\t\t\t\t{cfp.url}");
                    if (!string.IsNullOrEmpty(cfp.organization))stream.WriteLine($"Organization:\t\t\t{cfp.organization}");
                    if (!string.IsNullOrEmpty(cfp.title))stream.WriteLine($"Title:\t\t\t\t\t{cfp.title}");
                    if (!string.IsNullOrEmpty(cfp.@abstract))stream.WriteLine($"Abstract:\t\t\t\t{cfp.@abstract.Replace("<br>", "\n")}");
                    if (!string.IsNullOrEmpty(cfp.description))stream.WriteLine($"Description:\t\t\t{cfp.description.Replace("<br>", "\n")}");
                    if (!string.IsNullOrEmpty(cfp.audience_level)) stream.WriteLine($"Audience level:\t\t\t{cfp.audience_level}");
                    if (!string.IsNullOrEmpty(cfp.rating))stream.WriteLine($"Rating:\t\t\t\t\t{cfp.rating}");
                    if (!string.IsNullOrEmpty(cfp.talk_format))stream.WriteLine($"Type:\t\t\t\t\t{cfp.talk_format}");
                    if (!string.IsNullOrEmpty(cfp.tags))stream.WriteLine($"Tags:\t\t\t\t\t{cfp.tags}");
                    if (!string.IsNullOrEmpty(cfp.notes))stream.WriteLine($"Extra notes session:\t{cfp.notes.Replace("<br>", "\n")}");
                    if (!string.IsNullOrEmpty(cfp.additional_info)) stream.WriteLine($"Additional info:\t\t{cfp.additional_info.Replace("<br>", "\n")}");
                    stream.WriteLine($"Rating:");
                    stream.WriteLine("-------------------------------------------------");
                }
            }
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