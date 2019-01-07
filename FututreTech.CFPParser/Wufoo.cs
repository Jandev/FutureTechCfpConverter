using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using FututreTech.CFPParser.Model.Excel;

namespace FututreTech.CFPParser
{
    class Wufoo
    {
        internal static void ExtractCfps()
        {
            var filePath = @"D:\Temp\futuretech.xlsx";

            var cfpCollection = GetCfpCollection(filePath);
            WriteCfpCollectionToTextFile(cfpCollection, @"D:\Temp\FutureTech - Wufoo CFPs.txt");
        }

        private static void WriteCfpCollectionToTextFile(IEnumerable<Model.Excel.Cfp> cfpCollection, string fullPath)
        {
            File.Delete(fullPath);
            using(var stream = File.AppendText(fullPath))
            {
                foreach (var cfp in cfpCollection)
                {
                    Console.WriteLine($"Processing CFP {cfp.Id}.");
                    stream.WriteLine($"Id:\t\t\t\t\t\t{cfp.Id}");
                    stream.WriteLine($"Contactperson:\t \t\t{cfp.NameContactPerson}");
                    stream.WriteLine($"Contactperson e-mail:\t{cfp.EmailContactPerson}");
                    stream.WriteLine($"Speaker:\t\t\t\t{cfp.Id}");
                    stream.WriteLine($"Speaker e-mail\t\t\t{cfp.Id}");
                    stream.WriteLine($"Mobile phonenumber:\t\t{cfp.MobilePhoneSpeaker}");
                    stream.WriteLine($"Extra notes speaker:\t{cfp.ExtraNotesSpeaker}");
                    stream.WriteLine($"Title:\t\t\t\t\t{cfp.TitleTalk}");
                    stream.WriteLine($"Description:\t\t\t{cfp.DescriptionTalk}");
                    stream.WriteLine($"Type:\t\t\t\t\t{cfp.SessionType}");
                    stream.WriteLine($"Tracks:\t\t\t\t\t{cfp.Tracks}");
                    stream.WriteLine($"Difficulty:\t\t\t\t{cfp.Difficulty}");
                    stream.WriteLine("Developer:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupDeveloper) ? "Yes" : "No");
                    stream.WriteLine("Architect:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupArchitect) ? "Yes" : "No");
                    stream.WriteLine("Management:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupManagement) ? "Yes" : "No");
                    stream.WriteLine("Disruptor:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupDisruptor) ? "Yes" : "No");
                    stream.WriteLine("Expert:\t\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupExpert) ? "Yes" : "No");
                    stream.WriteLine($"Extra notes session:\t{cfp.ExtraNotesSession}");
                    stream.WriteLine("-------------------------------------------------");
                }
            }
        }

        private static IEnumerable<Model.Excel.Cfp> GetCfpCollection(string filePath)
        {
            bool processedFirstLine = false;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            if (!processedFirstLine)
                            {
                                processedFirstLine = true;
                            }
                            else
                            {
                                yield return MapToModel(reader);
                            }
                        }
                    } while (reader.NextResult());
                }
            }
        }

        private static Model.Excel.Cfp MapToModel(IExcelDataReader reader)
        {
            var model = new Cfp
            {
                Id = reader.GetString(0),
                NameContactPerson = reader.GetString(1),
                EmailContactPerson = reader.GetString(2),
                NameSpeaker = reader.GetString(3),
                EmailSpeaker = reader.GetString(4),
                MobilePhoneSpeaker = reader.GetString(5),
                ExtraNotesSpeaker = reader.GetString(6),
                TitleTalk = reader.GetString(7),
                DescriptionTalk = reader.GetString(8),
                SessionType = reader.GetString(9),
                Tracks = reader.GetString(10),
                Difficulty = reader.GetString(11),
                TargetGroupDeveloper = reader.GetString(12),
                TargetGroupArchitect = reader.GetString(13),
                TargetGroupManagement = reader.GetString(14),
                TargetGroupDisruptor = reader.GetString(15),
                TargetGroupExpert = reader.GetString(16),
                ExtraNotesSession = reader.GetString(17)  
            };

            return model;
        }
    }
}