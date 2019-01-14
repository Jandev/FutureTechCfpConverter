using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                foreach (var cfp in cfpCollection.OrderBy(c => c.NameContactPerson))
                {
                    Console.WriteLine($"Processing CFP {cfp.Id}.");
                    stream.WriteLine($"Id:\t\t\t\t\t\t{cfp.Id}");
                    if(!string.IsNullOrEmpty(cfp.NameContactPerson)) stream.WriteLine($"Contactperson:\t \t\t{cfp.NameContactPerson}");
                    if(!string.IsNullOrEmpty(cfp.EmailContactPerson)) stream.WriteLine($"Contactperson e-mail:\t{cfp.EmailContactPerson}");
                    if(!string.IsNullOrEmpty(cfp.NameSpeaker)) stream.WriteLine($"Speaker:\t\t\t\t{cfp.NameSpeaker}");
                    if(!string.IsNullOrEmpty(cfp.EmailSpeaker)) stream.WriteLine($"Speaker e-mail\t\t\t{cfp.EmailSpeaker}");
                    if(!string.IsNullOrEmpty(cfp.MobilePhoneSpeaker)) stream.WriteLine($"Mobile phonenumber:\t\t{cfp.MobilePhoneSpeaker}");
                    if(!string.IsNullOrEmpty(cfp.ExtraNotesSpeaker)) stream.WriteLine($"Extra notes speaker:\t{cfp.ExtraNotesSpeaker}");
                    if(!string.IsNullOrEmpty(cfp.TitleTalk)) stream.WriteLine($"Title:\t\t\t\t\t{cfp.TitleTalk}");
                    if(!string.IsNullOrEmpty(cfp.DescriptionTalk)) stream.WriteLine($"Description:\t\t\t{cfp.DescriptionTalk}");
                    if(!string.IsNullOrEmpty(cfp.SessionType)) stream.WriteLine($"Type:\t\t\t\t\t{cfp.SessionType}");
                    if(!string.IsNullOrEmpty(cfp.Tracks)) stream.WriteLine($"Tracks:\t\t\t\t\t{cfp.Tracks}");
                    if(!string.IsNullOrEmpty(cfp.Difficulty)) stream.WriteLine($"Difficulty:\t\t\t\t{cfp.Difficulty}");
                    if(!string.IsNullOrEmpty(cfp.TargetGroupDeveloper)) stream.WriteLine("Developer:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupDeveloper) ? "Yes" : "No");
                    if(!string.IsNullOrEmpty(cfp.TargetGroupArchitect)) stream.WriteLine("Architect:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupArchitect) ? "Yes" : "No");
                    if(!string.IsNullOrEmpty(cfp.TargetGroupManagement)) stream.WriteLine("Management:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupManagement) ? "Yes" : "No");
                    if(!string.IsNullOrEmpty(cfp.TargetGroupDisruptor)) stream.WriteLine("Disruptor:\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupDisruptor) ? "Yes" : "No");
                    if(!string.IsNullOrEmpty(cfp.TargetGroupExpert)) stream.WriteLine("Expert:\t\t\t\t\t{0}", !string.IsNullOrWhiteSpace(cfp.TargetGroupExpert) ? "Yes" : "No");
                    if (!string.IsNullOrEmpty(cfp.ExtraNotesSession)) stream.WriteLine($"Extra notes session:\t{cfp.ExtraNotesSession}");
                    stream.WriteLine($"Rating:");
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