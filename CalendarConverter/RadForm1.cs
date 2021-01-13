
namespace CalendarConverter
{
   using System;
   using System.Collections.Generic;
   using System.Drawing.Text;
   using System.IO;
   using System.Windows.Forms;
   using Telerik.Windows.Documents.Flow.Model;

   public partial class RadForm1 : Telerik.WinControls.UI.RadForm
   {
      private const string Header = "VERANSTALTUNG,ORT,VERANSTALTER,STARTDATUM,STARTZEIT,ENDDATUM,ENDZEIT,GANZTÄGIG?,KATEGORIE,BESCHREIBUNG";
      private const string Veranstalter = "GKG Kirchheim unter Teck";
      private const string Category = "Gottesdienste";
      private const int ShortDuration = 30;
      private const int StdDuration = 60;

      private RadFlowDocument m_Document;
      private string m_Filename;
      public RadForm1()
      {
         InitializeComponent();
      }

      private void radButton1_Click(object sender, EventArgs e)
      {
         if (openFileDialog1.ShowDialog() == DialogResult.OK)
         {
            OpenWord(openFileDialog1.FileName);
            ConvertTables();
         }
      }

      private void ConvertTables()
      {
         Log($"{Header}");
         foreach (var section in m_Document.Sections)
         {
            foreach (var table in section.EnumerateChildrenOfType<Table>())
            {
               if (table.GridColumnsCount == 4)
               {
                  var date = GetDateAboveTable(table);
                  TimeSpan startTime = TimeSpan.MinValue;
                  foreach (var row in table.Rows)
                  {
                     bool abort = false;
                     string[] cellText = new string[4];
                     int i = 0;

                     foreach (var cell in row.Cells)
                     {
                        var text = string.Empty;
                        foreach (var run in cell.EnumerateChildrenOfType<Run>())
                           text = text + run.Text;

                        if (i == 0)
                        {
                           if (!text.ToUpper().Contains("UHR"))
                              abort = true;
                           else
                           {
                              text = text.Replace("Uhr", "");
                              text = text.Replace(".", ":");
                              startTime = TimeSpan.Parse(text);
                           }
                        }

                        if (i != 1)
                           cellText[i++] = text.Trim();
                        else
                           cellText[i++] = ExpandLocation(text.Trim());
                     }

                     if (abort)
                        break;

                     var duration = cellText[2].StartsWith("Rosen") ? ShortDuration : StdDuration;
                     TimeSpan endTime = startTime.Add(TimeSpan.FromMinutes(duration));
                     Log($"{cellText[2]},{cellText[1]},{Veranstalter},{date:d},{startTime:g},{date:d},{endTime:g},,{Category},{cellText[3]}");
                  }
               }
            }
         }
      }

      private DateTime GetDateAboveTable(Table table)
      {
         if (table.Parent is Section section)
         {
            for (int i = 0; i < section.Blocks.Count; ++i)
            {
               if (section.Blocks[i] == table)
               {
                  for (int j = i; j > 0; --j)
                  {
                     if (section.Blocks[j] is Paragraph p)
                     {
                        var text = string.Empty;
                        foreach (var run in p.EnumerateChildrenOfType<Run>())
                           text = text + run.Text;

                        if (TextStartsWithDay(text.ToUpper()))
                        {
                           text = RemoveTrailingText(text);
                           var date = DateTime.Parse(text);

                           return date;
                        }
                     }
                  }
               }
            }
         }

         return DateTime.Now;
      }

      private string RemoveTrailingText(string text)
      {
         var pos = text.IndexOf("–", StringComparison.Ordinal);
         if (pos > 0)
            return text.Substring(0, pos - 1);
         
         pos = text.IndexOf("-", StringComparison.Ordinal);
         if (pos > 0)
            return text.Substring(0, pos - 1);

         return text;
      }

      private bool TextStartsWithDay(string text)
      {
         return text.StartsWith("SONNTAG")
                || text.StartsWith("MONTAG")
                || text.StartsWith("DIENSTAG")
                || text.StartsWith("MITTWOCH")
                || text.StartsWith("DONNERSTAG")
                || text.StartsWith("FREITAG")
                || text.StartsWith("SAMSTAG");
      }

      private string ExpandLocation(string text)
      {
         if (text.StartsWith("SU,"))
            return text.Replace("SU,", "St. Ulrich");
         if (text.StartsWith("SM,"))
            return text.Replace("SM,", "St. Markus");
         if (text.StartsWith("PP,"))
            return text.Replace("PP,", "Peter und Paul");
         if (text.StartsWith("HK,"))
            return text.Replace("HK,", "Heilig Kreuz");
         if (text.StartsWith("MK,"))
            return text.Replace("MK,", "Maria Königin");
         if (text.StartsWith("SN,"))
            return text.Replace("SN,", "St. Nikolaus");
         if (text.StartsWith("SL,"))
            return text.Replace("SL,", "St. Lukas");

         return text;
      }

      private void Log(string log)
      {
         listBox1.SelectedIndex = listBox1.Items.Add(log);
      }

      private void OpenWord(string fileName)
      {
         Telerik.Windows.Documents.Flow.FormatProviders.Docx.DocxFormatProvider provider = new Telerik.Windows.Documents.Flow.FormatProviders.Docx.DocxFormatProvider(); 
         using (Stream input = File.OpenRead(fileName)) 
         { 
            m_Document = provider.Import(input); 
         }

         m_Filename = fileName;
         Text = "Calendar Converter  - " + fileName;
      }

      private void radButton2_Click(object sender, EventArgs e)
      {
         saveFileDialog1.FileName = Path.GetFileNameWithoutExtension(m_Filename) + ".csv";
         if (saveFileDialog1.ShowDialog() == DialogResult.OK)
         {
            var list = new List<string>();
            foreach (var listBox1Item in listBox1.Items)
            {
               list.Add(listBox1Item.ToString());
            }
            File.WriteAllLines(saveFileDialog1.FileName, list);
         }
      }
   }
}
