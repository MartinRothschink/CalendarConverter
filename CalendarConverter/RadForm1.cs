// ***********************************************************************
// Assembly         : CalendarConverter
// Author           : Martin Rothschink
// Created          : 01-13-2021
//
// Last Modified By : Martin Rothschink
// Last Modified On : 01-13-2021
// ***********************************************************************
// <copyright file="RadForm1.cs" company="AxoNet Software GmbH">
//     Copyright © AxoNet Software GmbH 2021
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CalendarConverter
{
   using System;
   using System.Collections.Generic;
   using System.IO;
   using System.Windows.Forms;
   using Telerik.WinControls.UI;
   using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
   using Telerik.Windows.Documents.Flow.Model;

   public partial class RadForm1 : RadForm
   {
      private const string Category = "Gottesdienste";

      private const string Header =
         "VERANSTALTUNG,ORT,VERANSTALTER,STARTDATUM,STARTZEIT,ENDDATUM,ENDZEIT,GANZTÄGIG?,KATEGORIE,BESCHREIBUNG";

      private const string Organizer = "GKG Kirchheim unter Teck";

      private const int ShortDuration = 30;
      private const int StdDuration = 60;

      private RadFlowDocument m_Document;
      private string m_Filename;

      public RadForm1()
      {
         InitializeComponent();
      }

      /// <summary>
      ///    Save clicked, save CSV.
      /// </summary>
      /// <param name="sender">The sender.</param>
      /// <param name="e">The <see cref="EventArgs" /> instance containing the event data.</param>
      private void BtnSaveClick(object sender, EventArgs e)
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

      /// <summary>
      ///    Converts the tables.
      /// </summary>
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
                  var startTime = TimeSpan.MinValue;
                  foreach (var row in table.Rows)
                  {
                     var cellText = new string[4];

                     if (!ProcessRow(row, cellText, ref startTime))
                     {
                        break;
                     }

                     var duration = cellText[2].StartsWith("Rosen") ? ShortDuration : StdDuration;
                     var endTime = startTime.Add(TimeSpan.FromMinutes(duration));
                     Log(
                        $"{cellText[2]},{cellText[1]},{Organizer},{date:d},{startTime:g},{date:d},{endTime:g},,{Category},{cellText[3]}");
                  }
               }
            }
         }
      }

      /// <summary>
      ///    Expands the location.
      /// </summary>
      /// <param name="text">The text.</param>
      /// <returns>System.String.</returns>
      private string ExpandLocation(string text)
      {
         if (text.StartsWith("SU,"))
         {
            return text.Replace("SU,", "St. Ulrich");
         }

         if (text.StartsWith("SM,"))
         {
            return text.Replace("SM,", "St. Markus");
         }

         if (text.StartsWith("PP,"))
         {
            return text.Replace("PP,", "Peter und Paul");
         }

         if (text.StartsWith("HK,"))
         {
            return text.Replace("HK,", "Heilig Kreuz");
         }

         if (text.StartsWith("MK,"))
         {
            return text.Replace("MK,", "Maria Königin");
         }

         if (text.StartsWith("SN,"))
         {
            return text.Replace("SN,", "St. Nikolaus");
         }

         if (text.StartsWith("SL,"))
         {
            return text.Replace("SL,", "St. Lukas");
         }

         return text;
      }

      /// <summary>
      ///    Gets the date above table. Go back some paragraphs until we find a day name.
      /// </summary>
      /// <param name="table">The table.</param>
      /// <returns>DateTime.</returns>
      private DateTime GetDateAboveTable(Table table)
      {
         if (table.Parent is Section section)
         {
            for (var i = 0; i < section.Blocks.Count; ++i)
            {
               if (section.Blocks[i] == table)
               {
                  for (var j = i; j > 0; --j)
                  {
                     if (section.Blocks[j] is Paragraph p)
                     {
                        var text = string.Empty;
                        foreach (var run in p.EnumerateChildrenOfType<Run>())
                        {
                           text = text + run.Text;
                        }

                        if (TextStartsWithDayName(text))
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

      private void Log(string log)
      {
         listBox1.SelectedIndex = listBox1.Items.Add(log);
      }

      private void OpenWord(string fileName)
      {
         var provider = new DocxFormatProvider();
         using (Stream input = File.OpenRead(fileName))
         {
            m_Document = provider.Import(input);
         }

         m_Filename = fileName;
         Text = "Calendar Converter  - " + fileName;
      }

      /// <summary>
      ///    Processes one row of a table. A row has 4 cells with
      ///    Start time, location, event, optional description
      /// </summary>
      /// <param name="row">The row.</param>
      /// <param name="cellText">The cell text.</param>
      /// <param name="startTime">The start time.</param>
      /// <returns><c>true</c> if row is valid, <c>false</c> otherwise.</returns>
      private bool ProcessRow(TableRow row, string[] cellText, ref TimeSpan startTime)
      {
         var i = 0;
         var abort = false;

         foreach (var cell in row.Cells)
         {
            var text = string.Empty;
            foreach (var run in cell.EnumerateChildrenOfType<Run>())
            {
               text = text + run.Text;
            }

            if (i == 0)
            {
               if (!text.Contains("Uhr"))
               {
                  abort = true;
                  break;
               }

               text = text.Replace("Uhr", "");
               text = text.Replace(".", ":");
               startTime = TimeSpan.Parse(text);
            }

            cellText[i++] = i != 1 ? text.Trim() : ExpandLocation(text.Trim());
         }

         return abort;
      }

      /// <summary>
      ///    Handles the Click event of the radButton1 control.
      /// </summary>
      /// <param name="sender">The source of the event.</param>
      /// <param name="e">The <see cref="EventArgs" /> instance containing the event data.</param>
      private void radButton1_Click(object sender, EventArgs e)
      {
         if (openFileDialog1.ShowDialog() == DialogResult.OK)
         {
            OpenWord(openFileDialog1.FileName);
            ConvertTables();
         }
      }

      /// <summary>
      ///    Removes the trailing text after a date.
      /// </summary>
      /// <param name="text">The text.</param>
      /// <returns>System.String.</returns>
      private string RemoveTrailingText(string text)
      {
         var pos = text.IndexOf("–", StringComparison.Ordinal);
         if (pos > 0)
         {
            return text.Substring(0, pos - 1);
         }

         pos = text.IndexOf("-", StringComparison.Ordinal);
         if (pos > 0)
         {
            return text.Substring(0, pos - 1);
         }

         return text;
      }

      /// <summary>
      ///    Chechs if the text starts with a day name.
      /// </summary>
      /// <param name="text">The text.</param>
      /// <returns><c>true</c> if text starts with a day name, <c>false</c> otherwise.</returns>
      private bool TextStartsWithDayName(string text)
      {
         text = text.ToUpper();
         return text.StartsWith("SONNTAG")
                || text.StartsWith("MONTAG")
                || text.StartsWith("DIENSTAG")
                || text.StartsWith("MITTWOCH")
                || text.StartsWith("DONNERSTAG")
                || text.StartsWith("FREITAG")
                || text.StartsWith("SAMSTAG");
      }
   }
}