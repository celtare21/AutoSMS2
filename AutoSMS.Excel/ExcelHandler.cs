using AutoSMS.Common;
using GemBox.Spreadsheet;
using System;
using System.Data;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoSMS.Excel
{
    public static class ExcelMain
    {
        public static bool ExcelBusy;

        public static void LoadExcelData()
        {
            ExcelBusy = true;

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            string path;

            using (var dialog = new OpenFileDialog())
            {
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                dialog.Filter = @"Excel files (*.xlsx)|*.xlsx";
                dialog.FilterIndex = 1;

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                path = dialog.FileName;
            }

            var workbook = ExcelFile.Load(path);
            var worksheet = workbook.Worksheets[0];
            ExcelColumn nameColumn = null, phoneNumberColumn = null, birthDateColumn = null;

            foreach (var cell in worksheet.GetUsedCellRange(true))
            {
                if (cell.ValueType == CellValueType.Null)
                    continue;

                if (cell.StringValue.Contains("Participant Name"))
                    nameColumn = cell.Column;
                else if (cell.StringValue.Contains("Participants' Birth Date"))
                    birthDateColumn = cell.Column;
                else if (cell.StringValue.Contains("Phones"))
                    phoneNumberColumn = cell.Column;

                if (nameColumn != null && phoneNumberColumn != null && birthDateColumn != null)
                    break;
            }

            if (nameColumn == null || phoneNumberColumn == null || birthDateColumn == null)
                throw new NoNullAllowedException();

            var nameStrings = (from nameCell in nameColumn.Cells
                               where nameCell.ValueType != CellValueType.Null &&
                                     !nameCell.StringValue.Contains("Participant Name")
                               select nameCell.StringValue).ToList();

            var phoneStrings = (from phoneCell in phoneNumberColumn.Cells
                                where phoneCell.ValueType != CellValueType.Null &&
                                      !phoneCell.StringValue.Contains("Phones")
                                select phoneCell.StringValue).ToList();

            var birthStrings = (from birthCell in birthDateColumn.Cells
                                where birthCell.ValueType != CellValueType.Null &&
                                      !birthCell.StringValue.Contains("Participants' Birth Date")
                                select birthCell.StringValue).ToList();

            using (var db = new BirthdayEntriesContext())
            {
                db.Database.ExecuteSqlCommand("TRUNCATE TABLE [BirthdayEntries]");

                for (int i = 0; i < nameStrings.Count; i++)
                {
                    if (!string.IsNullOrWhiteSpace(nameStrings[i]) && !string.IsNullOrWhiteSpace(phoneStrings[i]) &&
                        !string.IsNullOrWhiteSpace(birthStrings[i]))
                        db.BirthdayEntries.Add(new BirthdayEntry
                        {
                            Name = nameStrings[i],
                            Number = ProcessPhone(phoneStrings[i]),
                            Birthday = DateTime.Parse(ProcessDates(birthStrings[i])),
                            SmsDate = DateTime.Parse(SqlDateTime.MinValue.ToString())
                        });
                }

                db.SaveChanges();
            }

            ExcelBusy = false;
        }

        private static string ProcessDates(string date)
        {
            var sb = new StringBuilder(date);
            var newDate = new StringBuilder();

            for (int i = 0; i < sb.Length; i++)
            {
                if (sb[i].Equals(char.Parse(@"""")))
                    continue;

                if (sb[i].Equals(char.Parse("-")))
                {
                    newDate.Append("/");
                    continue;
                }

                if (sb[i].Equals(char.Parse("T")))
                    return newDate.ToString();

                newDate.Append(sb[i]);
            }

            throw new ArgumentException();
        }

        private static string ProcessPhone(string number)
        {
            var sb = new StringBuilder(number);
            var newNumber = new StringBuilder();

            for (int i = 0; i < sb.Length; i++)
            {
                if (sb[i].Equals(char.Parse(",")))
                    break;

                newNumber.Append(sb[i]);
            }

            return newNumber.ToString();
        }
    }
}
