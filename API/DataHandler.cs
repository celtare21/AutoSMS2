using AutoSMS.Common;
using AutoSMS.Excel;
using Twilio.Rest.Api.V2010.Account;
using Twilio.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Timer = System.Timers.Timer;
// ReSharper disable ConditionalInvocation

namespace AutoSMS2.API
{
    public static class DataHandler
    {
        public static IEnumerable<BirthdayEntry> Entries;
        public static event EventHandler NewData;
        private static readonly Timer Timer;
        public static bool DbBusy;

        static DataHandler()
        {
            Timer = new Timer(TimeSpan.FromMinutes(30).TotalMilliseconds);
            Timer.Elapsed += Timer_Elapsed;

            if (Program.LoadExcel)
                ExcelMain.LoadExcelData();

            Task.Run(LoadData).ContinueWith(_ => NewData?.Invoke(default, EventArgs.Empty)).ContinueWith(_ => Timer_Elapsed(null, null));
        }

        private static void LoadData()
        {
            CheckForInternetConnection();

            using (var context = new BirthdayEntriesContext())
            {
                Entries = (from bd in context.BirthdayEntries
                           where bd.Birthday.Day == DateTime.Today.Day
                           where bd.Birthday.Month == DateTime.Today.Month
                           select bd).ToList();
            }

            Timer.Start();
        }

        private static void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Entries == null || DbBusy)
                return;

            DbBusy = true;

            bool update = false;

            LoadData();

            using (var context = new BirthdayEntriesContext())
            {
                foreach (var entry in Entries)
                {
                    if (entry.SmsDate < DateTime.Today)
                    {
                        {
                            var dbEntry = context.BirthdayEntries.Find(entry.Id);

                            if (dbEntry != null)
                            {
                                _ = MessageResource.Create
                                (
                                    body: $"La multi ani lui {entry.Name}",
                                    from: new PhoneNumber("+"),
                                    to: new PhoneNumber(entry.Number)
                                );

                                dbEntry.SmsDate = entry.SmsDate = DateTime.Today;
                                update = true;
                            }
                        }
                    }
                }

                if (update)
                {
                    context.SaveChanges();

                    try
                    {
                        NewData?.Invoke(default, EventArgs.Empty);
                    }
                    catch
                    {
                        MessageBox.Show(@"There was an error!");
                        Application.Exit();
                    }
                }
            }

            DbBusy = false;
        }

        public static void CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                {
                    using (client.OpenRead("http://google.com/generate_204"))
                    {
                        // Do nothing.
                    }
                }
            }
            catch
            {
                MessageBox.Show(@"No internet connection!");
                Application.Exit();
            }
        }
    }
}
