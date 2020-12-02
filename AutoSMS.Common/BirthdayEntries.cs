using System;
using System.Data.Entity;
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace AutoSMS.Common
{
    public class BirthdayEntry
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Number { get; set; }
        public DateTime Birthday { get; set; }
        public DateTime SmsDate { get; set; }
    }

    public class BirthdayEntriesContext : DbContext
    {
        public BirthdayEntriesContext() : base("name=BirthdayEntriesContextString")
        {
        }

        public DbSet<BirthdayEntry> BirthdayEntries { get; set; }
    }
}