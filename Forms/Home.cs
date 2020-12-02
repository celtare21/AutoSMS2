using AutoSMS2.API;
using AutoSMS.Excel;
using Twilio;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoSMS2.Forms
{
    public partial class Home : MetroFramework.Forms.MetroForm
    {
        private readonly SynchronizationContext _syncContext;

        public Home()
        {
            InitializeComponent();

            _syncContext = SynchronizationContext.Current;

            var menuItem = new MenuItem { Index = 0, Text = @"E&xit" };
            menuItem.Click += menuItem_Click;

            var contextMenu = new ContextMenu();
            contextMenu.MenuItems.AddRange(new[] { menuItem });

            notifyIcon.ContextMenu = contextMenu;

            metroListView.BeginUpdate();
            metroListView.Columns.Add("Nume", 180);
            metroListView.Columns.Add("Trimis", 54);
            metroListView.View = View.Details;
            metroListView.GridLines = true;
            metroListView.EndUpdate();

            DataHandler.NewData += DataHandler_NewData;

            FormClosing += Home_FormClosing;
        }

        private void Home_FormClosing(object sender, FormClosingEventArgs e)
        {
            // ReSharper disable once SwitchStatementMissingSomeEnumCasesNoDefault
            switch (e.CloseReason)
            {
                case CloseReason.UserClosing:
                    e.Cancel = true;
                    WindowState = FormWindowState.Minimized;
                    ShowInTaskbar = false;
                    notifyIcon.Visible = true;
                    Hide();
                    break;
                case CloseReason.WindowsShutDown:
                    if (DataHandler.DbBusy || ExcelMain.ExcelBusy)
                    {
                        MessageBox.Show(@"There is ongoing database work!");
                        return;
                    }

                    e.Cancel = true;
                    break;
            }
        }

        private void DataHandler_NewData(object sender, EventArgs e)
        {
            if (DataHandler.Entries == null)
                return;
            
            _syncContext.Post(_ =>
            {
                metroListView.BeginUpdate();

                metroListView.Items.Clear();

                foreach (var entry in DataHandler.Entries)
                {
                    metroListView.Items.Add(new ListViewItem(new[]
                                {entry.Name, entry.SmsDate.Equals(DateTime.Today).ToString()}));
                }

                metroListView.EndUpdate();
            }, null);
        }

        private async void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
            ShowInTaskbar = true;
            await Task.Delay(100);
            notifyIcon.Visible = false;
        }

        private static void menuItem_Click(object sender, EventArgs e) =>
                Application.Exit();

        private void Home_Load(object sender, EventArgs e)
        {
            DataHandler.CheckForInternetConnection();

            TwilioClient.Init("", "");
        }
    }
}
