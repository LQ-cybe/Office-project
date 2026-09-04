using System.Windows;
using System.Windows.Controls;

namespace SCADAServis.Controls
{
    /// <summary>
    /// Interaction logic for ganttCalendar.xaml
    /// </summary>
    public partial class ganttCalendar : UserControl
    {
        ganttUC gantt { get { return common.MainControl; } }
        public ganttCalendar()
        {
            InitializeComponent();
        }

        private void calendar_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            gantt.GoTo(calendar.SelectedDate.Value);
            this.Visibility = Visibility.Collapsed;
        }

        private void bClose_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Collapsed;
        }
    }
}
