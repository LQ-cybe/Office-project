using System.Windows;
using System.Windows.Controls;

namespace SCADAServis.Controls
{
    /// <summary>
    /// Interaction logic for ganttRowUC.xaml
    /// </summary>
    public partial class ganttRowUC : UserControl
    {
        public ganttRowUC()
        {
            InitializeComponent();
            VisualStateManager.GoToState(this, "IsMouseOut", false);
            this.MouseLeave += (o, e) => VisualStateManager.GoToState(this, "IsMouseOut", true);
            this.MouseMove += (o, e) => VisualStateManager.GoToState(this, "IsMouseIn", true);
            this.MouseEnter += (o, e) => VisualStateManager.GoToState(this, "IsMouseIn", true);
        }

        private void userControl_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
