using System.Windows.Controls;

namespace SCADAServis.Controls
{
    /// <summary>
    /// Interaction logic for ganttTimeTextUC.xaml
    /// </summary>
    public partial class ganttTimeTextUC : UserControl
    {
        public string Time { set { lCaption.Text = value; } }

        public ganttTimeTextUC()
        {
            InitializeComponent();
        }
    }
}
