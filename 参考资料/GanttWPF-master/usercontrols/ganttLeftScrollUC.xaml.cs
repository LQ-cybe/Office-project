using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SCADAServis.Controls
{
    /// <summary>
    /// Interaction logic for leftScrollUC.xaml
    /// </summary>
    public partial class GanttLeftScrollUC : UserControl
    {
        public GanttLeftScrollUC()
        {
            this.InitializeComponent();
            try
            {
                VisualStateManager.GoToState(this, "IsInvisible", false);
                VisualStateManager.GoToState(this, "IsMouseOut", false);
                root.MouseEnter += delegate { VisualStateManager.GoToState(this, "IsShown", true); };
                root.MouseLeave += delegate { VisualStateManager.GoToState(this, "IsInvisible", true); };
                path.MouseEnter += delegate { VisualStateManager.GoToState(this, "IsMouseIn", true); };
                path.MouseLeave += delegate { VisualStateManager.GoToState(this, "IsMouseOut", true); };
                path.MouseLeftButtonDown += new MouseButtonEventHandler(path_MouseLeftButtonDown);
            }
            catch { }//design mode
        }

        void path_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            common.MainControl.MoveLeft(5);
        }
    }
}