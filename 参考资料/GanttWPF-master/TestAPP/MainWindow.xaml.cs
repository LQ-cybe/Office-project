using SCADAServis.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TestAPP
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            List<GanttGroup> list = new List<GanttGroup>();

            for (int i = 0; i < 10; ++i) {
                GanttGroup ganttGroup = new GanttGroup();
                ganttGroup.Caption = "Group " + i;

                    GanttRow ganttRow = new GanttRow();
                    ganttRow.Caption = "test_" + i;

                    GanttItem ganttItem = new GanttItem();
                    ganttItem.Caption = "testItem_" + i;
                    ganttItem.Start = DateTime.Now;
                    ganttItem.End = ganttItem.Start.AddDays(7);
                    ganttRow.Items.Add(ganttItem);

                    ganttGroup.Rows.Add(ganttRow);
 
                list.Add(ganttGroup);
            }

            ganttChart.Groups = list;


            ganttChart.Start = DateTime.Now;
            ganttChart.End = ganttChart.Start.AddDays(14);

            ganttChart.Refresh();

        }
    }
}
