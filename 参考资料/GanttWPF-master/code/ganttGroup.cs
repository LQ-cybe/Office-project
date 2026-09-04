using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SCADAServis.Controls
{
    public class GanttGroup
    {
        public GanttGroup()
        {
            Rows = new List<GanttRow>();
        }

        public string Caption { get; set; }
        public UIElement Control { get; set; }
        public ganttUC Gantt { get; set; }
        public UIElement PaintControl { get; set; }
        public Grid ContentGrid { get; set; }
        Grid grid { get { return ContentGrid; } }
        public List<GanttRow> Rows { get; set; }
        ganttPointerUC pointer;
        ganttSelectorUC selector;

        public void InitializeGroup()
        {
            Expander expander = Control as Expander;

            var borderContainer = new Border { Style = common.Style("ganttGroupContentBorder") };
            ContentGrid = new Grid { ShowGridLines = false, Opacity = .9 };
            var grid = ContentGrid;
            grid.MouseMove += new System.Windows.Input.MouseEventHandler(grid_MouseMove);
            grid.MouseLeftButtonDown += new System.Windows.Input.MouseButtonEventHandler(grid_MouseLeftButtonDown);
            grid.MouseLeftButtonUp += new System.Windows.Input.MouseButtonEventHandler(grid_MouseLeftButtonUp);
            grid.Background = new SolidColorBrush { Color = Colors.White, Opacity = .1 };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(Gantt.FirstColumnWidth) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.1, GridUnitType.Star) });
            grid.MouseLeave += new System.Windows.Input.MouseEventHandler(grid_MouseLeave);

            ContentGrid = grid;
            var sp = new StackPanel();
            var graphicsContainer = new Canvas { Height = 0 };
            PaintControl = graphicsContainer;
            sp.Children.Add(graphicsContainer);
            sp.Children.Add(grid);

            borderContainer.Child = sp;
            expander.Content = borderContainer;
        }

        public void UpdateRows()
        {
            if (grid == null) InitializeGroup();

            grid.RowDefinitions.Clear();

            for (int i = 0; i < Rows.Count(); i++)
            {
                Rows[i].Items.ForEach(s => { s.IsVisible = false; s.Control = null; });
                Rows[i].Group = this;
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(0, GridUnitType.Auto) });

                var rowUC = new ganttRowUC();
                Grid.SetRow(rowUC, i);
                Grid.SetColumnSpan(rowUC, 2);
                grid.Children.Add(rowUC);

                //set text header of each row //
                var tb = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    Style = common.Style("ganttRowCaptionText"),
                    Text = Rows[i].Caption
                };
                var headerControl = new Border
                {
                    Child = tb,
                    Style = common.Style("ganttRowCaptionBorder")

                };
                Grid.SetRow(headerControl, i);
                grid.Children.Add(headerControl);

                ///// now load content of gantt ////////
                UIElement content = Rows[i].GetGanttContent();
                grid.Children.Add(content);
                Grid.SetRow(content as FrameworkElement, i);
                Grid.SetColumn(content as FrameworkElement, 1);
            }


        }

        void grid_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (Gantt.IsSelectionMode)
            {
                Gantt.ZoomTo(Gantt.SelectionStart, Gantt.SelectionEnd);
                //Gantt.Refresh();
                Gantt.IsSelectionMode = false;
            }
        }


        void grid_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Gantt.IsSelectionMode = false;
        }

        void grid_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var mouseX = e.GetPosition(ContentGrid).X;
            var leftStartX = (ContentGrid.ColumnDefinitions[0].ActualWidth);
            Gantt.SelectionStartX = mouseX - leftStartX;
            e.Handled = false;
        }
        public void UpdatePointer(double left)
        {
            if (pointer == null)
            {
                pointer = new ganttPointerUC();
                pointer.Width = 2;
                pointer.HorizontalAlignment = HorizontalAlignment.Left;
                pointer.VerticalAlignment = VerticalAlignment.Stretch;
                ContentGrid.Children.Insert(0, pointer);
                Grid.SetColumn(pointer, 1);
                if (Rows.Count > 0) Grid.SetRowSpan(pointer, Rows.Count);
            }
            pointer.Margin = new Thickness(left + 1, 0, 0, 0);
        }

        public void UpdateSelector(bool visible, double left, double right, bool headerVisible)
        {
            if (selector == null)
            {
                selector = new ganttSelectorUC();
                selector.Width = 2;
                selector.HorizontalAlignment = HorizontalAlignment.Left;
                selector.VerticalAlignment = VerticalAlignment.Stretch;
                ContentGrid.Children.Add(selector);
                Grid.SetColumn(selector, 1);
                if (Rows.Count > 0) Grid.SetRowSpan(selector, Rows.Count);
            }
            if (left > right) visible = false;
            selector.lStart.Visibility = headerVisible ? Visibility.Visible : Visibility.Collapsed;
            selector.lEnd.Visibility = headerVisible ? Visibility.Visible : Visibility.Collapsed;
            if (!visible) selector.Visibility = Visibility.Collapsed;
            else
            {
                selector.lStart.Text = Gantt.SelectionStart.ToString("HH:mm:ss"); // Gantt.Start.Add(new TimeSpan((long)(left / Gantt.Ratio))).ToString("HH:mm:ss");
                selector.lEnd.Text = Gantt.SelectionEnd.ToString("HH:mm:ss");//Gantt.Start.Add(new TimeSpan((long)(right / Gantt.Ratio))).ToString("HH:mm:ss");
                selector.Visibility = Visibility.Visible;
                selector.Margin = new Thickness(left + 1, 0, 0, 0);
                selector.Width = right - left;
            }
        }


        void grid_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var mouseX = e.GetPosition(ContentGrid).X;
            var leftStartX = (ContentGrid.ColumnDefinitions[0].ActualWidth);
            if (mouseX < leftStartX) { }//hide
            else
            {
                Gantt.Position = Gantt.Start.Add(new TimeSpan((long)((mouseX - leftStartX) / Gantt.Ratio)));

                //setSelectionMode(e.LeftButton, e.GetPosition(ContentGrid));
                Gantt.SelectionEndX = mouseX - leftStartX;
                for (int i = 0; i < Gantt.Groups.Count; i++)
                //update pointer in all groups
                {
                    var group = Gantt.Groups[i];
                    group.UpdatePointer(mouseX - leftStartX);

                    if (Gantt.IsSelectionMode)
                    {
                        group.UpdateSelector(true, Gantt.SelectionStartX, mouseX - leftStartX, i == 0);
                    }
                    //else UpdateSelector(false, 0, 0);
                }

                Gantt.TimeHeader.SetTimePosition(mouseX - leftStartX);
            }
        }

        //private void setSelectionMode(MouseButtonState leftButtonState, Point mousePos)
        //{
        //    if (leftButtonState == System.Windows.Input.MouseButtonState.Pressed)
        //        Gantt.IsSelectionMode = true;
        //    else Gantt.IsSelectionMode = false;


        //}
    }
}