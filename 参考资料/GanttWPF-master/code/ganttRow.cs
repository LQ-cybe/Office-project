using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SCADAServis.Controls
{
    public class GanttRow
    {
        public GanttRow()
        {
            Items = new List<GanttItem>();
        }
        public GanttGroup Group { get; set; }
        public List<GanttItem> Items { get; set; }
        public ganttUC Gantt { get { return Group.Gantt; } }
        public Canvas Control { get; set; }
        public string Caption { get; set; }
        public double? Height { get; set; }
        //ganttPointerUC pointer;

        private double defaultItemVerticalMargin = 6; //new ganttItemUC().Height + new ganttItemUC().Margin.Top + new ganttItemUC().Margin.Bottom;

        public double getItemHeight(GanttRow ganttRow)
        {
            var height = ganttRow.Height ?? Gantt.DefaultItemHeight;
            return height;
        }

        public void Clear()
        {
            Items.Clear();
        }

        public UIElement GetGanttContent()
        {
            //ganttRowBorder
            var container = new Border { Style = common.Style("ganttContentRowBorder") };
            //container.MouseMove += grid_MouseMove;
            var content = new Canvas { Height = getItemHeight(this) + defaultItemVerticalMargin };
            Control = content;
            Control.SizeChanged += new SizeChangedEventHandler(Control_SizeChanged);
            //Control.MouseMove += new System.Windows.Input.MouseEventHandler(Control_MouseMove);
            container.Child = content;
            //content.SizeChanged += (o, e) => content_SizeChanged(o, e, ganttRow);
            //Update();

            return container;

        }

        //void Control_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        //{
        //    if (pointer == null)
        //    {
        //        pointer = new ganttPointerUC();
        //        Control.Children.Add(pointer);
        //        Canvas.SetZIndex(pointer, 30);
        //        pointer.Height = Control.Height;
        //    }

        //    var pointX = e.GetPosition(Control).X;
        //    Canvas.SetLeft(pointer, pointX);            
        //}

        void Control_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateItems();
        }

        double containerWidth { get { return Control.ActualWidth; } }
        double containerHeight { get { return Control.ActualHeight; } }

        double timeRange { get { return (Gantt.End - Gantt.Start).Ticks; } }
        double pointRange { get { return (containerWidth); } }
        double ratio { get { return pointRange / timeRange; } }

        private void updateTimeSeparators()
        {
            Control.GetChildren(typeof(ganttSeparatorUC)).ForEach(s => Control.Children.Remove(s));

            foreach (var point in Gantt.TimeHeader.GetTimePoints().Where(s => s.IsMajor))
            {
                var line = new ganttSeparatorUC
                {
                    Height = Control.Height
                };

                Control.Children.Add(line);
                Canvas.SetLeft(line, point.Point);
                Canvas.SetZIndex(line, 10);
            }
        }

        public void UpdateItems()
        {
            updateTimeSeparators();

            var previousVisibleItems = Items.Where(s => s.IsVisible).ToArray();
            var previousInvisibleItems = Items.Except(previousVisibleItems).ToArray();
            var visibleItems = Items.Where(s => (s.Start < Gantt.End && s.End > Gantt.Start)).ToArray();
            var invisibleItems = Items.Except(visibleItems).ToArray();

            foreach (var item in visibleItems)
            {
                DateTime itemStartTime = item.Start;
                if (itemStartTime < Gantt.Start) itemStartTime = Gantt.Start;
                DateTime itemEndTime = item.End;
                if (itemEndTime > Gantt.End) itemEndTime = Gantt.End;

                double itemstart = (itemStartTime - Gantt.Start).Ticks * Gantt.Ratio;
                double itemend = (itemEndTime - Gantt.Start).Ticks * Gantt.Ratio;
                if (item.Caption == null) item.Caption = "[prázdné]";

                ganttItemUC itemUC;
                if (item.IsVisible)
                {
                    itemUC = item.Control;
                }
                else
                {
                    itemUC = new ganttItemUC
                    {
                        Data = item,
                        Caption = item.Caption
                    };

                    Control.Children.Add(itemUC);
                    if (!item.Color2.HasValue) item.Color2 = item.Color;

                    var itemBrush = new LinearGradientBrush(
                      new GradientStopCollection {
                         new GradientStop { Color = item.Color },
                         new GradientStop { Color = item.Color2.Value, Offset = 1 }
                       }, 90);

                    itemUC.border.Background = itemBrush;
                    item.Control = itemUC;
                }
                itemUC.Width = itemend - itemstart;
                item.IsVisible = true;
                itemUC.lLeft.Visibility = item.Start < Gantt.Start ? Visibility.Visible : Visibility.Collapsed;
                itemUC.lRight.Visibility = item.End > Gantt.End ? Visibility.Visible : Visibility.Collapsed;
                Canvas.SetLeft(itemUC, itemstart);
                Canvas.SetTop(itemUC, defaultItemVerticalMargin / 2);
                Canvas.SetZIndex(itemUC, 20);
            }

            foreach (var item in invisibleItems)
            {
                item.IsVisible = false;
                if (Control.Children.Contains(item.Control)) Control.Children.Remove(item.Control);
            }
        }

    }
}
