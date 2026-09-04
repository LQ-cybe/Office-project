using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SCADAServis.Controls
{
    /////////////////////////////////////////////////////////////////////////////////
    public partial class ganttUC : UserControl
    {
        public event timeLineArgs OnTimeLineChanged;
        //public event EventHandler OnMoveLeft;
        //public event EventHandler OnMoveRight;

        private StackPanel groupHolder = new StackPanel();

        public bool IsSelectionMode { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public DateTime Middle { get { return Start.Add((new TimeSpan((long)((End.Ticks - Start.Ticks) / 2)))); } }
        private DateTime _pointerDateTime { get; set; }
        private DateTime position;
        public double SelectionStartX { get; set; }
        public double SelectionEndX { get; set; }
        public DateTime SelectionStart { get { return this.Start.Add(new TimeSpan((long)(SelectionStartX / Ratio))); } }
        public DateTime SelectionEnd { get { return this.Start.Add(new TimeSpan((long)(SelectionEndX / Ratio))); } }

        public DateTime Position
        {
            get { return position; }
            set
            {
                position = value;
                //if (position < Start) position = Start; if (position > End) position = End;
            }
        }
        //public DateTime Position { get; set; }
        double defaultItemHeight = 0;
        public double DefaultItemHeight
        {
            get
            {
                if (defaultItemHeight == 0) defaultItemHeight = new ganttItemUC().Height;
                return defaultItemHeight;
            }
            set { defaultItemHeight = value; }
        }

        public List<GanttGroup> Groups { get; set; }

        public delegate void PointerMoved(PointerMovedArgs args);
        //public event PointerMoved OnPointerMoved;

        /// <summary>
        /// zobrazený čas na časové ose
        /// </summary>
        public double TimeRange { get { return (End - Start).Ticks; } }

        public double FirstColumnWidth = 200;
        /// <summary>
        /// celkový počet pixelů v časové ose
        /// </summary>
        public double PointRange
        {
            get
            {
                if (Groups != null && Groups.Count > 0 && Groups[0].ContentGrid != null)
                    //return (Groups[0].ContentGrid.ActualWidth - FirstColumnWidth);
                    return TimeHeader.ActualWidth;
                else return 0;
            }
        }

        /// <summary>
        ///  poměr mezi pixely a časem
        /// </summary>
        public double Ratio { get { return PointRange / TimeRange; } }

        public ganttUC()
        {
            common.MainControl = this;
            InitializeComponent();

            if (System.ComponentModel.DesignerProperties.GetIsInDesignMode(this))
            {
                Content = new TextBlock { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center, FontSize = 25, Text = "Gantt chart" };
            }
            else
            { }

            try  // will catch exception when in designMode
            {
                Groups = new List<GanttGroup>();
                this.MouseRightButtonDown += (o, e) => calendar.Visibility = Visibility.Visible;
                //this.MouseLeftButtonDown += (o, e) => calendar.Visibility = Visibility.Hidden;
            }
            catch { }
        }
        void updateGroups(GanttGroup[] groups)
        {
            this.Visibility = (groups != null && groups.Count() > 0) ? Visibility.Visible : Visibility.Hidden;
            //nejprve smazat vše, co už je nepotřebné

            foreach (var group in groupsContainer.Children.OfType<Expander>().Where(s => !groups.Any(g => g.Caption == (s as FrameworkElement).Tag.ToString())))
            {
                groupsContainer.Children.Remove(group);
            }

            // a pak přidat všechny, dosud nepřidané, skupiny :)
            foreach (var group in groups)
            {
                group.Gantt = this;
                Expander groupExpander = groupsContainer.FindControlWithTag<Expander>(group.Caption);
                if (groupExpander == null)
                {
                    groupExpander = new Expander
                    {
                        Style = FindResource("ganttGroupExpander") as Style,
                        IsExpanded = true,
                        Tag = group.Caption,
                        Header = group.Caption,
                    };

                    group.Control = groupExpander;
                    //group.InitializeGroup();
                    groupsContainer.Children.Add(groupExpander);
                }

                group.Control = groupExpander;

                group.UpdateRows();
            }
        }

        private void timeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            switch ((int)e.NewValue)
            {
                case 10: End = Start.AddMinutes(1); break;
                case 9: End = Start.AddMinutes(5); break;
                case 8: End = Start.AddMinutes(10); break;
                case 7: End = Start.AddMinutes(30); break;
                case 6: End = Start.AddHours(1); break;
                case 5: End = Start.AddHours(2); break;
                case 4: End = Start.AddHours(4); break;
                case 3: End = Start.AddHours(12); break;
                case 2: End = Start.AddHours(24); break;
                case 1: End = Start.AddHours(48); break;
                case 0: End = Start.AddHours(64); break;

            }
        }
        public TimeSpan Duration { get { return End - Start; } }

        public TimeSpan SmallStep
        {
            get
            {
                if (Duration.TotalHours > 10) return new TimeSpan(1, 0, 0);
                if (Duration.TotalHours > 5) return new TimeSpan(0, 30, 0);
                if (Duration.TotalMinutes > 50) return new TimeSpan(0, 10, 0);
                if (Duration.TotalMinutes > 10) return new TimeSpan(0, 1, 0);
                return new TimeSpan(0, 0, 30);
            }
        }

        public TimeSpan BigStep
        {
            get
            {
                if (Duration.TotalHours > 10) return new TimeSpan(10, 0, 0);
                if (Duration.TotalHours > 5) return new TimeSpan(0, 60, 0);
                if (Duration.TotalMinutes > 50) return new TimeSpan(0, 20, 0);
                if (Duration.TotalMinutes > 10) return new TimeSpan(0, 5, 0);
                return new TimeSpan(0, 0, 60);
            }
        }

        public void MoveLeft(int value)
        {
            if (value == 5)
            {
                Start = Start.Subtract(BigStep);
                End = End.Subtract(BigStep);
            }
            else
            {
                Start = Start.Subtract(SmallStep);
                End = End.Subtract(SmallStep);
            }

            if (OnTimeLineChanged != null)
            {
                OnTimeLineChanged(Start, End);
            }
            //Recreate();
        }

        public void MoveRight(int value)
        {
            if (value == 5)
            {
                Start = Start.Add(BigStep);
                End = End.Add(BigStep);
            }

            else
            {
                Start = Start.Add(SmallStep);
                End = End.Add(SmallStep);
            }

            if (OnTimeLineChanged != null)
            {
                OnTimeLineChanged(Start, End);
            }
            //Recreate();
        }

        public void ZoomTo(DateTime start, DateTime end)
        {
            Start = start;
            End = end;
            if (OnTimeLineChanged != null)
            {
                OnTimeLineChanged(start, end);
            }
        }

        public void GoTo(DateTime time)
        {
            var half = new TimeSpan((End - Start).Ticks / 2); //nejprve spočítá kolik dní je v rozsahu aktuálně
            ZoomTo(time.Subtract(half), time.Add(half));
            if (OnTimeLineChanged != null)
            {
                OnTimeLineChanged(Start, End);
            }
        }


        private void ganttContainer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (e.Delta < 0) if (Keyboard.IsKeyDown(Key.LeftShift)) MoveLeft(3); else MoveLeft(5);
            else
                if (e.Delta > 0) if (Keyboard.IsKeyDown(Key.LeftShift)) MoveRight(3); else MoveRight(5);
        }


        private void bScrollLeft_Click(object sender, RoutedEventArgs e)
        {
            MoveLeft(5);
        }

        private void bScrollRight_Click(object sender, RoutedEventArgs e)
        {
            MoveRight(5);
        }

        public void Refresh()
        {
            updateGroups(Groups.ToArray());
            TimeHeader.Update();
        }


        //týden
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var half = new TimeSpan(new TimeSpan(7, 0, 0, 0).Ticks / 2);
            ZoomTo(Middle.Subtract(half), Middle.Add(half));
        }

        //den
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var half = new TimeSpan(new TimeSpan(24, 0, 0).Ticks / 2);
            ZoomTo(Middle.Subtract(half), Middle.Add(half));
        }

        //hodina
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var half = new TimeSpan(new TimeSpan(0, 60, 0).Ticks / 2);
            ZoomTo(Middle.Subtract(half), Middle.Add(half));
        }

        //minuta
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var half = new TimeSpan(new TimeSpan(0, 0, 60).Ticks / 2);
            ZoomTo(Middle.Subtract(half), Middle.Add(half));
        }
    }

    public class PointerMovedArgs
    {
        public PointerMovedArgs(DateTime args)
        //, double x)
        {
            this.Value = args; //this.ValueDouble = x;
            //this.FirstContentCanvasLeft = firstContentCanvas;
        }
        public DateTime Value;
        public double ValueDouble;
        public double FirstContentCanvasLeft;
    }
}