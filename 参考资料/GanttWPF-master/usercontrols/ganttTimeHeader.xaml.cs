using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace SCADAServis.Controls
{
    /// <summary>
    /// Interaction logic for ganttTimeHeader.xaml
    /// </summary>
    public delegate void timeLineArgs(DateTime start, DateTime end);
    public partial class ganttTimeHeader : UserControl
    {
        public ganttTimeHeader()
        {
            InitializeComponent();
            this.SizeChanged += (o, e) => { try { Update(); } catch { } };
        }

        DateTime start;
        DateTime end;
        public ganttUC Gantt { get { return common.MainControl; } }
        //public event EventHandler OnMoveLeft;
        //public event EventHandler OnMoveRight;
        //public event timeLineArgs TimeLineChanged;

        TimeSpan length { get { return Gantt.End - Gantt.Start; } }

        double timeRange { get { return (Gantt.End - Gantt.Start).Ticks; } }
        double pointRange { get { return (this.ActualWidth); } }
        double ratio { get { return pointRange / timeRange; } }

        public void SetTimePosition(double pointX)
        {
            timePointerMini.Margin = new Thickness(pointX - (timePointerMini.Width / 2) + 3, 35, 0, 0);
            timePointerMini.lTime.Text = Gantt.Position.ToString("HH:mm:ss");
        }


        public IEnumerable<ganttTimePoint> GetTimePoints()
        {
            TimeSpan major;
            TimeSpan minor;
            //DateTime firstTime;

            if (length.TotalMinutes < 20)
            {
                major = new TimeSpan(0, 1, 0);
                minor = new TimeSpan(0, 0, 10);
            }
            else
                if (length.TotalHours < 2)
                {
                    major = new TimeSpan(0, 10, 0);
                    minor = new TimeSpan(0, 1, 0);
                }
                else
                {
                    major = new TimeSpan(1, 0, 0);
                    minor = new TimeSpan(0, 10, 0);
                }

            var majorPointCount = major.Ticks * ratio;
            var minorPointCount = minor.Ticks * ratio;

            double pointX;
            DateTime timePosition;

            // M I N O R  points
            pointX = 0;
            timePosition = start;
            while (pointX < pointRange)  //opakuj dokd nejsi na časové osy konci
            {
                ganttTimePoint timePoint = new ganttTimePoint();

                pointX += minorPointCount;
                timePoint.Point = pointX;
                timePoint.IsMajor = false;
                timePoint.DateTime = start.Add(minor);
                yield return timePoint;
            }

            // M A J O R   points
            pointX = 0;
            timePosition = start;
            while (pointX < pointRange)  //opakuj dokd nejsi na časové osy konci
            {
                ganttTimePoint timePoint = new ganttTimePoint();

                pointX += majorPointCount;
                timePoint.Point = pointX;
                timePoint.IsMajor = true;
                timePoint.DateTime = start.Add(major);
                yield return timePoint;
            }

        }


        public void Update()
        {
            start = Gantt.Start;
            end = Gantt.End;
            lStart.Text = start.ToString("HH:mm:ss");
            lEnd.Text = end.ToString("HH:mm:ss");
            timeLine.Children.Clear();

            foreach (var point in GetTimePoints())
            {
                timeLine.Children.Add(new Line
                {
                    X1 = point.Point,
                    Y1 = point.IsMajor ? 0 : 3,
                    X2 = point.Point,
                    Y2 = 8,
                    Stroke = point.IsMajor ? new SolidColorBrush(Colors.Gray) : new SolidColorBrush(Colors.LightGray),
                    StrokeThickness = 1
                });
            }
        }

        private void UserControl_MouseWheel(object sender, MouseWheelEventArgs e)
        {

        }
    }
}
