using System;

namespace SCADAServis.Controls
{
    public class ganttTimePoint
    {
        public DateTime DateTime { get; set; }
        public double Point { get; set; }
        public bool IsMajor { get; set; }
    }
}
