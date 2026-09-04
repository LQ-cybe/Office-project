using System;
using System.Collections.Generic;

namespace SCADAServis.Controls
{
    public class GanttItem
    {
        public string Identity { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public ganttItemUC Control { get; set; }
        public Dictionary<string, object> Data { get; set; }
        public int? StyleID { get; set; }
        public System.Windows.Media.Color Color { get; set; }
        public System.Windows.Media.Color? Color2 { get; set; }
        public string Caption { get; set; }
        public object Tag { get; set; }
        public bool IsVisible { get; set; }
    }
}
