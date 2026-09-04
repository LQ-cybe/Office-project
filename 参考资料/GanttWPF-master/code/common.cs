using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace SCADAServis.Controls
{
    static class common
    {
        public static ganttUC MainControl { get; set; }
        public static Style Style(string styleName)
        {
            return MainControl.Resources[styleName] as Style;
        }
        public static Brush Brush(string styleName)
        {
            return MainControl.Resources[styleName] as Brush;
        }
        public static object Resource(string styleName)
        {
            return MainControl.Resources[styleName];
        }

        public static List<UIElement> GetChildren(this DependencyObject parent, Type targetType)
        {
            List<UIElement> elements = new List<UIElement>();

            int count = VisualTreeHelper.GetChildrenCount(parent);
            if (count > 0)
            {
                for (int i = 0; i < count; i++)
                {
                    UIElement child = (UIElement)VisualTreeHelper.GetChild(parent, i);
                    if (child.GetType() == targetType || targetType.IsAssignableFrom(child.GetType()))
                    {
                        elements.Add(child);
                    }
                    elements.AddRange(GetChildren(child, targetType));
                }
            }
            return elements;
        }

        public static T FindControlWithTag<T>(this DependencyObject parent, string tag) where T : UIElement
        {
            List<UIElement> elements = new List<UIElement>();

            int count = VisualTreeHelper.GetChildrenCount(parent);
            if (count > 0)
            {
                for (int i = 0; i < count; i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                    if (typeof(FrameworkElement).IsAssignableFrom(child.GetType())
                        && ((string)((FrameworkElement)child).Tag == tag))
                    {
                        return child as T;
                    }
                    var item = FindControlWithTag<T>(child, tag);
                    if (item != null) return item as T;
                }
            }
            return null;
        }



    }
}
