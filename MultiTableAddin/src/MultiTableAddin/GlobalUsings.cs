// 解决 WPF + WinForms 同时启用时的类型歧义
// 注意: UserControl 和 TextBox 别名需要在各 WPF 视图文件中单独声明，
//       因为 TaskPaneHostControl.cs 需要使用 WinForms 版本
global using Brush = System.Windows.Media.Brush;
global using Brushes = System.Windows.Media.Brushes;
global using Color = System.Windows.Media.Color;
global using Point = System.Windows.Point;
global using Rectangle = System.Windows.Shapes.Rectangle;
global using Binding = System.Windows.Data.Binding;
global using Button = System.Windows.Controls.Button;
global using ComboBox = System.Windows.Controls.ComboBox;
global using Orientation = System.Windows.Controls.Orientation;
global using DragDropEffects = System.Windows.DragDropEffects;
global using HorizontalAlignment = System.Windows.HorizontalAlignment;
global using FlowDirection = System.Windows.FlowDirection;
global using MouseEventArgs = System.Windows.Input.MouseEventArgs;
global using DragEventArgs = System.Windows.DragEventArgs;
