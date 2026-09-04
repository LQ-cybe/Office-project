using System;

namespace MultiTableAddin.Views;

public interface IRequestClose
{
    event EventHandler? RequestClose;
}
