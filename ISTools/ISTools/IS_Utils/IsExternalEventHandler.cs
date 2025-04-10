using Autodesk.Revit.UI;
using System;

/// <summary>
/// Class for create ExternalEventHandler with custom external action
/// </summary>
public class IsExternalEventHandler : IExternalEventHandler
{
    private Action<UIApplication> _action;
    private ExternalEvent _externalEvent;

    public IsExternalEventHandler(Action<UIApplication> action, ExternalEvent externalEvent)
    {
        _action = action;
        _externalEvent = externalEvent;
    }

    public void Execute(UIApplication app)
    {
        try
        {
            _action?.Invoke(app);
        }
        finally
        {
            _externalEvent?.Dispose();
        }
    }

    public string GetName()
    {
        return "My Dynamic External Event Handler";
    }
}