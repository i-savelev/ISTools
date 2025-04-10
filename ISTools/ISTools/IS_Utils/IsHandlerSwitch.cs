using Autodesk.Revit.UI;
using System;

/// <summary>
/// Class for transfer custom action in ExternalEventHandler
/// in order to trigger an ExternalEventHandler you need write the folowing code:
/// IS_Handler_Switch my_handelr = new IS_Handler_Switch((app) =>
/// {
///     any action
///     ... 
/// });
/// my_handelr.External_event.Raise()
/// </summary>
public class IsHandlerSwitch
{
    public Action<UIApplication> CustomAction { get; set; }
    public ExternalEvent ExternalEvent { get; set; }

    public IsHandlerSwitch(Action<UIApplication> customAction)
    {
        ExternalEvent externalEvent = null;
        CustomAction = customAction;
        var eventHandler = new IsExternalEventHandler(CustomAction, externalEvent);
        ExternalEvent = ExternalEvent.Create(eventHandler);
    }
}
