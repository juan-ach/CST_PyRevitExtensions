using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CST.PyRevitCSharpTooling.Commands
{
    [Transaction(TransactionMode.Manual)]
    public abstract class CommandBase : IExternalCommand
    {
        protected abstract string CommandKey { get; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return NativeAutomationRunner.Execute(CommandKey, commandData, ref message);
        }
    }
}
