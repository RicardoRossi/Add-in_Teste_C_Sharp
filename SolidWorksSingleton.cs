using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;

namespace Taskpane_ZeroGrau
{
    internal static class SolidWorksSingleton
    {
        private static SldWorks swApp;

        internal static SldWorks GetApplication()
        {
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            return swApp;
        }
    }
}
