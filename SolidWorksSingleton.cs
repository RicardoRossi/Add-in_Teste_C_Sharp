using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;

namespace Add_in_Teste_C_Sharp
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
