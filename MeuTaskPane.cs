using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Taskpane_ZeroGrau {
    [ComVisible(true)]
    [ProgId("TaskPane2")]
    public partial class MeuTaskPane : UserControl {
        ISldWorks swApp;
        public MeuTaskPane() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            SldWorks swApp = SolidWorksSingleton.GetApplication();
            ModelDoc2 swModel;
            ModelDocExtension swModelExt;
            FeatureManager swFeatureMgr;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelExt = swModel.Extension;

            InsertBOM(swModel);
            swFeatureMgr = swModel.FeatureManager;
            swFeatureMgr.UpdateFeatureTree();
        }

        public void GetSwApp(ISldWorks swAppIn) {
            swApp = swAppIn;
        }

        private static void InsertBOM(ModelDoc2 swmodel) {
            DrawingDoc swDraw = (DrawingDoc)swmodel;
            Sheet currentSheet = swDraw.GetCurrentSheet();
            var views = currentSheet.GetViews();
            SolidWorks.Interop.sldworks.View swSheetView;
            swSheetView = views[0];

            //Check if the view is linked to a BOM
            //var exiteBOM = swSheetView.GetKeepLinkedToBOM();
            var bomName = swSheetView.GetKeepLinkedToBOMName();
            if (String.IsNullOrEmpty(bomName) ) {
                swmodel = swSheetView.ReferencedDocument;
                string config = swSheetView.ReferencedConfiguration;

                var bomPart = @"C:\Users\Ricar\Documents\Add-in Ricardo\Taskpane_ZeroGrau\bin\Release\PART.sldbomtbt";
                var bomAssembly = @"C:\Users\Ricar\Documents\Add-in Ricardo\Taskpane_ZeroGrau\bin\Release\ASSEMBLY.sldbomtbt";

                if ((swDocumentTypes_e)swmodel.GetType() == swDocumentTypes_e.swDocASSEMBLY) {
                    swSheetView.InsertBomTable4(false, 0.413, 0.049, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight,
                       (int)swBomType_e.swBomType_TopLevelOnly, config, bomAssembly, false, (int)swNumberingType_e.swNumberingType_None, false);
                }
                else if ((swDocumentTypes_e)swmodel.GetType() == swDocumentTypes_e.swDocPART) {
                    swSheetView.InsertBomTable4(false, 0.413, 0.049, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight,
                       (int)swBomType_e.swBomType_PartsOnly, config, bomPart, false, (int)swNumberingType_e.swNumberingType_None, false);
                }
            }
        }
    }
}
