using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.DatabaseServices;

namespace WolviesTools
{
    class LuxCloseout
    {
        [CommandMethod("LuxCloseout")]
        public void Closeout()
        {
            DocumentCollection acDocs = Application.DocumentManager;
            foreach (Document acDoc in acDocs)
            {

                //Skip if there is a Drawing1, Drawing2, etc open 
                if(acDoc.Name.Contains("Drawing"))
                {
                    continue;
                }

                //activate the document if not already active
                if(acDocs.MdiActiveDocument != acDoc)
                {
                    acDocs.MdiActiveDocument = acDoc;
                }
                
                Database acCurDb = acDoc.Database;
                using(Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    XrefGraph DbXrGraph = acCurDb.GetHostDwgXrefGraph(false);
                    for (int i = 1; i < DbXrGraph.NumNodes -1; i++)
                    {
                        XrefGraphNode XrGrphNode = DbXrGraph.GetXrefNode(i);

                    }
                    

                }
            }
        }
    }
}
