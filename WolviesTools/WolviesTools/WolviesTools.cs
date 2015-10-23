using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Geometry;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;
using acMenu = Autodesk.AutoCAD.Windows.MenuItem;

namespace WolviesTools
{
    //basic tools
    public class Wtools : IExtensionApplication
    {       
        //add context menu
        public void Initialize()
        {
            WolviesContextMenu.Attach();
        }
        //remove context menu
        public void Terminate()
        {
            WolviesContextMenu.Detach();
        }

        //give the total of several distances
        [CommandMethod("rundist")]
        public void RunDist()
        {
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            Editor acEd = acDoc.Editor;
            
            PromptDistanceOptions pdo = new PromptDistanceOptions("\nSelect first point:");
            
            int prec = (Int16) acApp.GetSystemVariable("DIMDEC");
            // List of the points selected and our result object

            List<double> pts = new List<double>();           
            PromptDoubleResult dist;

            // The selection loop
            pdo.AllowNone = true;
            do
            {               
                dist = acEd.GetDistance(pdo);                
                pts.Add(dist.Value);
            }
            while (dist.Status == PromptStatus.OK);

            if (dist.Status == PromptStatus.None)
            {
                double totDist = 0;
                if (pts.Count >= 1)
                {
                    foreach (double item in pts)
                    {
                        totDist = totDist + item;
                    }
                }                
                string arch_dist = Converter.DistanceToString(totDist, DistanceUnitFormat.Architectural, prec);                
                string dec_dist = Converter.DistanceToString(totDist, DistanceUnitFormat.Decimal, prec);                
                
                MessageBox.Show("Total distance is: " + arch_dist + " (" + dec_dist + "\")");

            }
        }        

        //close all open drawings without saving
        [CommandMethod("CloseAllNoSave", CommandFlags.Session)]
        public void CloseAllNoSave()
        {
            Editor ed = acApp.DocumentManager.MdiActiveDocument.Editor;            
            foreach (Document acDoc in acApp.DocumentManager)
            {
                try
                {
                    acDoc.CloseAndDiscard();
                }
                catch (System.Exception e)
                {                        
                    ed.WriteMessage(e.ToString());
                }                               
            }
        }

        //close open drawing without saving
        [CommandMethod("CloseNoSave", CommandFlags.Session)]
        public void CloseNoSave()
        {
            Editor ed = acApp.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            try
            {
                acDoc.CloseAndDiscard();
            }
            catch (System.Exception e)
            {
                ed.WriteMessage(e.ToString());
            }
        }

        //combine groups of text or mtext into one mtext block
        [CommandMethod("jointext")]
        public void JoinText()
        {
            var acObjList = new List<Entity>();
            var acObjList_Mtext = new List<String>();
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                    OpenMode.ForRead) as BlockTable;
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                // Request Mtext object to be selected
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();

                // If OK
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    
                    SelectionSet acSSet = acSSPrompt.Value;

                    // Step through the objects
                    foreach (SelectedObject acSSobj in acSSet)
                    {
                        // Check if valid
                        if (acSSobj != null)
                        {
                            // Open object for read
                            Entity acEnt = acTrans.GetObject(acSSobj.ObjectId,
                                OpenMode.ForRead) as Entity;

                            if (acEnt != null)
                            {
                                // Seperate out text and mtext entities
                                if (acEnt.GetType() == typeof(MText) || acEnt.GetType() == typeof(DBText))
                                {
                                    // add entity to array
                                    acObjList.Add(acEnt);
                                }
                            }
                        }
                    }
                }
             
                if (acObjList.Count != 0)
                {                   
                    //TO-DO:
                    //iterate through acObjList if object is mtext save over to acObjMtext
                    //if objuect is DBtext then convert to mtext and save over to acObjMtext
                    MText acNewMText = new MText();
                    ObjectId objId;
                    Point3d acPosition = new Point3d();

                    //get the insertion point of the new mtext block based on the first selected block of text
                    //TO-DO in future figure out which text is the on top in the drawing and use that point
                    if (acObjList[0].GetType() == typeof(DBText))
                    {
                        DBText tempDBtext = acTrans.GetObject(acObjList[0].ObjectId, OpenMode.ForRead) as DBText;
                        acPosition = tempDBtext.Position;
                    }
                    else if (acObjList[0].GetType() == typeof(MText))
                    {
                        MText tempMtext = acTrans.GetObject(acObjList[0].ObjectId, OpenMode.ForRead) as MText;
                        acPosition = tempMtext.Location;
                    }
                    else
                    {
                        acPosition = new Point3d(0, 0, 0);
                    }                   
                    //iterate though the list of entities and use properties of the first Mtext entity found
                    //for the new mtext entity
                    
                    try
                    {
                        MText firstMtext = (MText)acObjList.Find(x => x.GetType() == typeof(MText));
                        //set relevant properties to the new mtext entity
                        acNewMText.SetDatabaseDefaults();
                        acNewMText.Location = acPosition;
                        acNewMText.TextHeight = firstMtext.TextHeight;
                        acNewMText.TextStyle = firstMtext.TextStyle;
                        acNewMText.Width = firstMtext.Width;
                        acNewMText.Layer = firstMtext.Layer;
                    }
                    catch (System.Exception)
                    {                       
                        //set relevant properties to the new mtext entity
                        MText firstMtext = new MText();
                        acNewMText.SetDatabaseDefaults();
                        acNewMText.Location = acPosition;                          
                    }

                    //iterate though each entity add the entities text to the acObjList_Mtext based on
                    //if the entity is DBText or Mtext
                    foreach (Entity acEnt in acObjList)
                    {  
                        //test to see if acEnt is Mtext
                        if(acEnt.GetType() == typeof(MText))
                        {
                            //add text contents to acObjList_Mtest
                            MText acMtextTemp = acTrans.GetObject(acEnt.ObjectId, OpenMode.ForRead) as MText;
                            acObjList_Mtext.Add(acMtextTemp.Text);
                        }
                            //if acEnt is not mtext
                        else if (acEnt.GetType() == typeof(DBText))
                        {                            
                            //add text contents to acObjList_Mtext
                            DBText acDBText = acTrans.GetObject(acEnt.ObjectId, OpenMode.ForWrite) as DBText;
                            acObjList_Mtext.Add(acDBText.TextString);
                        }
                    }
                    
                    //check to make sure that the List acObjList_Mtext is not empty
                    if (acObjList_Mtext.Count != 0)
                    {
                        //add all strings stored in acObjList_Mtext to the new acNewMtext entity
                        string tempStr = "";
                        foreach (string str in acObjList_Mtext)
                        {
                            tempStr += str;
                            tempStr += "\\P";
                        }
                        acNewMText.Contents = tempStr;
                        objId = acBlkTblRec.AppendEntity(acNewMText);
                        acTrans.AddNewlyCreatedDBObject(acNewMText, true);
                    }
                    
                    //remove initially selected objects from the database.
                    for (int i = 0; i < acObjList.Count; i++)
                    {
                        Entity acEnt = acTrans.GetObject(acObjList[i].ObjectId, OpenMode.ForWrite) as Entity;
                        acEnt.Erase();
                        acEnt.Dispose();
                    }
                    acTrans.Commit();
                }
            }
        }

        [CommandMethod("dirr")]
        public void dirr()
        {
            OpenDir();
        }

        //open current working directory in explorer
        [CommandMethod("opendir")]
        public void OpenDir()
        {
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;

            string dir = acDoc.Name;
            try
            {
                Process.Start(Path.GetDirectoryName(dir));
            }
            catch (System.Exception)
            {
                MessageBox.Show("It appears the directory requested is not available\nSave the drawing and try again", "Directory not available");
            }
         }

        [CommandMethod("breakatpoint", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public static void Breakatpoint()
        {
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            Editor acEd = acDoc.Editor;


            PromptSelectionResult psr = acEd.GetSelection();
            if (psr.Status != PromptStatus.OK)
            {
                return;
            }

            foreach (ObjectId id in psr.Value.GetObjectIds())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId,
                        OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite) as BlockTableRecord;

                    Line line = (Line)acTrans.GetObject(id, OpenMode.ForRead);
                    PromptPointResult pr = acEd.GetPoint("Select break point");
                    if (pr.Status != PromptStatus.Cancel)
                    {
                        ObjectId objId;
                        Point3d pointPr = pr.Value;
                        //Test to ensure that point selected actually lies on the line
                        if (IsPointOnCurve(line, pointPr))
                        {
                            //create new line from selected line start point to point chosen
                            
                            var lineseg = new LineSegment3d(line.StartPoint, pointPr);
                            var vec = lineseg.Direction.MultiplyBy(2).Negate();
                            Line line_seg1 = new Line(line.StartPoint, pointPr.Add(vec));
                            line_seg1.Layer = line.Layer;                          
                            objId = acBlkTblRec.AppendEntity(line_seg1);
                            acTrans.AddNewlyCreatedDBObject(line_seg1, true);
                            
                            //create new line from point chosen to end point on selected line
                            lineseg = new LineSegment3d(pointPr, line.EndPoint);
                            vec = lineseg.Direction.MultiplyBy(2).Negate();
                            Line line_seg2 = new Line(pointPr.Subtract(vec), line.EndPoint);
                            line_seg2.Layer = line.Layer;
                            objId = acBlkTblRec.AppendEntity(line_seg2);
                            acTrans.AddNewlyCreatedDBObject(line_seg2, true);

                            //remove origionally selected line
                            Entity acEnt = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;
                            acEnt.Erase();
                            acEnt.Dispose();
                        }
                        else
                        {
                            MessageBox.Show("Point chosen does not lie on the line selected");
                        }
                    }
                    else
                    {
                        return;
                    }
                    acTrans.Commit();
                }
            }
        }

        [CommandMethod("nestedfreeze")]
        public static void NestedFreeze()
        {

        }
        [CommandMethod("openxref_readonly", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public static void openxref_readonly()
        {
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            Editor acEd = acDoc.Editor;
           
            PromptSelectionResult psr = acEd.GetSelection();
            if (psr.Status != PromptStatus.OK)
            {
                return;
            }
            using (Transaction acTr = acDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in psr.Value.GetObjectIds())
                {
                    BlockReference acBr = (BlockReference)acTr.GetObject(id, OpenMode.ForRead);
                    BlockTableRecord acBtr = (BlockTableRecord)acTr.GetObject(acBr.BlockTableRecord, OpenMode.ForRead);
                    if (acBtr.IsFromExternalReference)
                    {
                        Document doc = acApp.DocumentManager.Open(acBtr.GetXrefDatabase(true).Filename, true);
                        if (doc != null)
                        {
                            acApp.DocumentManager.MdiActiveDocument = doc; 
                        }
                    }
                    else
                    {
                        acEd.WriteMessage("Selection must be external reference");
                    }
                }
            }
        }

        [CommandMethod("gout")]
        public static void Greyout()
        {
            Document acDoc = acApp.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            Editor acEd = acDoc.Editor;

            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
            {
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                List<ObjectId> oLyrs = new List<ObjectId>();
                foreach (ObjectId acObjId in acLyrTbl)
                {
                    LayerTableRecord acLyrTblRec;
                    acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForRead) as LayerTableRecord;
                    if (acLyrTblRec.Name.Contains("XREF"))
                    {
                        if (acLyrTblRec.IsWriteEnabled == false)
                        {
                            acLyrTblRec.UpgradeOpen();
                        }
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8);
                    }
                }
                acTrans.Commit();
                acEd.Regen();
            }
        }

        [CommandMethod("jaaopen")]
        public void jaaopen()
        {
            Editor ed = acApp.DocumentManager.MdiActiveDocument.Editor;

            // First let's use the editor method, GetFileNameForOpen()

            PromptOpenFileOptions opts = new PromptOpenFileOptions("Select File");
            opts.Filter =
              "Drawing (*.dwg)|*.dwg|Design Web Format (*.dwf)|*.dwf|" +
              "Drawing Template (*.dwt)|*.dwt|Standards (*.dws)|*.dws|" +
              "DXF (*.dxf)|*.dxf|(All files (*.*)|*.*";
            
            
            PromptFileNameResult pr = ed.GetFileNameForOpen(opts);
            DocumentCollection acDocMgr = acApp.DocumentManager;
            Document doc = null;
            if (pr.Status == PromptStatus.OK)
            {
                foreach (string file in acDocMgr)
                {
                    if (pr.ReadOnly == true)
                    {
                        // Open document readonly
                        doc = acDocMgr.Open(file, true);
                    }
                    else
                    {
                        // Open document for write 
                        doc = acDocMgr.Open(file, false);
                    }
                }
            }
            
            if (doc != null)
            {
                acApp.DocumentManager.MdiActiveDocument = doc;
            }            
        }

        [CommandMethod("jopen")]
        public void jopen()
        {
            System.Windows.Forms.OpenFileDialog dlg = new System.Windows.Forms.OpenFileDialog();
            dlg.FileName = " ";
            dlg.DefaultExt = ".dwg";
            dlg.Filter = "Drawing|*.dwg|Standards|*.dws|DXF|*.dxf|Template|*.dwt";
            dlg.Multiselect = true;
            dlg.ShowReadOnly = true;

            DialogResult result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == DialogResult.OK)
            {
                DocumentCollection acDocMgr = acApp.DocumentManager;
                Document doc = null;
                foreach (string file in dlg.FileNames)
                {
                    if (dlg.ReadOnlyChecked)
                    {
                        // Open document readonly
                        doc = acDocMgr.Open(file, true);
                    }
                    else
                    {
                        // Open document for write                       
                        doc = acDocMgr.Open(file, false);
                    }
                }

                if (doc != null)
                {
                    acApp.DocumentManager.MdiActiveDocument = doc;
                }
            }
        }

        [CommandMethod("testopen", CommandFlags.Modal)]
        public void testopen()
        {
            Form2 dia = new Form2();
            dia.Show();            
        }
        private static bool IsPointOnCurve(Line line, Point3d pointPr)
        {
            try
            {
                line.GetDistAtPoint(pointPr);
                return true;
            }
            catch { }
            
            return false;
        }
    }

    class WolviesContextMenu
    {
        private static ContextMenuExtension menuExtension_WT;
        private static ContextMenuExtension menuExtension_WTLT;
        private static ContextMenuExtension LinemenuExtension;
        private static ContextMenuExtension BlockmenuExtension;
        
        internal static void Attach()
        {
            menuExtension_WT = new ContextMenuExtension();
            menuExtension_WTLT = new ContextMenuExtension();
            LinemenuExtension = new ContextMenuExtension();
            BlockmenuExtension = new ContextMenuExtension();    

            //Default menu items
            menuExtension_WT.Title = "Wolvies Tools";
            acMenu item1 = new acMenu("Close All (Don't Save)");           
            item1.Click += new EventHandler(item1_Click);
            menuExtension_WT.MenuItems.Add(item1);
            acMenu item2 = new acMenu("Close (Don't Save)");
            item2.Click += new EventHandler(item2_Click);
            menuExtension_WT.MenuItems.Add(item2);
            acMenu item3 = new acMenu("Running Distance");
            item3.Click += new EventHandler(item3_Click);
            menuExtension_WT.MenuItems.Add(item3);
            acMenu item4 = new acMenu("Join Text");
            item4.Click += new EventHandler(item4_Click);
            menuExtension_WT.MenuItems.Add(item4);
            acMenu item5 = new acMenu("Open current directory");
            item5.Click += new EventHandler(item5_Click);
            menuExtension_WT.MenuItems.Add(item5);

            menuExtension_WTLT.Title = "Layer tools";
            acMenu item6 = new acMenu("Grey out XREF");
            item6.Click += new EventHandler(item6_Click);
            menuExtension_WTLT.MenuItems.Add(item6);

            //Menu item for lines
            acMenu Line_item1 = new acMenu("Break line at point");
            Line_item1.Click += new EventHandler(Lineitem1_Click);
            LinemenuExtension.MenuItems.Add(Line_item1);

            //Menu item for blocks
            acMenu block_item1 = new acMenu("Open XREF (READ ONLY)");
            block_item1.Click += new EventHandler(Blockitem1_Click);
            BlockmenuExtension.MenuItems.Add(block_item1);            

            //settting up RXClasses for entity tyles
            RXClass rxClass_Line = Entity.GetClass(typeof(Line));
            RXClass rxClass_xref = Entity.GetClass(typeof(BlockReference));

            //Adding in the menu extensions
            acApp.AddDefaultContextMenuExtension(menuExtension_WT);
            acApp.AddDefaultContextMenuExtension(menuExtension_WTLT);
            acApp.AddObjectContextMenuExtension(rxClass_Line, LinemenuExtension);
            acApp.AddObjectContextMenuExtension(rxClass_xref, BlockmenuExtension);
        }

        //Close all no save
        private static void item1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to close all open documents without saving?", "Close All (No Save)", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Document doc = acApp.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute("CloseAllNoSave ", false, false, true);
            }
            else
            {
                return;
            }
        }

        //Close no save
        private static void item2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to close this document without saving?", "Close (No Save)", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Document doc = acApp.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute("CloseNoSave ", false, false, true);
            }
            else
            {
                return;
            }
        }

        private static void item3_Click(object sender, EventArgs e)
        {
            Document doc = acApp.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("rundist ", false, false, true);
        }

        private static void item4_Click(object sender, EventArgs e)
        {
            Document doc = acApp.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("jointext ", false, false, true);
        }

        private static void item5_Click(object sender, EventArgs e)
        {
            Document doc = acApp.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("opendir ", false, false, true);
        }

        private static void item6_Click(object sender, EventArgs e)
        {
            Document doc = acApp.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("gout ", false, false, true);
        }

        private static void Blockitem1_Click(object sender, EventArgs e)
        {
            Document doc = acApp.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("openxref_readonly ", false, false, true);
        }
        
        private static void Lineitem1_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.Windows.MenuItem mItem = sender as Autodesk.AutoCAD.Windows.MenuItem;
            if (mItem != null)
            {
                if (mItem.Text == "Break line at point")
                {
                    Document doc = acApp.DocumentManager.MdiActiveDocument;
                    doc.SendStringToExecute("breakatpoint ", false, false, true);
                }
            }
        }

        internal static void Detach()
        {
            acApp.RemoveDefaultContextMenuExtension(menuExtension_WT);
        }
    }
}
