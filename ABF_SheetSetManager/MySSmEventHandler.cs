using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using ACSMCOMPONENTS23Lib;

namespace ABF_SheetSetManager
{
    public class MySSmEventHandler : IAcSmEvents
    {
        private void IAcSmEvents_OnChanged(AcSmEvent ev, IAcSmPersist comp)
        {
            Autodesk.AutoCAD.ApplicationServices.Document activeDocument = default(Autodesk.AutoCAD.ApplicationServices.Document);
            activeDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            AcSmSheet oSheet = default(AcSmSheet);
            AcSmSubset oSubset = default(AcSmSubset);

            if (ev == AcSmEvent.ACSM_DATABASE_OPENED)
            {
                activeDocument.Editor.WriteMessage("\n" + comp.GetDatabase().GetFileName() + " was opened.");
            }
            if (ev == AcSmEvent.ACSM_DATABASE_CHANGED)
            {
                activeDocument.Editor.WriteMessage("\n" + comp.GetDatabase().GetFileName() + " database changed.");
            }
            if (ev == AcSmEvent.SHEET_DELETED)
            {
                oSheet = (AcSmSheet)comp;
                activeDocument.Editor.WriteMessage("\n" + oSheet.GetName() + " was deleted.");
            }
            if (ev == AcSmEvent.SHEET_SUBSET_CREATED)
            {
                oSubset = (AcSmSubset)comp;
                activeDocument.Editor.WriteMessage("\n" + oSubset.GetName() + " was created.");
            }
            if (ev == AcSmEvent.SHEET_SUBSET_DELETED)
            {
                oSubset = (AcSmSubset)comp;
                activeDocument.Editor.WriteMessage("\n" + oSubset.GetName() + " was deleted.");
            }
        }
        void IAcSmEvents.OnChanged(AcSmEvent ev, IAcSmPersist comp)
        {
            IAcSmEvents_OnChanged(ev, comp);
        }
    }
}
