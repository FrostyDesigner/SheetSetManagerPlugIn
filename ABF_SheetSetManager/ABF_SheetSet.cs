using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using ACSMCOMPONENTS23Lib;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

// Instructions:
// 1) Add references: 
// AcCoreMgd
// AcDbMgd
// AcMgd
// ACSMCOMPONENTS23Lib
// System.Windows.Forms

// 2) Right click Project > Properties > Debug > Start External Program > "C:\Program Files\Autodesk\AutoCAD 2020\acad.exe"

// 3) Create MySSmEventHandler class

// 4) Add Using statements

// 5) Make sure references do not copy local: Select Reference > right click > properties > Copy Local = False

namespace ABF_SheetSetManager
{
    public class ABF_SheetSet
    {
        MySSmEventHandler eventHandler;
        Int32 eventSSMCookie;
        Int32 eventDbCookie;
        Int32 eventSSetCookie;
        IAcSmSheetSetMgr m_sheetSetManager;
        IAcSmDatabase m_sheetSetDatabase;
        IAcSmSheetSet m_sheetSet;

        // Open a Sheet Set 
        [CommandMethod("ABF_OpenSheetSet")]
        public void OpenSheetSet()
        {
            // User Input: editor equals command line
            // To talk to the user you use the command line, aka the editor
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            PromptStringOptions pso = new PromptStringOptions("\nHello Nadia! \nWhat Sheet Set would you like to open?");
            pso.DefaultValue = @"C:\Users\rhale\Documents\AutoCAD Sheet Sets\Squid666.dst";
            pso.UseDefaultValue = true;
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);

            // Get a reference to the Sheet Set Manager object 
            IAcSmSheetSetMgr sheetSetManager = default(IAcSmSheetSetMgr);
            sheetSetManager = new AcSmSheetSetMgr();

            // Open a Sheet Set file 
            AcSmDatabase sheetSetDatabase = default(AcSmDatabase);
            //sheetSetDatabase = sheetSetManager.OpenDatabase(@"C:\Users\Robert\Documents\AutoCAD Sheet Sets\Expedia.dst", false);
            sheetSetDatabase = sheetSetManager.OpenDatabase(pr.StringResult, false);

            // Return the namd and description of the sheet set
            MessageBox.Show("Sheet Set Name: " + sheetSetDatabase.GetSheetSet().GetName() + "\nSheet Set Description: " + sheetSetDatabase.GetSheetSet().GetDesc());

            // Close the sheet set 
            sheetSetManager.Close(sheetSetDatabase);
        }

        // Create a new sheet set 
        [CommandMethod("ABF_CreateSheetSet")]
        public void CreateSheetSet()
        {
            // User Input: editor equals command line
            // To talk to the user you use the command line, aka the editor
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            PromptResult pr = ed.GetFileNameForSave("Hello Nadia, Where would you would like to save the new Sheet Set?");

            PromptStringOptions pso = new PromptStringOptions("\nHello Nadia! \nWhat will be the name of the new Sheet Set?");
            pso.AllowSpaces = true;
            PromptResult prSheetSetName = ed.GetString(pso);

            PromptStringOptions psoDescription = new PromptStringOptions("\nHello Nadia! \nWhat will be the description of the new Sheet Set?");
            psoDescription.AllowSpaces = true;
            PromptResult prSheetSetDescription = ed.GetString(psoDescription);


            //pso.DefaultValue = @"C:\Users\Robert\Documents\AutoCAD Sheet Sets\Expedia.dst";
            //pso.UseDefaultValue = true;

            // works...

            // Get a reference to the Sheet Set Manager object 
            IAcSmSheetSetMgr sheetSetManager = new AcSmSheetSetMgr();

            // Create a new sheet set file 
            //AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase(@"C:\Users\Robert\Documents\AutoCAD Sheet Sets\ExpediaSheetSetDemo.dst", "", true);
            AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase(pr.StringResult, "", true);

            // Get the sheet set from the database 
            AcSmSheetSet sheetSet = sheetSetDatabase.GetSheetSet();

            // Attempt to lock the database
            if (LockDatabase(ref sheetSetDatabase, true) == true)
            {
                // Set the name and description of the sheet set 
                //sheetSet.SetName("ExpediaSheetSetTest");
                sheetSet.SetName(prSheetSetName.StringResult);

                //sheetSet.SetDesc("Aluminum Bronze Fabricator's Sheet Set Object Demo");
                sheetSet.SetDesc(prSheetSetDescription.StringResult);

                // Unlock the database 
                LockDatabase(ref sheetSetDatabase, false);

                // Return the name and description of the sheet set 
                MessageBox.Show("Sheet Set Name: " + sheetSetDatabase.GetSheetSet().GetName() + "\nSheet Set Description: " + sheetSetDatabase.GetSheetSet().GetDesc());
            }
            else
            {
                // Display error message 
                MessageBox.Show("Sheet set could not be opened for write.");
            }

            // Close the sheet set 
            sheetSetManager.Close(sheetSetDatabase);
        }

        // Step through all open sheet sets 
        [CommandMethod("ABF_StepThroughOpenSheetSets")]
        public void StepThroughOpenSheetSets()
        {
            // Get a reference to the Sheet Set Manager object 
            IAcSmSheetSetMgr sheetSetManager = new AcSmSheetSetMgr();

            // Get the loaded databases 
            IAcSmEnumDatabase enumDatabase = sheetSetManager.GetDatabaseEnumerator();

            // Get the first open database 
            IAcSmPersist item = enumDatabase.Next();

            string customMessage = "";

            // If a database is open continue 
            if ((item != null))
            {
                int count = 0;

                // Step through the database enumerator 
                while ((item != null))
                {
                    // Append the file name of the open sheet set to the output string 
                    customMessage = customMessage + "\n" + item.GetDatabase().GetFileName();

                    // Get the next open database and increment the counter 
                    item = enumDatabase.Next();
                    count = count + 1;
                }

                customMessage = "Sheet sets open: " + count.ToString() + customMessage;
            }
            else
            {
                customMessage = "No sheet sets are currently open.";
            }

            // Display the custom message 
            MessageBox.Show(customMessage);
        }

        // Create a new sheet set with custom subsets
        [CommandMethod("ABF_CreateSheetSetWithSubsets")]
        public void CreateSheetSet_WithSubset()
        {
            // User Input: editor equals command line
            // To talk to the user you use the command line, aka the editor
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            PromptResult pr = ed.GetFileNameForSave("Hello Nadia, Where would you would like to save the new Sheet Set?");

            PromptStringOptions pso = new PromptStringOptions("\nHello Nadia! \nWhat will be the name of the new Sheet Set?");
            pso.AllowSpaces = true;
            PromptResult prSheetSetName = ed.GetString(pso);

            PromptStringOptions psoDescription = new PromptStringOptions("\nHello Nadia! \nWhat will be the description of the new Sheet Set?");
            psoDescription.AllowSpaces = true;
            PromptResult prSheetSetDescription = ed.GetString(psoDescription);

            // Get a reference to the Sheet Set Manager object 
            IAcSmSheetSetMgr sheetSetManager = new AcSmSheetSetMgr();

            // Create a new sheet set file 
            //AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase("C:\\Datasets\\CP318-4\\CP318-4.dst", "", true);
            AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase(pr.StringResult, "", true);

            // Get the sheet set from the database 
            AcSmSheetSet sheetSet = sheetSetDatabase.GetSheetSet();

            // Attempt to lock the database 
            if (LockDatabase(ref sheetSetDatabase, true) == true)
            {
                // Set the name and description of the sheet set 
                sheetSet.SetName(prSheetSetName.StringResult);
                sheetSet.SetDesc(prSheetSetDescription.StringResult);

                // Create two new subsets 
                AcSmSubset subset = default(AcSmSubset);
                //subset = CreateSubset(sheetSetDatabase, "Plans", "Building Plans", "", "C:\\Datasets\\CP318-4\\CP318-4.dwt", "Sheet", false);
                subset = CreateSubset(sheetSetDatabase, "Submittals", "Project Submittals", "", @"C:\Users\Robert\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet", false);

                //subset = CreateSubset(sheetSetDatabase, "Elevations", "Building Elevations", "", "C:\\Datasets\\CP318-4\\CP318-4.dwt", "Sheet", true);
                subset = CreateSubset(sheetSetDatabase, "As Builts", "Project As Builts", "", @"C:\Users\Robert\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet", true);

                // Unlock the database 
                LockDatabase(ref sheetSetDatabase, false);
            }
            else
            {
                // Display error message 
                MessageBox.Show("Sheet set could not be opened for write.");
            }

            // Close the sheet set 
            sheetSetManager.Close(sheetSetDatabase);
        }

        // Create a new sheet set 
        [CommandMethod("ABF_SheetSet_AddCustomProperty")]
        public void CreateSheetSet_AddCustomProperty()
        {
            // Get a reference to the Sheet Set Manager object 
            IAcSmSheetSetMgr sheetSetManager = new AcSmSheetSetMgr();

            // Create a new sheet set file 
            //AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase("C:\\Datasets\\CP318-4\\CP318-4.dst", "", true);
            //AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase("C:\\Datasets\\CP318-4\\CP318-4.dst", "", true);
            //AcSmDatabase sheetSetDatabase = sheetSetManager.OpenDatabase(@"C:\Users\rhale\Documents\AutoCAD Sheet Sets\Expedia.dst", true);           
            //AcSmDatabase sheetSetDatabase = sheetSetManager.(@"C:\Users\rhale\Documents\AutoCAD Sheet Sets\Expedia.dst", true);

            // Open a Sheet Set file 
            AcSmDatabase sheetSetDatabase = default(AcSmDatabase);
            //sheetSetDatabase = sheetSetManager.OpenDatabase("C:\\Program Files\\AutoCAD 2010\\Sample\\Sheet Sets\\Architectural\\IRD Addition.dst", false); 
            sheetSetDatabase = sheetSetManager.OpenDatabase(@"C:\Users\rhale\Documents\AutoCAD Sheet Sets\Expedia.dst", false);

            // Get the sheet set from the database 
            AcSmSheetSet sheetSet = sheetSetDatabase.GetSheetSet();

            // Attempt to lock the database 
            if (LockDatabase(ref sheetSetDatabase, true) == true)
            {
                // Get the folder the sheet set is stored in 
                string sheetSetFolder = null;
                sheetSetFolder = sheetSetDatabase.GetFileName().Substring(0, sheetSetDatabase.GetFileName().LastIndexOf("\\"));

                // Set the default values of the sheet set 
                SetSheetSetDefaults(sheetSetDatabase, "My Expedia Sheet Set", "A&B Fabricators Sheet Set Object Demo", sheetSetFolder, @"C:\Users\rhale\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet");

                // Create a sheet set property 
                SetCustomProperty(sheetSet, "Project Approved By", "Erin Tasche", PropertyFlags.CUSTOM_SHEETSET_PROP);

                // Create sheet properties 
                SetCustomProperty(sheetSet, "Checked By", "NK", PropertyFlags.CUSTOM_SHEET_PROP);

                SetCustomProperty(sheetSet, "Complete Percentage", "90%", PropertyFlags.CUSTOM_SHEET_PROP);

                SetCustomProperty(sheetSet, "Version #", "D", PropertyFlags.CUSTOM_SHEET_PROP);

                SetCustomProperty(sheetSet, "Version Issue Date", "08/29/19", PropertyFlags.CUSTOM_SHEET_PROP);

                //AddSheet(sheetSetDatabase, "Title Page", "Project Title Page", "Title Page", "T1"); 

                // Create two new subsets 
                AcSmSubset subset = default(AcSmSubset);
                subset = CreateSubset(sheetSetDatabase, "Plans", "Building Plans", "", @"C:\Users\rhale\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet", false);

                //AddSheet(subset, "North Plan", "Northern section of building plan", "North Plan", "P1"); 

                subset = CreateSubset(sheetSetDatabase, "Elevations", "Building Elevations", "", @"C:\Users\rhale\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet", true);

                // Sync the properties of the sheet set with the sheets and subsets 
                SyncProperties(sheetSetDatabase);

                // Unlock the database 
                LockDatabase(ref sheetSetDatabase, false);
            }
            else
            {
                // Display error message 
                MessageBox.Show("Sheet set could not be opened for write.");
            }

            // Close the sheet set 
            sheetSetManager.Close(sheetSetDatabase);
        }

        // Create a new sheet set 
        [CommandMethod("ABF_SyncProperties")]
        public void SyncProperties()
        {
            // User Input: editor equals command line
            // To talk to the user you use the command line, aka the editor
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            PromptResult pr1 = ed.GetFileNameForOpen("Hello Nadia, Please select the Sheet Set you would like to edit.");

            PromptStringOptions pso = new PromptStringOptions("\nHello Nadia! \nWhat version # is the drawing at?");
            //pso.DefaultValue = @"C:\Users\Robert\Documents\AutoCAD Sheet Sets\Expedia.dst";
            //pso.UseDefaultValue = true;
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);
            string versionNumber = pr.StringResult;

            pso = new PromptStringOptions("\nHello Nadia! \nWhat is the version issue date?");
            pr = ed.GetString(pso);
            string versionIssueDate = pr.StringResult;

            // Get a reference to the Sheet Set Manager object 
            IAcSmSheetSetMgr sheetSetManager = new AcSmSheetSetMgr();

            // Create a new sheet set file 
            //AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase("C:\\Datasets\\CP318-4\\CP318-4.dst", "", true); 
            //AcSmDatabase sheetSetDatabase = sheetSetManager.CreateDatabase(@"C:\Users\rhale\Documents\AutoCAD Sheet Sets\Expedia.dst", "", true);

            // Open a Sheet Set file 
            AcSmDatabase sheetSetDatabase = default(AcSmDatabase);
            //sheetSetDatabase = sheetSetManager.OpenDatabase("C:\\Program Files\\AutoCAD 2010\\Sample\\Sheet Sets\\Architectural\\IRD Addition.dst", false); 
            //sheetSetDatabase = sheetSetManager.OpenDatabase(@"C:\Users\rhale\Documents\AutoCAD Sheet Sets\Expedia.dst", false);
            sheetSetDatabase = sheetSetManager.OpenDatabase(pr1.StringResult, false);

            // Get the sheet set from the database 
            AcSmSheetSet sheetSet = sheetSetDatabase.GetSheetSet();

            // Attempt to lock the database 
            if (LockDatabase(ref sheetSetDatabase, true) == true)
            {
                // Get the folder the sheet set is stored in 
                string sheetSetFolder = null;
                sheetSetFolder = sheetSetDatabase.GetFileName().Substring(0, sheetSetDatabase.GetFileName().LastIndexOf("\\"));

                // Set the default values of the sheet set 
                //SetSheetSetDefaults(sheetSetDatabase, "CP318-4", "AU2009 Sheet Set Object Demo", sheetSetFolder, "C:\\Datasets\\CP318-4\\CP318-4.dwt", "Sheet"); 
                SetSheetSetDefaults(sheetSetDatabase, "Expedia Sheet Set", "ABF Sheet Set Object Demo", sheetSetFolder, @"C:\Users\rhale\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet");

                // Create a sheet set property 
                SetCustomProperty(sheetSet, "Project Approved By", "Erin Tasche", PropertyFlags.CUSTOM_SHEETSET_PROP);

                // Create sheet properties 
                //SetCustomProperty(sheetSet, "Checked By", "LAA", PropertyFlags.CUSTOM_SHEET_PROP);

                //SetCustomProperty(sheetSet, "Complete Percentage", "0%", PropertyFlags.CUSTOM_SHEET_PROP);

                SetCustomProperty(sheetSet, "Version #", versionNumber, PropertyFlags.CUSTOM_SHEET_PROP);

                SetCustomProperty(sheetSet, "Version Issue Date", versionIssueDate, PropertyFlags.CUSTOM_SHEET_PROP);


                //AddSheet(sheetSetDatabase, "Title Page", "Project Title Page", "Title Page", "T1"); 

                // Create two new subsets 
                AcSmSubset subset = default(AcSmSubset);
                //subset = CreateSubset(sheetSetDatabase, "Plans", "Building Plans", "", "C:\\Datasets\\CP318-4\\CP318-4.dwt", "Sheet", false);
                subset = CreateSubset(sheetSetDatabase, "Submittals", "Project Submittals", "", @"C:\Users\Robert\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet", false);

                //subset = CreateSubset(sheetSetDatabase, "Elevations", "Building Elevations", "", "C:\\Datasets\\CP318-4\\CP318-4.dwt", "Sheet", true);
                subset = CreateSubset(sheetSetDatabase, "As Builts", "Project As Builts", "", @"C:\Users\Robert\Documents\AutoCAD Sheet Sets\ABFStylesSS.dwt", "Sheet", true);

                // Add a sheet property 
                //SetCustomProperty(sheetSet, "Drafted By", "KMA", PropertyFlags.CUSTOM_SHEET_PROP);

                // Add a subset property 
                SetCustomProperty(sheetSet, "Count", "0", PropertyFlags.CUSTOM_SHEETSET_PROP);

                // Sync the properties of the sheet set with the sheets and subsets 
                SyncProperties(sheetSetDatabase);

                // Unlock the database 
                LockDatabase(ref sheetSetDatabase, false);
            }
            else
            {
                // Display error message 
                MessageBox.Show("Sheet set could not be opened for write.");
            }

            // Close the sheet set 
            sheetSetManager.Close(sheetSetDatabase);
        }




        // Helper Methods ////////////////////////////////////////////////////////////////////////////////////////

        // Used to lock/unlock a sheet set database 
        private bool LockDatabase(ref AcSmDatabase database, bool lockFlag)
        {
            bool dbLock = false;
            // If lockFalg equals True then attempt to lock the database, otherwise 
            // attempt to unlock it. 
            if (lockFlag == true & database.GetLockStatus() == AcSmLockStatus.AcSmLockStatus_UnLocked)
            {
                database.LockDb(database);
                dbLock = true;
            }
            else if (lockFlag == false && database.GetLockStatus() == AcSmLockStatus.AcSmLockStatus_Locked_Local)
            {
                database.UnlockDb(database, true);
                dbLock = true;
            }
            else
            {
                dbLock = false;
            }

            return dbLock;
        }

        private AcSmSubset CreateSubset(AcSmDatabase sheetSetDatabase, string name, string description, string newSheetLocation,
                                string newSheetDWTLocation, string newSheetDWTLayout, bool promptForDWT)
        {
            // Create a subset with the provided name and description 
            AcSmSubset subset = (AcSmSubset)sheetSetDatabase.GetSheetSet().CreateSubset(name, description);

            // Get the folder the sheet set is stored in 
            string sheetSetFolder = sheetSetDatabase.GetFileName().Substring(0, sheetSetDatabase.GetFileName().LastIndexOf("\\"));

            // Create a reference to a File Reference object 
            IAcSmFileReference fileReference = subset.GetNewSheetLocation();

            // Check to see if a path was provided, if not default 
            // to the location of the sheet set 
            if (!string.IsNullOrEmpty(newSheetLocation))
            {
                fileReference.SetFileName(newSheetLocation);
            }
            else
            {
                fileReference.SetFileName(sheetSetFolder);
            }

            // Set the location for new sheets added to the subset 
            subset.SetNewSheetLocation(fileReference);

            // Create a reference to a Layout Reference object 
            AcSmAcDbLayoutReference layoutReference = default(AcSmAcDbLayoutReference);
            layoutReference = subset.GetDefDwtLayout();

            // Check to see that a default DWT location and name was provided 
            if (!string.IsNullOrEmpty(newSheetDWTLocation))
            {
                // Set the template location and name of the layout 
                //for the Layout Reference object 
                layoutReference.SetFileName(newSheetDWTLocation);
                layoutReference.SetName(newSheetDWTLayout);

                // Set the Layout Reference for the subset 
                subset.SetDefDwtLayout(layoutReference);
            }

            // Set the Prompt for Template option of the subset 
            subset.SetPromptForDwt(promptForDWT);
            return subset;
        }

        // Set/create a custom sheet or sheet set property 
        private void SetCustomProperty(IAcSmPersist owner, string propertyName, object propertyValue, PropertyFlags sheetSetFlag)
        {
            // Create a reference to the Custom Property Bag
            AcSmCustomPropertyBag customPropertyBag = default(AcSmCustomPropertyBag);

            if (owner.GetTypeName() == "AcSmSheet")
            {
                AcSmSheet sheet = (AcSmSheet)owner;
                customPropertyBag = sheet.GetCustomPropertyBag();
            }
            else
            {
                AcSmSheetSet sheetSet = (AcSmSheetSet)owner;
                customPropertyBag = sheetSet.GetCustomPropertyBag();
            }

            // Create a reference to a Custom Property Value 
            AcSmCustomPropertyValue customPropertyValue = new AcSmCustomPropertyValue();
            customPropertyValue.InitNew(owner);

            // Set the flag for the property 
            customPropertyValue.SetFlags(sheetSetFlag);

            // Set the value for the property 
            customPropertyValue.SetValue(propertyValue);

            // Create the property 
            customPropertyBag.SetProperty(propertyName, customPropertyValue);
        }

        // Synchronize the properties of a sheet with the sheet set 
        private void SyncProperties(AcSmDatabase sheetSetDatabase)
        {
            // Get the objects in the sheet set 
            IAcSmEnumPersist enumerator = sheetSetDatabase.GetEnumerator();

            // Get the first object in the Enumerator 
            IAcSmPersist item = enumerator.Next();

            // Step through all the objects in the sheet set 
            while ((item != null))
            {
                IAcSmSheet sheet = null;

                // Check to see if the object is a sheet 
                if (item.GetTypeName() == "AcSmSheet")
                {
                    sheet = (IAcSmSheet)item;

                    // Create a reference to the Property Enumerator for 
                    // the Custom Property Bag 
                    IAcSmEnumProperty enumeratorProperty = item.GetDatabase().GetSheetSet().GetCustomPropertyBag().GetPropertyEnumerator();

                    // Get the values from the Sheet Set to populate to the sheets 
                    string name = "";
                    AcSmCustomPropertyValue customPropertyValue = null;

                    // Get the first property 
                    enumeratorProperty.Next(out name, out customPropertyValue);

                    // Step through each of the properties 
                    while ((customPropertyValue != null))
                    {
                        // Check to see if the property is for a sheet 
                        if (customPropertyValue.GetFlags() == PropertyFlags.CUSTOM_SHEET_PROP)
                        {

                            //// Create a reference to the Custom Property Bag 
                            //AcSmCustomPropertyBag customSheetPropertyBag = sheet.GetCustomPropertyBag(); 

                            //// Create a reference to a Custom Property Value 
                            //AcSmCustomPropertyValue customSheetPropertyValue = new AcSmCustomPropertyValue();
                            //customSheetPropertyValue.InitNew(sheet); 

                            //// Set the flag for the property 
                            //customSheetPropertyValue.SetFlags(customPropertyValue.GetFlags()); 

                            //// Set the value for the property 
                            //customSheetPropertyValue.SetValue(customPropertyValue.GetValue()); 

                            //// Create the property 
                            //customSheetPropertyBag.SetProperty(name, customSheetPropertyValue); 

                            SetCustomProperty(sheet, name, customPropertyValue.GetValue(), customPropertyValue.GetFlags());
                        }

                        // Get the next property 
                        enumeratorProperty.Next(out name, out customPropertyValue);
                    }
                }

                // Get the next Sheet 
                item = enumerator.Next();
            }
        }


        // Set the default properties of a sheet set
        private void SetSheetSetDefaults(AcSmDatabase sheetSetDatabase, string name)
        {
            SetSheetSetDefaults(sheetSetDatabase, name, "", "", "", "", false);
        }

        private void SetSheetSetDefaults(AcSmDatabase sheetSetDatabase, string name, string description)
        {
            SetSheetSetDefaults(sheetSetDatabase, name, description, "", "", "", false);
        }

        private void SetSheetSetDefaults(AcSmDatabase sheetSetDatabase, string name, string description,
                                         string newSheetLocation)
        {
            SetSheetSetDefaults(sheetSetDatabase, name, description, newSheetLocation, "", "", false);
        }

        private void SetSheetSetDefaults(AcSmDatabase sheetSetDatabase, string name, string description,
                                         string newSheetLocation, string newSheetDWTLocation)
        {
            SetSheetSetDefaults(sheetSetDatabase, name, description, newSheetLocation, newSheetDWTLocation, "", false);
        }

        private void SetSheetSetDefaults(AcSmDatabase sheetSetDatabase, string name, string description,
                                         string newSheetLocation, string newSheetDWTLocation, string newSheetDWTLayout)
        {
            SetSheetSetDefaults(sheetSetDatabase, name, description, newSheetLocation, newSheetDWTLocation, newSheetDWTLayout, false);
        }

        private void SetSheetSetDefaults(AcSmDatabase sheetSetDatabase, string name, string description,
                                         string newSheetLocation, string newSheetDWTLocation, string newSheetDWTLayout,
                                         bool promptForDWT)
        {
            // Set the Name and Description for the sheet set 
            sheetSetDatabase.GetSheetSet().SetName(name);
            sheetSetDatabase.GetSheetSet().SetDesc(description);

            // Check to see if a Storage Location was provided 
            if (!string.IsNullOrEmpty(newSheetLocation))
            {
                // Get the folder the sheet set is stored in 
                string sheetSetFolder = sheetSetDatabase.GetFileName().Substring(0, sheetSetDatabase.GetFileName().LastIndexOf("\\"));

                // Create a reference to a File Reference object 
                IAcSmFileReference fileReference = default(IAcSmFileReference);
                fileReference = sheetSetDatabase.GetSheetSet().GetNewSheetLocation();

                // Set the default storage location based on the location of the sheet set 
                fileReference.SetFileName(sheetSetFolder);

                // Set the new Sheet location for the sheet set 
                sheetSetDatabase.GetSheetSet().SetNewSheetLocation(fileReference);
            }

            // Check to see if a Template was provided 
            if (!string.IsNullOrEmpty(newSheetDWTLocation))
            {
                // Set the Default Template for the sheet set 
                AcSmAcDbLayoutReference layoutReference = default(AcSmAcDbLayoutReference);
                layoutReference = sheetSetDatabase.GetSheetSet().GetDefDwtLayout();

                // Set the template location and name of the layout 
                // for the Layout Reference object 
                layoutReference.SetFileName(newSheetDWTLocation);
                layoutReference.SetName(newSheetDWTLayout);

                // Set the Layout Reference for the sheet set 
                sheetSetDatabase.GetSheetSet().SetDefDwtLayout(layoutReference);
            }

            // Set the Prompt for Template option of the subset 
            sheetSetDatabase.GetSheetSet().SetPromptForDwt(promptForDWT);
        }

    }
}
