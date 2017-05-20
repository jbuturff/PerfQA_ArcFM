using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Diagnostics;
using ESRI.ScriptEngine;
using System.ComponentModel.Composition;
using Miner.Interop;
using Miner.Interop.Process;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.ArcMap;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.EditorExt;
using ESRI.ArcGIS.Editor;
using ESRI.ArcGIS.NetworkAnalysis;
using ESRI.ArcGIS.NetworkAnalyst;
using ESRI.ArcGIS.NetworkAnalystTools;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Output;

using ADODB;

namespace PerfQA_ArcFM
{
    [Export(typeof(IScriptCommand))] // export for MEF contract
    public class Commands : IScriptCommand
    {
        private IMMStoredDisplayManager _SDMGR;
        private IMMPageTemplateManager _PTMGR;
        private CommandParser _parser;
        private System.Diagnostics.Stopwatch SW1;
        private System.Diagnostics.Stopwatch StepSW;

        /// <summary>
        /// Get or set the Logger used to report results or errors. 
        /// Custom ScriptCommands can use the Logger's WriteLine method to log output. 
        /// The ESRI Script Engine assigns the Logger to each ScriptCommand. 
        /// </summary>
        public Logger Logger { get; set; }

        public IWorkspace Workspace { get; set; }

        public IFeatureClass FeatureClass { get; set; }

        private ISelection pSelection;
        private IGeometricNetwork m_pGN;
        private ScriptEngine pSC;
        public IEditor pEditor { get; set; }
        private string sStempName;

        public Commands()
        {
            _parser = new CommandParser();
            _parser.CommentIdentifier = "//";

            // The ESRI ScriptEngine automatically sets the Logger when a 
            // script is executed; however, initialize it here in case a user 
            // wants to create and set the LogPath before executing a script. 
            this.Logger = new Logger();

            SW1 = new System.Diagnostics.Stopwatch();
            pSC = new ScriptEngine();
            StepSW = new System.Diagnostics.Stopwatch();
        }

        /// <summary>
        /// Execute the command and return one of the following values: 
        /// 3 - Command executed successfully. 
        /// 2 - Command executed and failed. 
        /// 1 - Command is empty or contains comments only (not executed). 
        /// 0 - Command is not recognized. 
        /// </summary>
        public int Execute(string command)
        {
            string sCommandLine;
            string[] lstArgs;
            string[] lstGroups;
            string[] lstPoints;
            string[] lstValues;
            string verb = _parser.GetVerb(command, true);

            if (command.Length > verb.Length)
            {
                sCommandLine = command.Substring(verb.Length + 1);
            }
            else
            {
                sCommandLine = "";
            }
            //this.Logger.WriteLine(sCommandLine);
            switch (verb.Trim())
            {
                case "FORCETIMESTAMP":
                    this.Logger.LogTimeStamp = true;
                    return 3;
                //
                // Used to set the current workspace of the same workspace as the ArcFM Login Workspace
                //
                case "SETARCFMWORKSPACE":
                    if (SetArcFMWorkspace())
                    {
                        return 3; // success
                    }
                    else
                    {
                        return 2; // failed
                    }
                //
                // Used to Open an ArcFM Stored Display, either a user or system stored display
                //
                //Parameters
                // 1 - Stored Display name exactly as it appears in the database
                // 2 - Stored Display type (SYSTEM or USER)
                case "OPENSTOREDDISPLAY":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example: OPENSTOREDDISPLAY Stored Display Name,SYSTEM");
                        this.Logger.WriteLine("Example: OPENSTOREDDISPLAY Stored Display Name,USER");
                        return 2;
                    }
                    else
                    {
                        if (OpenStoredDisplay(lstArgs[0], lstArgs[1]))
                        {
                            return 3;
                        }
                        else
                        {
                            return 2;
                        }
                    }
                //
                // Used to Open an ArcFM Stored Display, either a user or system stored display
                //
                //Parameters
                // 1 - Stored Display name exactly as it appears in the database
                // 2 - Stored Display type (SYSTEM or USER)
                case "OPENPAGETEMPLATE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example: OPENPAGETEMPLATE Stored Display Name,SYSTEM");
                        this.Logger.WriteLine("Example: OPENPAGETEMPLATE Stored Display Name,USER");
                        return 2;
                    }
                    else
                    {
                        if (OpenPageTemplate(lstArgs[0], lstArgs[1]))
                        {
                            return 3;
                        }
                        else
                        {
                            return 2;
                        }
                    }

                // 
                // Used to trace up stream from a selected feature.
                //  Must first set class with FEATURECLASS command
                //
                // 1 - ObjectID of Feature.
                // 2 - Select Results ("False" - Default should be to draw traced features, "True" - Select Traced Features )
                //
                case "TRACEUPSTREAM":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        ArcFMTrace(lstArgs[0], lstArgs[1], 2);
                    }
                    return 3;
                // 
                // Used to trace down stream from a selected feature.
                //  Must first set class with FEATURECLASS command
                //
                // 1 - ObjectID of Feature.
                // 2 - Select Results ("False" - Default should be to draw traced features, "True" - Select Traced Features )
                //
                case "TRACEDOWNSTREAM":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:TRACEDOWNSTREAM OBJECTID,TRUE");
                        this.Logger.WriteLine("Example:TRACEDOWNSTREAM OBJECTID,FALSE");
                    }
                    else
                    {
                        ArcFMTrace(lstArgs[0], lstArgs[1], 1);
                    }
                    return 3;
                // 
                // Used to perform an isolation trace in all directions from feature.
                //  Must first set class with FEATURECLASS command
                //
                // 1 - ObjectID of Feature.
                // 2 - Select Results ("False" - Default should be to draw traced features, "True" - Select Traced Features )
                //
                case "TRACEISOLATING":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:TRACEISOLATION OBJECTID,TRUE");
                        this.Logger.WriteLine("Example:TRACEISOLATION OBJECTID,FALSE");
                    }
                    else
                    {
                        ArcFMTrace(lstArgs[0], lstArgs[1], 3);
                    }
                    return 3;
                //
                // Disable ArcFM Auto-Updaters
                // Must have called SETARCFMWORKSPACE
                // - No Paramaters
                case "DISABLEAUTOUPDATERS":
                    DisableAutoUpdateres();
                    return 3;
                //
                // Enable ArcFM Auto-Updaters
                // Must have called SETARCFMWORKSPACE
                // - No parameters
                case "ENABLEAUTOUPDATERS":
                    EnableAutoUpdaters();
                    return 3;
                //
                // Executes ArcFM Auto-Assign Worklocation to CU Association
                // Must have called SETARCFMWORKSPACE
                //
                case "AUTO_ASSIGN_WL":
                    ARCFM_AUTO_ASSIGN();
                    return 3;
                //
                // Opens an existing ArcFM Session managed by ArcFM Session Manager
                // Must have called SETARCFMWORKSPACE
                // 1 - Session Name
                // 2 - ODBC Connection String
                //
                case "OPEN_ARCFM_SESSION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:OPEN_ARCFM_SESSION SESSION_name,ODBC_CONNECTION_STRING");
                    }
                    else
                    {
                        OpenSession(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Opens an existing ArcFM Session managed by ArcFM Session Manager
                // Must have called SETARCFMWORKSPACE
                // 1 - ODBC Connection String
                // 2 - Save (TRUE/FALSE)
                //
                case "CLOSE_ARCFM_SESSION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:CLOSE_ARCFM_SESSION ODBC_CONNECTION_STRING,TRUE");
                        this.Logger.WriteLine("Example:CLOSE_ARCFM_SESSION ODBC_CONNECTION_STRING,FALSE");
                    }
                    else
                    {
                        CLOSE_ARCFM_SESSION(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Opens an existing ArcFM Session managed by ArcFM Session Manager
                // Must have called SETARCFMWORKSPACE
                // 1 - ODBC Connection String
                //
                case "CREATE_ARCFM_SESSION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:CREATE_ARCFM_SESSION ODBC_CONNECTION_STRING");
                    }
                    else
                    {
                        CREATE_ARCFM_SESSION(lstArgs[0]);
                        //CreateSession(lstArgs[0]);
                    }
                    return 3;
                //
                // Executes the ArcFM zoom to function.
                // 1 - Scale value as whole number of the ratio 100 for 1:100
                // 2 - X coordinate
                // 3 - Y coordinate
                //
                case "ARCFMZOOMTO":
                    // Scale, X, Y
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 3)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:ARCFMZOOMTO X,Y,SCALE *SCALE is a whole number");
                    }
                    else
                    {
                        ARCFM_ZOOMTO(lstArgs[2], lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Execute ArcFM Gas Trace
                // 1 - Trace Type ("ISO")
                //
                case "ARCFMGasTrace":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:ARCFMGasTrace ISO");
                    }
                    else
                    {
                        ArcFMGasTrace(lstArgs[0]);
                    }
                    return 3;
                //
                // Perform a search of a feature class with a where clause
                //   1 - Class Name
                //   2 - Where Clause
                //
                case "SEARCHBYATTRIBUTE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SEARCHBYATTRIBUTE Class Name,Where_Clause");
                    }
                    else
                    {
                        QueryByAttribute(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Perform a search of a layer with a where clause
                //   1 - Class Name
                //   2 - Where Clause
                //
                case "SELECTLAYERBYATTRIBUTE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SELECTLAYERBYATTRIBUTE LAYER_NAME,Where_Clause");
                    }
                    else
                    {
                        SelectByAttribute(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                //
                case "SELECTFEATURES":
                    lstArgs = GetParameters(sCommandLine);
                    // Parameters
                    //   1 - Class Name
                    //   2 - Envelope                    
                    if (lstArgs.Count() != 5)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SELECTFEATURES CLASS_NAME,WHERE_CLAUSE");
                    }
                    else
                    {
                        SpatialSelect(lstArgs[0], lstArgs[1], lstArgs[2], lstArgs[3], lstArgs[4]);
                    }
                    return 3;
                //
                // Select features from the map in the provided extent from the specified layer
                // 1 - Layer name
                // 2 - MinX
                // 3 - MinY
                // 4 - MaxX
                // 5 - MaxY
                // 
                case "SPATIALATTRIBUTESELECT":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 6)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SPATIALATTRIBUTESELECT LAYER_NAME,Query_Where_Clause,MinX,MinY,MaxX,MaxY");
                    }
                    else
                    {
                        SpatialSelectByAttribute(lstArgs[0], lstArgs[1], lstArgs[2], lstArgs[3], lstArgs[4], lstArgs[5]);
                    }
                    return 3;
                //
                // Select features from the map in the provided extent.  Respects map scales and selectability
                // 1 - MinX
                // 2 - MinY
                // 3 - MaxX
                // 4 - MaxY
                //
                case "MAPSPATIALSELECT":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() == 4)
                    {
                        MapSpatialSelect(lstArgs[0], lstArgs[1], lstArgs[2], lstArgs[3]);
                    }
                    else
                    {
                        //if (lstArgs.Count() == 0)
                        //{
                            IMxDocument pMXDoc = GetDoc();
                            IEnvelope pEnv = pMXDoc.ActiveView.Extent;
                            string sXmin = Convert.ToString(pEnv.XMin);
                            string sYmin = Convert.ToString(pEnv.YMin);
                            string sXmax = Convert.ToString(pEnv.XMax);
                            string sYmax = Convert.ToString(pEnv.YMax);
                            MapSpatialSelect(sXmin, sYmin, sXmax, sYmax);
                        //}
                        //else
                        //{
                          //  this.Logger.WriteLine("Invalid number of parameters");
                          //  this.Logger.WriteLine("Example:MAPSPATIALSELECT MinX,MinY,MaxX,MaxY");
                        //}
                    }
                    return 3;
                // 
                // Zoom to selected features
                // No Parameters
                //
                case "ZOOMTOSELECTED":
                    ZOOMTOSELECTED();
                    return 3;
                // 
                // Clear Selection Set
                // No Parameters
                //
                case "CLEARSELECTION":
                    CLEARSELECTION();
                    return 3;
                // 
                // Update Attribute Value of feature/row
                // 1 - ClassName
                // 2 - Column Name
                // 3 - Value
                // 4 - Where Clause
                case "UPDATE_ATTRIBUTE":
                    lstGroups = GetParameterGroups(sCommandLine);
                    if (lstGroups.Count() != 3)
                    {
                        lstArgs = GetParameters(sCommandLine);
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:UPDATE_ATTRUBUTE Class_Name;where_clause;Column_Name,Value  *can have multiple pairs of column_name and value");
                    }
                    else
                    {
                        UpdateStringAttribute(lstGroups[0], lstGroups[1], lstGroups[2]);
                    }
                    return 3;
                //
                // Set the Map Scale to the ratio
                // 1 - Whole number of the 1 to something ratio 
                //
                case "SETSCALE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SETSCALE 1000");
                    }
                    else
                    {
                        SetScale(lstArgs[0]);
                    }
                    return 3;
                //
                // Execute the specified command button.  The button must be visible in the map
                // 1 - Command Button GUID
                //
                case "COMMANDBUTTON":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:COMMANDBUTTON GUID");
                    }
                    else
                    {
                        CommandClick(lstArgs[0]);
                    }
                    return 3;
                // 
                // Delete the selected features.  Must be editing and must have editable features or rows selected
                // no parameters
                //
                case "DELETESELECTED":
                    DELETESELECTED();
                    return 3;
                //
                // Create a version
                // Workspace must have been set
                // 1 - Version Name
                // 2 - Add GUID to make version uniquey (TRUE/FALSE)
                //
                case "CREATEVERSION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:CREATEVERSION Version_Name,TRUE");
                        this.Logger.WriteLine("Example:CREATEVERSION Version_Name,FALSE");
                    }
                    else
                    {
                        CreateVersion(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                case "CHANGEVERSION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:CHANGEVERSION VERSION_NAME");
                    }
                    else
                    {
                        ChangeVersion(lstArgs[0]);
                    }
                    return 3;
                // 
                // Zoom to the mas full extents
                // no parameters
                //
                case "ZOOMFULL":
                    ZoomToFullExtents();
                    return 3;
                // 
                // Set one or all layers visibility
                // 1 - Layer name or * for all layers
                // 2 - Status (On or Off)
                //
                case "SETLAYERVISIBILITY":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SETLAYERVISIBILITY Layer_Name,ON");
                        this.Logger.WriteLine("Example:SETLAYERVISIBILITY Layer_Name,OFF");
                    }
                    else
                    {
                        SetLayersVisibility(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                // 
                // Clear placed Network Flags
                // no parameters
                //
                case "CLEARFLAGS":
                    ClearFlags();
                    return 3;
                // 
                // Clear placed Network Barriers
                // no parameters
                //
                case "CLEARBARRIERS":
                    ClearBarriers();
                    return 3;
                // 
                // Clear Network Trace results
                // no parameters
                //
                case "CLEARRESULTS":
                    ClearResults();
                    return 3;
                //
                // Place a network flag on a feature
                // 1 - Feature Class Name
                // 2 - ObjectID
                //
                case "PLACEFLAG":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:PLACEFLAG FeatureClassName,ObjectID");
                    }
                    else
                    {
                        PlaceFlag(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Place a network barrier on a feature
                // 1 - Feature Class Name
                // 2 - ObjectID
                //
                case "PLACEBARRIER":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:PLACEBARRIER FeatureClassName,ObjectID");
                    }
                    else
                    {
                        PlaceBarrier(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Execute Esri Network Trace
                // 1 - Trace Name
                // 2 - Do Trace
                case "EXECUTENETWORKTRACE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:EXECUTENETWORKTRACE Trace_Name,???");
                    }
                    else
                    {
                        ExecuteNetworkTrace(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // set Layers Selectability property
                // 1 - Layer name or * for all layers
                // 2 - Status (On or Off)
                //
                case "SETLAYERSELECTABLE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:SETLAYERSELECTABLE Layer_Name,ON");
                        this.Logger.WriteLine("Example:SETLAYERSELECTABLE Layer_Name,OFF");
                    }
                    else
                    {
                        SetSelectable(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Add feature(s) to selection set for layer
                // 1 - Layer Name to search for
                // 2 - ObjectID of feature to add
                //
                case "ADDTOLAYERSELECTION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        AddToSelection(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // remove feature(s) to selection set for layer
                // 1 - Layer Name to search for
                // 2 - ObjectID of feature to add
                //
                case "REMOVEFROMLAYERSELECTION":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        RemoveFromSelecdtion(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                //
                // Start the Editor Extension
                // Workspace to edit
                //
                case "STARTEDITOR":
                    StartEditor(Workspace);
                    return 3;
                //
                // Start the Editor Extension
                // Layer name to get workspace from
                //
                case "EDITWORKSPACE":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:EDITWORKSPACE LAYER_NAME");
                    }
                    else
                    {
                        EditWorkspace(lstArgs[0]);
                    }
                    return 3;
                //
                // Stop the Editor Extension from editing
                // 1 - Save Edits (TRUE/FALSE)
                //
                case "STOPEDITOR":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        StopEditing(lstArgs[0]);
                    }
                    return 3;
                case "CREATEPOINT":
                    lstGroups = GetParameterGroups(sCommandLine);
                    if (lstGroups.Count() != 3)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        lstArgs = GetParameters(lstGroups[0]);
                        lstPoints = GetParameters(lstGroups[1]);
                        lstValues = GetParameters(lstGroups[2]);
                        if (lstPoints.Count() != 2)
                        {
                            this.Logger.WriteLine("Invalid number of coordindates");
                            this.Logger.WriteLine("Example:");
                        }
                        else
                        {
                            CREATEPOINTFEATURE(lstGroups[0], lstGroups[1], lstGroups[2]);
                        }
                    }
                    return 3;
                case "CREATELINE":
                    lstGroups = GetParameterGroups(sCommandLine);
                    if (lstGroups.Count() != 3)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        lstArgs = GetParameters(lstGroups[0]);
                        if (lstGroups.Length >= 2) lstPoints = GetParameters(lstGroups[1]);
                        if (lstGroups.Length == 3) lstValues = GetParameters(lstGroups[2]);
                        CREATELINEFEATURE(lstGroups[0], lstGroups[1], lstGroups[2]);
                    }
                    return 3;
                case "CREATEPOLYGON":
                    lstGroups = GetParameterGroups(sCommandLine);
                    if (lstGroups.Count() != 3)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        lstArgs = GetParameters(lstGroups[0]);
                        if (lstGroups.Length >= 2) lstPoints = GetParameters(lstGroups[1]);
                        if (lstGroups.Length == 3) lstValues = GetParameters(lstGroups[2]);
                        CREATEPOLYGONFEATURE(lstGroups[0], lstGroups[1], lstGroups[2]);
                    }
                    return 3;
                //
                // Execute Provided SQL Statement
                // 1 - SQL To Execute
                //
                case "EXECUTESQL":
                    EXECUTESQL(sCommandLine);
                    return 3;
                // 
                // Shutdown ArcMap
                // no Parameters
                //
                case "SHUTDOWN":
                    ShutdownArcMap();
                    return 3;
                // 
                // Pan up or down, left or right a specified distance in map units
                // 1 - Direction (UP,DOWN,LEFT,RIGHT)
                // 2 - Distance in Map Units
                //
                case "PAN":
                    // Pass in direction (UP,DOWN,LEFT,RIGHT) and distance in map units
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 2)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        Pan(lstArgs[0], lstArgs[1]);
                    }
                    return 3;
                // 
                // Create a PDF Export Map
                // 1 - Output File Name, including Path
                //
                case "PRINT":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        ExportToPDF(lstArgs[0]);
                    }
                    return 3;
                // 
                // Places the ArcMap process in a sleep state until the specified time
                // 1 - Wakeup Clock time
                //
                case "WAKEUP_AT":
                    lstArgs = GetParameters(sCommandLine);
                    if (lstArgs.Count() != 1)
                    {
                        this.Logger.WriteLine("Invalid number of parameters");
                        this.Logger.WriteLine("Example:");
                    }
                    else
                    {
                        WakeUpAt(lstArgs[0]);
                    }
                    return 3;
                case "STEP_START":
                    lstArgs = GetParameters(sCommandLine);
                    if (StepSW == null)
                    {
                        StepSW = new System.Diagnostics.Stopwatch();
                    }
                    else
                    {
                        StepSW.Reset();
                    }
                    sStempName = lstArgs[0];
                    StepSW.Start();
                    return 3;
                case "STEP_END":
                    StepSW.Stop();
                    //lstArgs = GetParameters(sCommandLine);
                    RecordActionTime("JOB STEP:" + sStempName, StepSW.ElapsedMilliseconds);
                    return 3;
                default:
                    if (string.IsNullOrEmpty(_parser.RemoveComments(command)))
                        return 1; // command is empty or contains comments only
                    else
                        return 0; // unrecognized command
            }
        }

        #region ArcFMTools

        private bool SetArcFMWorkspace()
        {
            IMMLoginUtils pMMLogin;

            try
            {
                SW1.Reset();
                SW1.Start();
                //                IMMStandardWorkspaces pMMWKS
                //                pMMWKS = new MM
                //                pMMLogin = new MMDefaultLoginObject();
                //                ESRI.ArcGIS.Framework.IAppROT pAppRot = new ESRI.ArcGIS.Framework.AppROT();
                //                ESRI.ArcGIS.Framework.IApplication pApp;
                //                if (pAppRot.Count > 0)
                //                    {
                //                        for (int i = 0; i < pAppRot.Count; i++)
                //                        {
                //                            if (pAppRot.get_Item(i) is ESRI.ArcGIS.ArcMapUI.IMxApplication)
                //                            {
                //                                pApp = (ESRI.ArcGIS.Framework.IApplication)pAppRot;
                //                                pMMProps = (IMMPropertiesExt3)pApp.FindExtensionByName("MMPropertiesExt");
                //                                pMMProps.DefaultWorkspace = Workspace;
                //                            }
                //                        }
                //                    }
                //                this.Logger.WriteLine("SetArcFMWorkspace");
                _SDMGR = new MMStoredDisplayManager();
                pMMLogin = new MMLoginUtils();
                this.Workspace = pMMLogin.LoginWorkspace;
                ScriptEngine.BroadcastProperty("Workspace", this.Workspace, null);

                SW1.Stop();
                _SDMGR.Workspace = this.Workspace;
                RecordActionTime("SETARCFMWORKSPACE:", SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in SetArcFMWorkspace:" + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private bool OpenStoredDisplay(string SDName, string SDType)
        {
            try
            {
                this.Logger.WriteLine("OpenStoredDisplay:" + SDName);
                IMMStoredDisplayName pSDName;
                IMMEnumStoredDisplayName pESD;

                if (SDName == "")
                {
                    this.Logger.WriteLine("StoredDisplay Name not provided");
                    return false;
                }
                else
                {
                    if (_SDMGR == null)
                    {
                        _SDMGR = new MMStoredDisplayManager();
                    }
                    _SDMGR.Workspace = this.Workspace;
                    pSDName = new MMStoredDisplayName();
                    pSDName.Name = SDName;

                    if (SDType.ToUpper() == "SYSTEM")
                    {
                        pESD = _SDMGR.GetStoredDisplayNames(mmStoredDisplayType.mmSDTSystem);
                    }
                    else
                    {
                        pESD = _SDMGR.GetStoredDisplayNames(mmStoredDisplayType.mmSDTUser);
                    }
                    // Get list of Stored Displays.  We have to do this because we have not logged in to ArcFM.
                    pSDName = pESD.Next();
                    while (pSDName != null)
                    {
                        if (pSDName.Name.Trim().ToUpper() == SDName.Trim().ToUpper())
                        {
                            SW1.Reset();
                            SW1.Start();
                            try
                            {
                                _SDMGR.OpenStoredDisplay(pSDName);
                                SW1.Stop();
                                RecordActionTime("Open Stored Display :",SW1.ElapsedMilliseconds);
                            }
                            catch (Exception EX)
                            {
                                this.Logger.WriteLine("Error Opening Stored Display :" + EX.Message + ":" + EX.StackTrace);
                                SW1.Stop();
                            }
                            break;
                        }
                        pSDName = pESD.Next();
                    }
                    if (pSDName == null)
                    {
                        this.Logger.WriteLine("StoredDisplay not found:" + SDName);
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in OpenStoredDisplay:" + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private bool OpenPageTemplate(string SDName, string SDType)
        {
            try
            {
                this.Logger.WriteLine("OpenPageTemplate:" + SDName);
               
                IMMPageTemplateName pPTName;
                IMMEnumPageTemplateName pEPT;
                
                if (SDName == "")
                {
                    this.Logger.WriteLine("PageTemplate Name not provided");
                    return false;
                }
                else
                {
                    if (_PTMGR == null)
                    {
                        _PTMGR = (IMMPageTemplateManager) new MMPageTemplateManager();
                    }
                    _PTMGR.Workspace = this.Workspace;
                    pPTName = new MMPageTemplateNameClass();
                    pPTName.Name = SDName;

                    if (SDType.ToUpper() == "SYSTEM")
                    {
                        pEPT = _PTMGR.GetPageTemplateNames(mmPageTemplateType.mmPTTSystem);
                    }
                    else
                    {
                        pEPT = _PTMGR.GetPageTemplateNames(mmPageTemplateType.mmPTTUser);
                    }
                    // Get list of Stored Displays.  We have to do this because we have not logged in to ArcFM.
                    pPTName = pEPT.Next();
                    while (pPTName != null)
                    {
                        if (pPTName.Name.Trim().ToUpper() == SDName.Trim().ToUpper())
                        {
                            SW1.Reset();
                            SW1.Start();
                            try
                            {
                                _PTMGR.OpenPageTemplate(pPTName);
                                SW1.Stop();
                                RecordActionTime("Open Page Template :", SW1.ElapsedMilliseconds);
                            }
                            catch (Exception EX)
                            {
                                this.Logger.WriteLine("Error Opening Page Tempate :" + EX.Message + ":" + EX.StackTrace);
                                SW1.Stop();
                            }
                            break;
                        }
                        pPTName = pEPT.Next();
                    }
                    if (pPTName == null)
                    {
                        this.Logger.WriteLine("PageTemplate not found:" + SDName);
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in OpenPageTempate:" + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private bool DisableAutoUpdateres()
        {
            Type type = Type.GetTypeFromProgID("mmGeodatabase.MMAutoUpdater");
            try
            {
                SW1.Reset();
                SW1.Start();
                object obj = Activator.CreateInstance(type);

                IMMAutoUpdater autoupdater = obj as IMMAutoUpdater;

                //autoupdater.AutoUpdaterMode = mmAutoUpdaterMode.mmAUMNotSet;
                autoupdater.AutoUpdaterMode = mmAutoUpdaterMode.mmAUMNoEvents;
                SW1.Stop();
                RecordActionTime("Disable Autoupdaters:",SW1.ElapsedMilliseconds);
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in Disabling Autoupdaters:" + EX.Message);
            }

            return true;
        }

        private bool EnableAutoUpdaters()
        {
            Type type = Type.GetTypeFromProgID("mmGeodatabase.MMAutoUpdater");

            try
            {
                SW1.Reset();
                SW1.Start();
                object obj = Activator.CreateInstance(type);

                IMMAutoUpdater autoupdater = obj as IMMAutoUpdater;

                autoupdater.AutoUpdaterMode = mmAutoUpdaterMode.mmAUMArcMap;
                SW1.Stop();
                RecordActionTime("Enable Autoupdaters:",SW1.ElapsedMilliseconds);
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in Enabling Autoupdaters:" + EX.Message);
            }
            return true;
        }

        private bool TraceIsolating(string sOID, string sSelect)
        {
            IFeature pFeat;
            INetworkClass NetCLS;
            IGeometricNetwork pGN;
            ISimpleJunctionFeature pJuncFeat;
            Miner.Interop.IMMElectricTracingEx pElecTrace;
            Miner.Interop.IMMElectricTraceSettings pElecTraceSettings;
            IMMTracedElements pJunctions;
            IMMTracedElements pEdges;
            IMMCurrentStatus pCurrentStatus;
            int[] iJunctionBarriers;
            int[] iEdgeBarriers;

            iJunctionBarriers = new int[0];
            iEdgeBarriers = new int[0];
            pCurrentStatus = null;
            try
            {
                pElecTrace = new MMFeederTracerClass();

                pElecTraceSettings = new Miner.Interop.MMElectricTraceSettingsClass();
                pElecTraceSettings.RespectConductorPhasing = false;
                pElecTraceSettings.RespectEnabledField = false;

                SW1.Reset();
                SW1.Start();
                try
                {
                    if (FeatureClass is INetworkClass)
                    {
                        NetCLS = (INetworkClass)FeatureClass;
                        pGN = NetCLS.GeometricNetwork;
                        pFeat = FeatureClass.GetFeature(Convert.ToInt32(sOID));
                        if (pFeat is IJunctionFeature)
                        {
                            pJuncFeat = (ISimpleJunctionFeature)pFeat;

                            pElecTrace.TraceDownstream(pGN, pElecTraceSettings, pCurrentStatus,
                                pJuncFeat.EID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction,
                                Miner.Interop.SetOfPhases.abc, Miner.Interop.mmDirectionInfo.establishBySourceSearch, 0,
                                ESRI.ArcGIS.Geodatabase.esriElementType.esriETEdge, iJunctionBarriers,
                                iEdgeBarriers, false, out pJunctions, out pEdges);
                        }
                    }
                    SW1.Stop();
                    RecordActionTime("TraceIsolating Execution Time:",SW1.ElapsedMilliseconds);
                    return true;
                }
                catch (Exception EX)
                {
                    this.Logger.WriteLine("TraceDownStream Failed: " + EX.Message);
                    return false;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("TraceDownStream Failed: " + EX.Message);
                return false;
            }
        }

        public bool ARCFM_AUTO_ASSIGN()
        {
            SW1.Reset();
            SW1.Start();
            SW1.Stop();
            RecordActionTime("ArcFM_AUTO_ASSIGN:", SW1.ElapsedMilliseconds);
            return true;
        }

        public bool UPDATE_String(string sValue, string sWhere)
        {

            SW1.Reset();
            SW1.Start();
            SW1.Stop();
            RecordActionTime("UPDATE_STRING-Not Functioning:", SW1.ElapsedMilliseconds);
            return true;
        }

        public bool CREATE_ARCFM_SESSION(string sConnStr)
        {
            IMxDocument pMXDoc;
            IApplication pApp;
            IMMLoginUtils pMMLogin;
            IMMPxLogin pPxLogin;
            IMMSessionManager2 pmmSessionMangerExt;
            IMMPxIntegrationCache pmmSessionMangerIntegrationExt;
            IMMSessionVersion pMMSessVer;
            IMMSession pMMSession;
            IMMPxApplication pPXApp;
            IWorkspace pWKS;
            IExtension pExt;
            IVersion pVersion;
            IVersion pNewVersion;
            ADODB.Connection pPXConnection;

            try
            {
                SW1.Reset();
                SW1.Start();
                pMMLogin = new MMLoginUtils();
                Logger.WriteLine("Get LoginWorkspace");
                pWKS = pMMLogin.LoginWorkspace;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                Logger.WriteLine("Find PX Framework");
                pExt = pApp.FindExtensionByName("Session Manager Integration Extension");
                pmmSessionMangerIntegrationExt = (IMMPxIntegrationCache)pExt;
                pPXApp = pmmSessionMangerIntegrationExt.Application;
                pPxLogin = pPXApp.Login;
                
                
                Logger.WriteLine("open Connection");
                if (pPxLogin == null)
                {
                    if (sConnStr != "")
                    {
                        pPxLogin = new PxLoginClass();
                        pPXConnection = new Connection();
                        pPXConnection.ConnectionString = sConnStr;
                        pPXConnection.Open();
                        pPxLogin.Connection = pPXConnection;
                        pPXApp.Startup(pPxLogin);
                    }
                    else
                    {
                        Logger.WriteLine("No PX Connection String provided");
                    }
                }
                else
                {
                    pPXConnection = pPxLogin.Connection;
                }

                Logger.WriteLine("Start PX");
                //pPXApp.Startup(pPxLogin);
                pmmSessionMangerExt = (IMMSessionManager2)pPXApp.FindPxExtensionByName("MMSessionManager");
                pMMSession = pmmSessionMangerExt.CreateSession();
                pMMSessVer = (IMMSessionVersion)pMMSession;

                pVersion = (IVersion)pWKS;
                Logger.WriteLine("Create PX version");
                pNewVersion = pVersion.CreateVersion(pMMSessVer.get_Name());

                //pCV = new ChangeDatabaseVersion();
                //pVSet = pCV.Execute(pVersion, pNewVersion, (IBasicMap)pMXDoc.FocusMap);
                SwizzleDatasets(pApp, (IFeatureWorkspace)pVersion, (IFeatureWorkspace)pNewVersion);
                this.Workspace = (IWorkspace)pNewVersion;

                //pMXDoc.ActiveView.Refresh();

                SW1.Stop();
                RecordActionTime("CREATE_ARCFM_SESSION:", SW1.ElapsedMilliseconds);
                //pMXDoc.ActiveView.Refresh();
                ScriptEngine.BroadcastProperty("Workspace", Workspace, this);
                StartEditor(Workspace);
                pMXDoc.ActiveView.Refresh();
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("CREATE_ARCFM_SESSION Failed: " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }
        public bool OPEN_ARCFM_SESSION(string sSessionName, string sConnStr)
        {
            IMxDocument pMXDoc;
            IApplication pApp;
            IMMLoginUtils pMMLogin;
            IMMPxLogin pPxLogin;
            IMMSessionManager2 pmmSessionMangerExt;
            IMMPxIntegrationCache pmmSessionMangerIntegrationExt;
            IMMSessionVersion pMMSessVer;
            IMMSession pMMSession;
            IMMPxApplication pPXApp;
            IWorkspace pWKS;
            IVersionedWorkspace pVWKS;
            IExtension pExt;
            IVersion pVersion;
            IVersion pNewVersion;
            IChangeDatabaseVersion pCV;
            ADODB.Connection pPXConnection;
            int iSessionID;

            SW1.Reset();
            SW1.Start();

            try
            {
                pMMLogin = new MMLoginUtils();
                pWKS = pMMLogin.LoginWorkspace;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                pExt = pApp.FindExtensionByName("Session Manager Integration Extension");
                pmmSessionMangerIntegrationExt = (IMMPxIntegrationCache)pExt;
                pPXApp = pmmSessionMangerIntegrationExt.Application;
                pPxLogin = pPXApp.Login;
                pmmSessionMangerExt = (IMMSessionManager2)pPXApp.FindPxExtensionByName("MMSessionManager");

                if (pPxLogin.Connection == null)
                {
                    pPXConnection = new Connection();
                    pPxLogin.ConnectionString = sConnStr;
                    pPXConnection.Open();
                    pPXApp.Startup(pPxLogin);
                }
                else
                {
                    pPXConnection = pPxLogin.Connection;
                }

                iSessionID = Convert.ToInt32(sSessionName);
                pMMSession = pmmSessionMangerExt.GetSession(iSessionID, false);
                pMMSessVer = (IMMSessionVersion)pMMSession;

                pVersion = (IVersion)pWKS;
                pVWKS = (IVersionedWorkspace)Workspace;
                Logger.WriteLine("Verion:" + pMMSessVer.get_Name());
                pNewVersion = pVWKS.FindVersion(pMMSessVer.get_Name());

                pCV = new ChangeDatabaseVersion();
                //pVSet = pCV.Execute(pVersion, pNewVersion, (IBasicMap)pMXDoc.FocusMap);
                SwizzleDatasets(pApp, (IFeatureWorkspace)pVersion, (IFeatureWorkspace)pNewVersion);
                Workspace = (IWorkspace)pNewVersion;
                StartEditor(Workspace);
                ScriptEngine.BroadcastProperty("Workspace", Workspace, this);

                SW1.Stop();
                RecordActionTime("OPEN_ARCFM_SESSION:", SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("TraceDownStream Failed: " + EX.Message);
                return false;
            }
        }

        public bool CLOSE_ARCFM_SESSION(string sODBC, string sSave)
        {
            IVersion pVersion;
            IVersion pNewVersion;
            IWorkspace pWKS;
            ESRI.ArcGIS.Framework.IAppROT pAppRot = new ESRI.ArcGIS.Framework.AppROT();
            ESRI.ArcGIS.Framework.IApplication pApp;

            //Set pExt = pApp.FindExtensionByName("MMSessionManager")

            //pMMSM = new mmsession
            SW1.Reset();
            SW1.Start();
            pWKS = pEditor.EditWorkspace;
            pVersion = (IVersion)pWKS;

            IVersionInfo pVersionInfo = pVersion.VersionInfo;
            IVersionedWorkspace pVWKS = (IVersionedWorkspace)pWKS;
            pNewVersion = pVWKS.DefaultVersion;
            if (sSave.ToUpper() == "TRUE")
            {
                pEditor.StopEditing(true);
            }
            else
            {
                pEditor.StopEditing(false);
            }
            //Return to Default version
            Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
            System.Object obj = Activator.CreateInstance(t);
            pApp = obj as IApplication;
            
            SwizzleDatasets(pApp, (IFeatureWorkspace)pVersion, (IFeatureWorkspace)pNewVersion);
            SW1.Stop();
            RecordActionTime("CLOSE_ARCFM_SESSION:", SW1.ElapsedMilliseconds);
            return true;
        }

        public bool ARCFM_ZOOMTO(string sScale, string sX, string sY)
        {
            IPoint pPoint;
            double dScale;
            IMxDocument pMXDoc;
            Miner.Interop.IMMMapUtilities pMMMapUtils;
            ESRI.ArcGIS.Framework.IAppROT pAppRot = new ESRI.ArcGIS.Framework.AppROT();
            ESRI.ArcGIS.Framework.IApplication pApp;

            //Set pExt = pApp.FindExtensionByName("MMSessionManager")
            try
            {
                pMMMapUtils = new mmMapUtilsClass();
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                // Convert X,Y to point
                pPoint = new Point();
                pPoint.SpatialReference = pMXDoc.FocusMap.SpatialReference;

                pPoint.X = System.Convert.ToDouble(sX);
                pPoint.Y = System.Convert.ToDouble(sY);
                dScale = 1.0 / System.Convert.ToDouble(sScale);
                // Convert Scale to double
                Logger.WriteLine("Scale:" + dScale.ToString() + " X:" + pPoint.X.ToString() + " Y:" + pPoint.Y.ToString());
                SW1.Reset();
                SW1.Start();
                pMMMapUtils.ZoomTo(pPoint, pMXDoc.ActiveView, dScale);
                pMXDoc.ActiveView.Refresh();
                SW1.Stop();
                RecordActionTime("ARCFM_ZOOMTO:" ,SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("ArcFM_ZoomTo" + EX.Message + " " + EX.StackTrace);

                return false;
            }
        }

        private bool CreateSession(string sConnStr)
        {
            IMMLoginUtils pMMLogin;
            IMMPxLogin pPxLogin;
            IMMSessionManager2 pmmSessionMangerExt;
            IMMPxIntegrationCache pmmSessionMangerIntegrationExt;
            IMMSessionVersion pMMSessVer;
            IMMSession pMMSession;
            IMMPxApplication pPXApp;
            IWorkspace pWKS;
            IApplication pApp;
            IMxDocument pMXDoc;
            IExtension pExt;
            IVersion pVersion;
            IVersion pNewVersion;
            IPropertySet pOldPS;
            object[] pOldPropNames;
            object[] pOldPropValues;
            ADODB.Connection pPXConnection;

            try
            {
                SW1.Stop();
                SW1.Reset();
                pMMLogin = new MMLoginUtils();
                pWKS = pMMLogin.LoginWorkspace;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                pExt = pApp.FindExtensionByName("Session Manager Integration Extension");
                pmmSessionMangerIntegrationExt = (IMMPxIntegrationCache)pExt;
                pPXApp = pmmSessionMangerIntegrationExt.Application;
                pPxLogin = pPXApp.Login;

                if (pPxLogin == null)
                {
                    pPxLogin = (IMMPxLogin)new PxLogin();
                    pPXConnection = new Connection();
                    pPXConnection.ConnectionString = sConnStr;
                    pPXConnection.Open();
                    pPxLogin.Connection = pPXConnection;
                    pPXApp.Startup(pPxLogin);
                }
                else
                {
                    pPXConnection = pPxLogin.Connection;
                }
                pmmSessionMangerExt = (IMMSessionManager2)pPXApp.FindPxExtensionByName("MMSessionManager");
                pMMSession = pmmSessionMangerExt.CreateSession();
                pMMSessVer = (IMMSessionVersion)pMMSession;

                this.Logger.WriteLine("Session:" + pMMSessVer.get_Name());
                pVersion = (IVersion)pWKS;
                pOldPS = pWKS.ConnectionProperties;
                pOldPropNames = new object[1];
                pOldPropValues = new object[2];
                pOldPS.GetAllProperties(out pOldPropNames[0], out pOldPropValues[0]);

                pNewVersion = pVersion.CreateVersion(pMMSessVer.get_Name());

                SwizzleDatasets(pApp, (IFeatureWorkspace)pVersion, (IFeatureWorkspace)pNewVersion);
                this.Workspace = (IWorkspace)pNewVersion;
                ScriptEngine.BroadcastProperty("Workspace", this.Workspace, this);
                SW1.Stop();
                RecordActionTime("CreateSession:", SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception Ex)
            {
                this.Logger.WriteLine("CreateArcFMSession Error" + Ex.Message + " " + Ex.StackTrace);
                return false;
            }
        }

        private bool CloseSession()
        {
            IChangeDatabaseVersion pCV;
            ISet pVSet;
            IVersion pVersion;
            IVersion pNewVersion;
            IWorkspace pWKS;
            IVersionInfo pVInfo;
            IApplication pApp;
            IMxDocument pMXDoc;
            IMMLoginUtils pMMLogin;

            try
            {
                SW1.Reset();
                SW1.Start();
                pMMLogin = new MMLoginUtils();
                pWKS = pMMLogin.LoginWorkspace;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                pVersion = (IVersion)pWKS;
                pVInfo = pVersion.VersionInfo;
                pNewVersion = (IVersion)pVInfo.Parent;
                pCV = new ChangeDatabaseVersion();
                pVSet = pCV.Execute(pVersion, pNewVersion, (IBasicMap)pMXDoc.FocusMap);
                SW1.Stop();
                RecordActionTime("CloseSession:", SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("CloseSession Error" + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool OpenSession(string sSessionName, string sConnStr)
        {
            IMMLoginUtils pMMLogin;
            IMMPxLogin pPxLogin;
            IMMSessionManager2 pmmSessionMangerExt;
            IMMPxIntegrationCache pmmSessionMangerIntegrationExt;
            IMMSessionVersion pMMSessVer;
            IMMSession pMMSession;
            IMMPxApplication pPXApp;
            IWorkspace pWKS;
            IApplication pApp;
            IMxDocument pMXDoc;
            IExtension pExt;
            IVersion pVersion;
            IVersion pNewVersion;
            ADODB.Connection pPXConnection;
            int iSessionID;
            try
            {
                SW1.Reset();
                SW1.Start();
                pMMLogin = new MMLoginUtils();
                pWKS = pMMLogin.LoginWorkspace;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                pExt = pApp.FindExtensionByName("Session Manager Integration Extension");
                pmmSessionMangerIntegrationExt = (IMMPxIntegrationCache)pExt;
                pPXApp = pmmSessionMangerIntegrationExt.Application;
                pPxLogin = pPXApp.Login;
                pmmSessionMangerExt = (IMMSessionManager2)pPXApp.FindPxExtensionByName("MMSessionManager");

                if (pPxLogin.Connection == null)
                {
                    pPXConnection = new Connection();
                    pPxLogin.ConnectionString = sConnStr;
                    pPXConnection.Open();
                    pPXApp.Startup(pPxLogin);
                }
                else
                {
                    pPXConnection = pPxLogin.Connection;
                }

                
                iSessionID = Convert.ToInt32(sSessionName);
                pMMSession = pmmSessionMangerExt.GetSession(iSessionID, false);
                pMMSessVer = (IMMSessionVersion)pMMSession;

                pVersion = (IVersion)pWKS;
                pNewVersion = pVersion.CreateVersion(pMMSessVer.get_Name());

                // Copy Properties
                //pOldPS = pWKS.ConnectionProperties;
                //pNewPS = new PropertySet();
                //pOldPropNames = new object[1];
                //pOldPropValues = new object[1];
                //pOldPS.GetAllProperties(out pOldPropNames[0], out pOldPropValues[0]);
                //for (int i = 0; i < pOldPS.Count; i++)
                //{
                //    if (pOldPropNames[i] != "VERSION")
                //    {
                //        pNewPS.SetProperty(pOldPropNames[i].ToString(), pOldPropValues[i].ToString());
                //    }
                //        else
                //    {
                //        pNewPS.SetProperty("VERSION", pNewVersion.VersionName);
                //    }
                //}
                /// Change Version
                SwizzleDatasets(pApp, (IFeatureWorkspace)pVersion, (IFeatureWorkspace)pNewVersion);
                Workspace = (IWorkspace)pNewVersion;
                ScriptEngine.BroadcastProperty("Workspace", this.Workspace, this);
                SW1.Stop();
                RecordActionTime("OpenSession", SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception Ex)
            {
                Logger.WriteLine("Error in OpenSession:" + Ex.Message);
                return false;
            }
        }

        private bool ArcFMQA(string sSelectEdits)
        {
            IWorkspaceEdit pWKSEdit;

            try
            {
                SW1.Reset();
                SW1.Start();
                if (sSelectEdits.ToUpper() == "TRUE")
                {
                    pWKSEdit = (IWorkspaceEdit)this.Workspace;
                    if (pWKSEdit.IsBeingEdited() == true)
                    {

                    }
                }
                CommandClick("{6C0261C0-E400-11D3-B4A0-006008AD9A5E}");
                SW1.Stop();
                RecordActionTime("ArcFMQA", SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("ArcFMQA Error" + EX.Message + " " + EX.StackTrace);

                return false;
            }
        }

        private bool ArcFMGasTrace(string sTraceType)
        {
            ESRI.ArcGIS.EditorExt.INetworkAnalysisExt _pNetworkAnalysisExt = null;
            ESRI.ArcGIS.EditorExt.INetworkAnalysisExtFlags _pNetworkAnalysisExtFlags = null;
            ESRI.ArcGIS.EditorExt.INetworkAnalysisExtResults _pNetworkAnalysisExtResults = null;
            ESRI.ArcGIS.EditorExt.ITraceTask _pArcFMGasValveIsolationTask = null;
            ESRI.ArcGIS.Geodatabase.IGeometricNetwork _pGeometricNetwork = null;
            ESRI.ArcGIS.Geodatabase.IEnumNetEID _pResultJunctions = null;
            ESRI.ArcGIS.Geodatabase.IEnumNetEID _pResultEdges = null;
            Miner.Interop.IMMTraceUIUtilities _pTraceUtils = null;
            ArrayList _pBackFeedStarterEdges = null;
            ArrayList _pResultEdgeEIDs = null;
            ArrayList _pResultJunctionEIDs = null;

            try
            {
                SW1.Reset();
                SW1.Start();
                IMxDocument pMXDoc;
                IApplication pApp;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;

                _pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByName("Utility Network Analyst");
                if (_pNetworkAnalysisExt != null)
                {
                    _pNetworkAnalysisExtFlags = (INetworkAnalysisExtFlags)_pNetworkAnalysisExt;
                    _pNetworkAnalysisExtResults = (INetworkAnalysisExtResults)_pNetworkAnalysisExt;
                    _pGeometricNetwork = _pNetworkAnalysisExt.CurrentNetwork;
                    _pTraceUtils = new Miner.Interop.MMTraceUIUtilities();
                    if (_pTraceUtils != null)
                    {
                        _pArcFMGasValveIsolationTask = _pTraceUtils.GetTraceTaskByName(_pNetworkAnalysisExt, "ArcFM Gas Valve Isolation");
                        if (_pArcFMGasValveIsolationTask != null)
                        {
                            // Set current task on the network.
                            ITraceTasks pTraceTasks = (ITraceTasks)_pNetworkAnalysisExt;
                            pTraceTasks.CurrentTask = _pArcFMGasValveIsolationTask;
                            // Set up the flag symbol. (Only necessary for ITraceTask integration,
                            // since flags are not displayed as part of this application.)
                            // Allocate supporting structures.
                            _pBackFeedStarterEdges = new ArrayList();
                            _pBackFeedStarterEdges.Clear();
                            INetElements pNetElements = (INetElements)_pGeometricNetwork.Network;

                            // Start from scratch.
                            _pResultEdgeEIDs = new ArrayList();
                            _pResultEdgeEIDs.Clear();
                            _pResultJunctionEIDs = new ArrayList();
                            _pResultJunctionEIDs.Clear();

                            // Accumulate results.
                            ITraceTaskResults pArcFMGasValveIsolationTaskResults = (ITraceTaskResults)_pArcFMGasValveIsolationTask;
                            _pResultJunctions = pArcFMGasValveIsolationTaskResults.ResultJunctions;
                            _pResultJunctions.Reset();
                            for (int iCntr = 0; iCntr < _pResultJunctions.Count; iCntr++)
                            {
                                int iEID = _pResultJunctions.Next();
                                if (!(_pResultJunctionEIDs.Contains(iEID)))
                                    _pResultJunctionEIDs.Add(iEID);
                            }
                            _pResultEdges = pArcFMGasValveIsolationTaskResults.ResultEdges;
                            _pResultEdges.Reset();
                            for (int iCntr = 0; iCntr < _pResultEdges.Count; iCntr++)
                            {
                                int iEID = _pResultEdges.Next();
                                if (!_pResultEdgeEIDs.Contains(iEID))
                                    _pResultEdgeEIDs.Add(iEID);
                            }
                        }
                    }
                    SW1.Stop();
                    RecordActionTime("ArcFMGasTrace", SW1.ElapsedMilliseconds);
                }
                else
                {
                    this.Logger.WriteLine("Missing Utility Network AnalysisExt");
                    return false;
                }
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in ArcFMGasTrace " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        #endregion
#region Esri_STuff
        private bool QueryByAttribute(string sClassName, string sQuery)
        {
            IObjectClass pOC;
            ITable pTable;
            IFeatureWorkspace pFWKS;
            IRow pRow;
            IQueryFilter pQF;
            ICursor pCursor;
            long QueryTime;
            long FetchTime;
            int ReturnCount;
            this.Logger.WriteLine("Query by Attribute");
            try
            {
                SW1.Reset();
                SW1.Start();
                try
                {
                    pFWKS = (IFeatureWorkspace)Workspace;
                    pOC = (IObjectClass)pFWKS.OpenTable(sClassName);
                    try
                    {
                        pQF = new QueryFilter();
                        pQF.WhereClause = sQuery;
                        pTable = (ITable)pOC;
                        try
                        {
                            ReturnCount = 0;
                            SW1.Reset();
                            SW1.Start();
                            pCursor = pTable.Search(pQF, true);
                            QueryTime = SW1.ElapsedMilliseconds;
                            pRow = pCursor.NextRow();
                            while (pRow != null)
                            {
                                ReturnCount = ReturnCount + 1;
                                pRow = pCursor.NextRow();
                            }
                            SW1.Stop();
                            FetchTime = SW1.ElapsedMilliseconds - QueryTime;
                            RecordActionTime("QuerByAttribute Query Time:" , QueryTime);
                            RecordActionTime("QuerByAttribute Fetch Time:" , FetchTime);
                            RecordActionTime("QueryByAttribute Total", SW1.ElapsedMilliseconds);
                            return true;
                        }
                        catch (Exception EX)
                        {
                            this.Logger.WriteLine("Failed to query class:" + EX.Message);
                            return false;
                        }
                    }
                    catch (Exception EX2)
                    {
                        this.Logger.WriteLine("Failed to obtain Field:" + EX2.Message);
                        return false;
                    }
                }
                catch (Exception EX)
                {
                    this.Logger.WriteLine("Failed to open Class:" + EX.Message);
                    return false;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Failed:" + EX.Message);
                return false;
            }
        }

        private bool SelectByAttribute(string sLayerName, string sQuery)
        {

            IMxDocument pMXDoc;
            IEnumLayer pLayers;
            ILayer pLayer;
            IFeatureLayer pFLayer;
            IFeatureSelection pFLSel;
            IQueryFilter pQF;
            ISelectionSet pSelSet;

            try
            {
                SW1.Reset();
                SW1.Start();
                //Find Layer
                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if (pLayer.Name.ToUpper() == sLayerName.ToUpper())
                    {
                        if (pLayer is IFeatureLayer)
                        {
                            pFLayer = (IFeatureLayer)pLayer;
                            pFLSel = (IFeatureSelection)pFLayer;
                            pQF = new QueryFilter();
                            pQF.WhereClause = sQuery;
                            pFLSel.SelectFeatures(pQF, esriSelectionResultEnum.esriSelectionResultNew, false);
                            pSelSet = pFLSel.SelectionSet;
                            SW1.Stop();
                            RecordActionTime("SelectLayerByAttribute Features ("+ pSelSet.Count + ")", SW1.ElapsedMilliseconds);
                            return true;
                        }
                    }
                    pLayer = pLayers.Next();
                }
                this.Logger.WriteLine("SelectLayerByAttriute:Layer Not Found");
                return false;
            }
            catch (COMException ComEX)
            {
                this.Logger.WriteLine("Failed:"+ ComEX.ErrorCode + " " + ComEX.Message + " " + ComEX.StackTrace);
                return false;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Failed:" + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool SpatialSelectByAttribute(string sLayerName, string sQuery, string sMinX, string sMinY, string sMaxX, string sMaxY)
        {
            IEnumLayer pLayers;
            ILayer pLayer;
            IFeatureLayer pFLayer;
            IFeatureSelection pFLSel;
            ISelectionSet pSelSet;
            IEnvelope pEnv;
            IMxDocument pMXDoc;
            ISpatialFilter pSF;
            IFeatureClass pFC;


            try
            {
                SW1.Reset();
                SW1.Start();
                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if (pLayer.Name == sLayerName)
                    {
                        if (pLayer is IFeatureLayer)
                        {
                            pFLayer = (IFeatureLayer)pLayer;
                            pFLSel = (IFeatureSelection)pFLayer;
                            pFC = pFLayer.FeatureClass;
                            pEnv = new EnvelopeClass();
                            //pEnv.SpatialReference = pFC.
                            pEnv.PutCoords(System.Convert.ToDouble(sMinX), System.Convert.ToDouble(sMinY), System.Convert.ToDouble(sMaxX), System.Convert.ToDouble(sMaxY));
                            pSF = new SpatialFilter();
                            pSF.WhereClause = sQuery;
                            pSF.Geometry = pEnv;
                            pSF.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
                            pFLSel.SelectFeatures(pSF, esriSelectionResultEnum.esriSelectionResultNew, false);
                            pSelSet = pFLSel.SelectionSet;
                            SW1.Stop();
                            RecordActionTime ("SpatialSelectbyAttribute Features (" + pSelSet.Count + ")" ,SW1.ElapsedMilliseconds);
                            return true;
                        }
                    }
                    pLayer = pLayers.Next();
                }
                return false;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Failed:" + EX.Message);
                return false;
            }
        }

        private bool SpatialSelect(string sClassName, string sMinX, string sMinY, string sMaxX, string sMaxY)
        {
            IObjectClass pOC;
            ITable pTable;
            IFeatureWorkspace pFWKS;
            ISpatialFilter pSQF;
            ISelectionSet2 pSelSet;
            IFeatureClass pFC;
            IEnvelope pEnv;

            this.Logger.WriteLine("Spatial Select by Attribute");
            try
            {
                try
                {
                    pFWKS = (IFeatureWorkspace)Workspace;
                    pOC = (IObjectClass)pFWKS.OpenTable(sClassName);
                    try
                    {
                        pFC = (IFeatureClass)pOC;
                        pSQF = new SpatialFilter();
                        pSQF.GeometryField = pFC.ShapeFieldName;
                        pEnv = new EnvelopeClass();
                        pEnv.PutCoords(System.Convert.ToDouble(sMinX), System.Convert.ToDouble(sMinY), System.Convert.ToDouble(sMaxX), System.Convert.ToDouble(sMaxY));
                        pTable = (ITable)pOC;
                        try
                        {
                            SW1.Reset();
                            SW1.Start();
                            pSelSet = (ISelectionSet2)pTable.Select(pSQF, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, Workspace);
                            SW1.Stop();
                            RecordActionTime("SpatialSelect Features (" + pSelSet.Count + ")" , SW1.ElapsedMilliseconds);
                            return true;
                        }
                        catch (Exception EX)
                        {
                            this.Logger.WriteLine("Failed to query class:" + EX.Message);
                            return false;
                        }
                    }
                    catch (Exception EX2)
                    {
                        this.Logger.WriteLine("Failed to obtain Field:" + EX2.Message);
                        return false;
                    }
                }
                catch (Exception EX)
                {
                    this.Logger.WriteLine("Failed to open Class:" + EX.Message);
                    return false;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Failed:" + EX.Message);
                return false;
            }
        }

        private bool MapSpatialSelect(string sMinX, string sMinY, string sMaxX, string sMaxY)
        {
            IEnvelope pEnv;
            IMxDocument pDoc;
            IMap pMap;

            this.Logger.WriteLine("Map Spatial Select");
            try
            {
                try
                {
                    pDoc = GetDoc();
                    pMap = (IMap)pDoc.FocusMap;
                    pEnv = new EnvelopeClass();
                    pEnv.PutCoords(System.Convert.ToDouble(sMinX), System.Convert.ToDouble(sMinY), System.Convert.ToDouble(sMaxX), System.Convert.ToDouble(sMaxY));
                    ISelectionEnvironment pSelEnv = new SelectionEnvironment();
                    pSelEnv.AreaSelectionMethod = esriSpatialRelEnum.esriSpatialRelIntersects;
                    pSelEnv.CombinationMethod = esriSelectionResultEnum.esriSelectionResultNew;

                    SW1.Reset();
                    SW1.Start();
                    pMap.SelectByShape(pEnv, pSelEnv, false);
                    SW1.Stop();
                    RecordActionTime("MapSpatialSelect Features (" + pMap.SelectionCount + ")", SW1.ElapsedMilliseconds);
                    return true;
                }
                catch (Exception EX)
                {
                    this.Logger.WriteLine("Failed to open Class:" + EX.Message);
                    return false;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Failed:" + EX.Message);
                return false;
            }
        }

        public bool ZOOMTOSELECTED()
        {
            IMxDocument pMXDoc;
            IDocument pDoc;
            IApplication pApp;
            IEnvelope pPoly;
            IMap pMap;
            IActiveView pAV;
            try
            {
                SW1.Reset();
                SW1.Start();

                //Type t = Type.GetTypeFromCLSID(typeof(AppRefClass).GUID);
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                pDoc = pApp.Document;
                pMap = pMXDoc.FocusMap;
                    if (pMap.SelectionCount > 0)
                    {
                    pAV = pMXDoc.ActiveView;
                    pPoly = new EnvelopeClass();
                    ISelection selection = pMap.FeatureSelection;
                    IEnumFeature enumFeature = selection as IEnumFeature;
                    enumFeature.Reset();
                    IFeature feature;
                    feature = enumFeature.Next();
                    while (feature != null)
                    {
                        pPoly.Union(feature.ShapeCopy.Envelope);
                        feature = enumFeature.Next();
                    }

                    //pPoly = (IEnvelope)pGeomBag;
                    if (pPoly == null)
                    {
                        this.Logger.WriteLine("No Geometry");
                    }
                    else
                    {
                        this.Logger.WriteLine("X1:" + pPoly.XMin + " Y1:" + pPoly.YMin + " X2:" + pPoly.XMax + " Y2:" + pPoly.YMax);
                        if (pPoly.XMin == pPoly.XMax)
                        {
                            pPoly.Expand(1000, 1000,false);
                        }
                        pAV.Extent = (IEnvelope)pPoly;
                        pAV.ContentsChanged();
                        pAV.Refresh();
                    }
                    SW1.Stop();
                    RecordActionTime("ZoomToSelected Time", SW1.ElapsedMilliseconds);
                    return true;
                }
                else
                {
                    SW1.Stop();
                    this.Logger.WriteLine("ZoomToSelected:no features selected:");
                    return false;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("ZoomToSelected error:" + EX.Message + " Stack:" + EX.StackTrace);
                return false;
            }
        }

        public bool CLEARSELECTION()
        {
            IMxDocument pMXDoc;
            IApplication pApp;
            ICommandItem pCmdItem;
            IDocument pDoc;

            SW1.Reset();
            SW1.Start();
            //Type t = Type.GetTypeFromCLSID(typeof(AppRefClass).GUID);
            Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
            System.Object obj = Activator.CreateInstance(t);
            pApp = obj as IApplication;
            pMXDoc = (IMxDocument)pApp.Document;
            pDoc = pApp.Document;
            //pMXDoc.FocusMap.ClearSelection();
            UID pUID = new UID();
            pUID.Value = "{37C833F3-DBFD-11D1-AA7E-00C04FA37860}";
            pCmdItem = pDoc.CommandBars.Find(pUID, false, false);
            if (pCmdItem != null)
            {
                pCmdItem.Execute();
                pMXDoc.ActiveView.Refresh();
            }
            else
            {
                this.Logger.WriteLine("Command not found");
            }

            SW1.Stop();
            RecordActionTime("ClearSelection" ,SW1.ElapsedMilliseconds);
            return true;
        }

        public bool UpdateStringAttribute(string sClassName,string sWhere, string sUpdates)
        {
            IApplication pApp;

            IObjectClass pOC;
            ITable pTable;
            IRow pRow;
            ICursor pCursor;
            IFeatureWorkspace pFWKS;
            IQueryFilter pQF;
            UID eUID = new ESRI.ArcGIS.esriSystem.UIDClass();
            IWorkspaceEdit pWKSEdit;
            string[] sUpdateArray;
            bool bAborted;
            try
            {
                sUpdateArray = GetParameters(sUpdates);
                if (sUpdateArray.Count() % 1 != 0)
                {
                }
                else
                {
                    try
                    {
                        Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                        System.Object obj = Activator.CreateInstance(t);
                        pApp = obj as IApplication;

                        SW1.Reset();
                        SW1.Start();
                        pFWKS = (IFeatureWorkspace)Workspace;
                        pWKSEdit = (IWorkspaceEdit)Workspace;
                        pOC = (IObjectClass)pFWKS.OpenTable(sClassName);
                        try
                        {
                            pQF = new QueryFilter();
                            pQF.WhereClause = sWhere;
                            pTable = (ITable)pOC;
                            try
                            {
                                pCursor = pTable.Search(pQF, false);
                                Logger.WriteLine("Got Cursor");
                                pRow = pCursor.NextRow();
                                Logger.WriteLine("Got Row");
                                //pWKSEdit.StartEditOperation();
                                bAborted = false;
                                pEditor.StartOperation();
                                Logger.WriteLine("Started Editor");
                                int iPos;
                                while (pRow != null)
                                {
                                    iPos = 0;
                                    for (int i = 0; i < ((sUpdateArray.Count()+1)/2); i++)
                                    {
                                        int iFld = pRow.Fields.FindField(sUpdateArray[iPos]);
                                        if (iFld > 0)
                                        {
                                            Logger.WriteLine("Field:" + iFld);
                                            try
                                            {
                                                pRow.set_Value(iFld, sUpdateArray[iPos + 1]);
                                            }
                                            catch (Exception EX2)
                                            {
                                                Logger.WriteLine("Error in Update Attribute:" + EX2.Message);
                                                bAborted = true;
                                            }
                                        }
                                        else
                                        {
                                            Logger.WriteLine("Field not found:" + sUpdateArray[i]);
                                        }
                                        iPos = iPos + 2;
                                    }
                                    if (!bAborted)
                                    {
                                        pRow.Store();
                                    }
                                    pRow = pCursor.NextRow();
                                }
                                //pWKSEdit.StopEditOperation();
                                pEditor.StopOperation("Update Attribute");
                                SW1.Stop();
                                RecordActionTime("UPDATE_ATTRIBUTE", SW1.ElapsedMilliseconds);
                                return true;
                            }
                            catch (Exception EX)
                            {
                                this.Logger.WriteLine("UpdateStringAttribute error:" + EX.Message + " " + EX.StackTrace);
                                return false;
                            }
                        }
                        catch (Exception EX)
                        {
                            this.Logger.WriteLine("UpdateStringAttribute error:" + EX.Message + " " + EX.StackTrace);
                            return false;
                        }
                    }
                    catch (Exception EX)
                    {
                        this.Logger.WriteLine("UpdateStringAttribute error:" + EX.Message + " " + EX.StackTrace);
                        return false;
                    }
                }
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("UpdateStringAttribute error:" + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool CommandClick(string sCommandID)
        {
            // Programmatically "clicks" an ArcMap command button.

            try
            {
                UID pCUID;
                ICommandBars pCommandBars;
                ICommandItem pCommandItem1;
                ICommand pCommand;
                IApplication pApp;
                string sName;
                bool bReturn;

                SW1.Reset();
                SW1.Start();

                bReturn = false;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;

                try
                {

                    pCUID = new UID();
                    pCUID.Value = sCommandID;
                }
                catch (Exception EX2)
                {
                    Logger.WriteLine("Error on CommandButton on finding Tool:" + EX2.Message);
                    return false;
                }
                pCommandBars = pApp.Document.CommandBars;
                pCommandItem1 = pCommandBars.Find(pCUID);
                if (pCommandItem1 != null)
                {
                    if (pCommandItem1.Type == esriCommandTypes.esriCmdTypeCommand)
                    {
                        pCommand = (ICommand)pCommandItem1;
                        pCommand.OnClick();
                        SW1.Stop();
                        try
                        {
                            sName = pCommand.Name;
                        }
                        catch (Exception EX)
                        {
                            this.Logger.WriteLine("CommandClick Error" + EX.Message + " " + EX.StackTrace);

                            sName = "Not Found";
                        }
                        RecordActionTime("COMMANDBUTTON (" + sName + ")",SW1.ElapsedMilliseconds);
                        bReturn = true;
                    }
                    else
                    {
                        Logger.WriteLine("Invalid Type");
                    }
                }
                else
                {
                    Logger.WriteLine("Not found");
                }
                return bReturn;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("CommandButton error:" + EX.Message);
                return false;
            }
        }

        private bool SetScale(string sScale)
        {
            try
            {
                IApplication pApp;
                IMap pMap;
                IMxDocument pMXDoc;
                bool bReturn;

                SW1.Reset();
                SW1.Start();

                bReturn = false;
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                pMap = pMXDoc.FocusMap;
                pMap.MapScale = System.Convert.ToDouble(sScale);
                pMXDoc.ActiveView.Refresh();
                SW1.Stop();
                RecordActionTime("SETSCALE" ,SW1.ElapsedMilliseconds);

                return bReturn;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("SetScale error:" + EX.Message);
                return false;
            }
        }

        private bool CreateVersion(string sName, string sAddGUID)
        {
            IVersion pVersion;
            IVersion pNewVersion;
            IChangeDatabaseVersion pCV;
            ISet pVSet;
            IApplication pApp;
            IMxDocument pMXDoc;
            System.Guid VGuid;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;

                VGuid = System.Guid.NewGuid();
                pVersion = (IVersion)Workspace;
                if (sAddGUID.ToUpper() == "TRUE")
                {
                    pNewVersion = pVersion.CreateVersion(sName + "_" + VGuid.ToString());
                }
                else
                {
                    pNewVersion = pVersion.CreateVersion(sName);
                }
                pCV = new ChangeDatabaseVersion();
                pVSet = pCV.Execute(pVersion, pNewVersion, (IBasicMap)pMXDoc.FocusMap);
                Workspace = (IWorkspace)pNewVersion;
                ScriptEngine.BroadcastProperty("Workspace", Workspace, this);

                SW1.Stop();

                RecordActionTime("CREATEVERSION:",SW1.ElapsedMilliseconds);

                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("CreateVersion error:" + EX.Message);
                return false;
            }
        }

        private bool ChangeVersion(string sName)
        {
            IVersion pVersion;
            IVersion pNewVersion;
            IVersionedWorkspace pVWKS;
            IChangeDatabaseVersion pCV;
            ISet pVSet;
            IApplication pApp;
            IMxDocument pMXDoc;
            IWorkspace pWKS;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;

                pWKS = Workspace;
                pVWKS = (IVersionedWorkspace)Workspace;
                pVersion = (IVersion)pWKS;
                pNewVersion = pVWKS.FindVersion(sName);
                pCV = new ChangeDatabaseVersion();
                pVSet = pCV.Execute(pVersion, pNewVersion, (IBasicMap)pMXDoc.FocusMap);
                Workspace = (IWorkspace)pNewVersion;
                //SwizzleDatasets(pApp, (IFeatureWorkspace)pVersion, (IFeatureWorkspace)pNewVersion);
                ScriptEngine.BroadcastProperty("Workspace", Workspace, this);

                SW1.Stop();
                RecordActionTime("CHANGEVERSION" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("ChangeVersion error:" + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool ZoomToFullExtents()
        {
            IMxDocument pMXDoc;
            IEnvelope pEnv;

            try
            {
                SW1.Reset();
                SW1.Start();

                pMXDoc = GetDoc();
                pEnv = pMXDoc.ActiveView.FullExtent;
                pMXDoc.ActiveView.Extent = pEnv;
                pMXDoc.ActiveView.Refresh();

                SW1.Stop();
                RecordActionTime("ZOOMTOFULLEXTENT" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("ZoomToFullExtents error:" + EX.Message);
                return false;
            }
        }

        private bool ZoomToScale(string sX, string sY, string sScale)
        {
            IMxDocument pMXDoc;
            IEnvelope pEnv;
            IPoint pPoint;
            double dScale;
            ESRI.ArcGIS.Framework.IAppROT pAppRot = new ESRI.ArcGIS.Framework.AppROT();
            ESRI.ArcGIS.Framework.IApplication pApp;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXDoc = (IMxDocument)pApp.Document;
                // Convert X,Y to point
                pPoint = new Point();
                pPoint.SpatialReference = pMXDoc.FocusMap.SpatialReference;

                pPoint.X = System.Convert.ToDouble(sX);
                pPoint.Y = System.Convert.ToDouble(sY);
                dScale = 1.0 / System.Convert.ToDouble(sScale);
                // Convert Scale to double
                Logger.WriteLine("Scale:" + dScale.ToString() + " X:" + pPoint.X.ToString() + " Y:" + pPoint.Y.ToString());

                pMXDoc = GetDoc();
                pEnv = pMXDoc.ActiveView.FullExtent;
                pMXDoc.ActiveView.Extent = pEnv;
                pMXDoc.ActiveView.Refresh();

                SW1.Stop();
                RecordActionTime("ZOOMTOFULLEXTENT" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("ZoomToFullExtents error:" + EX.Message);
                return false;
            }
        }

        private bool SetLayersVisibility(string sLayerName, string sStatus)
        {
            IMxDocument pMXDoc;
            IEnumLayer pLayers;
            ILayer pLayer;

            try
            {
                SW1.Reset();
                SW1.Start();

                //Find Layer
                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if ((pLayer.Name == sLayerName) || (sLayerName == "*"))
                    {
                        if (sStatus == "ON")
                        {
                            pLayer.Visible = true;
                        }
                        else
                        {
                            pLayer.Visible = false;
                        }
                        if (sLayerName != "*")
                        {
                            SW1.Stop();
                            RecordActionTime("SETLAYERVISIBILITY" , SW1.ElapsedMilliseconds);
                            return true;
                        }
                    }
                    pLayer = pLayers.Next();
                }
                SW1.Stop();
                RecordActionTime("SETLAYERVISIBILITY" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("SetLayersVisibility error:" + EX.Message);
                return false;
            }
        }

        private bool SetSelectable(string sLayerName, string sStatus)
        {
            IMxDocument pMXDoc;
            IEnumLayer pLayers;
            ILayer pLayer;
            IFeatureLayer pFLayer;

            try
            {
                SW1.Reset();
                SW1.Start();

                //Find Layer
                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if ((pLayer.Name == sLayerName) || (sLayerName == "*"))
                    {
                        if (pLayer is IFeatureLayer)
                        {
                            pFLayer = (IFeatureLayer)pLayer;
                            if (sStatus == "ON")
                            {
                                pFLayer.Selectable = true;
                            }
                            else
                            {
                                pFLayer.Selectable = false;
                            }
                        }
                        if (sLayerName != "*")
                        {
                            SW1.Stop();
                            RecordActionTime("SETSELECTABLE:" , SW1.ElapsedMilliseconds);
                            return true;
                        }
                    }
                    pLayer = pLayers.Next();
                }
                SW1.Stop();
                RecordActionTime("SETSELECTABLE:" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("SetSelectable error:" + EX.Message);
                return false;
            }
        }

        private bool AddToSelection(string sLayerName, string sOID)
        {
            IMxDocument pMXDoc;
            IEnumLayer pLayers;
            ILayer pLayer;
            IFeatureLayer pFLayer;
            IFeatureSelection pFSel;
            IFeatureClass pFC;
            IFeature pFeat;

            try
            {
                SW1.Reset();
                SW1.Start();

                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if (pLayer.Name == sLayerName)
                    {
                        if (pLayer is IFeatureLayer)
                        {
                            pFLayer = (IFeatureLayer)pLayer;
                            pFSel = (IFeatureSelection)pFLayer;
                            pFC = pFLayer.FeatureClass;
                            pFeat = pFC.GetFeature(Convert.ToInt32(sOID));
                            if (pFeat != null)
                            {
                                pFSel.Add(pFeat);
                            }
                        }
                    }
                    pLayer = pLayers.Next();
                }
                SW1.Stop();
                RecordActionTime("ADDTOSELECTION:" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("AddToSelection error:" + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool RemoveFromSelecdtion(string sLayerName, string sOID)
        {
            IMxDocument pMXDoc;
            IEnumLayer pLayers;
            ILayer pLayer;
            IFeatureLayer pFLayer;
            IFeatureSelection pFSel;
            ISelectionSet pSelSet;
            IGeoDatabaseBridge2 pGDBBridge;
            int[] OIDList;

            try
            {
                SW1.Reset();
                SW1.Start();

                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if (pLayer.Name == sLayerName)
                    {
                        if (pLayer is IFeatureLayer)
                        {
                            pFLayer = (IFeatureLayer)pLayer;
                            pFSel = (IFeatureSelection)pFLayer;
                            pSelSet = pFSel.SelectionSet;
                            pGDBBridge = new GeoDatabaseHelperClass();
                            OIDList = new int[1];
                            OIDList[0] = Convert.ToInt32(sOID);
                            pGDBBridge.RemoveList(pSelSet, OIDList);
                        }
                    }
                    pLayer = pLayers.Next();
                }
                SW1.Stop();
                RecordActionTime("REMOVEFROMSELECTION:" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("RemoveFromSelection error:" + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool ClearFlags()
        {
            INetworkAnalysisExt pNetworkAnalysisExt;
            IMxDocument pMXDoc;
            IApplication pApp;
            IMxApplication pMXApp;
            INetworkAnalysisExtFlags pNetworkAnalysisExtFlags;
            UID pUID;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXApp = (IMxApplication)pApp;
                pMXDoc = (IMxDocument)pApp.Document;
                pUID = new UID();
                pUID.Value = "esriEditorExt.UtilityNetworkAnalysisExt";
                pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(pUID);
                pNetworkAnalysisExtFlags = (INetworkAnalysisExtFlags)pNetworkAnalysisExt;
                pNetworkAnalysisExtFlags.ClearFlags();
                SW1.Stop();
                RecordActionTime("CLEARFLAGS" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in ClearFlags " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool ClearBarriers()
        {
            INetworkAnalysisExt pNetworkAnalysisExt;
            IMxDocument pMXDoc;
            IApplication pApp;
            IMxApplication pMXApp;
            INetworkAnalysisExtBarriers pNetworkAnalysisExtBarriers;
            UID pUID;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXApp = (IMxApplication)pApp;
                pMXDoc = (IMxDocument)pApp.Document;
                pUID = new UID();
                pUID.Value = "esriEditorExt.UtilityNetworkAnalysisExt";
                pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(pUID);
                pNetworkAnalysisExtBarriers = (INetworkAnalysisExtBarriers)pNetworkAnalysisExt;
                pNetworkAnalysisExtBarriers.ClearBarriers();
                SW1.Stop();
                RecordActionTime("CLEARBARRIERS" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in ClearBarriers " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool ClearResults()
        {
            INetworkAnalysisExt pNetworkAnalysisExt;
            IMxDocument pMXDoc;
            IApplication pApp;
            IMxApplication pMXApp;
            INetworkAnalysisExtResults pNetworkAnalysisExtResults;
            UID pUID;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXApp = (IMxApplication)pApp;
                pMXDoc = (IMxDocument)pApp.Document;
                pUID = new UID();
                pUID.Value = "esriEditorExt.UtilityNetworkAnalysisExt";
                pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(pUID);
                pNetworkAnalysisExtResults = (INetworkAnalysisExtResults)pNetworkAnalysisExt;
                pNetworkAnalysisExtResults.ClearResults();
                SW1.Stop();
                RecordActionTime("CLEARRESULTS" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in ClearResults " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool PlaceFlag(string sClassName, string sOID)
        {
            IFeatureClass pFC;
            IFeatureWorkspace pFWKS;
            IFeature pFeat;
            INetElements pNetElements;
            ISimpleJunctionFeature pJunc;
            INetworkAnalysisExt pNetworkAnalysisExt;
            IMxDocument pMXDoc;
            IApplication pApp;
            IMxApplication pMXApp;
            INetworkAnalysisExtFlags pNetworkAnalysisExtFlags;
            UID pUID;
            ITraceFlowSolverGEN pTraceFlowSolver;
            INetSolver pNetSolver;
            Miner.Interop.IMMTraceUIUtilities pTraceUtils;
            IPolycurve pPoly;
            int iCLSID;
            int iOID;
            int iSUBID;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXApp = (IMxApplication)pApp;
                pMXDoc = (IMxDocument)pApp.Document;
                pUID = new UID();
                pUID.Value = "esriEditorExt.UtilityNetworkAnalysisExt";
                pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(pUID);
                pNetworkAnalysisExtFlags = (INetworkAnalysisExtFlags)pNetworkAnalysisExt;
                pNetElements = (INetElements)pNetworkAnalysisExt.CurrentNetwork.Network;
                pTraceUtils = new Miner.Interop.MMTraceUIUtilities();

                pFWKS = (IFeatureWorkspace)Workspace;
                pFC = pFWKS.OpenFeatureClass(sClassName);
                if (pFC != null)
                {
                    pFeat = pFC.GetFeature(Convert.ToInt32(sOID));
                    INetworkFeature pNetFeat = (INetworkFeature)pFeat;
                    m_pGN = pNetFeat.GeometricNetwork;
                    if (pFeat != null)
                    {
                        pTraceFlowSolver = (ITraceFlowSolverGEN)new TraceFlowSolver();
                        pNetSolver = (INetSolver)pTraceFlowSolver;
                        if (pFeat.FeatureType == esriFeatureType.esriFTSimpleJunction)
                        {
                            pJunc = (ISimpleJunctionFeature)pFeat;
                            pNetElements.QueryIDs(pJunc.EID, esriElementType.esriETEdge, out iCLSID, out iOID, out iSUBID);
                            //pJFlag = (INetFlag)new JunctionFlag();
                            //pJFlag.UserClassID = iCLSID;
                            //pJFlag.UserID = iOID;
                            //pJFlag.UserSubID = iSUBID;
                            IFlagDisplay pFlagDisplay = new JunctionFlagDisplay();
                            pFlagDisplay.Symbol = (ISymbol)CreateMarkerSymbol(0);
                            pFlagDisplay.Geometry = (IPoint)pFeat.ShapeCopy;
                            //Set pNetSolver.SourceNetwork = pNetElements
                            pNetworkAnalysisExtFlags.AddJunctionFlag((IJunctionFlagDisplay)pFlagDisplay);
                        }
                        else
                        {
                            IPoint pPoint;
                            pPoint = new Point();
                            double dDist;
                            dDist = .5;
                            pPoly = (IPolycurve)pFeat.ShapeCopy;
                            pPoly.QueryPoint(esriSegmentExtension.esriNoExtension, dDist, true, pPoint);
                            IFlagDisplay pFlagDisplay = new EdgeFlagDisplay();
                            IEdgeFeature pEdge = (IEdgeFeature)pFeat;
                            IObjectClass pOC = (IObjectClass)pFeat.Table;
                            int iEID = pNetElements.GetEID(pOC.ObjectClassID, pFeat.OID, 0, esriElementType.esriETEdge);
                            //pNetElements.QueryIDs(pEdge.ToJunctionEID, esriElementType.esriETEdge, out iCLSID, out iOID, out iSUBID);
                            pFlagDisplay.FeatureClassID = pOC.ObjectClassID;
                            pFlagDisplay.FID = pFeat.OID;
                            pFlagDisplay.SubID = 0;
                            //pFlagDisplay.Geometry = pPoint;
                            pFlagDisplay.Symbol = (ISymbol)CreateMarkerSymbol(0);
                            ISimpleMarkerSymbol pSMS = (ISimpleMarkerSymbol)pFlagDisplay.Symbol;
                            IColor pColor;
                            pColor = new RgbColor();
                            pColor.RGB = 100;
                            pSMS.Size = 10;
                            pNetworkAnalysisExtFlags.AddEdgeFlag((IEdgeFlagDisplay)pFlagDisplay);
                        }
                    }
                }
                SW1.Stop();
                RecordActionTime("PLACEFLAG" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in PlaceFlags " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool PlaceBarrier(string sClassName, string sOID)
        {
            IFeatureClass pFC;
            IFeatureWorkspace pFWKS;
            IFeature pFeat;
            INetElements pNetElements;
            ISimpleJunctionFeature pJunc;
            INetworkAnalysisExt pNetworkAnalysisExt;
            INetElementBarriers pJunctionElementBarriers;
            INetElementBarriers pEdgeElementBarriers;
            IMxDocument pMXDoc;
            IApplication pApp;
            IMxApplication pMXApp;
            INetworkAnalysisExtBarriers pNetworkAnalysisExtBarriers;
            UID pUID;
            INetFlag pJFlag;
            IFlagDisplay pJFlagDisplay;
            ITraceFlowSolverGEN pTraceFlowSolver;
            INetSolver pNetSolver;
            ISelectionSetBarriers pSelectionSetBarriers;
            IPolycurve pPoly;
            int[] iOIDs;
            int iCLSID;
            int iOID;
            int iSUBID;
            double dDist;
            INetworkFeature pNetFeat;
            try
            {
                SW1.Reset();
                SW1.Start();

                try
                {
                    Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                    System.Object obj = Activator.CreateInstance(t);
                    pApp = obj as IApplication;
                    pMXApp = (IMxApplication)pApp;
                    pMXDoc = (IMxDocument)pApp.Document;
                    pUID = new UID();
                    pUID.Value = "esriEditorExt.UtilityNetworkAnalysisExt";
                    pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(pUID);
                    pNetworkAnalysisExtBarriers = (INetworkAnalysisExtBarriers)pNetworkAnalysisExt;
                    pNetElements = (INetElements)pNetworkAnalysisExt.CurrentNetwork.Network;
                    pNetworkAnalysisExtBarriers.CreateElementBarriers(out pJunctionElementBarriers, out pEdgeElementBarriers);
                    pNetworkAnalysisExtBarriers.CreateSelectionBarriers(out pSelectionSetBarriers);
                    iOIDs = new int[1];
                }
                catch (Exception EX2)
                {
                    Logger.WriteLine("Error in PlaceBarriers getting extension " + EX2.Message + " " + EX2.StackTrace);
                    return false;
                }

                pNetworkAnalysisExtBarriers.SelectionSemantics = esriAnalysisType.esriAnalysisOnAllFeatures;

                pFWKS = (IFeatureWorkspace)Workspace;
                pFC = pFWKS.OpenFeatureClass(sClassName);
                if (pFC != null)
                {
                    pFeat = pFC.GetFeature(Convert.ToInt32(sOID));
                    if (pFeat != null)
                    {
                        iOIDs[0] = pFeat.OID;
                        pNetFeat = (INetworkFeature)pFeat;

                        pTraceFlowSolver = (ITraceFlowSolverGEN)new TraceFlowSolver();
                        pNetSolver = (INetSolver)pTraceFlowSolver;
                        if (pFeat.FeatureType == esriFeatureType.esriFTSimpleJunction)
                        {
                            try
                            {
                                IPoint pPoint;
                                pPoint = (IPoint)pFeat.ShapeCopy;
                                pJunc = (ISimpleJunctionFeature)pFeat;
                                pNetElements.QueryIDs(pJunc.EID, esriElementType.esriETJunction, out iCLSID, out iOID, out iSUBID);
                                pJFlag = (INetFlag)new JunctionFlag();
                                pJFlagDisplay = new JunctionFlagDisplay();
                                pJFlagDisplay.Geometry = pPoint;
                                pJFlagDisplay.Symbol = (ISymbol)CreateMarkerSymbol(1);
                                pNetworkAnalysisExtBarriers.AddJunctionBarrier((IJunctionFlagDisplay)pJFlagDisplay);
                            }
                            catch (Exception EX2)
                            {
                                Logger.WriteLine("Error in PlaceBarriers.Creating Flag " + EX2.Message + " " + EX2.StackTrace);
                                return false;
                            }
                        }
                        else
                        {
                            try
                            {
                                IPoint pPoint;
                                pPoint = new Point();
                                dDist = .5;
                                pPoly = (IPolycurve)pFeat.ShapeCopy;
                                pPoly.QueryPoint(esriSegmentExtension.esriExtendAtFrom, dDist, true, pPoint);
                                IFlagDisplay pFlagDisplay = new EdgeFlagDisplay();
                                IEdgeFeature pEdge = (IEdgeFeature)pFeat;
                                pNetElements.QueryIDs(pEdge.ToJunctionEID, esriElementType.esriETJunction, out iCLSID, out iOID, out iSUBID);
                                pFlagDisplay.Geometry = pPoint;
                                pFlagDisplay.Symbol = (ISymbol)CreateMarkerSymbol(1);
                                pNetworkAnalysisExtBarriers.AddEdgeBarrier((IEdgeFlagDisplay)pFlagDisplay);
                            }
                            catch (Exception EX2)
                            {
                                Logger.WriteLine("Error in PlaceBarriers placing barrier on edge feature " + EX2.Message + " " + EX2.StackTrace);
                                return false;
                            }
                        }
                    }
                }

                SW1.Stop();
                RecordActionTime("PLACEBARRIER" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in PlaceBarriers " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool ExecuteNetworkTrace(string sTraceName, string sDoTrace)
        {
            ITraceFlowSolverGEN pTraceFlowSolver;
            INetworkAnalysisExt pNetworkAnalysisExt;
            IMxDocument pMXDoc;
            IApplication pApp;
            IMxApplication pMXApp;
            UID pUID;
            int i;
            bool bFound;
            INetElements pNetElements;
            INetSolver pNetSolver;
            ITraceTask pTraceTask;
            ITraceTasks pTraceTasks;

            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;
                pMXApp = (IMxApplication)pApp;
                pMXDoc = (IMxDocument)pApp.Document;
                pUID = new UID();
                pUID.Value = "esriEditorExt.UtilityNetworkAnalysisExt";
                pNetworkAnalysisExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(pUID);
                pNetElements = (INetElements)m_pGN.Network;
                pTraceFlowSolver = (ITraceFlowSolverGEN)new TraceFlowSolver();
                pNetSolver = pTraceFlowSolver as INetSolver;
                pTraceFlowSolver.TraceIndeterminateFlow = true;
                pNetSolver.SourceNetwork = m_pGN.Network;
                bFound = false;
                for (i = 0; i < pNetworkAnalysisExt.NetworkCount; i++)
                {
                    if (pNetworkAnalysisExt.Network[i].OrphanJunctionFeatureClass.FeatureClassID == m_pGN.OrphanJunctionFeatureClass.FeatureClassID)
                    {
                        pNetworkAnalysisExt.CurrentNetwork = pNetworkAnalysisExt.Network[i];
                        Logger.WriteLine("Network Set");
                        bFound = true;
                        break;
                    }
                }
                if (bFound != true)
                {
                    Logger.WriteLine("No Network Found");
                    return false;
                }
                pTraceTasks = (ITraceTasks)pNetworkAnalysisExt;
                //ITraceTask pTraceTask = new FindConnectedTask();
                bFound = false;
                pTraceTask = null;
                for (i = 0; i < pTraceTasks.TaskCount; i++)
                {
                    if (pTraceTasks.Task[i].Name.ToUpper() == sTraceName.ToUpper())
                    {
                        pTraceTasks.CurrentTask = pTraceTasks.Task[i];
                        pTraceTask = pTraceTasks.Task[i];
                        Logger.WriteLine("Task Set");
                        bFound = true;
                        break;
                    }
                }
                if (bFound != true)
                {
                    Logger.WriteLine("No Task Set");
                    return false;
                }

                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(500);
                if (pTraceTask != null)
                {
                    if (pTraceTask.EnableSolve)
                    {
                        Logger.WriteLine("Execute");
                        //pMXDoc.ActiveView.Refresh();
                        CommandClick("{D6CD856D-BF8F-11D2-BABE-00C04FA33C20}");
                        Logger.WriteLine("GetResults");
                        INetworkAnalysisExtResults pNAResults = (INetworkAnalysisExtResults)pNetworkAnalysisExt;
                        Logger.WriteLine("Found:" + pNAResults.ResultFeatureCount);
                    }
                }
                //}
                SW1.Stop();
                RecordActionTime("EXECUTENETWORKTRACE(" + sTraceName + ")" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (COMException cEx)
            {
                Logger.WriteLine("COMError in EXECUTENETWORKTRACE " + cEx.Message + " " + cEx.StackTrace);
                return false;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in EXECUTENETWORKTRACE " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

#endregion
        private string[] GetParameters(string sInArgs)
        {
            string[] lstArgs;
            if (sInArgs.Length > 0)
            {
                lstArgs = sInArgs.Split(',');
                for (int i = 0; i < lstArgs.GetUpperBound(0); i++)
                {
                    lstArgs[i] = lstArgs[i].Trim();
                    //this.Logger.WriteLine("Arg:[" + i + "]" + lstArgs[i]);
                }
            }
            else
            {
                lstArgs = new string[2] { "", "" };
                this.Logger.WriteLine("No Parameters");
            }
            return lstArgs;
        }

        private IMxDocument GetDoc()
        {
            IMxDocument pMXDoc;
            IApplication pApp;
            Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
            System.Object obj = Activator.CreateInstance(t);
            pApp = obj as IApplication;
            pMXDoc = (IMxDocument)pApp.Document;
            return pMXDoc;
        }
        #region Utilities
        private IDictionary<int, SelectedObjects> GetClassIDs(IGeometricNetwork pGN)
        {
            IDictionary<int, SelectedObjects> dicIDs;
            IEnumFeatureClass pEFC;
            IFeatureClass pFC;
            IDataset pDS;
            SelectedObjects pSelObjs;

            try
            {
                dicIDs = new Dictionary<int, SelectedObjects>();
                pEFC = pGN.get_ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTComplexEdge);
                pFC = pEFC.Next();
                while (pFC != null)
                {
                    if (!dicIDs.ContainsKey(pFC.FeatureClassID))
                    {
                        pDS = (IDataset)pFC;
                        pSelObjs = new SelectedObjects(pFC);
                        dicIDs.Add(pFC.FeatureClassID, pSelObjs);
                    }
                    pFC = pEFC.Next();
                }

                pEFC = pGN.get_ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleEdge);
                pFC = pEFC.Next();
                while (pFC != null)
                {
                    if (!dicIDs.ContainsKey(pFC.FeatureClassID))
                    {
                        pDS = (IDataset)pFC;
                        pSelObjs = new SelectedObjects(pFC);
                        dicIDs.Add(pFC.FeatureClassID, pSelObjs);
                    }
                    pFC = pEFC.Next();
                }

                pEFC = pGN.get_ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleJunction);
                pFC = pEFC.Next();
                while (pFC != null)
                {
                    if (!dicIDs.ContainsKey(pFC.FeatureClassID))
                    {
                        pDS = (IDataset)pFC;
                        pSelObjs = new SelectedObjects(pFC);
                        dicIDs.Add(pFC.FeatureClassID, pSelObjs);
                    }
                    pFC = pEFC.Next();
                }
                Marshal.ReleaseComObject(pEFC);
                return dicIDs;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("GetClassIDs Error" + EX.Message + " " + EX.StackTrace);

                return null;
            }
        }

        private void AddSelection(IMxDocument pMXDoc, IDictionary<int, SelectedObjects> dicClasses)
        {
            IEnumLayer pELayers;
            ILayer pLayer;
            UID pUID;
            IFeatureLayer pFLayer;
            IFeatureClass pFC;
            SelectedObjects pSelObjs;
            IFeatureSelection pFSel;
            IMap pMap;

            try
            {
                pMXDoc = GetDoc();
                pMap = pMXDoc.FocusMap;
                pUID = new UIDClass();
                pUID.Value = "{40A9E885-5533-11d0-98BE-00805F7CED21}";
                pELayers = pMap.get_Layers(pUID, true);
                pLayer = pELayers.Next();
                while (pLayer != null)
                {
                    if (pLayer is IFeatureLayer)
                    {
                        pFLayer = (IFeatureLayer)pLayer;
                        pFC = pFLayer.FeatureClass;
                        if (dicClasses.TryGetValue(pFC.FeatureClassID, out pSelObjs))
                        {
                            if (pSelObjs != null)
                            {
                                try
                                {
                                    pFSel = (IFeatureSelection)pFLayer;
                                    pFSel.SelectionSet = pSelObjs.SelectionSet;
                                }
                                catch (Exception EX2)
                                {
                                    Logger.WriteLine("Error in Selection:" + EX2.Source);
                                }
                            }
                            else
                            {
                                Logger.WriteLine("No Selection Object found");
                            }
                        }
                    }
                    pLayer = pELayers.Next();
                }
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in AddSelection :" + EX.Message + " " + EX.StackTrace);
            }
        }

        private bool ArcFMTrace(string sOID, string sSelect, int iTraceType)
        {
            IFeature pFeat;
            INetworkClass NetCLS;
            IGeometricNetwork pGN;
            ISimpleJunctionFeature pJuncFeat;
            Miner.Interop.IMMElectricTracingEx pElecTrace;
            Miner.Interop.IMMElectricTraceSettings pElecTraceSettings;
            IMMTracedElements pJunctions;
            IMMTracedElements pEdges;
            IMMCurrentStatus pCurrentStatus;
            IMMEnumFeedPath pFeedPaths;
            int[] iJunctionBarriers;
            int[] iEdgeBarriers;
            IApplication pApp;
            IMxDocument pMXDoc;
            IObjectClass pOC;
            ISelection pSel;
            SelectedObjects pSelObjs;
            IDictionary<int, SelectedObjects> dicClasses;
            int iSelectionCount;
            ESRI.ArcGIS.Geodatabase.esriElementType pElemType;
            int iEID;
            IEdgeFeature pEdge;
            IMMEnumTraceStopper pStopperJunctions;
            IMMEnumTraceStopper pStopperEdges;

            iJunctionBarriers = new int[0];
            iEdgeBarriers = new int[0];
            pCurrentStatus = null;
            iSelectionCount = 0;
            try
            {
                pElecTrace = new MMFeederTracerClass();

                pElecTraceSettings = new Miner.Interop.MMElectricTraceSettingsClass();
                pElecTraceSettings.RespectConductorPhasing = true;
                pElecTraceSettings.RespectEnabledField = false;

                SW1.Reset();
                SW1.Start();
                try
                {
                    if (FeatureClass is INetworkClass)
                    {
                        Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                        System.Object obj = Activator.CreateInstance(t);
                        pApp = obj as IApplication;
                        pMXDoc = (IMxDocument)pApp.Document;
                        pSel = pMXDoc.FocusMap.FeatureSelection;

                        NetCLS = (INetworkClass)FeatureClass;
                        pGN = NetCLS.GeometricNetwork;
                        dicClasses = GetClassIDs(pGN);
                        pFeat = FeatureClass.GetFeature(Convert.ToInt32(sOID));
                        if (pFeat is IJunctionFeature)
                        {
                            pElemType = ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction;
                            pJuncFeat = (ISimpleJunctionFeature)pFeat;
                            iEID = pJuncFeat.EID;
                        }
                        else
                        {
                            pElemType = ESRI.ArcGIS.Geodatabase.esriElementType.esriETEdge;
                            pEdge = (IEdgeFeature)pFeat;
                            iEID = pEdge.FromJunctionEID;
                        }

                        pOC = (IObjectClass)pFeat.Class;
                        if (dicClasses.TryGetValue(pOC.ObjectClassID, out pSelObjs))
                        {
                            try
                            {
                                //pSelObjs.SelectionSet.Add(pFeat.OID);
                            }
                            catch (Exception EX2)
                            {
                                this.Logger.WriteLine("ArcFMTrace Error" + EX2.Message + " " + EX2.StackTrace);

                                Logger.WriteLine("Error adding Breaker:" + pFeat.OID);
                            }
                        }

                        //Down Stream
                        if (iTraceType == 1)
                        {
                            pElecTrace.TraceDownstream(pGN, pElecTraceSettings, pCurrentStatus,
                                iEID, pElemType, Miner.Interop.SetOfPhases.abc, Miner.Interop.mmDirectionInfo.establishBySourceSearch, 0,
                                ESRI.ArcGIS.Geodatabase.esriElementType.esriETEdge, iJunctionBarriers,
                                iEdgeBarriers, false, out pJunctions, out pEdges);
                            Logger.WriteLine("Edges Found:" + pEdges.Count);
                            Logger.WriteLine("Junctions Found:" + pJunctions.Count);
                            iSelectionCount = pEdges.Count + pJunctions.Count;
                            if (sSelect.ToUpper() == "TRUE")
                            {
                                SelectFeatures(pEdges, pJunctions, pGN.Network, true);
                            }
                            else
                            {
                                SelectFeatures(pEdges, pJunctions, pGN.Network, false);
                            }
                        }
                        //Up Stream
                        else if (iTraceType == 2)
                        {
                            try
                            {
                                pElecTrace.FindFeedPaths(pGN, pElecTraceSettings, pCurrentStatus,
                                    iEID, pElemType, Miner.Interop.SetOfPhases.abc, iJunctionBarriers,
                                    iEdgeBarriers, out pFeedPaths, out pStopperJunctions, out pStopperEdges);
                                if (pFeedPaths == null)
                                {
                                    iSelectionCount = 0;
                                }
                                else
                                {
                                    iSelectionCount = pFeedPaths.Count;
                                    Logger.WriteLine("Paths:" + pFeedPaths.Count);
                                    IMMFeedPathEx pFeedPath;
                                    IMMEnumPathElement pEPathElem;

                                    pFeedPath = pFeedPaths.Next();
                                    while (pFeedPath != null)
                                    {
                                        pEPathElem = pFeedPath.GetPathElementEnum();
                                        iSelectionCount = iSelectionCount + pEPathElem.Count;
                                        pFeedPath = pFeedPaths.Next();
                                    }
                                }

                                Logger.WriteLine("Count:" + iSelectionCount);
                                if (sSelect.ToUpper() == "TRUE")
                                {
                                    //pFeedPath
                                    //SelectFeatures((IMMTracedElements)pStopperEdges, (IMMTracedElements)pStopperJunctions, pGN.Network, true);
                                }
                                else
                                {
                                    //SelectFeatures((IMMTracedElements)pStopperEdges, (IMMTracedElements)pStopperJunctions, pGN.Network, false);
                                }
                            }
                            catch (Exception EX2)
                            {
                                this.Logger.WriteLine("TraceUpStream Failed: " + EX2.Message + ": " + EX2.StackTrace);
                                return false;
                            }
                        }
                        else
                        {
                            try
                            {
                                MMFeederTracerClass pFTC;
                                IMMNetworkAnalysisExtForFramework pMMNetExt;
                                pFTC = new MMFeederTracerClass();
                                pMMNetExt = new MMNetworkAnalysisExtForFrameworkClass();
                                IEnumNetEID pJunctionsEID;
                                IEnumNetEID pEdgesEID;

                                Miner.Interop.IMMElectricTraceSettingsEx pElecTraceSettingsEX;
                                pElecTraceSettingsEX = new Miner.Interop.MMElectricTraceSettingsClass();
                                pElecTraceSettingsEX.RespectEnabledField = false;
                                pElecTraceSettingsEX.RespectESRIBarriers = false;
                                pElecTraceSettingsEX.UseFeederManagerCircuitSources = true;
                                pElecTraceSettingsEX.UseFeederManagerProtectiveDevices = true;


                                pFTC.DistributionCircuitTrace(pGN, pMMNetExt, pElecTraceSettingsEX, iEID, pElemType, mmPhasesToTrace.mmPTT_AnySinglePhase, out pJunctionsEID, out pEdgesEID);
                                iSelectionCount = 0;
                                if (pJunctionsEID != null)
                                {
                                    iSelectionCount = pJunctionsEID.Count;
                                }
                                if (pEdgesEID != null)
                                {
                                    iSelectionCount = iSelectionCount + pEdgesEID.Count;
                                }
                                Logger.WriteLine("Found:" + iSelectionCount);
                                UID uid = new UIDClass();
                                uid.Value = "esriNetworkAnalystUI.NetworkAnalystExtension"; //ESRI Network Analyst extension
                                //IExtension ext = extMgr.FindExtension(uid);
                                IExtension ext = pApp.FindExtensionByCLSID(uid);
                                if (ext != null)
                                {
                                    IUtilityNetworkAnalysisExt utilityNetworkAnalystExt = ext as IUtilityNetworkAnalysisExt;
                                    if (utilityNetworkAnalystExt != null)
                                    {
                                        INetworkAnalysisExtResults networkAnalysisResults = (INetworkAnalysisExtResults)utilityNetworkAnalystExt;
                                        if (networkAnalysisResults != null)
                                        {
                                            networkAnalysisResults.ClearResults();                                
                                            networkAnalysisResults.SetResults(pJunctionsEID, pEdgesEID);
                                            //determine if results should be drawn

                                            if (sSelect.ToUpper() == "TRUE")
                                            {
                                                networkAnalysisResults.ResultsAsSelection = true;
                                                networkAnalysisResults.CreateSelection(pJunctionsEID, pEdgesEID);
                                            }
                                            else
                                            {
                                                // temporarily toggle the user's setting
                                                networkAnalysisResults.ResultsAsSelection = true;
                                                networkAnalysisResults.CreateSelection(pJunctionsEID, pEdgesEID);
                                                networkAnalysisResults.ResultsAsSelection = false;
                                            }
                                        }
                                        else
                                        {
                                            this.Logger.WriteLine("NetworkAnalysisResults Null");
                                        }
                                    }
                                    else
                                    {
                                        this.Logger.WriteLine("UtilityNetworkAnalysisExt Null");
                                    }
                                }
                            }
                            catch (Exception EX2)
                            {
                                this.Logger.WriteLine("TraceIsolation Failed: " + EX2.Message + ": " + EX2.StackTrace);
                                return false;
                            }
                        }
                        //AddSelection(pMXDoc, dicClasses);
                        //iSelectionCount = pMXDoc.FocusMap.SelectionCount;
                    }
                    SW1.Stop();
                    if (iTraceType == 1)
                    {
                        RecordActionTime("TraceDownStream Execution Time: (" + iSelectionCount.ToString() + ")" , SW1.ElapsedMilliseconds);
                    }
                    else if (iTraceType == 2)
                    {
                        RecordActionTime("TraceUpStream Execution Time: (" + iSelectionCount.ToString() + ")", SW1.ElapsedMilliseconds);
                    }
                    else
                    {
                        RecordActionTime("TraceIsolating Execution Time: (" + iSelectionCount.ToString() + ")", SW1.ElapsedMilliseconds);
                    }
                    pMXDoc = GetDoc();
                    pMXDoc.ActiveView.Refresh();
                    return true;
                }
                catch (Exception EX)
                {
                    this.Logger.WriteLine("TraceDownStream Failed: " + EX.Message + ": " + EX.StackTrace);
                    return false;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("TraceDownStream Failed: " + EX.Message);
                return false;
            }
        }
        //---------------------------

        private void SwizzleDatasets(IApplication pApp, IFeatureWorkspace pOldWorkspace, IFeatureWorkspace pNewWorkspace)
        {
            IVersion pOldVersion;
            IVersion pNewVersion;
            IMxDocument pMxDoc;
            IMaps pMaps;
            IMap pMap;
            ITableCollection pTableCollection;
            int iCurrTable;
            IEnumLayer pEnumLayer;
            ILayer pLayer;
            IFeatureLayer pFeatureLayer;
            IFeatureClass pFeatureClass;
            IFeatureClass pNewFeatureClass;
            IDataset pDataset;
            IMapAdmin pMapAdmin;
            int ICnt;

            try
            {
                pMxDoc = (IMxDocument)pApp.Document;
                pOldVersion = (IVersion)pOldWorkspace;
                pNewVersion = (IVersion)pNewWorkspace;

                ICnt = 0;
                pMaps = pMxDoc.Maps;
                //this.Logger.WriteLine("Maps:" + pMaps.Count);
                pMapAdmin = (IMapAdmin)pMxDoc.FocusMap;
                for (int iCurrMap = 0; iCurrMap < pMaps.Count; iCurrMap++)
                {
                    //this.Logger.WriteLine("Map:" + iCurrMap);
                    pMap = (IMap)pMaps.get_Item(iCurrMap);
                    if (pMap.LayerCount > 0)
                    {
                        pEnumLayer = pMap.get_Layers(null, true);
                        pLayer = pEnumLayer.Next();
                        while (pLayer != null)
                        {
                            if (pLayer is IFeatureLayer)
                            {
                                pFeatureLayer = (IFeatureLayer)pLayer;
                                if (pFeatureLayer.FeatureClass != null)
                                {
                                    pFeatureClass = pFeatureLayer.FeatureClass;
                                    pDataset = (IDataset)pFeatureClass;
                                    if (IsWorkspaceSame((IWorkspace)pDataset.Workspace, (IWorkspace)pOldWorkspace))
                                    //if (pDataset.Workspace.Equals(pOldWorkspace))
                                    {
                                        pNewFeatureClass = pNewWorkspace.OpenFeatureClass(pDataset.Name);
                                        pFeatureLayer.FeatureClass = pNewFeatureClass;
                                        pMapAdmin.FireChangeFeatureClass(pFeatureClass, pNewFeatureClass);
                                        ICnt++;
                                    }
                                }
                            }
                            pLayer = pEnumLayer.Next();
                        }
                    }

                    pTableCollection = (ITableCollection)pMap;
                    if (pTableCollection.TableCount > 0)
                    {
                        for (iCurrTable = 0; iCurrTable < pTableCollection.TableCount; iCurrTable++)
                        {
                            pDataset = (IDataset)pTableCollection.get_Table(iCurrTable);
                            if (IsWorkspaceSame((IWorkspace)pDataset.Workspace, (IWorkspace)pOldWorkspace))
                            //if (pDataset.Workspace.Equals(pOldWorkspace))
                            {
                                pTableCollection.RemoveTable((ITable)pDataset);
                                pTableCollection.AddTable(pNewWorkspace.OpenTable(pDataset.Name));
                                ICnt++;
                            }
                        }
                    }
                }

                pMxDoc.UpdateContents();
                //pMapAdmin.FireChangeFeatureClass (pOldVersion, pNewVersion);
                pMapAdmin.FireChangeVersion(pOldVersion, pNewVersion);
                pNewVersion.RefreshVersion();
                this.Logger.WriteLine("Changed Objects:" + ICnt);
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in SwizzleDatasets:" + EX.Message + " " + EX.StackTrace);
            }

        }

        private bool IsWorkspaceSame(IWorkspace pWKS1, IWorkspace pWKS2)
        {
            IPropertySet pPS1;
            IPropertySet pPS2;

            pPS1 = pWKS1.ConnectionProperties;
            pPS2 = pWKS2.ConnectionProperties;

            if (!IsPropertySame("INSTANCE", pPS1, pPS2))
            {
                return false;
            }

            //if (!IsPropertySame("USER", pPS1, pPS2))
            //{
            //    return false;
            //}
            //if (!IsPropertySame("SERVER", pPS1, pPS2))
            //{
            //    return false;
            //}
            //if (!IsPropertySame("DATABASE", pPS1, pPS2))
            //{
            //    return false;
            //}


            return true;
        }
        private bool IsPropertySame(string Property, IPropertySet pProp1, IPropertySet pProp2)
        {
            object oValue;
            string sValue1;
            string sValue2;

            try
            {
                oValue = pProp1.GetProperty(Property);
                if (oValue == null)
                {
                    return true;
                }
                sValue1 = Convert.ToString(oValue);
                oValue = pProp2.GetProperty(Property);
                if (oValue == null)
                {
                    return true;
                }
                sValue2 = Convert.ToString(oValue);
                sValue2 = sValue2.ToUpper();
                if (Property == "INSTANCE")
                {
                    if (sValue2.StartsWith("SDE:ORACLE$"))
                    {
                        sValue2 = sValue2.Substring(11);
                    }
                }
                if (sValue1.ToUpper() != sValue2.ToUpper())
                {
                    //Logger.WriteLine("Property:" + sValue1 + " - " + sValue2);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("IsPropertySame Error" + EX.Message + " " + EX.StackTrace);
                return true;
            }
        }
        private ISimpleMarkerSymbol CreateMarkerSymbol(int iColor)
        {
            try
            {
                ISimpleMarkerSymbol markerSymbol = new SimpleMarkerSymbol();
                IRgbColor color = new RgbColor();
                switch (iColor)
                {
                    case 0:
                        color.Red = 0;
                        color.Green = 100;
                        color.Blue = 0;
                        break;
                    case 1:
                        color.Red = 100;
                        color.Green = 0;
                        color.Blue = 0;
                        break;
                    default:
                        color.Red = 0;
                        color.Green = 0;
                        color.Blue = 0;
                        break;
                }
                markerSymbol.Color = color;
                markerSymbol.Size = 10; //.Width = 2.5;
                markerSymbol.Style = ESRI.ArcGIS.Display.esriSimpleMarkerStyle.esriSMSSquare;
                return markerSymbol;
            }
            catch (System.Exception e)
            {
                this.Logger.WriteLine("CreateMarkerSymbol Error" + e.Message + " " + e.StackTrace);

                return null;
            }
        }

        private bool EXECUTESQL(string sCommand)
        {
            try
            {
                SW1.Reset();
                SW1.Start();
                Logger.WriteLine(sCommand);
                Workspace.ExecuteSQL(sCommand);

                SW1.Stop();
                RecordActionTime("EXECUTESQL:" , SW1.ElapsedMilliseconds);
                return true;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in EXECUTESQL " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private ISimpleLineSymbol CreateLineSymbol(int iColor)
        {
            try
            {
                ISimpleLineSymbol lineSymbol = new SimpleLineSymbol();
                IRgbColor color = new RgbColor();
                switch (iColor)
                {
                    case 0:
                        color.Red = 0;
                        color.Green = 0;
                        color.Blue = 0;
                        break;
                    case 1:
                        color.Red = 100;
                        color.Green = 0;
                        color.Blue = 0;
                        break;
                    default:
                        color.Red = 0;
                        color.Green = 0;
                        color.Blue = 0;
                        break;
                }
                lineSymbol.Color = color;
                lineSymbol.Width = 2.5;
                lineSymbol.Style = ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSSolid;
                return lineSymbol;
            }
            catch (System.Exception e)
            {
                this.Logger.WriteLine("CreateLineSymbol Error" + e.Message + " " + e.StackTrace);

                return null;
            }
        }

        private bool EditWorkspace (string sLayerName)
        {
            IMxDocument pMXDoc;
            IEnumLayer pLayers;
            ILayer pLayer;
            IFeatureLayer pFLayer;
            IDataset pDS;
            try
            {
                pMXDoc = GetDoc();
                pLayers = pMXDoc.FocusMap.get_Layers(null, true);
                pLayer = pLayers.Next();
                while (pLayer != null)
                {
                    if ((pLayer.Name.ToUpper() == sLayerName.ToUpper()))
                    {
                        pFLayer = (IFeatureLayer)pLayer;
                        pDS = (IDataset) pFLayer.FeatureClass;
                        Workspace = pDS.Workspace;
                        StartEditor(Workspace);
                        return true;
                    }
                    pLayer = pLayers.Next();
                }
                Logger.WriteLine("Layer not found");
            }
            catch
            {
            }
            return true;
        }

        private bool StartEditor(IWorkspace pWKS )
        {
            IApplication pApp;
            UID eUID = new ESRI.ArcGIS.esriSystem.UIDClass();
            IVersion pVer;
          
            try
            {
                SW1.Reset();
                SW1.Start();

                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;

                eUID.Value = "esriEditor.Editor";
                pEditor = pApp.FindExtensionByCLSID(eUID) as IEditor;

                pVer = (IVersion)Workspace;
                
                if (pVer == null)
                {
                    if (pWKS != null)
                    {
                        pVer = (IVersion)pWKS;
                        Workspace = pWKS;
                    }
                    else
                    {
                        Logger.WriteLine("No Valid workspace found");
                        return false;
                    }
                }
                this.Logger.WriteLine("Workspace Name:" + pVer.VersionName);
                if (pEditor != null)
                {
                    if (pEditor.EditState != esriEditState.esriStateEditing)
                    {
                        
                        pEditor.StartEditing(Workspace);
                        SW1.Stop();
                        RecordActionTime("StartEditing: " , SW1.ElapsedMilliseconds);
                    }
                    else
                    {
                        this.Logger.WriteLine("Already Editing");
                    }
                }
                else
                {
                    this.Logger.WriteLine("No Editor");
                }
                return true;
            }
            catch (System.Runtime.InteropServices.COMException ComEX)
            {
                Logger.WriteLine("Error in StartEditor: " + ComEX.Message + " " + ComEX.StackTrace);
                return false;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in StartEditor: " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }

        private bool StopEditing(string sSave)
        {
            try
            {
                SW1.Reset();
                SW1.Start();

                if (pEditor != null)
                {
                    if (pEditor.EditState == esriEditState.esriStateEditing)
                    {
                        if (sSave.ToUpper() == "TRUE")
                        {
                            pEditor.StopEditing(true);
                        }
                        else
                        {
                            pEditor.StopEditing(false);
                        }
                        SW1.Stop();
                        RecordActionTime("StopEditing: " , SW1.ElapsedMilliseconds);
                        return true;
                    }
                    else
                    {
                        Logger.WriteLine("Not Edidting");
                        return false;
                    }
                }
                else
                {
                    Logger.WriteLine("No Editor");
                    return false;
                }
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in StopEditor: " + EX.Message + " " + EX.StackTrace);
                return false;
            }
        }
        private bool CREATELINEFEATURE(string sClassName, string sPoints, string sVars)
        {
            IFeatureClass pFC;
            IFeatureWorkspace pFWKS;
            IFeature pFeat;
            IPolyline pLine;
            IPoint[] pPts;
            IMxDocument pMXDoc;
            IGeometryBridge2 pGeoBrg;
            IPointCollection4 pPointColl;
            string[,] sVar;

            try
            {
                SW1.Reset();
                SW1.Start();
                pMXDoc = GetDoc();
                Logger.WriteLine("Get Points");
                pPts = GetPoints(sPoints, pMXDoc.FocusMap.SpatialReference);
                if (pPts != null)
                {
                    Logger.WriteLine("at least 2 points");
                    if (pPts.GetUpperBound(0) > 0)
                    {
                        Logger.WriteLine("setup vars");
                        sVar = GetVars(sVars);
                        pGeoBrg = new GeometryEnvironment() as IGeometryBridge2;
                        pPointColl = new PolylineClass();
                        pGeoBrg.AddPoints(pPointColl, pPts);
                        if (pEditor != null)
                        {
                            Logger.WriteLine("Open Class");
                            pFWKS = (IFeatureWorkspace)pEditor.EditWorkspace;
                            pFC = pFWKS.OpenFeatureClass(sClassName);
                            if (pFC != null)
                            {
                                pEditor.StartOperation();
                                pFeat = pFC.CreateFeature();
                                pLine = (IPolyline)pPointColl;
                                pFeat.Shape = (IGeometry)pLine;
                                SetFields(pFeat, sVar);
                                pFeat.Store();
                                pEditor.StopOperation("Create Polyline");
                            }
                            else
                            {
                                Logger.WriteLine("Failed to open class:" + sClassName);
                            }
                        }
                        else
                        {
                            Logger.WriteLine("Editor not found");
                        }
                        SW1.Stop();
                        RecordActionTime("CreateLINEFeature: " , SW1.ElapsedMilliseconds);
                        return true;
                    }
                    else
                    {
                        Logger.WriteLine("Empty Geometry");
                        return false;
                    }
                }
                else
                {
                    Logger.WriteLine("Null Geometry");
                    return false;
                }
            }
            catch (COMException ComEX)
            {
                Logger.WriteLine("Error in CreateLINEFeature: " + ComEX.ErrorCode + " :" + ComEX.Message + ":" + ComEX.StackTrace);
                return false;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in CreateLINEFeature: " + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private bool CREATEPOINTFEATURE(string sClassName, string sPoints, string sVars)
        {
            IFeatureClass pFC;
            IFeatureWorkspace pFWKS;
            IFeature pFeat;
            IPoint pPt;
            IMxDocument pMXDoc;
            string[,] sVar;
            double dX;
            double dY;
            string[] sPts;

            try
            {
                SW1.Reset();
                SW1.Start();
                pMXDoc = GetDoc();
                sVar = GetVars(sVars);
                sPts = sPoints.Split(',');
                pPt = new Point();
                pPt.SpatialReference = pMXDoc.FocusMap.SpatialReference;
                dX = Convert.ToDouble(sPts[0]);
                dY = Convert.ToDouble(sPts[1]);
                pPt.PutCoords(dX, dY);
                if (pPt != null)
                {
                    pFWKS = (IFeatureWorkspace)Workspace;
                    pFC = pFWKS.OpenFeatureClass(sClassName);
                    if (pFC != null)
                    {
                        if (pEditor != null)
                        {
                            pEditor.StartOperation();
                            pFeat = pFC.CreateFeature();
                            pFeat.Shape = (IGeometry)pPt;
                            SetFields(pFeat, sVar);
                            pFeat.Store();
                            pEditor.StopOperation("Create Point");
                        }
                        else
                        {
                            Logger.WriteLine("Editor not found");
                        }
                    }
                    SW1.Stop();
                    RecordActionTime("CreatePointFeature: ", SW1.ElapsedMilliseconds);
                    return true;
                }
                else
                {
                    Logger.WriteLine("Null Geometry");
                    return false;
                }
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in CreatePointFeature: " + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private bool CREATEPOLYGONFEATURE(string sClassName, string sPoints, string sVars)
        {
            IFeatureClass pFC;
            IFeatureWorkspace pFWKS;
            IFeature pFeat;
            IPolygon pPolygon;
            IPoint[] pPts;
            IMxDocument pMXDoc;
            IGeometryBridge2 pGeoBrg;
            IPointCollection4 pPointColl;
            string[,] sVar;

            try
            {
                SW1.Reset();
                SW1.Start();
                pMXDoc = GetDoc();
                pPts = GetPoints(sPoints, pMXDoc.FocusMap.SpatialReference);
                if (pPts != null)
                {
                    if (pPts.GetUpperBound(0) > 0)
                    {
                        sVar = GetVars(sVars);
                        pGeoBrg = new GeometryEnvironment() as IGeometryBridge2;
                        pPointColl = new PolygonClass();
                        pGeoBrg.AddPoints(pPointColl, pPts);
                        pFWKS = (IFeatureWorkspace)Workspace;
                        pFC = pFWKS.OpenFeatureClass(sClassName);
                        if (pFC != null)
                        {
                            if (pEditor != null)
                            {
                                pEditor.StartOperation();
                                pFeat = pFC.CreateFeature();
                                pPolygon = (IPolygon)pPointColl;
                                pFeat.Shape = (IGeometry)pPolygon;
                                SetFields(pFeat, sVar);
                                pFeat.Store();
                                Logger.WriteLine("OID:" + pFeat.OID);
                                pEditor.StopOperation("Create Polygon");
                            }
                            else
                            {
                                Logger.WriteLine("Editor not found");
                            }
                        }
                        SW1.Stop();
                        RecordActionTime("CreatePOYGONFeature: " ,SW1.ElapsedMilliseconds);
                        return true;
                    }
                    else
                    {
                        Logger.WriteLine("Missing Geometry");
                        return false;
                    }
                }
                else
                {
                    Logger.WriteLine("Null Geometry");
                    return false;
                }
            }
            catch (COMException ComEX)
            {
                Logger.WriteLine("Error in CreatePOLYGONFeature: " + ComEX.ErrorCode + " :" + ComEX.Message + ":" + ComEX.StackTrace);
                return false;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in CreatePOLYGONFeature: " + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private bool DELETESELECTED()
        {
            IMxDocument pMXDoc;

            try
            {
                SW1.Reset();
                SW1.Start();
                pMXDoc = GetDoc();
                if (pEditor != null)
                {
                    if (pEditor.EditState == esriEditState.esriStateEditing)
                    {
                        CommandClick("{16CD71E5-98C3-11D1-873B-0000F8751720}");
                        SW1.Stop();
                        RecordActionTime("DeleteSelected: ",SW1.ElapsedMilliseconds);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    Logger.WriteLine("No Editor");
                    return false;   
                }
            }
            catch (COMException ComEX)
            {
                Logger.WriteLine("COMError in DeleteSelected: " + ComEX.ErrorCode + " :" + ComEX.Message + ":" + ComEX.StackTrace);
                return false;
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in DeleteSelected: " + EX.Message + ":" + EX.StackTrace);
                return false;
            }
        }

        private List<IESRIShape> GetShapes(string command)
        {
            List<IESRIShape> eSRIShapes = new List<IESRIShape>();
            List<string> argsList = this._parser.GetArgsList(command, false);
            if (argsList.Count == 0)
            {
                return eSRIShapes;
            }
            List<string> strs = new List<string>();
            bool flag = false;
            double num = 0;
            double num1 = 0;
            for (int i = 0; i < argsList.Count; i++)
            {
                flag = (argsList[i] == ";" ? true : i >= argsList.Count - 1);
                if (this._parser.IsNumeric(argsList[i]))
                {
                    strs.Add(argsList[i]);
                }
                if (flag)
                {
                    if (strs.Count == 0)
                    {
                        throw new Exception("No coordinates found.");
                    }
                    if (strs.Count % 2 != 0)
                    {
                        throw new Exception(string.Format("Invalid number [{0}] of coordinates (must be multiple of 2).", strs.Count));
                    }
                    switch (strs.Count)
                    {
                        case 2:
                            {
                                IPoint point = this.ToPoint(strs[0], strs[1]);
                                eSRIShapes.Add((IESRIShape)point);
                                if (i >= argsList.Count - 1)
                                {
                                    return eSRIShapes;
                                }
                                int multiplier = this._parser.GetMultiplier(argsList, ref i);
                                if (multiplier != 0)
                                {
                                    this._parser.ModifyValue("XOFFSET", argsList, ref num, ref i);
                                    this._parser.ModifyValue("YOFFSET", argsList, ref num1, ref i);
                                    for (int j = 1; j < multiplier; j++)
                                    {
                                        double x = point.X + num;
                                        double y = point.Y + num1;
                                        point = this.ToPoint(x, y);
                                        eSRIShapes.Add((IESRIShape)point);
                                    }
                                }
                                break;
                            }
                        case 3:
                            {
                                IPolygon polygon = this.ToPolygon(strs.ToArray());
                                eSRIShapes.Add((IESRIShape)polygon);
                                if (i >= argsList.Count - 1)
                                {
                                    return eSRIShapes;
                                }
                                int multiplier1 = this._parser.GetMultiplier(argsList, ref i);
                                if (multiplier1 == 0)
                                {
                                    this._parser.ModifyValue("XOFFSET", argsList, ref num, ref i);
                                    this._parser.ModifyValue("YOFFSET", argsList, ref num1, ref i);
                                    for (int k = 1; k < multiplier1; k++)
                                    {
                                        strs = new List<string>();
                                        IPointCollection pointCollection = (IPointCollection)polygon;
                                        for (int l = 0; l < pointCollection.PointCount; l++)
                                        {
                                            IPoint point1 = pointCollection.get_Point(l);
                                            string str = (point1.X + num).ToString();
                                            string str1 = (point1.Y + num1).ToString();
                                            strs.Add(str);
                                            strs.Add(str1);
                                        }
                                        polygon = this.ToPolygon(strs.ToArray());
                                        eSRIShapes.Add((IESRIShape)polygon);
                                    }
                                }
                                break;
                            }
                        case 4:
                            {
                                IPolyline polyline = this.ToPolyline(strs[0], strs[1], strs[2], strs[3]);
                                eSRIShapes.Add((IESRIShape)polyline);
                                if (i >= argsList.Count - 1)
                                {
                                    return eSRIShapes;
                                }
                                int multiplier2 = this._parser.GetMultiplier(argsList, ref i);
                                if (multiplier2 == 0)
                                {
                                    this._parser.ModifyValue("XOFFSET", argsList, ref num, ref i);
                                    this._parser.ModifyValue("YOFFSET", argsList, ref num1, ref i);
                                    for (int m = 1; m < multiplier2; m++)
                                    {
                                        double x1 = polyline.FromPoint.X + num;
                                        double y1 = polyline.FromPoint.Y + num1;
                                        double x2 = polyline.ToPoint.X + num;
                                        double y2 = polyline.ToPoint.Y + num1;
                                        polyline = this.ToPolyline(x1, y1, x2, y2);
                                        eSRIShapes.Add((IESRIShape)polyline);
                                    }
                                }
                                break;
                            }
                        default:
                            {
                                goto case 3;
                            }
                    }
                    strs = new List<string>();
                }
            }
            return eSRIShapes;
        }

        private IPoint ToPoint(List<string> args)
        {
            if (args == null || args.Count < 2)
            {
                throw new Exception("At least 2 coordinates are required to create a Point.");
            }
            return ToPoint(args[0], args[1]);
        }

        private IPoint ToPoint(string X, string Y)
        {
            double num;
            if (!double.TryParse(X, out num))
            {
                throw new Exception(string.Format("Point X coordinate [{0}] is not numeric.", X));
            }
            double num1 = num;
            if (!double.TryParse(Y, out num))
            {
                throw new Exception(string.Format("Point Y coordinate [{0}] is not numeric.", Y));
            }
            return this.ToPoint(num1, num);
        }

        private IPoint ToPoint(double x, double y)
        {
            IPoint pointClass = new PointClass();
            pointClass.PutCoords(x, y);
            pointClass.SpatialReference = this.GetSpatialReference(null);
            return pointClass;
        }

        public IPolygon ToPolygon(params string[] args)
        {
            if (args == null || (int)args.Length < 6)
            {
                throw new Exception("At least 6 coordinates are required to create a Polygon.");
            }
            if ((int)args.Length % 2 != 0)
            {
                throw new Exception(string.Format("Invalid number [{0}] of coordinates (must be multiple of 2).", (int)args.Length));
            }
            IPointCollection polygonClass = new PolygonClass();
            IPolygon spatialReference = (IPolygon)polygonClass;
            for (int i = 0; i < (int)args.Length; i = i + 2)
            {
                IPoint point = this.ToPoint(args[i], args[i + 1]);
                object value = Missing.Value;
                object obj = Missing.Value;
                polygonClass.AddPoint(point, ref value, ref obj);
            }
            if (!spatialReference.IsClosed)
            {
                spatialReference.Close();
            }
            spatialReference.SpatialReference = this.GetSpatialReference(null);
            return spatialReference;
        }

        private IPolyline ToPolyline(List<string> args)
        {
            if (args == null || args.Count < 4)
            {
                throw new Exception("At least 4 coordinates are required to create a Polyline.");
            }
            return this.ToPolyline(args[0], args[1], args[2], args[3]);
        }

        private IPolyline ToPolyline(string fromPointX, string fromPointY, string toPointX, string toPointY)
        {
            double num;

            if (!double.TryParse(fromPointX, out num))
            {
                throw new Exception(string.Format("Polyline FromPoint X coordinate [{0}] is not numeric.", fromPointX));
            }
            double num1 = num;
            if (!double.TryParse(fromPointY, out num))
            {
                throw new Exception(string.Format("Polyline FromPoint Y coordinate [{0}] is not numeric.", fromPointY));
            }
            double num2 = num;
            if (!double.TryParse(toPointX, out num))
            {
                throw new Exception(string.Format("Polyline ToPoint X coordinate [{0}] is not numeric.", toPointX));
            }
            double num3 = num;
            if (!double.TryParse(toPointY, out num))
            {
                throw new Exception(string.Format("Polyline ToPoint Y coordinate [{0}] is not numeric.", toPointY));
            }
            return this.ToPolyline(num1, num2, num3, num);
        }

        public IPolyline ToPolyline(double fromPointX, double fromPointY, double toPointX, double toPointY)
        {
            IPoint point = this.ToPoint(fromPointX, fromPointY);
            return this.ToPolyline(point, this.ToPoint(toPointX, toPointY));
        }

        public IPolyline ToPolyline(IPoint fromPoint, IPoint toPoint)
        {
            IPolyline polylineClass = new PolylineClass()
            {
                FromPoint = fromPoint,
                ToPoint = toPoint,
                SpatialReference = this.GetSpatialReference(null)
            };
            return polylineClass;
        }

        private ISpatialReference ToSpatialReference(int factoryCode)
        {
            ISpatialReferenceFactory spatialReferenceEnvironmentClass = new SpatialReferenceEnvironmentClass();
            if (this.IsGeographicType(factoryCode))
            {
                return spatialReferenceEnvironmentClass.CreateGeographicCoordinateSystem(factoryCode);
            }
            if (!this.IsProjectedType(factoryCode))
            {
                return null;
            }
            return spatialReferenceEnvironmentClass.CreateProjectedCoordinateSystem(factoryCode);
        }

        private ISpatialReference ToSpatialReference(string gcsEnumeration)
        {
            esriSRGeoCSType _esriSRGeoCSType;
            esriSRGeoCS2Type _esriSRGeoCS2Type;
            esriSRGeoCS3Type _esriSRGeoCS3Type;
            esriSRProjCSType _esriSRProjCSType;
            esriSRProjCS2Type _esriSRProjCS2Type;
            esriSRProjCS3Type _esriSRProjCS3Type;
            esriSRProjCS4Type _esriSRProjCS4Type;
            int num = 0;
            if (Enum.TryParse<esriSRGeoCSType>(gcsEnumeration, out _esriSRGeoCSType))
            {
                num = (int)_esriSRGeoCSType;
            }
            else if (Enum.TryParse<esriSRGeoCS2Type>(gcsEnumeration, out _esriSRGeoCS2Type))
            {
                num = (int)_esriSRGeoCS2Type;
            }
            else if (Enum.TryParse<esriSRGeoCS3Type>(gcsEnumeration, out _esriSRGeoCS3Type))
            {
                num = (int)_esriSRGeoCS3Type;
            }
            else if (Enum.TryParse<esriSRProjCSType>(gcsEnumeration, out _esriSRProjCSType))
            {
                num = (int)_esriSRProjCSType;
            }
            else if (Enum.TryParse<esriSRProjCS2Type>(gcsEnumeration, out _esriSRProjCS2Type))
            {
                num = (int)_esriSRProjCS2Type;
            }
            else if (!Enum.TryParse<esriSRProjCS3Type>(gcsEnumeration, out _esriSRProjCS3Type))
            {
                if (!Enum.TryParse<esriSRProjCS4Type>(gcsEnumeration, out _esriSRProjCS4Type))
                {
                    throw new Exception(string.Format("Invalid predefined geographic coordinate system enumeration: '{0}'", gcsEnumeration));
                }
                num = (int)_esriSRProjCS4Type;
            }
            else
            {
                num = (int)_esriSRProjCS3Type;
            }
            return this.ToSpatialReference(num);
        }

        private string ToSpatialReferenceEnumeration(int factoryCode)
        {
            if (Enum.IsDefined(typeof(esriSRGeoCSType), factoryCode))
            {
                return Enum.GetName(typeof(esriSRGeoCSType), factoryCode);
            }
            if (Enum.IsDefined(typeof(esriSRGeoCS2Type), factoryCode))
            {
                return Enum.GetName(typeof(esriSRGeoCS2Type), factoryCode);
            }
            if (Enum.IsDefined(typeof(esriSRGeoCS3Type), factoryCode))
            {
                return Enum.GetName(typeof(esriSRGeoCS3Type), factoryCode);
            }
            if (Enum.IsDefined(typeof(esriSRProjCSType), factoryCode))
            {
                return Enum.GetName(typeof(esriSRProjCSType), factoryCode);
            }
            if (Enum.IsDefined(typeof(esriSRProjCS2Type), factoryCode))
            {
                return Enum.GetName(typeof(esriSRProjCS2Type), factoryCode);
            }
            if (Enum.IsDefined(typeof(esriSRProjCS3Type), factoryCode))
            {
                return Enum.GetName(typeof(esriSRProjCS3Type), factoryCode);
            }
            if (!Enum.IsDefined(typeof(esriSRProjCS4Type), factoryCode))
            {
                return string.Format("Invalid predefined geographic coordinate system factory code: '{0}'", factoryCode);
            }
            return Enum.GetName(typeof(esriSRProjCS4Type), factoryCode);
        }
        private ISpatialReference GetSpatialReference(IFeatureClass featureClass = null)
        {
            IMxDocument pMXDoc;

            pMXDoc = GetDoc();
            if (featureClass == null)
            {
                featureClass = this.FeatureClass;
            }
            if (featureClass != null && featureClass is IGeoDataset)
            {
                return ((IGeoDataset)featureClass).SpatialReference;
            }
            if (pMXDoc.ActiveView == null || pMXDoc.ActiveView.FocusMap == null)
            {
                return null;
            }
            return pMXDoc.ActiveView.FocusMap.SpatialReference;
        }
        private bool IsGeographicType(int factoryCode)
        {
            if (Enum.IsDefined(typeof(esriSRGeoCSType), factoryCode) || Enum.IsDefined(typeof(esriSRGeoCS2Type), factoryCode))
            {
                return true;
            }
            return Enum.IsDefined(typeof(esriSRGeoCS3Type), factoryCode);
        }

        private bool IsProjectedType(int factoryCode)
        {
            if (Enum.IsDefined(typeof(esriSRProjCSType), factoryCode) || Enum.IsDefined(typeof(esriSRProjCS2Type), factoryCode) || Enum.IsDefined(typeof(esriSRProjCS2Type), factoryCode))
            {
                return true;
            }
            return Enum.IsDefined(typeof(esriSRProjCS3Type), factoryCode);
        }
        private string GetCoordinates(IESRIShape shape)
        {
            string str = "[]";
            if (shape is IPoint)
            {
                IPoint point = (IPoint)shape;
                str = string.Format("[{0}, {1}]", point.X, point.Y);
            }
            else if (shape is IPolyline)
            {
                IPolyline polyline = (IPolyline)shape;
                object[] x = new object[] { polyline.FromPoint.X, polyline.FromPoint.Y, polyline.ToPoint.X, polyline.ToPoint.Y };
                str = string.Format("[{0},{1} / {2},{3}]", x);
            }
            else if (shape is IPolygon)
            {
                IPointCollection pointCollection = (IPointCollection)((IPolygon)shape);
                str = "[";
                for (int i = 0; i < pointCollection.PointCount; i++)
                {
                    IPoint point1 = pointCollection.get_Point(i);
                    str = string.Concat(str, string.Format("{0},{1}", point1.X, point1.Y));
                    if (i < pointCollection.PointCount - 1)
                    {
                        str = string.Concat(str, ", ");
                    }
                }
                str = string.Concat(str, "]");
            }
            return str;
        }

        private bool ShutdownArcMap()
        {
            IApplication pApp;
            IDocument pDoc;
            IMxDocument pMXDoc;
            IDocumentDirty2 pDocDirty;

            Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
            System.Object obj = Activator.CreateInstance(t);
            pApp = obj as IApplication;
            pDoc = pApp.Document;
            pMXDoc = (IMxDocument)pDoc;
            pDocDirty = (IDocumentDirty2)pMXDoc;
            pDocDirty.SetClean();
            Workspace = null;
            ScriptEngine.BroadcastProperty("Workspace", this.Workspace, this);
            ScriptEngine.BroadcastProperty("FeatureClass", null, this);
            _SDMGR = null;
            _PTMGR = null;
            pSelection = null;
            m_pGN = null;
            pSC = null;
            pEditor = null;
            pDoc = null;
            pMXDoc = null;
            pApp.Shutdown();
            //Process iProc = Process.GetCurrentProcess();
            //iProc.Kill();

            Logger.WriteLine("Closing Arcmap");
            return true;
        }

        private bool Pan(string sDirection, string sDistance)
        {
            SW1.Reset();
            SW1.Start();
            IMxDocument pDoc;
            IEnvelope pEnv;
            IMap pMap;
            double dDistance;

            dDistance = Convert.ToDouble(sDistance);
            pDoc = GetDoc();
            pMap = pDoc.FocusMap;
            pEnv = pDoc.ActiveView.Extent;
            switch (sDirection.ToUpper())
            {
                case "UP":
                    pEnv.Offset(0, dDistance);
                    Logger.WriteLine(string.Format("Pan Up {0}", dDistance));
                    break;
                case "DOWN":
                    pEnv.Offset(0, dDistance * -1);
                    Logger.WriteLine(string.Format("Pan Down {0}", dDistance));
                    break;
                case "LEFT":
                    pEnv.Offset(dDistance * -1, 0);
                    Logger.WriteLine(string.Format("Pan Left {0}", dDistance));
                    break;
                case "RIGHT":
                    pEnv.Offset(dDistance, 0);
                    Logger.WriteLine(string.Format("Pan Right {0}",dDistance));
                    break;
            };
            
            pDoc.ActiveView.Extent = pEnv;
            pDoc.ActiveView.Refresh();
            Logger.WriteLine(string.Format("ActiveView extent set to: [XMin={0}, YMin={1}, XMax={2}, YMax={3}]", pEnv.XMin,pEnv.YMin,pEnv.XMax,pEnv.YMax));
            SW1.Stop();
            RecordActionTime("PAN " + sDirection, SW1.ElapsedMilliseconds);

            return true;
        }

        private bool WakeUpAt(string sWakeupAt)
        {
            int iSleepTime;
            DateTime dtNow;
            DateTime dtWake;

            try
            {
                dtNow = DateTime.Now;
                dtWake = Convert.ToDateTime(sWakeupAt);
                iSleepTime = (((dtWake.Hour - dtNow.Hour) * 60) * 60) * 1000;
                iSleepTime = iSleepTime + (((dtWake.Minute - dtNow.Minute) * 60) * 1000);
                iSleepTime = iSleepTime + ((dtWake.Second - dtNow.Second) * 1000);
                Logger.WriteLine("Wake At:" + dtWake);
                Logger.WriteLine("Milliseconds:" + iSleepTime);
                if (iSleepTime > (60 * 60 * 1000))
                {
                    Logger.WriteLine("Sleep time exceed 1 hour, assuming it is an error");
                }
                else
                {
                    System.Threading.Thread.Sleep(iSleepTime);
                }
                Logger.WriteLine("Woke Up:" + string.Format("{0:MM/dd/yy H:mm:ss}", DateTime.Now));
            }
            catch
            {
            }
            return true;
        }

        private bool ExportToPDF(string sFileName)
        {
            IMxDocument pMXDoc;
            IGraphicsContainer docGraphicsContainer;
            IElement docElement;
            IOutputRasterSettings docOutputRasterSettings;
            IMapFrame docMapFrame;
            IActiveView tmpActiveView;
            IActiveView pAV;
            IExport docExport;
            IPrintAndExport docPrintExport;
            System.Guid UID;

            try
            {
                pMXDoc = GetDoc();
                pAV = pMXDoc.ActiveView;

                /* This function sets the OutputImageQuality for the active view. If the active view is a pagelayout, then
                    * it must also set the output image quality for each of the maps in the pagelayout.
                    */
                if (pAV is IMap)
                {
                    docOutputRasterSettings = pAV.ScreenDisplay.DisplayTransformation
                        as IOutputRasterSettings;
                    docOutputRasterSettings.ResampleRatio = 1;
                    //pMXDoc.ActiveView = (IActiveView)pMXDoc.PageLayout;
                }
                else if (pAV is IPageLayout)
                {
                    //Assign ResampleRatio for PageLayout
                    docOutputRasterSettings = pAV.ScreenDisplay.DisplayTransformation
                        as IOutputRasterSettings;
                    docOutputRasterSettings.ResampleRatio = 1;
                    //and assign ResampleRatio to the maps in the PageLayout.
                    docGraphicsContainer = pAV as IGraphicsContainer;
                    docGraphicsContainer.Reset();
                    docElement = docGraphicsContainer.Next();
                    while (docElement != null)
                    {
                        if (docElement is IMapFrame)
                        {
                            docMapFrame = docElement as IMapFrame;
                            tmpActiveView = docMapFrame.Map as IActiveView;
                            docOutputRasterSettings =
                                tmpActiveView.ScreenDisplay.DisplayTransformation as
                                IOutputRasterSettings;
                            docOutputRasterSettings.ResampleRatio = 1;
                        }
                        docElement = docGraphicsContainer.Next();
                    }
                    docMapFrame = null;
                    docGraphicsContainer = null;
                    tmpActiveView = null;
                }
                //docOutputRasterSettings = null;
                //docOutputRasterSettings =(IOutputRasterSettings) pAV.ScreenDisplay.DisplayTransformation;
                //docOutputRasterSettings.ResampleRatio = 1;
                docExport = new ExportPDFClass();
                docPrintExport = new PrintAndExport();

                //Process myProc = new Process();
                //set the export filename (which is the nameroot + the appropriate file extension)
                UID = System.Guid.NewGuid();
                docExport.ExportFileName = sFileName + "_" + UID.ToString() + ".PDF";
                Logger.WriteLine("Export:" + docExport.ExportFileName);

                //Output Image Quality of the export.  The value here will only be used if the export
                // object is a format that allows setting of Output Image Quality, i.e. a vector exporter.
                // The value assigned to ResampleRatio should be in the range 1 to 5.
                // 1 corresponds to "Best", 5 corresponds to "Fast"

                // check if export is vector or raster
                if (docExport is IOutputRasterSettings)
                {
                    // for vector formats, assign the desired ResampleRatio to control drawing of raster layers at export time   
                    //RasterSettings = (IOutputRasterSettings)docExport;
                    //RasterSettings.ResampleRatio = 1;

                    // NOTE: for raster formats output quality of the DISPLAY is set to 1 for image export 
                    // formats by default which is what should be used
                }

                try
                {
                    docPrintExport.Export(pAV, docExport, 600, true, null);
                }
                catch (Exception EX)
                {
                    Logger.WriteLine("Error in 2ExportPDF " + EX.Message + "Stack" + EX.StackTrace);
                }
                //MessageBox.Show("Finished exporting " + sOutputDir + sNameRoot + "." + docExport.Filter.Split('.')[1].Split('|')[0].Split(')')[0] + ".", "Export Active View Sample");
                pMXDoc.ActiveView = (IActiveView)pMXDoc.FocusMap;
                Logger.WriteLine("ExportToPDF:");
            }
            catch (COMException ComEX)
            {
                Logger.WriteLine("Error in ExportPDF " + ComEX.ErrorCode + " " + ComEX.Message + "Stack" + ComEX.StackTrace);
            }
            catch (Exception EX)
            {
                Logger.WriteLine("Error in ExportPDF " + EX.Message + "Stack" + EX.StackTrace);
            }
            return true;
        }

        private string[] GetParameterGroups(string sInArgs)
        {
            string[] lstArgs;

            if (sInArgs.Length > 0)
            {
                lstArgs = sInArgs.Split(';');
                for (int i = 0; i < lstArgs.Count(); i++)
                {
                    this.Logger.WriteLine("Arg:[" + i + "]" + lstArgs[i]);
                }
            }
            else
            {
                lstArgs = new string[2] { "", "" };
                this.Logger.WriteLine("No Parameters");
            }
            return lstArgs;
        }
        private string[,] GetVars(string sParms)
        {
            int iRows;
            string[] sTemp;
            int iPos;
            string[,] sResults;

            sTemp = sParms.Split(',');
            iRows = sTemp.Length / 2;
            iPos = 0;
            sResults = new string[iRows, 2];
            for (int i = 0; i < iRows; i++)
            {
                sResults[i, 0] = sTemp[iPos];
                iPos++;
                sResults[i, 1] = sTemp[iPos];
                iPos++;
            }
            return sResults;
        }

        private void SetFields(IFeature pFeat, string[,] sVar)
        {
            IField pField;
            int iValue;
            double dValue;
            string sFieldName;

            for (int i = 0; i <= sVar.GetUpperBound(0); i++)
            {
                Logger.WriteLine("Field:" + sVar[i, 0]);
                sFieldName = sVar[i, 0].ToUpper();
                sFieldName = sFieldName.Trim();
                int iFld = pFeat.Fields.FindField(sFieldName);
                if (iFld > -1)
                {
                    pField = pFeat.Fields.Field[iFld];
                    switch (pField.Type)
                    {
                        case esriFieldType.esriFieldTypeInteger:
                            iValue = Convert.ToInt32(sVar[i, 1]);
                            pFeat.set_Value(iFld, iValue);
                            break;
                        case esriFieldType.esriFieldTypeDouble:
                            dValue = Convert.ToDouble(sVar[i, 1]);
                            pFeat.set_Value(iFld, dValue);
                            break;
                        case esriFieldType.esriFieldTypeString:
                            pFeat.set_Value(iFld, sVar[i, 1]);
                            break;
                    }
                }
                else
                {
                    Logger.WriteLine("Field not found:" + sVar[i, 0]);
                }
            }

        }

        private IPoint[] GetPoints(string sInput, ISpatialReference pSR)
        {
            int iStrCnt;
            int iPairCnt;
            string[] sNumbers;
            IPoint[] pPts;
            IPoint pPt;
            int i;
            int iPos;
            double dX;
            double dY;

            sNumbers = sInput.Split(',');
            iStrCnt = sNumbers.GetUpperBound(0) + 1;
            iPairCnt = iStrCnt / 2;
            Logger.WriteLine("Pairs:" + iPairCnt + "," + iStrCnt);
            if ((iPairCnt * 2) == iStrCnt)
            {
                pPts = new ESRI.ArcGIS.Geometry.IPoint[iPairCnt];
                iPos = 0;
                for (i = 0; i < iPairCnt; i++)
                {
                    pPt = new Point();
                    pPt.SpatialReference = pSR;
                    dX = Convert.ToDouble(sNumbers[iPos]);
                    dY = Convert.ToDouble(sNumbers[iPos + 1]);
                    Logger.WriteLine("(" + i + "DX,DY:" + dX + "," + dY);
                    pPt.PutCoords(dX, dY);
                    pPts[i] = pPt;
                    iPos = iPos + 2;
                }
                return pPts;
            }
            else
            {
                return null;
            }
        }

        private void SelectFeatures(IMMTracedElements pEdges, IMMTracedElements pJunctions, INetwork pNetwork, bool overrideSelectSetting)
        {
            IApplication pApp;
            IEnumNetEID junctionList;
            IEnumNetEID edgeList;

            try
            {
                Type t = Type.GetTypeFromProgID("esriFramework.AppRef");
                System.Object obj = Activator.CreateInstance(t);
                pApp = obj as IApplication;

                //Type type = Type.GetTypeFromProgID("esriSystem.ExtensionManager");
                //IExtensionManager extMgr = Activator.CreateInstance(type) as IExtensionManager;
                //if (extMgr != null)
                //{
                    UID uid = new UIDClass();
                    uid.Value = "esriNetworkAnalystUI.NetworkAnalystExtension"; //ESRI Network Analyst extension
                    //IExtension ext = extMgr.FindExtension(uid);
                    IExtension ext = pApp.FindExtensionByCLSID(uid);
                    if (ext != null)
                    {
                        IUtilityNetworkAnalysisExt utilityNetworkAnalystExt = ext as IUtilityNetworkAnalysisExt;
                        if (utilityNetworkAnalystExt != null)
                        {
                            INetworkAnalysisExtResults networkAnalysisResults = (INetworkAnalysisExtResults) utilityNetworkAnalystExt;
                            if (networkAnalysisResults != null)
                            {
                                networkAnalysisResults.ClearResults();
                                junctionList = ConvertTracedElementsCollectionToEnumNetEid(pJunctions, pNetwork);
                                edgeList = ConvertTracedElementsCollectionToEnumNetEid(pEdges, pNetwork);
                                networkAnalysisResults.SetResults(junctionList, edgeList);
                                //determine if results should be drawn
                                if (networkAnalysisResults.ResultsAsSelection)
                                {
                                    networkAnalysisResults.CreateSelection(junctionList, edgeList);
                                }
                                else if (overrideSelectSetting)
                                {
                                    // temporarily toggle the user's setting
                                    networkAnalysisResults.ResultsAsSelection = true;
                                    networkAnalysisResults.CreateSelection(junctionList, edgeList);
                                    networkAnalysisResults.ResultsAsSelection = false;
                                }
                            }
                            else
                            {
                                this.Logger.WriteLine("NetworkAnalysisResults Null");
                            }
                        }
                        else
                        {
                            this.Logger.WriteLine("UtilityNetworkAnalysisExt Null");
                        }
                    }
                    else
                    {
                        this.Logger.WriteLine("Network Extension Null");
                    }
                //}
                //else
                //{
                //    this.Logger.WriteLine("Extension Manager Null");
                //}
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine("Error in SelectFeatures:" + EX.Message + " " + EX.StackTrace);
            }
        }

        public IEnumNetEID ConvertTracedElementsCollectionToEnumNetEid(IMMTracedElements tracedElements, INetwork network)
        {
            IEnumNetEID resultEnum = null;

            try
            {
                // we have a feedpath, now we want to extract all the elements of one type and 
                // put them in an esri enumneteid so we can display it on the map or other fun stuff

                IEnumNetEIDBuilder enumNetEIDBuilder = new EnumNetEIDArrayClass();
                enumNetEIDBuilder.Network = network;
                enumNetEIDBuilder.ElementType = (esriElementType)tracedElements.ElementType;

                INetElements netElements = network as INetElements;

                if ((netElements != null) && (esriElementType.esriETEdge == (esriElementType)tracedElements.ElementType))
                {
                    tracedElements.Reset();
                    IMMTracedElement tracedElement = tracedElements.Next();
                    while (tracedElement != null)
                    {
                        IEnumNetEID tempEnumNetEID = netElements.GetEIDs(tracedElement.FCID, tracedElement.OID, esriElementType.esriETEdge);
                        tempEnumNetEID.Reset();
                        for (int eid = tempEnumNetEID.Next(); eid > 0; eid = tempEnumNetEID.Next())
                        {
                            enumNetEIDBuilder.Add(eid);
                        }
                        tracedElement = tracedElements.Next();
                    }
                }
                else
                {
                    tracedElements.Reset();
                    IMMTracedElement tracedElement = tracedElements.Next();
                    while (tracedElement != null)
                    {
                        enumNetEIDBuilder.Add(tracedElement.EID);
                        tracedElement = tracedElements.Next();
                    }
                }
                if (enumNetEIDBuilder == null)
                {
                    this.Logger.WriteLine("Builder is null");
                }
                resultEnum = enumNetEIDBuilder as IEnumNetEID;
            }
            catch (Exception EX)
            {
                this.Logger.WriteLine(" ConvertTracedElements:" + EX.Message + " " + EX.StackTrace);
            }
            return resultEnum;
        }


        #endregion
        public void RecordActionTime(string ActionText, long ActionMilliseconds)
        {
            if (pSC.ReportExecutionTimeInMilliseconds == true)
            {
                Logger.WriteLine(ActionText + ":" +  ActionMilliseconds.ToString());
            }
            else
            {
                Logger.WriteLine(ActionText + ":" + string.Format("{0:0.00000}",System.Convert.ToDouble( ActionMilliseconds)/1000));
            }
        }
    }
}
