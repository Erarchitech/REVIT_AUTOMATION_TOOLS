# coding: utf8
import os
import time
import StringIO
import System
import datetime
import json
from Autodesk.Revit.DB import *
from shutil import copyfile
from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon, TaskDialogCommonButtons, TaskDialogCommandLinkId, TaskDialogResult
from pyrevit import forms
from pyrevit import script
import datetime
import pyrevit
import codecs
import subprocess
from System.Collections.Generic import List


def on_failure_processing(sender, event):
    # print("Failure - " + str(event.GetProcessingResult()))
    # pass
    accessor = event.GetFailuresAccessor()
    FailureSev = accessor.GetSeverity()
    if FailureSev == FailureSeverity.None:
        event.SetProcessingResult(FailureProcessingResult.Continue)
    elif FailureSev == FailureSeverity.Warning:
        accessor.CommitPendingTransaction()
    elif FailureSev == FailureSeverity.Error:
        accessor.RollBackPendingTransaction()
    elif FailureSev == FailureSeverity.DocumentCorruption:
        accessor.RollBackPendingTransaction()


def on_dialog_open(sender, event):
    did = event.DialogId
    logger = script.get_logger()
    if did != "":
        logger.info("Dialog opened ID:" + did)
        if did == "TaskDialog_Model_Opened_By_Another_User":
            logger.info("Continue")
            event.OverrideResult(1)
        elif did == "TaskDialog_Missing_Third_Party_Updaters":
            logger.info("CommandLink1")
            event.OverrideResult(int(TaskDialogResult.CommandLink1))
        else:
            pass


def getNavisworksView():
    views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)
    for view in views:
        print(view.Title)
        if view.Name.upper() == "Navisworks".upper() and not(view.IsTemplate):
            return view
    return None

def checkDetailView(NWview):
    if NWview.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger() < 3:
        NWviewTemplateId = NWview.ViewTemplateId
        if NWviewTemplateId.IntegerValue == -1:
            View = NWview
        else:
            ViewTemplate = doc.GetElement(NWviewTemplateId)
            ViewTemplateNotControlledParameters = ViewTemplate.GetNonControlledTemplateParameterIds()
            param = ViewTemplate.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL)
            Template = True
            for p in ViewTemplateNotControlledParameters:
                if doc.GetElement(p) == param:
                    Template = False
            if Template:
                View = ViewTemplate
            else:
                View = NWview
        transAct = Transaction(doc, 'Change Detail Level')
        try:
            transAct.Start()
            View.DetailLevel = ViewDetailLevel.Fine
        finally:
            doc.Regenerate()
            transAct.Commit()
        # print(NWview.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger())


def ExportNWCtoTarget(Model, target):
    if Model is None:
        result = "Error - Model is null"
    else:
        doc = OpenModelWithDetach(Model, ["00_", "Link", "#_", "Temp", "Grid", "Level"])
        if doc is None:
            result = "Error - Model opening error"
        else:
            fileName = Model.replace("\\", "/").split("/")[-1].upper().replace(".RVT", "")
            view = getNavisworksView()
            checkDetailView(view)
            if view is None or view == None:
                # result = False
                WriteError("No view Named NavisWorks. Export to NWC was not done", targetFolder, "Convertion")
            elif NWview.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger() < 3:
                result = "View NavisWorks Failed to set Detail level to \"Fine\""
            else:
                #Navisworks Export Options
                navisworksExportOptions = NavisworksExportOptions()
                navisworksExportOptions.ConvertElementProperties = True
                navisworksExportOptions.Coordinates = NavisworksCoordinates.Shared
                navisworksExportOptions.DivideFileIntoLevels = True
                navisworksExportOptions.ExportElementIds = True
                navisworksExportOptions.ExportLinks = False
                navisworksExportOptions.ExportParts = False
                navisworksExportOptions.ExportRoomAsAttribute = False
                navisworksExportOptions.ExportRoomGeometry = False
                navisworksExportOptions.ExportScope = NavisworksExportScope.View
                navisworksExportOptions.ExportUrls = True
                navisworksExportOptions.FindMissingMaterials = False
                navisworksExportOptions.Parameters = NavisworksParameters.All
                navisworksExportOptions.ViewId = view.Id
                # print Application.VersionNumber 

                if Application.VersionNumber >= 2021:
                    navisworksExportOptions.FacetingFactor = 4
                    navisworksExportOptions.ConvertLights = False
                    navisworksExportOptions.ConvertLinkedCADFormats = False


                try:
                    doc.Export(target, fileName,navisworksExportOptions)
                except:
                    result = "Error occured on Export"
    return result

def ExportIFCtoTarget(Model, target):
    if Model is not None:
        # doc = OpenModelWithDetach(Model, ["Link", "#_", "Temp", "Grid", "Level"])
        doc = OpenModelWithDetach(Model, ["00_", "Link", "#_", "Temp", "Grid", "Level"])
        if doc is not None:
            ModelName = Model.replace("\\", "/").split("/")[-1].upper().replace(".RVT", "")
            NWview = getNavisworksView()
            # views = FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements()
            # NWview = None
            # for v in views:
            #     if v.Name.upper() == "NavisWorks".upper() and not(v.IsTemplate):
            #         NWview = v

            if NWview is not None:
                checkDetailView(NWview)
                # if NWview.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger() < 3:
                #     NWviewTemplateId = NWview.ViewTemplateId
                #     if NWviewTemplateId.IntegerValue == -1:
                #         View = NWview
                #     else:
                #         ViewTemplate = doc.GetElement(NWviewTemplateId)
                #         ViewTemplateNotControlledParameters = ViewTemplate.GetNonControlledTemplateParameterIds()
                #         param = ViewTemplate.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL)
                #         Template = True
                #         for p in ViewTemplateNotControlledParameters:
                #             if doc.GetElement(p) == param:
                #                 Template = False
                #         if Template:
                #             View = ViewTemplate
                #         else:
                #             View = NWview
                #     transAct = Transaction(doc, 'Change Detail Level')
                #     try:
                #         transAct.Start()
                #         View.DetailLevel = ViewDetailLevel.Fine
                #     finally:
                #         doc.Regenerate()
                #         transAct.Commit()
                #     # print(NWview.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger())
                if NWview.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger() < 3:
                    result = "View NavisWorks Failed to set Detail level to \"Fine\""
                else:
                    t = Transaction(doc, 'ExportIFC')
                    fho = t.GetFailureHandlingOptions()
                    fho.SetDelayedMiniWarnings(True)
                    t.SetFailureHandlingOptions(fho)
                    # Create an instance of IFCExportOptions
                    IFCOptions = IFCExportOptions()
                    IFCOptions.WallAndColumnSplitting = True
                    IFCOptions.ExportBaseQuantities = True
                    IFCOptions.FilterViewId = NWview.Id
                    IFCOptions.FileVersion = IFCVersion.IFC2x3CV2
                    IFCOptions.SpaceBoundaryLevel = 0
                    IFCOptions.WallAndColumnSplitting = True
                    IFCOptions.AddOption("ExportInternalRevitPropertySets", "true")
                    IFCOptions.AddOption("ExportIFCCommonPropertySets", "true")
                    IFCOptions.AddOption("ExportAnnotations", "false")
                    IFCOptions.AddOption("Use2DRoomBoundaryForVolume", "false")
                    IFCOptions.AddOption("UseFamilyAndTypeNameForReference", "false")
                    IFCOptions.AddOption("ExportVisibleElementsInView", "true")
                    IFCOptions.AddOption("ExportPartsAsBuildingElements", "false")
                    IFCOptions.AddOption("UseActiveViewGeometry", "true")
                    IFCOptions.AddOption("ExportSpecificSchedules", "false")
                    IFCOptions.AddOption("ExportBoundingBox", "false")
                    IFCOptions.AddOption("ExportSolidModelRep", "false")
                    IFCOptions.AddOption("ExportSchedulesAsPsets", "false")
                    IFCOptions.AddOption("ExportUserDefinedPsets", "false")
                    IFCOptions.AddOption("ExportUserDefinedParameterMapping", "false")
                    IFCOptions.AddOption("ExportLinkedFiles", "true")
                    IFCOptions.AddOption("IncludeSiteElevation", "true")
                    IFCOptions.AddOption("SitePlacement", "0")
                    IFCOptions.AddOption("TessellationLevelOfDetail", "0")
                    IFCOptions.AddOption("UseOnlyTriangulation", "false")
                    IFCOptions.AddOption("ActiveViewId", str(NWview.Id.IntegerValue))
                    IFCOptions.AddOption("StoreIFCGUID", "false")
                    IFCOptions.AddOption("FileType", "Ifc")
                    IFCOptions.AddOption("ExportUserDefinedPsetsFileName", "")
                    IFCOptions.AddOption("ExportUserDefinedParameterMappingFileName", "")
                    IFCOptions.AddOption("ExportRoomsInView", "false")
                    IFCOptions.AddOption("ExcludeFilter", "")
                    IFCOptions.AddOption("COBieCompanyInfo", "")
                    IFCOptions.AddOption("COBieProjectInfo", "")
                    IFCOptions.AddOption("IncludeSteelElements", "true")
                    IFCOptions.AddOption("UseTypeNameOnlyForIfcType", "false")
                    IFCOptions.AddOption("UseVisibleRevitNameAsEntityName", "false")

                    try:
                        t.Start()
                        ExportRes = doc.Export(target, ModelName, IFCOptions)
                        if ExportRes:
                            result = "Export done"
                        else:
                            result = "Error occured on Export"
                    except:
                        t.RollBack()
                        result = "Error occured on Export"
                    finally:
                        if t.HasEnded():
                            pass
                        else:
                            t.RollBack()

            else:
                result = "Error - No view Named NavisWorks"
            doc.Close(False)
        else:
            result = "Error - Model opening error"
    else:
        result = "Error - Model is null"
    return result


def OpenServerModelWithDetach(path, worksetmasktoclose=None):
    try:
        app = __revit__.Application
        oop = OpenOptions()
        oop.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        oop.Audit = True
        # oop.OpenForeignOption = OpenForeignOption.Open
        wsconf = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        pathlist = path.split("/")
        Server = pathlist[2]
        Path = path.replace("RSN://"+Server+"/","")
        mp = ServerPath(Server, Path)
        if worksetmasktoclose != None:
            if not isinstance(worksetmasktoclose,list):
                worksetmasktoclose = [worksetmasktoclose]
            try:
                worksets = WorksharingUtils.GetUserWorksetInfo(mp)
                WSToOpen = List[WorksetId]()
                for ws in worksets:
                    isclose = []
                    for m in worksetmasktoclose:
                        isclose.append(m in ws.Name)
                    if not (True in isclose):
                        WSToOpen.Add(ws.Id)
                # for i in WSToOpen:
                #     print(i)
                if len(WSToOpen) > 0:
                    wsconf.Open(WSToOpen)
            except:
                pass
        oop.SetOpenWorksetsConfiguration(wsconf)
        doc = app.OpenDocumentFile(mp, oop, )
        return doc
    except:
        return None

def OpenFileModelWithDetach(path,worksetmasktoclose = None):
    # print(worksetmasktoclose)
    try:
        app = __revit__.Application
        oop = OpenOptions()
        oop.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        # oop.OpenForeignOption = OpenForeignOption.Open
        wsconf = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        mp = FilePath (path)
        if worksetmasktoclose != None:
            if not isinstance(worksetmasktoclose,list):
                worksetmasktoclose = [worksetmasktoclose]
            try:
                worksets = WorksharingUtils.GetUserWorksetInfo(mp)
                WSToOpen = List[WorksetId]()
                for ws in worksets:
                    isclose = []
                    for m in worksetmasktoclose:
                        isclose.append(m in ws.Name)
                    if not (True in isclose):
                        WSToOpen.Add(ws.Id)
                # for i in WSToOpen:
                #     print(i)АвтоКАд015
                
                if len(WSToOpen) > 0:
                    wsconf.Open(WSToOpen)
            except:
                pass
        oop.SetOpenWorksetsConfiguration(wsconf)
        doc = app.OpenDocumentFile(mp, oop)
        return doc
    except:
        return None

def OpenModelWithDetach(path,worksetmasktoclose = None):
    __revit__.DialogBoxShowing += on_dialog_open 
    __revit__.Application.FailuresProcessing += on_failure_processing
    result = None
    try:
        if path !=None:
            if path.startswith("RSN"):
                result = OpenServerModelWithDetach(path,worksetmasktoclose)
            else:
                result = OpenFileModelWithDetach(path,worksetmasktoclose)
    finally:
        __revit__.DialogBoxShowing -= on_dialog_open 
        __revit__.Application.FailuresProcessing -= on_failure_processing
    return result

def OpenARModelWithDetach(path):
    __revit__.DialogBoxShowing += on_dialog_open 
    __revit__.Application.FailuresProcessing += on_failure_processing
    result = None
    try:
        if path !=None:
            modelpath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
            app = __revit__.Application
            oop = OpenOptions()
            oop.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            wsconf = WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets)
            oop.SetOpenWorksetsConfiguration(wsconf)
            doc = app.OpenDocumentFile(modelpath, oop)
            Links = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsElementType().ToElements()
            for l in Links:
                if l.IsNestedLink :
                    continue
                if RevitLinkType.IsLoaded(doc,l.Id):
                    pass
                else:
                    LinkName = l.get_Parameter(BuiltInParameter.RVT_LINK_FILE_NAME_WITHOUT_EXT).AsString().upper()
                    if ".IFC" in LinkName or "_M_" in LinkName or "_E_" in LinkName:
                        pass
                    else:
                        try:
                            l.Load()
                        except:
                            pass
                result = doc
    finally:
        __revit__.DialogBoxShowing -= on_dialog_open 
        __revit__.Application.FailuresProcessing -= on_failure_processing
    return result


def SaveAsServerModelAsCentral(doc, path):
    try:
        saop = SaveAsOptions()
        saop.OverwriteExistingFile = True
        saop.Compact = True
        saop.MaximumBackups = 1
        wssop = WorksharingSaveAsOptions()
        wssop.SaveAsCentral = True
        saop.SetWorksharingOptions(wssop)
        pathlist = path.split("/")
        Server = pathlist[2]
        Path = path.replace("RSN://" + Server + "/", "")
        mp = ServerPath(Server, Path)
        doc.SaveAs(mp, saop)
        # Relinquish all
        twcOpts = TransactWithCentralOptions()
        rOptions = RelinquishOptions(True)
        rOptions.UserWorksets = True
        WorksharingUtils.RelinquishOwnership(doc, rOptions, twcOpts)
        # syncopt = SynchronizeWithCentralOptions()
        # syncopt.SetRelinquishOptions(rOptions)
        # syncopt.SaveLocalBefore = False
        # syncopt.SaveLocalAfter = False
        # doc.SynchronizeWithCentral(twcOpts, syncopt)        
        return True
    except:
        return False


def SaveAsFileModelAsCentral(doc, path):
    try:
        folder = os.path.dirname(os.path.abspath(path))
        if not os.path.exists(folder):
            os.makedirs(folder)
        saop = SaveAsOptions()
        saop.OverwriteExistingFile = True
        saop.Compact = True
        saop.MaximumBackups = 1
        wssop = WorksharingSaveAsOptions()
        wssop.SaveAsCentral = True
        saop.SetWorksharingOptions(wssop)
        doc.SaveAs(path, saop)
        # Relinquish all
        twcOpts = TransactWithCentralOptions()
        rOptions = RelinquishOptions(True)
        rOptions.UserWorksets = True
        WorksharingUtils.RelinquishOwnership(doc, rOptions, twcOpts)
        # syncopt = SynchronizeWithCentralOptions()
        # syncopt.SetRelinquishOptions(rOptions)
        # syncopt.SaveLocalBefore = False
        # syncopt.SaveLocalAfter = False
        # doc.SynchronizeWithCentral(twcOpts, syncopt)        
        return True
    except:
        return False


def SaveAsModelWithDetach(doc, path):
    if path.startswith("RSN"):
        return SaveAsServerModelAsCentral(doc, path)
    else:
        return SaveAsFileModelAsCentral(doc, path)


def EasyCopyModelCMD(sourcepath, targetpath, overwrite=True):
    version = __revit__.Application.VersionNumber
    cmdline = "C:\\Program Files\\Autodesk\\Revit " + str(version) + "\\RevitServerToolCommand\\RevitServerTool.exe"
    cmdline += " createLocalRvt"
    sourcepathlist = sourcepath.split("/")
    Server = sourcepathlist[2]
    Path = sourcepath.replace("RSN://" + Server + "/", "")
    cmdline += " \"" + Path + "\""
    cmdline += " -s" + Server
    cmdline += "-d \"" + targetpath + "\""

    if overwrite:
        cmdline += " -o"
    cmdline += "|pause"
    if os.path.exists(targetpath):
        oldtime = os.path.getmtime(targetpath)
    else:
        oldtime = None
    subprocess.call(cmdline)
    if os.path.exists(targetpath) and os.path.getmtime(targetpath) != oldtime:
        return True
    else:
        return False


def EasyCopyModel(sourcepath, targetpath, overwrite=True):
    try:
        if targetpath.startswith("RSN"):
            return False
        else:
            if sourcepath.startswith("RSN"):
                sourcepathlist = sourcepath.split("/")
                Server = sourcepathlist[2]
                Path = sourcepath.replace("RSN://" + Server + "/", "")
                SourceModelPath = ServerPath(Server, Path)
            else:
                SourceModelPath = FilePath(sourcepath)
            TargetPathFolder = os.path.dirname(targetpath)
            if not os.path.exists(TargetPathFolder):
                os.mkdir(TargetPathFolder)
            # Try CMD utility
            # if not EasyCopyModelCMD(sourcepath, targetpath, overwrite):
            app = __revit__.Application
            app.CopyModel(SourceModelPath, targetpath, overwrite)
            return True
    except:
        return False


def OpenAndSaveAsModelWithDetach(sourcepath, targetpath):
    result = []
    pathfortransmit = []
    if isinstance(targetpath, list):
        pass
    else:
        targetpath = [targetpath]
    if sourcepath.lower().endswith(".rvt"):
        __revit__.DialogBoxShowing += on_dialog_open
        if len(targetpath) == 1 and not (targetpath[0].startswith("RSN")) and not("AR" in sourcepath or "_A_" in sourcepath):
            easyresult = EasyCopyModel(sourcepath, targetpath[0])
            if easyresult:
                pathfortransmit.append(targetpath[0])
        else:
            easyresult = False
        # print(easyresult)

        if len(targetpath) > 1 or easyresult is False:
            if "AR" in sourcepath or "_A_" in sourcepath:
                model = OpenARModelWithDetach(sourcepath)
            else:
                model = OpenModelWithDetach(sourcepath)
            for t in targetpath:
                if model is not None:
                    SaveResult = SaveAsModelWithDetach(model, t)
                    result.append(SaveResult)
                    if SaveResult and not(t.startswith("RSN")):
                        pathfortransmit.append(t)
                else:
                    result.append(False)
            if model is not None:
                model.Close(False)
        __revit__.DialogBoxShowing -= on_dialog_open
    else:
        for t in targetpath:
            try:
                copyfile(sourcepath, t)
                result.append(True)
            except:
                result.append(False)
    for p in pathfortransmit:
        ModelPath = FilePath(p)
        tmd = TransmissionData.ReadTransmissionData(ModelPath)
        tmd.IsTransmitted = True
        TransmissionData.WriteTransmissionData(ModelPath, tmd)
    return result


def TransmissionDataLinksCorrection(models, newlinksdict):
    Result = []
    for m in models:
        try:
            ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(m)
            tmd = TransmissionData.ReadTransmissionData(ModelPath)
            links = tmd.GetAllExternalFileReferenceIds()
            for li in links:
                ref = tmd.GetLastSavedReferenceData(li)
                LinkPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(ref.GetPath())
                LinkName = LinkPath.replace("/", "\\").split("\\")[-1]
                if LinkName in newlinksdict.keys():
                    LinkNewPath = newlinksdict[LinkName]
                    tmd.SetDesiredReferenceData(li, ModelPathUtils.ConvertUserVisiblePathToModelPath(LinkNewPath), PathType.Relative, False)
                # else:
                    # tmd.SetDesiredReferenceData(li,ModelPathUtils.ConvertUserVisiblePathToModelPath(LinkPath),ref.PathType,False)
            tmd.IsTransmitted = True
            TransmissionData.WriteTransmissionData(ModelPath, tmd)
            Result.append(True)
        except:
            Result.append(False)
    return Result


def RSContent(uri, rsi, folder="|"):
    urlcont = uri + folder + "/contents"
    try:
        request = System.Net.WebRequest.Create(urlcont)
        request.Method = "GET"
        request.Headers.Add("User-Name", System.Environment.UserName)
        request.Headers.Add("User-Machine-Name", System.Environment.MachineName)
        request.Headers.Add("Operation-GUID", System.Guid.NewGuid().ToString())
        response = request.GetResponse()
        responseStr = System.IO.StreamReader(response.GetResponseStream()).ReadToEnd()
        # print(responseStr)
        resultjson = json.loads(responseStr)
        fileslist = []
        folders = resultjson["Folders"]
        files = resultjson["Models"]
        for f in files:
            if folder == "|":
                filelink = f["Name"]
            else:
                filelink = folder.replace("|", "/") + "/" + f["Name"]
            fileslist.append(rsi + filelink)
        for f in folders:
            if folder == "|":
                subfolder = f["Name"]
            else:
                subfolder = folder + "|" + f["Name"]
            fileslist.extend(RSContent(uri, rsi, subfolder))
    except:
        fileslist = ["!ERROR! on folder - " + folder]
    return fileslist


def ReadServerModels(path):
    RevitServer = path.split("/")[2]
    RevitVersion = __revit__.Application.VersionNumber
    url = "http://" + RevitServer + "/RevitServerAdminRESTService" + RevitVersion + "/AdminRESTService.svc/"
    rsn = "RSN://" + RevitServer + "/"
    folder = path.replace(rsn, "")
    FolderForRequest = folder.replace("/", "|")
    if path + "/" == rsn or path == rsn:
        RSModels = RSContent(url, rsn)
    else:
        RSModels = RSContent(url, rsn, FolderForRequest)
    return RSModels


def ReadFileModels(path, extension):
    extension = extension.lower()
    tree = os.walk(path)
    FSModels = []
    for t in tree:
        folder = t[0]
        LevelFiles = t[2]
        for f in LevelFiles:
            if f.lower().endswith(extension) and "archive" not in folder.lower().replace(path, ""):
                FSModels.append(folder + "\\" + f)
    return FSModels


def ReadModels(path, extension=".rvt"):
    if extension.lower() == ".rvt" and path.startswith("RSN"):
        return ReadServerModels(path.replace("\\","/"))
    else:
        return ReadFileModels(path.replace("/", "\\"), extension)


def RSDate(uri, rsi, model):
    if model != None:
        try:
            urlmodel = uri + model.replace(rsi,"").replace("/","|") + "/modelInfo"
            request = System.Net.WebRequest.Create(urlmodel)
            request.Method = "GET"
            request.Headers.Add("User-Name", System.Environment.UserName);
            request.Headers.Add("User-Machine-Name", System.Environment.MachineName);
            request.Headers.Add("Operation-GUID", System.Guid.NewGuid().ToString()); 
            response = request.GetResponse()
            responseStr = System.IO.StreamReader(response.GetResponseStream()).ReadToEnd()
            # print(responseStr)
            resultjson = json.loads(responseStr)
            valstr = resultjson["DateModified"].replace("/Date(","").replace(")/","")
            val = float(valstr)/1000
            date = val
            return date
        except:
            return None    
    else:
        return None

def GetModelsDate(models):
    if isinstance(models,list):
        pass
    else:
        models = [models]
    datelist = []
    for m in models:
        if m == None:
            date = None
        elif m.startswith("RSN"):
            RevitServer = m.split("/")[2]
            RevitVersion = __revit__.Application.VersionNumber
            url =   "http://" + RevitServer + "/RevitServerAdminRESTService" + RevitVersion + "/AdminRESTService.svc/"
            rsn = "RSN://" + RevitServer + "/"
            ModelForRequest = m.replace(rsn,"").replace("/","|")
            date= RSDate(url,rsn,ModelForRequest)
            
        elif os.path.exists(m):
            date = float(os.path.getmtime(m))
        else:
            date = None
        datelist.append(date)
    return datelist

def FindTargetModelsByPath(sourcemodels, sourceserver, sourcefolder, targetserver, targetfolder):
    if isinstance(sourcemodels,list):
        pass
    else:
        sourcemodels = [sourcemodels]
    targetmodels = []
    for s in sourcemodels:
        targetmodel = s.replace(sourceserver + sourcefolder, targetserver + targetfolder)
        if targetmodel.startswith("RSN"):
            targetmodel = targetmodel.replace("\\","/")
        else:
            targetmodel = targetmodel.replace("/","\\")
        targetmodels.append(targetmodel)
    return targetmodels

def CompareModelsFromSourceAndTargets(SourceServer, SourceFolder, TargetServer, TargetFolder,extension = ".rvt"):
    SourceModels = ReadModels(SourceServer + SourceFolder,extension)
    
    SourceModelsDates = GetModelsDate(SourceModels)
    if isinstance(TargetServer,list):
        TargetServerlist = TargetServer
        TargetFolderlist = TargetFolder
    else:
        TargetServerlist = [TargetServer]
        TargetFolderlist = [TargetFolder]
    NumberOfTargets = len(TargetServer)
    TargetModelsList = []
    TargetModelsDatesList = []
    for i in range(NumberOfTargets):
        TargetModels = FindTargetModelsByPath(SourceModels, SourceServer, SourceFolder, TargetServerlist[i], TargetFolderlist[i])
        TargetModelsDates = GetModelsDate(TargetModels)
        TargetModelsList.append(TargetModels)
        TargetModelsDatesList.append(TargetModelsDates)
    TargetModelsList = map(list, zip(*TargetModelsList))
    TargetModelsDatesList = map(list, zip(*TargetModelsDatesList))
    ModelsWithTargets = {}
    for i in range(len(SourceModels)):
        ModelsWithTargets[SourceModels[i]] = TargetModelsList[i]
    Models = {}
    Models["All"] = sorted(SourceModels)
    ModelNeedToReload = []
    for i in range(len(SourceModels)):
        ReloadState = []
        #Разобраться почему SourceModelsDates становиться NoneType
        try:
            for t in range(NumberOfTargets):
                ReloadState.append(TargetModelsDatesList[i][t]<SourceModelsDates[i]+10*60)
        except:
            for t in range(NumberOfTargets):
                ReloadState.append(True)
        if True in ReloadState:
            ModelNeedToReload.append(SourceModels[i])
    Models["Not up to date"] = sorted(ModelNeedToReload)
    return [ModelsWithTargets,Models]

def CopyModels(Models,ModelsWithTargets,inheritanced = False):
    if __revit__.Application.Documents.IsEmpty:
        ModelsReport = {}
        if Models != None and len(Models)>0:
            Result = []
            count = 0
            number = len(Models)
            chanseled = False
            start = datetime.datetime.now()
            error = []
            errormodels = []
            with forms.ProgressBar(title='Copy models ' + str(count) + "/" + str(number),cancellable=True) as pb:
                lefttime = ""
                pb.update_progress(count, number)
                for m in Models:
                    modelname = m.replace("\\","/").split("/")[-1]
                    pb.title='Copy ' + modelname + " (" + str(count) + "/" + str(number) + lefttime + ")"
                    pb.update_progress(count, number)
                    if not(chanseled):
                        res = OpenAndSaveAsModelWithDetach(m,ModelsWithTargets[m])
                        Result.append(res)
                        if False in res:
                            error.append("Model copy error:\n" + m.replace("/","|").replace("\\","|").split("|")[-1])
                            errormodels.append(m)
                            ModelsReport[m] = "Error"
                        else:
                            ModelsReport[m] = "Ok"
                    else:
                        ModelsReport[m] = "Canceled"
                    count += 1
                    le = (datetime.datetime.now() - start).seconds
                    left = round(le/count*(number - count))
                    lefthours = left // 3600
                    leftminutes = int((left - lefthours*3600)/60)
                    lefttime = " average " + ((str(lefthours) + " hours and " ) if lefthours>0 else "") + str(leftminutes) + " minutes" + " left"
                    if pb.cancelled and not(chanseled):
                        canseldialog = TaskDialog("Cancel")
                        canseldialog.MainInstruction = "Are you sure you want to cansel process?"
                        canseldialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                        canseldialog.DefaultButton = TaskDialogResult.No
                        answer = canseldialog.Show()
                        if answer == TaskDialogResult.Yes:
                            chanseled = True
            if chanseled:
                dialog = TaskDialog("Operation Canceled")
            else:
                dialog = TaskDialog("Models copy")
            dialog.MainInstruction = str(len(Result)) + " models was processed."
            dialog.CommonButtons = TaskDialogCommonButtons.Close   
            if len(error)>0:
                dialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning
                dialog.MainContent = "On copy of " + str(len(error)) + " models errors occurred."
                dialog.ExpandedContent = "\n".join(error)
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Try copy error models again")
            answer = dialog.Show()
            if answer == TaskDialogResult.CommandLink1:
                report = CopyModels(errormodels,ModelsWithTargets,True)
                if report != None:
                    for k in report.keys():
                        ModelsReport[k] = ModelsReport[k] + "<br/>" + report[k]
            if  inheritanced == False:
                tabledata = []
                for m in sorted(ModelsReport.keys()):
                    model = m.replace("\\","/").split("/")[-1]
                    tabledata.append([model,ModelsReport[m],m,"<br/>".join(ModelsWithTargets[m])])
                output = pyrevit.output.get_output()
                output.print_table(table_data=tabledata,
                    title="Model copy report",
                    columns=["Model","State","Source Path","Target Path"]
                    )

        return ModelsReport
    else:
        reportdialog = TaskDialog("Model Copy")
        reportdialog.MainInstruction = "Funktion requires all Revit models to be closed. Please close all open models and try again."
        reportdialog.CommonButtons = TaskDialogCommonButtons.Close
        reportdialog.DefaultButton = TaskDialogResult.Close
        reportdialog.Show()


def ConvertToNWC(ModelsForNWC):
    if ModelsForNWC != None and isinstance(ModelsForNWC,list):
        if len(ModelsForNWC)>0:
            tempfilename = os.environ['temp'] + "\\NavisExport\\nwfexport" + str(time.time()) + ".txt"
            # print(tempfilename)
            cmdfilename = os.environ['temp'] + "\\NavisExport\\nwfexport" + str(time.time()) + ".cmd"
            vbsfilename = os.environ['temp'] + "\\NavisExport\\nwfexport" + str(time.time()) + ".vbs"
            logfilename = os.environ['temp'] + "\\NavisExport\\nwfexport" + str(time.time()) + ".log"
            if os.path.exists("C:\\Program Files\\Autodesk\\Navisworks Manage 2020\\FileToolsTaskRunner.exe"):
                command = "\"C:\\Program Files\\Autodesk\\Navisworks Manage 2020\\FileToolsTaskRunner.exe\"" + " /i \"" + tempfilename + "\" /of \"" + tempfilename.replace(".txt",".nwf") + "\" /inc /version 2018 >> " + logfilename
            else:
                command = "\"C:\\Program Files\\Autodesk\\Navisworks Manage 2019\\FileToolsTaskRunner.exe\"" + " /i \"" + tempfilename + "\" /of \"" + tempfilename.replace(".txt",".nwf") + "\" /inc /version 2018 >> " + logfilename
            folder = os.path.dirname(os.path.abspath(tempfilename))
            if not os.path.exists(folder):
                os.makedirs(folder)
            try:
                f = codecs.open(tempfilename, encoding='utf-8', mode='w+')
                for l in ModelsForNWC:
                    f.write(u'' + l.replace("/","\\") + '\n')
            finally:
                f.close()
            try:
                f = codecs.open(cmdfilename, encoding='utf-8', mode='w+')
                f.write(u'' + command)
            finally:
                f.close()
            try:
                f = codecs.open(vbsfilename, encoding='utf-8', mode='w+')
                f.write(u'Set WshShell = CreateObject("WScript.Shell")\n')
                f.write(u'WshShell.Run chr(34) & "' + cmdfilename + '" & chr(34), 0, true \n')
                f.write(u'Set WshShell = Nothing\n')
                f.write(u'Set WshShell = CreateObject("WScript.Shell")\n')
                f.write(u'WshShell.Run "cmd /q /k more " & "' + logfilename + '", 1, false\n')
                f.write(u'Set WshShell = Nothing\n')
                f.write(u'answer = MsgBox ("Details are in cmd",0,"Export Done")\n')
            finally:
                f.close()
            os.startfile(vbsfilename)



def ConvertToNWCstream(ModelsForNWC, stream = 2):
    try:
        stream = int(stream)
        if stream > 3:
            stream = 3
    except:
        stream = 1
    if ModelsForNWC != None and isinstance(ModelsForNWC,list):
        if len(ModelsForNWC)>5:
            StreamModelNumber = -(-len(ModelsForNWC)//stream)
            start = 0
            for s in range(stream):
                end = start + StreamModelNumber
                if end>=len(ModelsForNWC):
                    end = len(ModelsForNWC)
                ConvertToNWC(ModelsForNWC[start:end])
                start = end 
        elif len(ModelsForNWC)>0:
            ConvertToNWC(ModelsForNWC)

def SelectModelsForNWC(FolderPickName = "Select Folder",
                            FolderPickInstruction = "Please select path type",
                            SelectTitle = "Select Models",
                            SelectButtonText = "Select",
                            Multiselect = False,
                            pathtype = None):
    SelectedModels = None
    folder = SelectFolder(FolderPickName,FolderPickInstruction, pathtype)
    if folder != None:
        models = ReadModels(folder)
        RevitModels = ReadModels(folder)
        ModelsDict = {}
        ModelsDict["ALL"] = []
        ModelsDict["Not up to date"] = []
        for rm in RevitModels:
            nm = rm[:-3] + "nwc"
            if os.path.exists(rm):
                rvtdate = float(os.path.getmtime(rm))
            else:
                rvtdate = None
            if os.path.exists(nm):
                nwcdate = float(os.path.getmtime(nm))
            else:
                nwcdate = None
            if rvtdate != None:
                ModelsDict["ALL"].append(rm)
                if nwcdate == None or nwcdate<rvtdate+10*60:
                    # print(rm)
                    # print(nm)
                    # print(nwcdate)
                    # print(rvtdate+10*60)
                    # print(nwcdate - (rvtdate+10*60))
                    ModelsDict["Not up to date"].append(rm)
        SelectedModels = forms.SelectFromList.show(
            ModelsDict,
            title=SelectTitle,
            width=1000,
            button_name=SelectButtonText,
            multiselect=Multiselect,
            group_selector_title = 'Option',
            default_group = 'Not up to date'
            )
    if SelectedModels == None:
        return []
    elif isinstance(SelectedModels,list):
        return SelectedModels
    else:
        return [SelectedModels]


def SelectModelsFromFolder(FolderPickName = "Select Folder",
                           FolderPickInstruction = "Please select path type",
                           SelectTitle = "Select Models",
                           SelectButtonText = "Select",
                           Multiselect = False,
                           pathtype = None,
                           extension=".rvt"):
    SelectedModels = None
    folder = SelectFolder(FolderPickName, FolderPickInstruction, pathtype)
    if folder != None:
        models = ReadModels(folder, extension)
        SelectedModels = forms.SelectFromList.show(
            models,
            title=SelectTitle,
            width=1000,
            button_name=SelectButtonText,
            multiselect=Multiselect
            )
    if SelectedModels == None:
        return []
    elif isinstance(SelectedModels,list):
        return SelectedModels
    else:
        return [SelectedModels]

def SelectFolder(description = "Select Folder",instruct = "Please select path type", pathtype = None):
    if pathtype == None:
        dialog = TaskDialog(description)
        dialog.MainInstruction = instruct
        dialog.CommonButtons = TaskDialogCommonButtons.Close   
        dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "File Server Path")
        dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Revit Server")
        dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Insert Path As String")
        dialog.DefaultButton = TaskDialogResult.CommandLink1
        answer = dialog.Show()
    elif pathtype == "File":
       answer = TaskDialogResult.CommandLink1
    elif pathtype == "Revit":
        answer = TaskDialogResult.CommandLink2
    elif pathtype == "String":
        answer = TaskDialogResult.CommandLink3
    else:
        answer = None

    if answer == TaskDialogResult.CommandLink1:
        path = pyrevit.forms.pick_folder(description)
    elif answer == TaskDialogResult.CommandLink2:
        RevitVersion = __revit__.Application.VersionNumber
        RSNPath = os.environ['programdata'] + "\\Autodesk\\Revit Server %s\\Config\\RSN.ini"%RevitVersion
        RSNfile = open(RSNPath)
        RevitServer = RSNfile.read().rstrip().replace(" ", "")
        if "\n" in RevitServer:
            RevitServers = RevitServer.split("\n")
            RevitServer = forms.SelectFromList.show(RevitServers,
                                                    title="Select Revit Server",
                                                    width=1000,
                                                    Height=500,
                                                    button_name='Select',
                                                    multiselect=False)
            RevitServer = RevitServer
        if RevitServer is not None:
            url = "http://" + RevitServer + "/RevitServerAdminRESTService" + RevitVersion + "/AdminRESTService.svc/"
            seldict = {"Select Discipline":[""]}
            seldict["!!_ALL_SERVER"] = ["RSN://" + RevitServer]
            Projects = RSFolderContent(url)[0]
            # print(Projects)
            for p in Projects:
                Zones = RSFolderContent(url,p)[0]
                # print(Zones)
                subpath = "RSN://" + RevitServer + "/" +  p
                if "!!_WHOLE_PROJECT" in seldict.keys():
                    seldict["!!_WHOLE_PROJECT"].append(subpath)
                else:
                    seldict["!!_WHOLE_PROJECT"] = [subpath]
                for z in Zones:
                    Disciplines = RSFolderContent(url,p + "|" + z)[0]
                    # print(Disciplines)
                    subpath = "RSN://" + RevitServer + "/" +  p + "/" + z
                    if "00_ALL_DISCIPLINES" in seldict.keys():
                        seldict["00_ALL_DISCIPLINES"].append(subpath)
                    else:
                        seldict["00_ALL_DISCIPLINES"] = [subpath]
                    for d in Disciplines:
                        subpath = "RSN://" + RevitServer + "/" + p + "/" + z + "/" + d
                        if d in seldict.keys():
                            seldict[d].append(subpath)
                        else:
                            seldict[d] = [subpath]

            SelectedProject = forms.SelectFromList.show(
                seldict,
                title=description,
                width=1000,
                Height=500,
                button_name='Select',
                multiselect=False,
                group_selector_title = 'Discipline',
                default_group = 'Select Discipline'
                )
            if SelectedProject == "":
                path = None
            else:    
                path = SelectedProject
    elif answer == TaskDialogResult.CommandLink3:
        path = forms.ask_for_string(
            title = description,
            prompt = 'Insert Path',
            default = 'Path',
            width=2000
            )
        if path == 'Path':
            path = None
    else:
        path = None

    return path


def RSFolderContent(uri, folder="|"):
    urlcont = uri + folder + "/contents"
    try:
        request = System.Net.WebRequest.Create(urlcont)
        request.Method = "GET"
        request.Headers.Add("User-Name", System.Environment.UserName)
        request.Headers.Add("User-Machine-Name", System.Environment.MachineName)
        request.Headers.Add("Operation-GUID", System.Guid.NewGuid().ToString())
        response = request.GetResponse()
        responseStr = System.IO.StreamReader(response.GetResponseStream()).ReadToEnd()
        resultjson = json.loads(responseStr)
        fileslist = []
        folders = resultjson["Folders"]
        files = resultjson["Models"]
        folderNames = []
        filesNames = []
        for f in folders:
            folderNames.append(f["Name"])
        for f in files:
            filesNames.append(f["Name"])
        return [folderNames, filesNames]
    except:
        return None


# def GetCurrentProjectAndDiscipline():
#     uidoc = __revit__.ActiveUIDocument
#     if uidoc !=None:
#         doc = uidoc.Document
#         path = doc.PathName
#     else:
#         return None

# def GetCurrentDiscipline()
