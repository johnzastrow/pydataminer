# Author:  Ehren Hill - Tetra Tech
# Date:    November 20, 2013
# Version: ArcGIS 10.1 and 10.2
# Purpose: This script will iterate through each GIS file in a folder (and all recursive folders) and report information
# about each layers metadata.  The script is intended to run from a script tool inside an ArcGIS Toolbox.  The script requires three
# user input values:
#   1. An input folder
#   2. An output Excel File
#   3. An output folder to store temporary files

try:
    #Setup environment and import modules
    import xlwt, os, arcpy
    from xml.etree.ElementTree import ElementTree
    from xml.etree.ElementTree import Element, SubElement
    arcpy.env.overwriteOutput = True


    #Get user input parameters, please note that the inputFolder cannot contain any GIS files, or they will not be read
    inputFolder = arcpy.GetParameterAsText(0)
    outputFile = arcpy.GetParameterAsText(1)
    tempOutputFolder = arcpy.GetParameterAsText(2)
    #Input parameters below are for testing only
    #inputFolder = r"C:\workingProjectData\tetraTech\data\workingData\dataMiner\subfolder"
    #outputFile = r"C:\workingProjectData\tetraTech\tasks\dataMiner\version_2.0\output\output.xls"
    #tempOutputFolder = r"C:\workingProjectData\tetraTech\tasks\dataMiner\version_2.0\temp"


    #Create workbook with one sheet named Output
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Output')


    #Create list of column names in the workbook
    headerNamesList = ['ID', 'Feature Name', 'Feature Path', 'Feature Type', 'Object Type','Feature Count', 'Title', 'Keywords',
                   'Abstract', 'Purpose', 'Contact', 'Distribution', 'Coordinate System Type', 'Geographic Coordinate System',
                   'Projected Coordinate System', 'Horizontal Units', 'North Bounding Coordinate',
                   'East Bounding Coordinate', 'South Bounding Coordinate', 'West Bounding Coordinate',
                   'Created Date', 'Modified Date', 'MetaData Version']
    headerColumn = 0
    for name in headerNamesList:
        ws.write(0,headerColumn, name)
        headerColumn = headerColumn + 1


    #Create empty global lists and other global variables
    datasetList = []
    fcList = []
    tableList = []
    workspaceList = []
    rasterList = []
    listOfEnvWorkspaces = []
    excelRow = 1


    #Fucntion to test if metadata tags have values in XML file
    def verifyMetadata(tag):
        global tagReturn
        if tag is not None:
            tagReturn = tag.text
            return tagReturn
        else:
            tagReturn = "METADATA VALUE IS EMPTY"
            return tagReturn


    #Searches through all folders for GIS files
    for dir, folderNames, fileNames in os.walk(inputFolder):
        listOfEnvWorkspaces.append(dir)
    for subFolderName in listOfEnvWorkspaces:
        arcpy.env.workspace = subFolderName
        print "Processing Folder: " + subFolderName
        arcpy.AddMessage("Processing Folder: " + subFolderName)

        datasetList = arcpy.ListDatasets()
        fcList = arcpy.ListFeatureClasses()
        tableList = arcpy.ListTables()
        workspaceList = arcpy.ListWorkspaces('', 'Access')
        rasterList = arcpy.ListRasters()

        #Checks for feature classes in file geodatabases feature classes, coverages, and CAD files
        for dataset in datasetList:
            for fc in arcpy.ListFeatureClasses('', '', dataset):
                try:
                    #Creates a temporary empty XML file to store each files XML data
                    fcXML = tempOutputFolder + "/" + fc + ".xml"
                    tempXMLFile = open(fcXML, "w")
                    tempXMLFile.write("<metadata>\n")
                    tempXMLFile.write("</metadata>\n")
                    tempXMLFile.close()
                    #Copies metadata from file into temporary XML file
                    arcpy.MetadataImporter_conversion (fc, fcXML)

                    #Makes an ElementTree object and reads XML into ElementTree
                    tree = ElementTree()
                    tree.parse(fcXML)
                    #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                    if tree.find("Esri/ArcGISFormat") is not None:
                        print "     Processing File: " + os.path.join(dataset, fc) + " (Version 10.x XML)"
                        arcpy.AddMessage("     Processing File: " + os.path.join(dataset, fc) + " (Version 10.x XML)")
                        title = tree.find("dataIdInfo/idCitation/resTitle")
                        keywords = tree.find("dataIdInfo/searchKeys/keyword")
                        abstract = tree.find("dataIdInfo/idAbs")
                        purpose = tree.find("dataIdInfo/idPurp")
                        contact = tree.find("mdContact/rpIndName")
                        distribution = tree.find("distInfo/distFormat/formatName")
                        fcMetaDataVersion = "10.x"
                    else:
                        print "     Processing File: " + os.path.join(dataset, fc) + " (Version 9.x XML)"
                        arcpy.AddMessage ("     Processing File: " + os.path.join(dataset, fc) + " (Version 9.x XML)")
                        title = tree.find("idinfo/citation/citeinfo/title")
                        keywords = tree.find("idinfo/keywords/theme/themekey")
                        abstract = tree.find("idinfo/descript/abstract")
                        purpose = tree.find("idinfo/descript/purpose")
                        contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                        distribution = tree.find("distinfo/resdesc")
                        fcMetaDataVersion = "9.x"
                    created = tree.find("Esri/CreaDate")
                    modified = tree.find("Esri/ModDate")

                    #Setup variables for textToWriteList and check for valid metadata values
                    if arcpy.Describe(dataset).dataType == "Coverage" or arcpy.Describe(dataset).dataType == "CadDrawingDataset" or arcpy.Describe(dataset).dataType == "Workspace":
                        desc = arcpy.Describe(dataset + "/" + fc)
                    else:
                        desc = arcpy.Describe(fc)
                    fcPath = os.path.join(subFolderName, dataset)
                    fcDataType = desc.dataType
                    fcShapeType = desc.shapeType
                    if arcpy.Describe(dataset).dataType == "Coverage" or arcpy.Describe(dataset).dataType == "CadDrawingDataset":
                        fcFeatureCount = str(arcpy.GetCount_management(dataset + "/" + fc))
                    else:
                        fcFeatureCount = str(arcpy.GetCount_management(fc))
                    verifyMetadata (title)
                    fcTitle = tagReturn
                    verifyMetadata (keywords)
                    fcKeywords = tagReturn
                    verifyMetadata (abstract)
                    fcAbstract = tagReturn
                    verifyMetadata (purpose)
                    fcPurpose = tagReturn
                    verifyMetadata (contact)
                    fcContact = tagReturn
                    verifyMetadata (distribution)
                    fcDistribution = tagReturn
                    fcCoordType = desc.SpatialReference.type
                    if fcCoordType == "Unknown":
                        fcGCS = "No Coordinate System Assigned"
                    else:
                        fcGCS = desc.SpatialReference.GCS.name
                    if desc.SpatialReference.PCSname:
                        fcPCS = desc.SpatialReference.PCSname
                        fcUnit = desc.SpatialReference.linearUnitName
                    else:
                        fcPCS = "N/A"
                        fcUnit = "N/A"
                    if str(desc.extent) == "None" or int(fcFeatureCount) == 0 :
                        fcYMax = "No Features"
                        fcXMax = "No Features"
                        fcYMin = "No Features"
                        fcXMin = "No Features"
                    else:
                        fcYMax = str(desc.extent.YMax)
                        fcXMax = str(desc.extent.XMax)
                        fcYMin = str(desc.extent.YMin)
                        fcXMin = str(desc.extent.XMin)
                    verifyMetadata (created)
                    fcCreated = tagReturn
                    verifyMetadata (modified)
                    fcModified = tagReturn

                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, fc, fcPath, fcDataType, fcShapeType, fcFeatureCount, fcTitle, fcKeywords, fcAbstract,
                                       fcPurpose, fcContact, fcDistribution, fcCoordType, fcGCS, fcPCS, fcUnit, fcYMax, fcXMax, fcYMin,
                                       fcXMin, fcCreated, fcModified, fcMetaDataVersion]

                    #Start writing information on fc and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    os.remove(fcXML)

                    #Increment row counter by 1
                    excelRow = excelRow + 1

                except:
                    print "     Could not process: " + fc + " at: " + os.path.join(subFolderName, dataset) + "; Skipping"
                    arcpy.AddMessage("     Could not process: " + fc + " at: " + os.path.join(subFolderName, dataset) + "; Skipping")
                    fcPath = os.path.join(subFolderName, dataset)
                    fcDataType = "COULD NOT PROCESS FILE"
                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, fc, fcPath, fcDataType]

                    #Start writing information on fc and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    try:
                        os.remove(fcXML)
                    except:
                        print "     Could not remove file: " + str(fcXML)
                        arcpy.AddMessage("     Could not remove file: " + str(fcXML))

                    #Increment row counter by 1
                    excelRow = excelRow + 1


        #Checks for feature classes in file geodatabases (not in feature datesets)
        for fc in fcList:
            try:
                if arcpy.Describe(fc).dataType != "CoverageFeatureClass":
                    #Creates a temporary empty XML file to store each files XML data
                    fcXML = tempOutputFolder + "/" + fc + ".xml"
                    tempXMLFile = open(fcXML, "w")
                    tempXMLFile.write("<metadata>\n")
                    tempXMLFile.write("</metadata>\n")
                    tempXMLFile.close()
                    #Copies metadata from file into previous created temporary XML file
                    arcpy.MetadataImporter_conversion (fc, fcXML)

                    #Makes an ElementTree object and reads XML into ElementTree
                    tree = ElementTree()
                    tree.parse(fcXML)
                    #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                    if tree.find("Esri/ArcGISFormat") is not None:
                        print "     Processing File: " + fc + " (Version 10.x XML)"
                        arcpy.AddMessage("     Processing File: " + fc + " (Version 10.x XML)")
                        title = tree.find("dataIdInfo/idCitation/resTitle")
                        keywords = tree.find("dataIdInfo/searchKeys/keyword")
                        abstract = tree.find("dataIdInfo/idAbs")
                        purpose = tree.find("dataIdInfo/idPurp")
                        contact = tree.find("mdContact/rpIndName")
                        distribution = tree.find("distInfo/distFormat/formatName")
                        fcMetaDataVersion = "10.x"
                    else:
                        print "     Processing File: " + fc + " (Version 9.x XML)"
                        arcpy.AddMessage("     Processing File: " + fc + " (Version 9.x XML)")
                        title = tree.find("idinfo/citation/citeinfo/title")
                        keywords = tree.find("idinfo/keywords/theme/themekey")
                        abstract = tree.find("idinfo/descript/abstract")
                        purpose = tree.find("idinfo/descript/purpose")
                        contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                        distribution = tree.find("distinfo/resdesc")
                        fcMetaDataVersion = "9.x"
                    created = tree.find("Esri/CreaDate")
                    modified = tree.find("Esri/ModDate")

                    #Setup variables for textToWriteList and check for valid metadata values
                    desc = arcpy.Describe(fc)
                    fcPath = desc.path
                    fcDataType = desc.dataType
                    fcShapeType = desc.shapeType
                    fcFeatureCount = str(arcpy.GetCount_management(fc))
                    verifyMetadata (title)
                    fcTitle = tagReturn
                    verifyMetadata (keywords)
                    fcKeywords = tagReturn
                    verifyMetadata (abstract)
                    fcAbstract = tagReturn
                    verifyMetadata (purpose)
                    fcPurpose = tagReturn
                    verifyMetadata (contact)
                    fcContact = tagReturn
                    verifyMetadata (distribution)
                    fcDistribution = tagReturn
                    fcCoordType = desc.SpatialReference.type
                    if fcCoordType == "Unknown":
                        fcGCS = "No Coordinate System Assigned"
                    else:
                        fcGCS = desc.SpatialReference.GCS.name
                    if desc.SpatialReference.PCSname:
                        fcPCS = desc.SpatialReference.PCSname
                        fcUnit = desc.SpatialReference.linearUnitName
                    else:
                        fcPCS = "N/A"
                        fcUnit = "N/A"
                    if str(desc.extent) == "None":
                        fcYMax = "No Features"
                        fcXMax = "No Features"
                        fcYMin = "No Features"
                        fcXMin = "No Features"
                    else:
                        fcYMax = str(desc.extent.YMax)
                        fcXMax = str(desc.extent.XMax)
                        fcYMin = str(desc.extent.YMin)
                        fcXMin = str(desc.extent.XMin)
                    verifyMetadata (created)
                    fcCreated = tagReturn
                    verifyMetadata (modified)
                    fcModified = tagReturn

                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, fc, fcPath, fcDataType, fcShapeType, fcFeatureCount, fcTitle, fcKeywords, fcAbstract,
                                       fcPurpose, fcContact, fcDistribution, fcCoordType, fcGCS, fcPCS, fcUnit, fcYMax, fcXMax, fcYMin,
                                       fcXMin, fcCreated, fcModified, fcMetaDataVersion]

                    #Start writing information on fc and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    os.remove(fcXML)

                    #Increment row counter by 1
                    excelRow = excelRow + 1

            except:
                print "     Could not process: " + fc + " at: " + subFolderName + "; Skipping"
                arcpy.AddMessage("     Could not process: " + fc + " at: " + subFolderName + "; Skipping")
                fcPath = subFolderName
                fcDataType = "COULD NOT PROCESS FILE"
                #Create list to hold text values for writing to Excel
                textToWriteList = [excelRow, fc, fcPath, fcDataType]

                #Start writing information on fc and metadata to Excel
                column = 0
                for text in textToWriteList:
                    row = ws.row(excelRow)
                    row.write(column, text)
                    column = column + 1

                #Remove temporary empty XML file
                try:
                    os.remove(fcXML)
                except:
                    print "     Could not remove file: " + str(fcXML)
                    arcpy.AddMessage("     Could not remove file: " + str(fcXML))

                #Increment row counter by 1
                excelRow = excelRow + 1


        #Checks for tables in a directory and in file geodatabases
        for table in tableList:
            try:
                #Creates a temporary empty XML file to store each files XML data
                tableXML = tempOutputFolder + "/" + table + ".xml"
                tempXMLFile = open(tableXML, "w")
                tempXMLFile.write("<metadata>\n")
                tempXMLFile.write("</metadata>\n")
                tempXMLFile.close()
                #Copies metadata from file into previous created temporary XML file
                arcpy.MetadataImporter_conversion (table, tableXML)

                #Makes an ElementTree object and reads XML into ElementTree
                tree = ElementTree()
                tree.parse(tableXML)
                #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                if tree.find("Esri/ArcGISFormat") is not None:
                    print "     Processing File: " + table + " (Version 10.x XML)"
                    arcpy.AddMessage("     Processing File: " + table + " (Version 10.x XML)")
                    title = tree.find("dataIdInfo/idCitation/resTitle")
                    keywords = tree.find("dataIdInfo/searchKeys/keyword")
                    abstract = tree.find("dataIdInfo/idAbs")
                    purpose = tree.find("dataIdInfo/idPurp")
                    contact = tree.find("mdContact/rpIndName")
                    distribution = tree.find("distInfo/distFormat/formatName")
                    tableMetaDataVersion = "10.x"
                else:
                    print "     Processing File: " + table + " (Version 9.x XML)"
                    arcpy.AddMessage("     Processing File: " + table + " (Version 9.x XML)")
                    title = tree.find("idinfo/citation/citeinfo/title")
                    keywords = tree.find("idinfo/keywords/theme/themekey")
                    abstract = tree.find("idinfo/descript/abstract")
                    purpose = tree.find("idinfo/descript/purpose")
                    contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                    distribution = tree.find("distinfo/resdesc")
                    tableMetaDataVersion = "9.x"
                created = tree.find("Esri/CreaDate")
                modified = tree.find("Esri/ModDate")

                #Setup variables for textToWriteList and check for valid metadata values
                desc = arcpy.Describe(table)
                tablePath = desc.path
                tableDataType = desc.dataType
                tableShapeType = "N/A"
                tableFeatureCount = "N/A"
                verifyMetadata (title)
                tableTitle = tagReturn
                verifyMetadata (keywords)
                tableKeywords = tagReturn
                verifyMetadata (abstract)
                tableAbstract = tagReturn
                verifyMetadata (purpose)
                tablePurpose = tagReturn
                verifyMetadata (contact)
                tableContact = tagReturn
                verifyMetadata (distribution)
                tableDistribution = tagReturn
                tableCoordType = "N/A"
                tableGCS = "N/A"
                tablePCS = "N/A"
                tableUnit = "N/A"
                tableYMax = "N/A"
                tableXMax = "N/A"
                tableYMin = "N/A"
                tableXMin = "N/A"
                verifyMetadata (created)
                tableCreated = tagReturn
                verifyMetadata (modified)
                tableModified = tagReturn

                #Create list to hold text values for writing to Excel
                textToWriteList = [excelRow, table, tablePath, tableDataType, tableShapeType, tableFeatureCount, tableTitle, tableKeywords, tableAbstract,
                                   tablePurpose, tableContact, tableDistribution, tableCoordType, tableGCS, tablePCS, tableUnit, tableYMax, tableXMax, tableYMin,
                                   tableXMin, tableCreated, tableModified, tableMetaDataVersion]

                #Start writing information on table and metadata to Excel
                column = 0
                for text in textToWriteList:
                    row = ws.row(excelRow)
                    row.write(column, text)
                    column = column + 1

                #Remove temporary empty XML file
                os.remove(tableXML)

                #Increment row counter by 1
                excelRow = excelRow + 1

            except:
                print "     Could not process: " + table + " at: " + subFolderName + "; Skipping"
                arcpy.AddMessage("     Could not process: " + table + " at: " + subFolderName + "; Skipping")
                tablePath = subFolderName
                tableDataType = "COULD NOT PROCESS FILE"
                #Create list to hold text values for writing to Excel
                textToWriteList = [excelRow, table, tablePath, tableDataType]

                #Start writing information on fc and metadata to Excel
                column = 0
                for text in textToWriteList:
                    row = ws.row(excelRow)
                    row.write(column, text)
                    column = column + 1

                #Remove temporary empty XML file
                try:
                    os.remove(tableXML)
                except:
                    print "     Could not remove file: " + str(tableXML)
                    arcpy.AddMessage("Could not remove file: " + str(tableXML))

                #Increment row counter by 1
                excelRow = excelRow + 1



        #Checks for feature classes and tables in personal geodatabases and feature datasets in personal geodatabases
        for workspace in workspaceList:
            #set environmental workspace to personal geodatabase
            arcpy.env.workspace = workspace
            dataset = arcpy.ListDatasets()
            tableset = arcpy.ListTables()
            rasterset = arcpy.ListRasters()

            #Processes feature classes in personal geodatabase feature dataset
            for item in dataset:
                for fc in arcpy.ListFeatureClasses('', '', item):
                    try:
                        #Creates a temporary empty XML file to store each files XML data
                        fcXML = tempOutputFolder + "/" + fc + ".xml"
                        tempXMLFile = open(fcXML, "w")
                        tempXMLFile.write("<metadata>\n")
                        tempXMLFile.write("</metadata>\n")
                        tempXMLFile.close()
                        #update fc to include workspace path
                        fcAndWorkspace = workspace + '/' + item + '/' + fc
                        #Copies metadata from file into temporary XML file
                        arcpy.MetadataImporter_conversion (fcAndWorkspace, fcXML)
                        #Makes an ElementTree object and reads XML into ElementTree
                        tree = ElementTree()
                        tree.parse(fcXML)
                        #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                        if tree.find("Esri/ArcGISFormat") is not None:
                            print "     Processing File: " + os.path.join(item, fc) + " (Version 10.x XML)"
                            arcpy.AddMessage("     Processing File: " + os.path.join(item, fc) + " (Version 10.x XML)")
                            title = tree.find("dataIdInfo/idCitation/resTitle")
                            keywords = tree.find("dataIdInfo/searchKeys/keyword")
                            abstract = tree.find("dataIdInfo/idAbs")
                            purpose = tree.find("dataIdInfo/idPurp")
                            contact = tree.find("mdContact/rpIndName")
                            distribution = tree.find("distInfo/distFormat/formatName")
                            fcMetaDataVersion = "10.x"
                        else:
                            print "     Processing File: " + os.path.join(item, fc) + " (Version 9.x XML)"
                            arcpy.AddMessage ("     Processing File: " + os.path.join(item, fc) + " (Version 9.x XML)")
                            title = tree.find("idinfo/citation/citeinfo/title")
                            keywords = tree.find("idinfo/keywords/theme/themekey")
                            abstract = tree.find("idinfo/descript/abstract")
                            purpose = tree.find("idinfo/descript/purpose")
                            contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                            distribution = tree.find("distinfo/resdesc")
                            fcMetaDataVersion = "9.x"
                        created = tree.find("Esri/CreaDate")
                        modified = tree.find("Esri/ModDate")

                        #Setup variables for textToWriteList and check for valid metadata values
                        desc = arcpy.Describe(fcAndWorkspace)
                        fcPath = os.path.join(subFolderName, workspace, item)
                        fcDataType = desc.dataType
                        fcShapeType = desc.shapeType
                        fcFeatureCount = str(arcpy.GetCount_management(fcAndWorkspace))
                        verifyMetadata (title)
                        fcTitle = tagReturn
                        verifyMetadata (keywords)
                        fcKeywords = tagReturn
                        verifyMetadata (abstract)
                        fcAbstract = tagReturn
                        verifyMetadata (purpose)
                        fcPurpose = tagReturn
                        verifyMetadata (contact)
                        fcContact = tagReturn
                        verifyMetadata (distribution)
                        fcDistribution = tagReturn
                        fcCoordType = desc.SpatialReference.type
                        if fcCoordType == "Unknown":
                            fcGCS = "No Coordinate System Assigned"
                        else:
                            fcGCS = desc.SpatialReference.GCS.name
                        if desc.SpatialReference.PCSname:
                            fcPCS = desc.SpatialReference.PCSname
                            fcUnit = desc.SpatialReference.linearUnitName
                        else:
                            fcPCS = "N/A"
                            fcUnit = "N/A"
                        if str(desc.extent) == "None" or int(fcFeatureCount) == 0 :
                            fcYMax = "No Features"
                            fcXMax = "No Features"
                            fcYMin = "No Features"
                            fcXMin = "No Features"
                        else:
                            fcYMax = str(desc.extent.YMax)
                            fcXMax = str(desc.extent.XMax)
                            fcYMin = str(desc.extent.YMin)
                            fcXMin = str(desc.extent.XMin)
                        verifyMetadata (created)
                        fcCreated = tagReturn
                        verifyMetadata (modified)
                        fcModified = tagReturn

                        #Create list to hold text values for writing to Excel
                        textToWriteList = [excelRow, fc, fcPath, fcDataType, fcShapeType, fcFeatureCount, fcTitle, fcKeywords, fcAbstract,
                                           fcPurpose, fcContact, fcDistribution, fcCoordType, fcGCS, fcPCS, fcUnit, fcYMax, fcXMax, fcYMin,
                                           fcXMin, fcCreated, fcModified, fcMetaDataVersion]

                        #Start writing information on fc and metadata to Excel
                        column = 0
                        for text in textToWriteList:
                            row = ws.row(excelRow)
                            row.write(column, text)
                            column = column + 1

                        #Remove temporary empty XML file
                        os.remove(fcXML)

                        #Increment row counter by 1
                        excelRow = excelRow + 1

                    except:
                        print "     Could not process: " + fc + " at: " + os.path.join(subFolderName, workspace, item) + "; Skipping"
                        arcpy.AddMessage("     Could not process: " + fc + " at: " + os.path.join(subFolderName, workspace, item) + "; Skipping")
                        fcPath = os.path.join(subFolderName, workspace, item)
                        fcDataType = "COULD NOT PROCESS FILE"
                        #Create list to hold text values for writing to Excel
                        textToWriteList = [excelRow, fc, fcPath, fcDataType]

                        #Start writing information on fc and metadata to Excel
                        column = 0
                        for text in textToWriteList:
                            row = ws.row(excelRow)
                            row.write(column, text)
                            column = column + 1

                        #Remove temporary empty XML file
                        try:
                            os.remove(fcXML)
                        except:
                            print "     Could not remove file: " + str(fcXML)
                            arcpy.AddMessage("     Could not remove file: " + str(fcXML))

                        #Increment row counter by 1
                        excelRow = excelRow + 1


            #Processes feature classes in personal geodatbase (not in feature dataset)
            for fc in arcpy.ListFeatureClasses():
                try:
                    #Creates a temporary empty XML file to store each files XML data
                    fcXML = tempOutputFolder + "/" + fc + ".xml"
                    tempXMLFile = open(fcXML, "w")
                    tempXMLFile.write("<metadata>\n")
                    tempXMLFile.write("</metadata>\n")
                    tempXMLFile.close()
                    #update fc to include workspace path
                    fcAndWorkspace = workspace + '/' + fc
                    #Copies metadata from file into temporary XML file
                    arcpy.MetadataImporter_conversion (fcAndWorkspace, fcXML)
                    #Makes an ElementTree object and reads XML into ElementTree
                    tree = ElementTree()
                    tree.parse(fcXML)
                    #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                    if tree.find("Esri/ArcGISFormat") is not None:
                        print "     Processing File: " + fc + " (Version 10.x XML)"
                        arcpy.AddMessage("     Processing File: " + fc + " (Version 10.x XML)")
                        title = tree.find("dataIdInfo/idCitation/resTitle")
                        keywords = tree.find("dataIdInfo/searchKeys/keyword")
                        abstract = tree.find("dataIdInfo/idAbs")
                        purpose = tree.find("dataIdInfo/idPurp")
                        contact = tree.find("mdContact/rpIndName")
                        distribution = tree.find("distInfo/distFormat/formatName")
                        fcMetaDataVersion = "10.x"
                    else:
                        print "     Processing File: " + fc + " (Version 9.x XML)"
                        arcpy.AddMessage ("     Processing File: " + fc + " (Version 9.x XML)")
                        title = tree.find("idinfo/citation/citeinfo/title")
                        keywords = tree.find("idinfo/keywords/theme/themekey")
                        abstract = tree.find("idinfo/descript/abstract")
                        purpose = tree.find("idinfo/descript/purpose")
                        contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                        distribution = tree.find("distinfo/resdesc")
                        fcMetaDataVersion = "9.x"
                    created = tree.find("Esri/CreaDate")
                    modified = tree.find("Esri/ModDate")

                    #Setup variables for textToWriteList and check for valid metadata values
                    desc = arcpy.Describe(fcAndWorkspace)
                    fcPath = os.path.join(subFolderName, workspace)
                    fcDataType = desc.dataType
                    fcShapeType = desc.shapeType
                    fcFeatureCount = str(arcpy.GetCount_management(fcAndWorkspace))
                    verifyMetadata (title)
                    fcTitle = tagReturn
                    verifyMetadata (keywords)
                    fcKeywords = tagReturn
                    verifyMetadata (abstract)
                    fcAbstract = tagReturn
                    verifyMetadata (purpose)
                    fcPurpose = tagReturn
                    verifyMetadata (contact)
                    fcContact = tagReturn
                    verifyMetadata (distribution)
                    fcDistribution = tagReturn
                    fcCoordType = desc.SpatialReference.type
                    if fcCoordType == "Unknown":
                        fcGCS = "No Coordinate System Assigned"
                    else:
                        fcGCS = desc.SpatialReference.GCS.name
                    if desc.SpatialReference.PCSname:
                        fcPCS = desc.SpatialReference.PCSname
                        fcUnit = desc.SpatialReference.linearUnitName
                    else:
                        fcPCS = "N/A"
                        fcUnit = "N/A"
                    if str(desc.extent) == "None" or int(fcFeatureCount) == 0 :
                        fcYMax = "No Features"
                        fcXMax = "No Features"
                        fcYMin = "No Features"
                        fcXMin = "No Features"
                    else:
                        fcYMax = str(desc.extent.YMax)
                        fcXMax = str(desc.extent.XMax)
                        fcYMin = str(desc.extent.YMin)
                        fcXMin = str(desc.extent.XMin)
                    verifyMetadata (created)
                    fcCreated = tagReturn
                    verifyMetadata (modified)
                    fcModified = tagReturn

                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, fc, fcPath, fcDataType, fcShapeType, fcFeatureCount, fcTitle, fcKeywords, fcAbstract,
                                       fcPurpose, fcContact, fcDistribution, fcCoordType, fcGCS, fcPCS, fcUnit, fcYMax, fcXMax, fcYMin,
                                       fcXMin, fcCreated, fcModified, fcMetaDataVersion]

                    #Start writing information on fc and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    os.remove(fcXML)

                    #Increment row counter by 1
                    excelRow = excelRow + 1

                except:
                    print "     Could not process: " + fc + " at: " + os.path.join(subFolderName, workspace) + "; Skipping"
                    arcpy.AddMessage("     Could not process: " + fc + " at: " + os.path.join(subFolderName, workspace) + "; Skipping")
                    fcPath = os.path.join(subFolderName, workspace)
                    fcDataType = "COULD NOT PROCESS FILE"
                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, fc, fcPath, fcDataType]

                    #Start writing information on fc and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    try:
                        os.remove(fcXML)
                    except:
                        print "     Could not remove file: " + str(fcXML)
                        arcpy.AddMessage("     Could not remove file: " + str(fcXML))

                    #Increment row counter by 1
                    excelRow = excelRow + 1


            #Processes tables in personal geodatabase
            for table in tableset:
                try:
                    #Creates a temporary empty XML file to store each files XML data
                    tableXML = tempOutputFolder + "/" + table + ".xml"
                    tempXMLFile = open(tableXML, "w")
                    tempXMLFile.write("<metadata>\n")
                    tempXMLFile.write("</metadata>\n")
                    tempXMLFile.close()
                    #update table to include workspace path
                    tableAndWorkspace = workspace + '/' + table
                    #Copies metadata from file into previous created temporary XML file
                    arcpy.MetadataImporter_conversion (tableAndWorkspace, tableXML)

                    #Makes an ElementTree object and reads XML into ElementTree
                    tree = ElementTree()
                    tree.parse(tableXML)
                    #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                    if tree.find("Esri/ArcGISFormat") is not None:
                        print "     Processing File: " + table + " (Version 10.x XML)"
                        arcpy.AddMessage("     Processing File: " + table + " (Version 10.x XML)")
                        title = tree.find("dataIdInfo/idCitation/resTitle")
                        keywords = tree.find("dataIdInfo/searchKeys/keyword")
                        abstract = tree.find("dataIdInfo/idAbs")
                        purpose = tree.find("dataIdInfo/idPurp")
                        contact = tree.find("mdContact/rpIndName")
                        distribution = tree.find("distInfo/distFormat/formatName")
                        tableMetaDataVersion = "10.x"
                    else:
                        print "     Processing File: " + table + " (Version 9.x XML)"
                        arcpy.AddMessage("     Processing File: " + table + " (Version 9.x XML)")
                        title = tree.find("idinfo/citation/citeinfo/title")
                        keywords = tree.find("idinfo/keywords/theme/themekey")
                        abstract = tree.find("idinfo/descript/abstract")
                        purpose = tree.find("idinfo/descript/purpose")
                        contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                        distribution = tree.find("distinfo/resdesc")
                        tableMetaDataVersion = "9.x"
                    created = tree.find("Esri/CreaDate")
                    modified = tree.find("Esri/ModDate")

                    #Setup variables for textToWriteList and check for valid metadata values
                    desc = arcpy.Describe(tableAndWorkspace)
                    tablePath = os.path.join(subFolderName, workspace)
                    tableDataType = desc.dataType
                    tableShapeType = "N/A"
                    tableFeatureCount = "N/A"
                    verifyMetadata (title)
                    tableTitle = tagReturn
                    verifyMetadata (keywords)
                    tableKeywords = tagReturn
                    verifyMetadata (abstract)
                    tableAbstract = tagReturn
                    verifyMetadata (purpose)
                    tablePurpose = tagReturn
                    verifyMetadata (contact)
                    tableContact = tagReturn
                    verifyMetadata (distribution)
                    tableDistribution = tagReturn
                    tableCoordType = "N/A"
                    tableGCS = "N/A"
                    tablePCS = "N/A"
                    tableUnit = "N/A"
                    tableYMax = "N/A"
                    tableXMax = "N/A"
                    tableYMin = "N/A"
                    tableXMin = "N/A"
                    verifyMetadata (created)
                    tableCreated = tagReturn
                    verifyMetadata (modified)
                    tableModified = tagReturn

                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, table, tablePath, tableDataType, tableShapeType, tableFeatureCount, tableTitle, tableKeywords, tableAbstract,
                                       tablePurpose, tableContact, tableDistribution, tableCoordType, tableGCS, tablePCS, tableUnit, tableYMax, tableXMax, tableYMin,
                                       tableXMin, tableCreated, tableModified, tableMetaDataVersion]

                    #Start writing information on table and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    os.remove(tableXML)

                    #Increment row counter by 1
                    excelRow = excelRow + 1

                except:
                    print "     Could not process: " + table + " at: " + os.path.join(subFolderName, workspace) + "; Skipping"
                    arcpy.AddMessage("     Could not process: " + table + " at: " + os.path.join(subFolderName, workspace) + "; Skipping")
                    tablePath = os.path.join(subFolderName, workspace)
                    tableDataType = "COULD NOT PROCESS FILE"
                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, table, tablePath, tableDataType]

                    #Start writing information on fc and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    try:
                        os.remove(tableXML)
                    except:
                        print "     Could not remove file: " + str(tableXML)
                        arcpy.AddMessage("Could not remove file: " + str(tableXML))

                    #Increment row counter by 1
                    excelRow = excelRow + 1


            #Processes rasters in personal geodatabase
            for raster in rasterset:
                try:
                    #Creates a temporary empty XML file to store each files XML data
                    rasterXML = tempOutputFolder + "/" + raster + ".xml"
                    tempXMLFile = open(rasterXML, "w")
                    tempXMLFile.write("<metadata>\n")
                    tempXMLFile.write("</metadata>\n")
                    tempXMLFile.close()
                    #update raster to include workspace path
                    rasterAndWorkspace = workspace + '/' + raster
                    #Copies metadata from file into previous created temporary XML file
                    arcpy.MetadataImporter_conversion (rasterAndWorkspace, rasterXML)

                    #Makes an ElementTree object and reads XML into ElementTree
                    tree = ElementTree()
                    tree.parse(rasterXML)
                    #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                    if tree.find("Esri/ArcGISFormat") is not None:
                        print "     Processing File: " + raster + " (Version 10.x XML)"
                        arcpy.AddMessage("     Processing File: " + raster + " (Version 10.x XML)")
                        title = tree.find("dataIdInfo/idCitation/resTitle")
                        keywords = tree.find("dataIdInfo/searchKeys/keyword")
                        abstract = tree.find("dataIdInfo/idAbs")
                        purpose = tree.find("dataIdInfo/idPurp")
                        contact = tree.find("mdContact/rpIndName")
                        distribution = tree.find("distInfo/distFormat/formatName")
                        rasterMetaDataVersion = "10.x"
                    else:
                        print "     Processing File: " + raster + " (Version 9.x XML)"
                        arcpy.AddMessage("     Processing File: " + raster + " (Version 9.x XML)")
                        title = tree.find("idinfo/citation/citeinfo/title")
                        keywords = tree.find("idinfo/keywords/theme/themekey")
                        abstract = tree.find("idinfo/descript/abstract")
                        purpose = tree.find("idinfo/descript/purpose")
                        contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                        distribution = tree.find("distinfo/resdesc")
                        rasterMetaDataVersion = "9.x"
                    created = tree.find("Esri/CreaDate")
                    modified = tree.find("Esri/ModDate")

                    #Setup variables for textToWriteList and check for valid metadata values
                    desc = arcpy.Describe(rasterAndWorkspace)
                    rasterPath = os.path.join(subFolderName, workspace)
                    rasterDataType = desc.dataType
                    if rasterDataType == "RasterBand":
                        rasterShapeType = "No information Available"
                        rasterFeatureCount = "No information Available"
                    else:
                        rasterShapeType = desc.format
                        rasterFeatureCount = str(desc.bandCount) + " Bands"
                    verifyMetadata (title)
                    rasterTitle = tagReturn
                    verifyMetadata (keywords)
                    rasterKeywords = tagReturn
                    verifyMetadata (abstract)
                    rasterAbstract = tagReturn
                    verifyMetadata (purpose)
                    rasterPurpose = tagReturn
                    verifyMetadata (contact)
                    rasterContact = tagReturn
                    verifyMetadata (distribution)
                    rasterDistribution = tagReturn
                    rasterCoordType = desc.SpatialReference.type
                    rasterGCS = desc.SpatialReference.GCS.name
                    if desc.SpatialReference.PCSname:
                        rasterPCS = desc.SpatialReference.PCSname
                        rasterUnit = desc.SpatialReference.linearUnitName
                    else:
                        rasterPCS = "N/A"
                        rasterUnit = "N/A"
                    if rasterFeatureCount:
                        rasterYMax = str(desc.extent.YMax)
                        rasterXMax = str(desc.extent.XMax)
                        rasterYMin = str(desc.extent.YMin)
                        rasterXMin = str(desc.extent.XMin)
                    else:
                        rasterYMax = "N/A"
                        rasterXMax = "N/A"
                        rasterYMin = "N/A"
                        rasterXMin = "N/A"
                    verifyMetadata (created)
                    rasterCreated = tagReturn
                    verifyMetadata (modified)
                    rasterModified = tagReturn

                    #Create list to hold text values for writing to Excel
                    textToWriteList = [excelRow, raster, rasterPath, rasterDataType, rasterShapeType, rasterFeatureCount, rasterTitle, rasterKeywords, rasterAbstract,
                                       rasterPurpose, rasterContact, rasterDistribution, rasterCoordType, rasterGCS, rasterPCS, rasterUnit, rasterYMax, rasterXMax, rasterYMin,
                                       rasterXMin, rasterCreated, rasterModified, rasterMetaDataVersion]

                    #Start writing information on raster and metadata to Excel
                    column = 0
                    for text in textToWriteList:
                        row = ws.row(excelRow)
                        row.write(column, text)
                        column = column + 1

                    #Remove temporary empty XML file
                    os.remove(rasterXML)

                    #Increment row counter by 1
                    excelRow = excelRow + 1

                except:
                        print "     Could not process: " + raster + " at: " + os.path.join(subFolderName, workspace) + "; Skipping"
                        arcpy.AddMessage("     Could not process: " + raster + " at: " + os.path.join(subFolderName, workspace) + "; Skipping")
                        rasterPath = os.path.join(subFolderName, workspace)
                        rasterDataType = "COULD NOT PROCESS FILE"
                        #Create list to hold text values for writing to Excel
                        textToWriteList = [excelRow, raster, rasterPath, rasterDataType]

                        #Start writing information on fc and metadata to Excel
                        column = 0
                        for text in textToWriteList:
                            row = ws.row(excelRow)
                            row.write(column, text)
                            column = column + 1

                        #Remove temporary empty XML file
                        try:
                            os.remove(rasterXML)
                        except:
                            print "     Could not remove file: " + str(rasterXML)
                            arcpy.AddMessage("      Could not remove file: " + str(rasterXML))

                        #Increment row counter by 1
                        excelRow = excelRow + 1

        #Checks for rasters in a directory and in file geodatabases
        for raster in rasterList:
            try:
                #Creates a temporary empty XML file to store each files XML data
                rasterXML = tempOutputFolder + "/" + raster + ".xml"
                tempXMLFile = open(rasterXML, "w")
                tempXMLFile.write("<metadata>\n")
                tempXMLFile.write("</metadata>\n")
                tempXMLFile.close()
                #Copies metadata from file into previous created temporary XML file
                arcpy.MetadataImporter_conversion (raster, rasterXML)

                #Makes an ElementTree object and reads XML into ElementTree
                tree = ElementTree()
                tree.parse(rasterXML)
                #Checks on XML version (9.x vs. 10.x) and gathers information on metadata
                if tree.find("Esri/ArcGISFormat") is not None:
                    print "     Processing File: " + raster + " (Version 10.x XML)"
                    arcpy.AddMessage("     Processing File: " + raster + " (Version 10.x XML)")
                    title = tree.find("dataIdInfo/idCitation/resTitle")
                    keywords = tree.find("dataIdInfo/searchKeys/keyword")
                    abstract = tree.find("dataIdInfo/idAbs")
                    purpose = tree.find("dataIdInfo/idPurp")
                    contact = tree.find("mdContact/rpIndName")
                    distribution = tree.find("distInfo/distFormat/formatName")
                    rasterMetaDataVersion = "10.x"
                else:
                    print "     Processing File: " + raster + " (Version 9.x XML)"
                    arcpy.AddMessage("     Processing File: " + raster + " (Version 9.x XML)")
                    title = tree.find("idinfo/citation/citeinfo/title")
                    keywords = tree.find("idinfo/keywords/theme/themekey")
                    abstract = tree.find("idinfo/descript/abstract")
                    purpose = tree.find("idinfo/descript/purpose")
                    contact = tree.find("idinfo/ptcontac/cntinfo/cntperp/cntper")
                    distribution = tree.find("distinfo/resdesc")
                    rasterMetaDataVersion = "9.x"
                created = tree.find("Esri/CreaDate")
                modified = tree.find("Esri/ModDate")

                #Setup variables for textToWriteList and check for valid metadata values
                desc = arcpy.Describe(raster)
                rasterPath = os.path.join(subFolderName, raster)
                rasterDataType = desc.dataType
                if rasterDataType == "RasterBand":
                    rasterShapeType = "No information Available"
                    rasterFeatureCount = "No information Available"
                else:
                    rasterShapeType = desc.format
                    rasterFeatureCount = str(desc.bandCount) + " Bands"
                verifyMetadata (title)
                rasterTitle = tagReturn
                verifyMetadata (keywords)
                rasterKeywords = tagReturn
                verifyMetadata (abstract)
                rasterAbstract = tagReturn
                verifyMetadata (purpose)
                rasterPurpose = tagReturn
                verifyMetadata (contact)
                rasterContact = tagReturn
                verifyMetadata (distribution)
                rasterDistribution = tagReturn
                rasterCoordType = desc.SpatialReference.type
                rasterGCS = desc.SpatialReference.GCS.name
                if desc.SpatialReference.PCSname:
                    rasterPCS = desc.SpatialReference.PCSname
                    rasterUnit = desc.SpatialReference.linearUnitName
                else:
                    rasterPCS = "N/A"
                    rasterUnit = "N/A"
                if rasterFeatureCount:
                    rasterYMax = str(desc.extent.YMax)
                    rasterXMax = str(desc.extent.XMax)
                    rasterYMin = str(desc.extent.YMin)
                    rasterXMin = str(desc.extent.XMin)
                else:
                    rasterYMax = "N/A"
                    rasterXMax = "N/A"
                    rasterYMin = "N/A"
                    rasterXMin = "N/A"
                verifyMetadata (created)
                rasterCreated = tagReturn
                verifyMetadata (modified)
                rasterModified = tagReturn

                #Create list to hold text values for writing to Excel
                textToWriteList = [excelRow, raster, rasterPath, rasterDataType, rasterShapeType, rasterFeatureCount, rasterTitle, rasterKeywords, rasterAbstract,
                                   rasterPurpose, rasterContact, rasterDistribution, rasterCoordType, rasterGCS, rasterPCS, rasterUnit, rasterYMax, rasterXMax, rasterYMin,
                                   rasterXMin, rasterCreated, rasterModified, rasterMetaDataVersion]

                #Start writing information on raster and metadata to Excel
                column = 0
                for text in textToWriteList:
                    row = ws.row(excelRow)
                    row.write(column, text)
                    column = column + 1

                #Remove temporary empty XML file
                os.remove(rasterXML)

                #Increment row counter by 1
                excelRow = excelRow + 1

            except:
                print "     Could not process: " + raster + " at: " + subFolderName + "; Skipping"
                arcpy.AddMessage("     Could not process: " + raster + " at: " + subFolderName + "; Skipping")
                rasterPath = os.path.join(subFolderName, dataset)
                rasterDataType = "COULD NOT PROCESS FILE"
                #Create list to hold text values for writing to Excel
                textToWriteList = [excelRow, raster, rasterPath, rasterDataType]

                #Start writing information on fc and metadata to Excel
                column = 0
                for text in textToWriteList:
                    row = ws.row(excelRow)
                    row.write(column, text)
                    column = column + 1

                #Remove temporary empty XML file
                try:
                    os.remove(rasterXML)
                except:
                    print "     Could not remove file: " + str(rasterXML)
                    arcpy.AddMessage("      Could not remove file: " + str(rasterXML))

                #Increment row counter by 1
                excelRow = excelRow + 1


    #Saves output
    wb.save(outputFile)

except Exception, e:
  import traceback
  print e.message
  map(arcpy.AddError, traceback.format_exc().split("\n"))
  arcpy.AddError(str(e))
  arcpy.AddMessage("This tool encountered a serious error and could not finish.")