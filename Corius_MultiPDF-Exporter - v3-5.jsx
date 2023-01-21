//*******************************************************
// MultiPDF-Exporter.jsx
// Version 3.5
//
// Copyright 2022 Corius
// Comments or suggestions to contact@corius.fr
//
//*******************************************************

// Script version
var ExportScriptVersion = 'v3.5';
var ProfileScriptVersion = 'v1.0';

// AI document variables
var docObj = app.activeDocument;
var docName = docObj.fullName;
var docFolder = docName.parent.fsName;

// other files variables
var CoriusLogo;
var ProfileXMLFile;

// profile management variables
var newProfileXML = new XML("<profile><Name/><PDF/><outlineTXT/><prefix1_txt/><prefix1_custom/><prefix2_txt/><prefix2_custom/><prefix3_txt/><prefix3_custom/><suffix1_txt/><suffix1_custom/><suffix2_txt/><suffix2_custom/><suffix3_txt/><suffix3_custom/><savePath/></profile>");
var currentProfileXML = new XML("<profile><Name/><PDF/><outlineTXT/><prefix1_txt/><prefix1_custom/><prefix2_txt/><prefix2_custom/><prefix3_txt/><prefix3_custom/><suffix1_txt/><suffix1_custom/><suffix2_txt/><suffix2_custom/><suffix3_txt/><suffix3_custom/><savePath/></profile>");
var ProfileList;                  
var profileXML;
var profileNames = [];
var updateListNeeded = false;
var numberPrefSuff = 5; // to set the number of prefix/suffix components in the profile manager window

// export preparation variables
var txtFromSymbol =[];  // to store all texts contained in symbol. array of pairs like this [SymbolName, SymbolText]
var loadProfilePref = false; // to keep track of Prefix being loaded from selected export profile
var ignoreProfilePref = []; // to keep track of Prefix being overwritten by user before export. List is [StatusPrefixForExport_1, StatusPrefixForExport_2, …]
var loadProfileSuff = false;  // same as above, for Suffix
var ignoreProfileSuff = [];   // same as above, for Suffix

// export  process variables
var selectedProfileName = '';
var ArtboardToExportList = [];
var DatasetToExportList = [];
var RegularExportData = [];
var VectorExportData = [];

// dataset variables
var DatasetExists = false;
var currentDataset = 0;

// GUI variables
var exportConfigNumber = 1; //Quantity of export profile displayed
var blockH = 120;  // à voir si je conserve cette var globale
var blockW = 230; // à voir si je conserve cette var globale
// Exporter Window
this.ExporterWindow = new Window('dialog',"Corius' MultiPDF-Exporter  -  ["+ExportScriptVersion+"]");
var topGRP = this.ExporterWindow.add('group', undefined, '');
var topLeftGRP = topGRP.add('group', undefined, '');
var imgPanel = topLeftGRP.add('group', undefined, '');
var topCenterGRP = topGRP.add('group', undefined, '');
var managerBtn = topCenterGRP.add('button', undefined, 'OPEN PROFILE MANAGER', {name:'manager'});
var NbrExportGRP = topCenterGRP.add('group', undefined, '');
var topRightGRP = topGRP.add('group', undefined, '');
var DataSetNamingGRP = topRightGRP.add('group', undefined, '');
var DatasetNameExample = topRightGRP.add('statictext', undefined, ''); 
var middleGRP = this.ExporterWindow.add('group', undefined, '');
var PDFExpMainGRP = middleGRP.add('group', undefined, '');
var selectListGRP = middleGRP.add('group', undefined, '');
var ButtonGRP = this.ExporterWindow.add('group', undefined, '');
var alertFilename = ButtonGRP.add('statictext', undefined, "Tip: if the save path starts with '\\' the files will be saved in a subfolder of current Illustrator file"); 
// Profile manager window
this.ProfileWindow = new Window('dialog',"Corius' MultiPDF-Exporter  ---  Profile Manager  -  ["+ProfileScriptVersion+"]");
var ProfWinTopGRP = this.ProfileWindow.add('group', undefined, '');
var ProfWinTopLeftGRP = ProfWinTopGRP.add('group', undefined, '');
var ProfWinProfileGRP = ProfWinTopLeftGRP.add('group', undefined, '');
var ProfWinProfileNameGRP = ProfWinTopLeftGRP.add('group', undefined, '');
var ProfWinImgPanel = ProfWinTopGRP.add('panel', undefined, '');
var ProfWinConfigGRP = this.ProfileWindow.add('panel', undefined, '[Profile configuration]');
var configProfWinTopGRP = ProfWinConfigGRP.add('group', undefined, '');
var ProfWinPrefixGRP = ProfWinConfigGRP.add('panel', undefined, '[Prefix]');
var ProfWinSuffixGRP = ProfWinConfigGRP.add('panel', undefined, '[Suffix]');
var ProfWinPathGRP = ProfWinConfigGRP.add('panel', undefined, '[PDF Export Path]');
var ProfWinBtnGRP = this.ProfileWindow.add('group', undefined, '');


///////////////// Precheck function
function checkStatusBeforeRun(){
    if (docObj.saved == false){
        var myTxt = "Please save the document before running the script";
        popAlertWindow(myTxt,1,'save')
    } else {
        startPreparation();
    }
}




////////////// Preparation functions
function startPreparation(){
    docObj.selection = null;
    
    getTextFromAllSymbols();
    docObj.selection = null;
    
    //check if .AI file uses datasets
    if (docObj.dataSets.length > 0){
        DatasetExists = true;
    }

    checkLogoFileExists();
    
    checkProfileFileExists();
    
    prepareGUI();
}

function getTextFromAllSymbols(){
    var symbName = '';
    var symbString = '';
    var tempoLayer = docObj.layers.add();
    tempoLayer.name = 'MultiPDF-Export_temporary';
    var grpItem;
    var txtFrm;
    var i=0;

    if(docObj.symbols.length >0){
        var mysymbol = docObj.symbolItems.add(docObj.symbols[0]);
        for (i=0;i<docObj.symbols.length;i++){
            symbName = docObj.symbols[i].name;
            mysymbol = docObj.symbolItems.add(docObj.symbols[i]);
            mysymbol.breakLink();
            
            grpItem  = tempoLayer.pageItems[0];
            if (grpItem.groupItems){
                while(grpItem.groupItems.length > 0){
                   grpItem = grpItem.groupItems[0];
                }
            }
            if (grpItem.textFrames && grpItem.textFrames.length > 0){
                txtFrm = grpItem.textFrames[0];
                symbString = txtFrm.contents;
                if (symbString != ''){
                    txtFromSymbol.push([symbName, symbString]);
                }
            }
        }
    
        tempoLayer.remove();
    }
}

function checkLogoFileExists(){
    try{
        CoriusLogo = ScriptUI.newImage(Folder.userData.absoluteURI + "/CoriusScripts/CoriusLogo.png");
    } catch(e){
        myFileMaker();
        CoriusLogo = ScriptUI.newImage(Folder.userData.absoluteURI + "/CoriusScripts/CoriusLogo.png");
    }    
}

function checkProfileFileExists(){
    ProfileXMLFile = new File( Folder.userData.absoluteURI + "/CoriusScripts/Corius_PDF_Profiles.xml" );

    if (ProfileXMLFile.exists){
        ProfileXMLFile.open( "r" );
        ProfileList = new XML(ProfileXMLFile.read());
        ProfileXMLFile.close();
        currentProfileXML = ProfileList.child("profile")[0];
    } else {
        ProfileList = new XML("<profileList>"+currentProfileXML+"</profileList>");
        ProfileXMLFile.open( "w" );
        ProfileXMLFile.write( ProfileList.toString());
        ProfileXMLFile.close();
    }
                                     
    profileXML = ProfileList.child("profile");
    for (var i=0; i<profileXML.length(); ++i){
        if (profileXML[i].child("Name").toString() != '' && profileXML[i].child("Name").toString() != null){
            profileNames.push(profileXML[i].child("Name").toString());
        } else if (profileXML.length() == 1){
             profileNames.push('New');
        }
    }

}

///////////////////////////////// GUI functions
function prepareGUI(){
    prepareExporterWindow();
    
    prepareProfileManagerWindow();
    
    launchExporterWindow();
}

function prepareExporterWindow(){    
    // initialize Export window group containers
    // TOP GROUP
    topGRP.orientation = 'row';
    topGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP];
    //TOP LEFT
    topLeftGRP.orientation = 'column';
    topLeftGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    imgPanel.orientation = 'row';
    imgPanel.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    //TOP CENTER
    topCenterGRP.orientation = 'column';
    topCenterGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.BOTTOM];
    managerBtn.onClick = function() { createProfileManagerWindow() };
    NbrExportGRP.orientation = 'row';
    NbrExportGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.BOTTOM];
    //TOP RIGHT
    topRightGRP.orientation = 'column';
    topRightGRP.alignment = [ScriptUI.Alignment.RIGHT, ScriptUI.Alignment.TOP];
    DataSetNamingGRP.orientation = 'row';
    DataSetNamingGRP.alignment = [ScriptUI.Alignment.RIGHT, ScriptUI.Alignment.TOP];
    DatasetNameExample.size = [ 500,20 ]; 
    DatasetNameExample.alignment = [ScriptUI.Alignment.RIGHT, ScriptUI.Alignment.TOP];
    
    //MIDDLE GROUP
    middleGRP.orientation = 'row';
    //MIDDLE LEFT
    PDFExpMainGRP.orientation = 'column';
    PDFExpMainGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.TOP]
    //MIDDLE RIGHT
    selectListGRP.orientation = 'row';

    //BUTTON GROUP
    ButtonGRP.orientation = 'row';
    ButtonGRP.alignment = [ScriptUI.Alignment.RIGHT, ScriptUI.Alignment.BOTTOM]
    alertFilename.size = [ 800,20 ];
}

function prepareProfileManagerWindow(){
    //TOP GROUP
    ProfWinTopGRP.orientation = 'row';
    ProfWinTopGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP];
    ProfWinTopLeftGRP.orientation = 'column';
    ProfWinTopLeftGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    ProfWinProfileGRP.orientation = 'row';
    ProfWinProfileGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    ProfWinProfileNameGRP.orientation = 'row';
    ProfWinProfileNameGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    ProfWinImgPanel.orientation = 'row';
    ProfWinImgPanel.alignment = [ScriptUI.Alignment.RIGHT, ScriptUI.Alignment.TOP];
    ProfWinConfigGRP.orientation = 'column';
    ProfWinConfigGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    configProfWinTopGRP.orientation = 'row';
    configProfWinTopGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP];
    ProfWinPrefixGRP.orientation = 'row';
    ProfWinPrefixGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    ProfWinSuffixGRP.orientation = 'row';
    ProfWinSuffixGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    ProfWinPathGRP.orientation = 'row';
    ProfWinPathGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP];
    ProfWinBtnGRP.orientation = 'row';
    ProfWinBtnGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.TOP];
}


/////////// EXPORTER WINDOW GUI ///////////////////////
function launchExporterWindow(){
    // ARTBOARD SELECTION    
    var artboardSelectList = selectListGRP.add('listbox', [0, 0, 33 + blockW, blockH], undefined,{multiselect:true, numberOfColumns: 2, showHeaders: true, columnTitles:['#','Artboard name'], columnWidths:[30, blockW]});    
    var artboard;
    var myitem;
    for(var i=0; i<docObj.artboards.length; i++){
        artboard = docObj.artboards[i];
        myItem = artboardSelectList.add('item', i+1);
        myItem.subItems[0].text = artboard.name;
    }
    artboardSelectList.selection = 0;
    
    // DATASET SELECTION
    if (DatasetExists) {
        var datasetSelectList = selectListGRP.add('listbox', [0, 0, 33 + blockW, blockH], undefined,{multiselect:true, numberOfColumns: 2, showHeaders: true, columnTitles:['#','Dataset name'], columnWidths:[30, blockW]});    
        var dataset;
        for(var i=0; i<docObj.dataSets.length; i++){
            dataset = docObj.dataSets[i];
            myItem = datasetSelectList.add('item', i+1);
            myItem.subItems[0].text = dataset.name;
        }
        datasetSelectList.selection = 0;
        
        // NAMING CHOICES FOR DATASET EXPORT
        var namingLabel = DataSetNamingGRP.add('statictext', undefined, 'Naming pattern'); 
        namingLabel.size = [ 90,20 ];
        var namingList;
        if (docObj.artboards.length > 1){
            namingList = DataSetNamingGRP.add('dropdownlist',undefined,['Artboard[TXT]Dataset', 'Dataset[TXT]Artboard']); 
        } else {
            namingList = DataSetNamingGRP.add('dropdownlist',undefined,['Artboard[TXT]Dataset', 'Dataset[TXT]Artboard', 'Dataset']);
        }
        namingList.size = [ 150,20 ];
        namingList.selection = 0;
        namingList.onChange = updateDatasetExample;
        var separLabel = DataSetNamingGRP.add('statictext', undefined, 'Separator TXT'); 
        separLabel.size = [ 83,20 ];        
        var separTxt = DataSetNamingGRP.add('edittext', undefined, ''); 
        separTxt.size = [ 50,20 ];
        separTxt.onChanging = updateDatasetExample;
        
        updateDatasetExample();
    }

    // CORIUS LOGO
    var logoIMG = imgPanel.add('image',undefined, CoriusLogo);    

    // EXPORT SETTINGS
    var exportNumberLabel = NbrExportGRP.add('statictext', undefined, 'How many export batch ?'); 
    exportNumberLabel.size = [ 185,20 ];
    var exportNumberList = NbrExportGRP.add('dropdownlist',undefined,[1, 2, 3, 4, 5]); 
    exportNumberList.size = [ 45,20 ];
    exportNumberList.selection = 0;
    
    exportNumberList.onChange = updateExportConfigGRP;
    exportNumberList.onActivate = updateExportConfigDropDnList;
    
    // EXPORT MAIN GROUP
    updateExportConfigGRP();
    
    // BUTTONS ROW
    var cancelBtn = ButtonGRP.add('button', undefined, 'Cancel', {name:'cancel'});
    cancelBtn.onClick = function() { revertAndAbort() };
    var exportBtn = ButtonGRP.add('button', undefined, 'Export', {name:'export'});
    exportBtn.onClick = function() { checkSettingsBeforeExport() };
    
    
    //-------------------------------------------------------- SHOW THE EXPORTER WINDOW -----------------------------------------//
    ExporterWindow.spacing = 10;
    ExporterWindow.layout.layout(true);

    ExporterWindow.show();
}

/////////// PROFILE MGR WINDOW GUI ///////////////////////

function createProfileManagerWindow(){
    // CLEAN START
    ProfileWindow = new Window('dialog',"Corius' MultiPDF-Exporter  ---  Profile Manager  -  ["+ProfileScriptVersion+"]");
    ProfileWindow.preferredSize = [-1,-1];
    
    // initialize Profile Window group containers
    ProfWinTopGRP = ProfileWindow.add('group', undefined, '')
    ProfWinTopGRP.orientation = 'row';
    ProfWinTopGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP]
    ProfWinTopLeftGRP = ProfWinTopGRP.add('group', undefined, '')
    ProfWinTopLeftGRP.orientation = 'column';
    ProfWinTopLeftGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP]
    ProfWinProfileGRP = ProfWinTopLeftGRP.add('group', undefined, '')
    ProfWinProfileGRP.orientation = 'row';
    ProfWinProfileGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP]
    ProfWinProfileNameGRP = ProfWinTopLeftGRP.add('group', undefined, '')
    ProfWinProfileNameGRP.orientation = 'row';
    ProfWinProfileNameGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP]
    ProfWinImgPanel = ProfWinTopGRP.add('panel', undefined, '')
    ProfWinImgPanel.orientation = 'row';
    ProfWinImgPanel.alignment = [ScriptUI.Alignment.RIGHT, ScriptUI.Alignment.TOP]

    ProfWinConfigGRP = ProfileWindow.add('panel', undefined, '[Profile configuration]')
    ProfWinConfigGRP.orientation = 'column';
    ProfWinConfigGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP]
    configProfWinTopGRP = ProfWinConfigGRP.add('group', undefined, '')
    configProfWinTopGRP.orientation = 'row';
    configProfWinTopGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP]
    ProfWinPrefixGRP = ProfWinConfigGRP.add('panel', undefined, '[Prefix]')
    ProfWinPrefixGRP.orientation = 'row';
    ProfWinPrefixGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP]
    ProfWinSuffixGRP = ProfWinConfigGRP.add('panel', undefined, '[Suffix]')
    ProfWinSuffixGRP.orientation = 'row';
    ProfWinSuffixGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP]
    ProfWinPathGRP = ProfWinConfigGRP.add('panel', undefined, '[PDF Export Path]')
    ProfWinPathGRP.orientation = 'row';
    ProfWinPathGRP.alignment = [ScriptUI.Alignment.FILL, ScriptUI.Alignment.TOP]


    ProfWinBtnGRP = ProfileWindow.add('group', undefined, '')
    ProfWinBtnGRP.orientation = 'row';
    ProfWinBtnGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.TOP]
     
    // Top part    
    var ProfileLabel = ProfWinProfileGRP.add('statictext', undefined, 'Select Profile:'); 
    ProfileLabel.size = [ 100,20 ];
    
    updateProfileList(0);
    
    //Corius Logo
    var logoIMG = ProfWinTopGRP.add('image',undefined, CoriusLogo);

    //Name Group    
    var nameLabel = ProfWinProfileNameGRP.add('statictext', undefined, 'Profile name:'); 
    nameLabel.size = [ 100,20 ];
    var nameTxt = ProfWinProfileNameGRP.add('edittext', undefined, ''); 
    nameTxt.size = [ 200,20 ];                             

    //Config Group
    var presetLabel = configProfWinTopGRP.add('statictext', undefined, 'PDF preset:'); 
    presetLabel.size = [ 68,20 ];
    var presetDropList = configProfWinTopGRP.add('dropdownlist',undefined,app.PDFPresetsList);
    var vectorChkBx = configProfWinTopGRP.add('checkbox',undefined,'Vectorize all texts');
    vectorChkBx.value = false;
    
    // PREFIX AND SUFFIX
    var myPrefSuffGRP;
    var myPrefSuffRadio1;
    var myPrefSuffRadio2;
    var myPrefSuffRadio3;
    var myPrefSuffTxt;
    var myPrefSuffPlusLabel;
    var i=0;
    
    //Prefix Group
    for (i = 0; i < numberPrefSuff; i++){
        myPrefSuffGRP= ProfWinPrefixGRP.add('group', undefined, '');
        myPrefSuffGRP.orientation = 'column';
        myPrefSuffGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
        myPrefSuffRadio1 = myPrefSuffGRP.add('radiobutton', undefined, 'Custom text');
        myPrefSuffRadio2 = myPrefSuffGRP.add('radiobutton', undefined, 'Text in symbol (or artboard)');
        myPrefSuffRadio3 = myPrefSuffGRP.add('radiobutton', undefined, 'Text in artboard (or symbol)');
        myPrefSuffRadio1.value = true;
        myPrefSuffTxt = myPrefSuffGRP.add('edittext', undefined, ''); 
        myPrefSuffTxt.size = [ 200,20 ]; 
        
        //add the "+"  before next component of prefix
        if (i < numberPrefSuff - 1){
            myPrefSuffPlusLabel = ProfWinPrefixGRP.add('statictext', undefined, '[ + ]'); 
            myPrefSuffPlusLabel.size = [ 29,20 ];
        }
    }
    
    //Suffix Group
    for (i = 0; i < numberPrefSuff; i++){
        myPrefSuffGRP= ProfWinSuffixGRP.add('group', undefined, '');
        myPrefSuffGRP.orientation = 'column';
        myPrefSuffGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
        myPrefSuffRadio1 = myPrefSuffGRP.add('radiobutton', undefined, 'Custom text');
        myPrefSuffRadio2 = myPrefSuffGRP.add('radiobutton', undefined, 'Text in symbol (or artboard)');
        myPrefSuffRadio3 = myPrefSuffGRP.add('radiobutton', undefined, 'Text in artboard (or symbol)');
        myPrefSuffRadio1.value = true;
        myPrefSuffTxt = myPrefSuffGRP.add('edittext', undefined, ''); 
        myPrefSuffTxt.size = [ 200,20 ]; 
        
        //add the "+"  before next component of prefix
        if (i < numberPrefSuff - 1){
            myPrefSuffPlusLabel = ProfWinSuffixGRP.add('statictext', undefined, '[ + ]'); 
            myPrefSuffPlusLabel.size = [ 29,20 ];
        }
    }
    
    // Save Path Group
    var FolderTxt = ProfWinPathGRP.add('edittext', undefined, ''); 
    FolderTxt.size = [ 600,20 ];
    FolderTxt.text = '';    
    var BrowseProfBtn = ProfWinPathGRP.add('button', undefined, 'Browse...', {name:'BrowseBtn'}); 
	BrowseProfBtn.onClick = function() { FolderTxt.text = Folder.selectDialog().fsName; 
                                                             if (FolderTxt.text == 'null'){
                                                                FolderTxt.text = docFolder;
                                                             }
                                                          }
    
    // Buttons Group
    var deleteProfBtn = ProfWinBtnGRP.add('button', undefined, 'Delete profile', {name:'delete'});
    deleteProfBtn.onClick = function() { deleteProfile() };
    var cancelProfBtn = ProfWinBtnGRP.add('button', undefined, 'Cancel', {name:'cancel'});
    cancelProfBtn.onClick = function() { ProfileWindow.hide() };
    var saveProfBtn = ProfWinBtnGRP.add('button', undefined, 'Save profile', {name:'save'});
    saveProfBtn.onClick = function() { saveProfile() };
    
    updateProfileWindow();
    
    ProfileWindow.layout.layout(true);
    ProfileWindow.show();
}

function updateProfileWindow(){
    selectedProfileName = ProfWinProfileGRP.children[1].selection.text;
    if (selectedProfileName == "New"){
        currentProfileXML = newProfileXML;
    } else {
        currentProfileXML = ProfileList.child("profile")[ProfWinProfileGRP.children[1].selection.index];
    }
    ProfWinProfileNameGRP.children[1].text = currentProfileXML.child("Name").toString();
    
    // PDF preset
    var pdfsetting = currentProfileXML.child("PDF").toString();
    
    var j=0;
    for (var i=0;i<app.PDFPresetsList.length && j==0;i++){
        if (app.PDFPresetsList[i].toString()==pdfsetting){
            j = 1;
        }
    } 
    if (j == 1){
        configProfWinTopGRP.children[1].selection = i-1;
    } else {
        configProfWinTopGRP.children[1].selection = null;
    }
    
    //Outline text checkbox
    if (currentProfileXML.child("outlineTXT").toString() == "true"){
        configProfWinTopGRP.children[2].value = true;
    } else {
        configProfWinTopGRP.children[2].value = false;
    }
     
    //PREFIX 
    var k = 0;
    for (k=0; k < numberPrefSuff; k++){
        if (currentProfileXML.child("prefix"+(k+1)+"_custom").toString() == "artboardText"){
            ProfWinPrefixGRP.children[k*2].children[2].value = true;
        } else if (currentProfileXML.child("prefix"+(k+1)+"_custom").toString() == "symbol"){
            ProfWinPrefixGRP.children[k*2].children[1].value = true;
        } else {
            ProfWinPrefixGRP.children[k*2].children[0].value = true;
        }
        ProfWinPrefixGRP.children[k*2].children[3].text = currentProfileXML.child("prefix"+(k+1)+"_txt").toString();
    }
     
    //SUFFIX 
    for (k=0; k < numberPrefSuff ; k++){
        if (currentProfileXML.child("suffix"+(k+1)+"_custom").toString() == "artboardText"){
            ProfWinSuffixGRP.children[k*2].children[2].value = true;
        } else if (currentProfileXML.child("suffix"+(k+1)+"_custom").toString() == "symbol"){
            ProfWinSuffixGRP.children[k*2].children[1].value = true;
        } else {
            ProfWinSuffixGRP.children[k*2].children[0].value = true;
        }
        ProfWinSuffixGRP.children[k*2].children[3].text = currentProfileXML.child("suffix"+(k+1)+"_txt").toString();
    }
     
    //Export Path
    ProfWinPathGRP.children[0].text = currentProfileXML.child("savePath").toString();
    
}

function updateProfileList(selectIndex){
    ProfileXMLFile.open( "r" );
    ProfileList = new XML(ProfileXMLFile.read());
    ProfileXMLFile.close();
    profileXML = ProfileList.child("profile");
    profileNames = [];
    
    updateListNeeded = true;
    
    if (ProfWinProfileGRP.children.length == 2){
        ProfWinProfileGRP.remove(1);
    }

    for (var i=0; i<profileXML.length(); ++i){
        if (profileXML[i].child("Name").toString() != '' && profileXML[i].child("Name").toString() != null){
            profileNames.push(profileXML[i].child("Name").toString());
        }
    }

    profileNames.push('New');
    
    var profileSelectList = ProfWinProfileGRP.add('dropdownlist', undefined, profileNames);
    profileSelectList.selection = selectIndex;
    
    profileSelectList.onChange = updateProfileWindow;
    
    ProfileWindow.layout.layout(true);
}

///////////////////////////////// EXPORTER WINDOW FUNCTIONS
function updateExportConfigGRP(){
    var mygrpHeight = 0;
    exportConfigNumber = NbrExportGRP.children[1].selection.index + 1;
    if (exportConfigNumber > PDFExpMainGRP.children.length) {
        for (var i=PDFExpMainGRP.children.length; i<exportConfigNumber;i++){
            if(exportConfigNumber != 1){
                mygrpHeight = PDFExpMainGRP.children[0].size.height;
                PDFExpMainGRP.size.height = PDFExpMainGRP.size.height + mygrpHeight;
                ExporterWindow.size.height = ExporterWindow.size.height + mygrpHeight;
            }
            createProfileConfigGRP(i+1);
            ExporterWindow.layout.layout(true);
        }
    }
    
    if (exportConfigNumber < PDFExpMainGRP.children.length) {
        for (var i=PDFExpMainGRP.children.length; i>exportConfigNumber;i--){
            PDFExpMainGRP.remove(i-1);
            ignoreProfilePref.pop() ;
            ignoreProfileSuff.pop() ;
        }
    }

    selectListGRP.children[0].size.height = exportConfigNumber * PDFExpMainGRP.children[0].size.height + (exportConfigNumber - 1) * 10;
    selectListGRP.children[selectListGRP.children.length - 1].size.height = exportConfigNumber * PDFExpMainGRP.children[0].size.height + (exportConfigNumber - 1) * 10;
    
    ExporterWindow.layout.layout(true);
}

function updateExportConfigDropDnList(){
    if (updateListNeeded){
        for (var i=0;i<PDFExpMainGRP.children.length;i++){
            var mySelector = PDFExpMainGRP.children[i].children[0].children[0].children[1]; 
            if (mySelector.selection == null){
                var selectedIndex = null;
            } else {
                var selectedIndex = mySelector.selection.index;
            }
            updateSelectorList(mySelector, selectedIndex);   
        }
        updateListNeeded = false;
    }
}

function updateSelectorList(mySelector, selectIndex){
    mySelector.removeAll();
    for (var i=0; i<profileNames.length;i++){
        mySelector.add('item', profileNames[i]);
    }
    mySelector.selection = selectIndex;
}

function createProfileConfigGRP(num){
    var configGRP = PDFExpMainGRP.add('panel', undefined, 'Export ' + num);
    configGRP.orientation = 'column';
    configGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    
    ignoreProfilePref[num-1] = false;
    ignoreProfileSuff[num-1] = false;
    
    // GRP first row
    var paramsGRP = configGRP.add('group', undefined, '');
    paramsGRP.orientation = 'row';
    paramsGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.CENTER];
    // Child 1
    var profileSelectGRP = paramsGRP.add('group', undefined, '');
    profileSelectGRP.orientation = 'row';
    profileSelectGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.CENTER]; 
    var profileSelectLabel = profileSelectGRP.add('statictext', undefined, 'Export Profile:'); 
    profileSelectLabel.size = [ 83,20 ];
    var profileSelectDropList = profileSelectGRP.add('dropdownlist',undefined,profileNames);
    profileSelectDropList.onChange = updateExportSettings;
    profileSelectDropList.onActivate = updateExportConfigDropDnList;
    // Child 2
    var presetGRP = paramsGRP.add('group', undefined, '');
    presetGRP.orientation = 'row';
    presetGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.CENTER];  
    var presetLabel = presetGRP.add('statictext', undefined, 'PDF preset:'); 
    presetLabel.size = [ 68,20 ];
    var presetDropList = presetGRP.add('dropdownlist',undefined,app.PDFPresetsList);
    // Child 3
    var vectorChkBx = paramsGRP.add('checkbox',undefined,'Vectorize all texts');
    vectorChkBx.value = false;
    
    // GRP second row
    var prefsuffGRP = configGRP.add('group', undefined, '');
    prefsuffGRP.orientation = 'row';
    prefsuffGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.CENTER];
    
    var prefGRP = prefsuffGRP.add('group', undefined, '');
    prefGRP.orientation = 'row';
    prefGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.CENTER];    
    var prefLabel = prefGRP.add('statictext', undefined, 'Prefix:'); 
    prefLabel.size = [ 41,20 ];
    var prefTxt = prefGRP.add('edittext', undefined, ''); 
    prefTxt.size = [ 260,20 ];
    prefTxt.onChanging = launchUpdateExample;
    
    var suffGRP = prefsuffGRP.add('group', undefined, '');
    suffGRP.orientation = 'row';
    suffGRP.alignment = [ScriptUI.Alignment.CENTER, ScriptUI.Alignment.CENTER];    
    var suffLabel = suffGRP.add('statictext', undefined, 'Suffix:'); 
    suffLabel.size = [ 40,20 ];
    var suffTxt = suffGRP.add('edittext', undefined, ''); 
    suffTxt.size = [ 260,20 ];
    suffTxt.onChanging = launchUpdateExample;
    
    // GRP third row
    var testFilenameGRP = configGRP.add('group', undefined, '');
    testFilenameGRP.orientation = 'row';
    testFilenameGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
    var exampleLabel = testFilenameGRP.add('statictext', undefined, 'Example:'); 
    exampleLabel.size = [ 54,20 ];
    var exampleFilename = testFilenameGRP.add('statictext', undefined, ''); 
    exampleFilename.size = [ 320,20 ];
    var fileCreationMethod = testFilenameGRP.add('dropdownlist',undefined,['Multiple file (1 per artboard)', 'Single file (all artboards in 1 PDF file)']);
    fileCreationMethod.size = [ 240,20 ];
    fileCreationMethod.selection = 0;
    
    //GRP fourth row
    var folderGRP = configGRP.add('group', undefined, '');
    folderGRP.orientation = 'row';   
    folderGRP.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];    
    var FolderLabel = folderGRP.add('statictext', undefined, 'Folder:'); 
    FolderLabel.size = [ 41,20 ];
    var FolderTxt = folderGRP.add('edittext', undefined, ''); 
    FolderTxt.size = [ 500,20 ];
    FolderTxt.text = docFolder;    
    var BrowseBtn = folderGRP.add('button', undefined, 'Browse...', {name:'BrowseBtn'}); 
	BrowseBtn.onClick = function() { FolderTxt.text = Folder.selectDialog().fsName; 
                                                        if (FolderTxt.text == 'null'){
                                                                FolderTxt.text = docFolder;
                                                        }
                                                     }
    
   
}

function launchUpdateExample(){
    var exportGRP = this.parent.parent.parent;
    
    if (this.parent.children[0].text == 'Prefix:'){       
        if (loadProfilePref == true){
            loadProfilePref = false;
            for (var i=0; i<PDFExpMainGRP.children.length;i++){
                if (PDFExpMainGRP.children[i] == exportGRP){
                    ignoreProfilePref[i] = false;
                }
            }
        } else {
            for (var i=0; i<PDFExpMainGRP.children.length;i++){
                if (PDFExpMainGRP.children[i] == exportGRP){
                    ignoreProfilePref[i] = true;
                }
            }
        }
    } else {
        if (loadProfileSuff == true){
            loadProfileSuff = false;
            for (var i=0; i<PDFExpMainGRP.children.length;i++){
                if (PDFExpMainGRP.children[i] == exportGRP){
                    ignoreProfileSuff[i] = false;
                }
            }
        } else {
            for (var i=0; i<PDFExpMainGRP.children.length;i++){
                if (PDFExpMainGRP.children[i] == exportGRP){
                    ignoreProfileSuff[i] = true;
                }
            }
        }
    }
    
    updateExample(exportGRP);
}

function updateExample(exportGRP){
    var exampleFileName = '';
    
    if (selectListGRP.children[0].selection && selectListGRP.children[0].selection.length > 0){
        var myItem = selectListGRP.children[0].selection[0];
        exampleFileName = exportGRP.children[1].children[0].children[1].text;
        exampleFileName += myItem.subItems[0].text;
        exampleFileName += exportGRP.children[1].children[1].children[1].text;
    } else {
        exampleFileName = exportGRP.children[1].children[0].children[1].text;
        exampleFileName += docObj.artboards[0].name;
        exampleFileName += exportGRP.children[1].children[1].children[1].text;
    }

    exampleFileName = cleanPrefSuff(exampleFileName,'verbose');

    exportGRP.children[2].children[1].text = exampleFileName;
}

function updateDatasetExample(){
    var tempname = '';
    
    if (DataSetNamingGRP.children[1].selection.index == 0){
        tempName = docObj.artboards[0].name+DataSetNamingGRP.children[3].text+docObj.dataSets[0].name;
    } else if (DataSetNamingGRP.children[1].selection.index == 1){
        tempName = docObj.dataSets[0].name+DataSetNamingGRP.children[3].text+docObj.artboards[0].name;
    } else {
        tempName = docObj.dataSets[0].name;
    }
    
    tempName = cleanPrefSuff(tempName,'verbose');
    
    DatasetNameExample.text = 'Example: ' + tempName;
}

function fillPrefixSuffix(myProfileXML, mystring, artboardName){
    var mystringtxt = mystring + '_txt';
    var mystringcustom = mystring + '_custom';
    var myResultString = '';
    var swapSymbolArtboardSearch = false;
    
    if (myProfileXML.child(mystringcustom).toString() == 'custom'){
        myResultString = myProfileXML.child(mystringtxt).toString();
    } else if (myProfileXML.child(mystringcustom).toString() == 'symbol'){
        myResultString = searchTXTfromSymbol(myProfileXML.child(mystringtxt).toString());
        if (myResultString == '' && mystringtxt != ''){
            swapSymbolArtboardSearch = true;
        }
    } else if (artboardName != ''){
        var myArtboard = docObj.artboards.getByName(artboardName);
        myResultString = searchTextInArtboard(myArtboard, myProfileXML.child(mystringtxt).toString());
        if (myResultString == '' && mystringtxt != ''){
            swapSymbolArtboardSearch = true;
        }
    }

    if (swapSymbolArtboardSearch == true && artboardName != ''){
        if (myProfileXML.child(mystringcustom).toString() != 'symbol'){
            myResultString = searchTXTfromSymbol(myProfileXML.child(mystringtxt).toString());
        } else {
            myResultString = searchTextInArtboard(myArtboard, myProfileXML.child(mystringtxt).toString());
        }
    }

    return myResultString;
}

function cleanPrefSuff(myStr, mode){
    if (mode == 'verbose'){
        // VERBOSE mode is to let the user know the special caracter is
        // acknowledged but won't be used
        //'[/!\\ NO☢☣♻☠⦿◘⛔⛝⛞⭕🚫🚯🚳🛇❗⚠]'
        myStr = myStr.replace(/\|/g,'[⚠]');
        myStr = myStr.replace(/\//g,'[⚠]');
        myStr = myStr.replace(/\\/g,'[⚠]');
        myStr = myStr.replace(/:/g,'[⚠]');
        myStr = myStr.replace(/\*/g,'[⚠]');
        myStr = myStr.replace(/\?/g,'[⚠]');
        myStr = myStr.replace(/"/g,'[⚠]');
        myStr = myStr.replace(/</g,'[⚠]');
        myStr = myStr.replace(/>/g,'[⚠]');
    } else {
        myStr = myStr.replace(/\|/g,'');
        myStr = myStr.replace(/\//g,'');
        myStr = myStr.replace(/\\/g,'');
        myStr = myStr.replace(/:/g,'');
        myStr = myStr.replace(/\*/g,'');
        myStr = myStr.replace(/\?/g,'');
        myStr = myStr.replace(/"/g,'');
        myStr = myStr.replace(/</g,'');
        myStr = myStr.replace(/>/g,'');
    }
    
    return myStr;
}

function searchTXTfromSymbol(mySymbolName){
    var symbolFound = false;
    var myResultString = '';
    
    for (var i=0;i<txtFromSymbol.length;i++){
        if (txtFromSymbol[i][0].toLowerCase() == mySymbolName.toLowerCase()) {
            myResultString = txtFromSymbol[i][1];
        }
    }

    return myResultString;
}

function searchTextInArtboard(myArtboard,myTxt){
    var txtToAdd = '';
    var myitem;
    var tmp = 0;
    var i=0;
    var artboardIndex;
    var idFound = false;
    var txtFound = false;
    
        for(tmp=0;tmp<docObj.textFrames.length && !txtFound;tmp++){
            myitem = docObj.textFrames[tmp];
                                
            if (myitem.hidden == false && myitem.name.toLowerCase() == myTxt.toLowerCase()){
                if ((myitem.position[0] >= myArtboard.artboardRect[0] && myitem.position[0] <= myArtboard.artboardRect[2])||(myitem.position[0] <= myArtboard.artboardRect[0] && myitem.position[0] >= myArtboard.artboardRect[2])){
                    if ((myitem.position[1] >= myArtboard.artboardRect[1] && myitem.position[1] <= myArtboard.artboardRect[3])||(myitem.position[1] <= myArtboard.artboardRect[1] && myitem.position[1] >= myArtboard.artboardRect[3])){
                        //$.writeln('myitem.position X & Y OK'); 
                        txtFound = true;
                        txtToAdd = cleanPrefSuff(myitem.contents, 'CLEAN');
                    }                                 
                }
            }
        }

    return txtToAdd;
}

function updateExportSettings(){
    var profileSelectList = this;
    var exportGRP = this.parent.parent.parent;
    
    
    if (profileSelectList.selection != null && profileSelectList.selection.text != '' && profileSelectList.selection.text != null && profileSelectList.selection.text != 'New'){
        var selectedProfileXML =  ProfileList.child("profile")[profileSelectList.selection.index];
        
        // pdf preset
        var settingFound = false;
        for (var i=0;i<app.PDFPresetsList.length && !settingFound;i++){
            if (app.PDFPresetsList[i] == selectedProfileXML.child("PDF").toString()){
                settingFound = true;
                profileSelectList.parent.parent.children[1].children[1].selection = i;
            }
        }
    
        //outline txt
        if(selectedProfileXML.child("outlineTXT").toString() == 'true'){
            profileSelectList.parent.parent.children[2].value = true;
        } else {
            profileSelectList.parent.parent.children[2].value = false;
        }
    
        //prefix - suffix
        loadProfilePref = true;
        loadProfileSuff = true;
        var myprefix = '';
        for (var j=1;j<numberPrefSuff+1;j++){
            myprefix += fillPrefixSuffix(selectedProfileXML, 'prefix'+j, '');
        }
        exportGRP.children[1].children[0].children[1].text = myprefix;
        
        var mysuffix = '';
        for (var j=1;j<numberPrefSuff+1;j++){
            mysuffix += fillPrefixSuffix(selectedProfileXML, 'suffix'+j, '');
        }
        exportGRP.children[1].children[1].children[1].text = mysuffix;
        
        //path
        if(selectedProfileXML.child("savePath").toString() != '' && selectedProfileXML.child("savePath").toString() != null){
            exportGRP.children[3].children[1].text = selectedProfileXML.child("savePath").toString();
        } else {
            exportGRP.children[3].children[1].text = docFolder;
        }
        
        
        updateExample(exportGRP);
        loadProfilePref = false;
        loadProfileSuff = false;
    }
}


///////////////////////////////// PROFILE MANAGER WINDOW FUNCTIONS

function deleteProfile(){
    if (ProfWinProfileNameGRP.children[1].text == ""){
        Window.alert("Please select the profile to delete");
    } else {
        var newProfileList = new XML("<profileList></profileList>")
        var nameFound = false;
        for (var i=0;i<ProfileList.child("profile").length();i++){
            if (ProfileList.child("profile")[i].child("Name").toString() == ProfWinProfileNameGRP.children[1].text){
                nameFound = true;
            } else {
                newProfileList.appendChild(ProfileList.child("profile")[i]);
            }
        }
        
        if (nameFound){
            //Save the file
            ProfileXMLFile.open( "w" );
            ProfileXMLFile.write( newProfileList.toString());
            ProfileXMLFile.close();
        
            //Update select list
            updateListNeeded = true;
            updateProfileList(0);
            updateProfileWindow();
        } else {
            Window.alert("The selected profile doesn't exist");
        }
    }
}

function saveProfile(){
    // save only if the profile has a name
    if (ProfWinProfileNameGRP.children[1].text == ""){
        Window.alert("Please input a name for this profile");
    } else {
        var isNewName = true;
        var selectedIndex = ProfWinProfileGRP.children[1].selection.index;
        
        for (var i=0;i<ProfileList.child("profile").length() && isNewName;i++){
            if (ProfileList.child("profile")[i].child("Name").toString() == ProfWinProfileNameGRP.children[1].text && i != selectedIndex){
                isNewName = false;
                Window.alert("This name is already used by another profile");
            }
        }
        
        if (isNewName && ProfWinProfileGRP.children[1].selection == ProfWinProfileGRP.children[1].items.length - 1){
            //Window.alert("New profile to save");
            var myNewProfileXML  = new XML("<profile><Name/><PDF/><outlineTXT/><savePath/></profile>");
            makeProfileXML(myNewProfileXML);
            // create the relevant number of prefix/suffix nodes            
            if(ProfileList.length() == 1 && ProfileList.child("profile")[0].child("Name").toString() == ''){
                ProfileList = new XML("<profileList></profileList>");
            }
            
            ProfileList.appendChild(myNewProfileXML);
        
            //Save the file
            ProfileXMLFile.open( "w" );
            ProfileXMLFile.write( ProfileList.toString());
            ProfileXMLFile.close();
        } else if (isNewName) {
            makeProfileXML(currentProfileXML);
        
            //Save the file
            ProfileXMLFile.open( "w" );
            ProfileXMLFile.write( ProfileList.toString());
            ProfileXMLFile.close();
        }
        
        //Update select list
        updateProfileList(selectedIndex);
    }    
}

function makeProfileXML(myXML){
        //Name
        myXML.replace("Name", new XML("<Name>"+ProfWinProfileNameGRP.children[1].text+"</Name>"));
        
        //PDF preset
        if (configProfWinTopGRP.children[1].selection != null){
            myXML.replace("PDF", new XML("<PDF>"+configProfWinTopGRP.children[1].selection.toString()+"</PDF>"));
        }
        
        //Outline text
        myXML.replace("outlineTXT", new XML("<outlineTXT>"+configProfWinTopGRP.children[2].value+"</outlineTXT>"));
        
        //PREFIX
        var i=0;
        var k = i+1;
        var customtxt = "custom";
        var xmlString;
        
        for (i=0; i<numberPrefSuff; i++){
            customtxt = "custom";
             if (ProfWinPrefixGRP.children[i*2].children[1].value == true){
                customtxt = "symbol";
            }
            if (ProfWinPrefixGRP.children[i*2].children[2].value == true){
                customtxt = "artboardText";
            }
            k = i+1;
            xmlString = "<prefix"+k+"_custom>"+customtxt+"</prefix"+k+"_custom>";
            myXML.replace("prefix"+k+"_custom", new XML(xmlString));
            xmlString = "<prefix"+k+"_txt>"+ProfWinPrefixGRP.children[i*2].children[3].text+"</prefix"+k+"_txt>";
            myXML.replace("prefix"+k+"_txt", new XML(xmlString));
        }
        
        //SUFFIX        
        for (i=0; i<numberPrefSuff; i++){
            customtxt = "custom";
             if (ProfWinSuffixGRP.children[i*2].children[1].value == true){
                customtxt = "symbol";
            }
            if (ProfWinSuffixGRP.children[i*2].children[2].value == true){
                customtxt = "artboardText";
            }
            k = i+1;
            xmlString = "<suffix"+k+"_custom>"+customtxt+"</suffix"+k+"_custom>";
            myXML.replace("suffix"+k+"_custom", new XML(xmlString));
            xmlString = "<suffix"+k+"_txt>"+ProfWinSuffixGRP.children[i*2].children[3].text+"</suffix"+k+"_txt>";
            myXML.replace("suffix"+k+"_txt", new XML(xmlString));
        }
     
        //Export Path
        myXML.replace("savePath", new XML("<savePath>"+ProfWinPathGRP.children[0].text+"</savePath>"));    
}


/////////////////////////////////////////////////////////////////////////////////////////////// LAUNCH THE SCRIPT ////////////////////////////////////////////////////////////////////////////////////////////////////////////

checkStatusBeforeRun();


///////////////////////////////// EXPORT PROCESS FUNCTIONS

function prepareExport(){
    var i=0;
    var j=0;
    var myExportGRP;
    var fileMode;
    var myPresetIndex;
    var myArtboards = [];
    var myPDFOptions;
    var myPDFVectoOptions;
    var myPref = [];
    var mySuff = [];
    var myPath = '';
    var myTempoFolder = '';
    var vectorWanted;
    RegularExportData = [];
    VectorExportData = [];
    
    for (i=0; i<exportConfigNumber; i++){
        myExportGRP = PDFExpMainGRP.children[i];
        fileMode = myExportGRP.children[2].children[2].selection.index;
        myPresetIndex = myExportGRP.children[0].children[1].children[1].selection.index;
        myPref = [ignoreProfilePref[i], myExportGRP.children[1].children[0].children[1].text];
        if (!ignoreProfilePref[i]){
            myPref[1] = getPrefSuffDataToArray(myExportGRP.children[0].children[0].children[1].selection.index, 'prefix');
        }
        mySuff = [ignoreProfileSuff[i], myExportGRP.children[1].children[1].children[1].text];
        if (!ignoreProfileSuff[i]){
            mySuff[1] = getPrefSuffDataToArray(myExportGRP.children[0].children[0].children[1].selection.index, 'suffix');
        }
        myTempoFolder = myExportGRP.children[3].children[1].text;
        vectorWanted =myExportGRP.children[0].children[2].value;
        
        if (myTempoFolder.charAt(0) == '\\') {
            myPath = docFolder + myTempoFolder;
            var myDataFolder = new Folder( myPath );
            // make certain the folder exists
            myDataFolder.create();
        } else {
            myPath = myTempoFolder;
        }
        
        if (fileMode == 0){
            // CASE TO EXPORT INTO MULTIPLE PDF FILES (1 FILE PER ARTBOARD)
            for (j=0; j<ArtboardToExportList.length; j++){
                myArtboards = [ArtboardToExportList[j]];
                if(!vectorWanted){
                    myPDFOptions = makePDFOptions(myPresetIndex,j);
                    RegularExportData[RegularExportData.length] = [myExportGRP, myPDFOptions, myPref, mySuff, myPath, myArtboards];
                } else {
                    myPDFOptions = makePDFOptions(0,j);
                    myPDFVectoOptions = makePDFOptions(myPresetIndex,j);
                    VectorExportData[VectorExportData.length] = [myExportGRP, myPDFOptions, myPref, mySuff, myPath, myArtboards, myPDFVectoOptions];
                }
            }
        } else {
            myArtboards = ArtboardToExportList;
            if(!vectorWanted){
                myPDFOptions = makePDFOptions(myPresetIndex,-1);
                RegularExportData[RegularExportData.length] = [myExportGRP, myPDFOptions, myPref, mySuff, myPath, myArtboards];
            } else {
                myPDFOptions = makePDFOptions(0,-1);
                myPDFVectoOptions = makePDFOptions(myPresetIndex,-1);
                VectorExportData[VectorExportData.length] = [myExportGRP, myPDFOptions, myPref, mySuff, myPath, myArtboards, myPDFVectoOptions];
            }
        }
    }

    launchExport();
}

function getPrefSuffDataToArray(profileIndex, nodeName){
    var myArr = [];
    var i =0;
    
    for (i=0;i < numberPrefSuff;i++){
        myArr[i] = [ProfileList.child("profile")[profileIndex].child(nodeName+i+'_txt').toString(),ProfileList.child("profile")[profileIndex].child(nodeName+i+'_custom').toString()];
    }
    
    return myArr;
}

function makePDFOptions(presetIndex,artboardIndex){
    var SELECTION = selectListGRP.children[0].selection;
    var options = new PDFSaveOptions(); 
    options.pDFPreset = app.PDFPresetsList[presetIndex];
    
    if (artboardIndex == -1){
        options.artboardRange = '';
        for (var j=0; j<SELECTION.length;j++){
            if (j > 0){ options.artboardRange +=  ',';}
            options.artboardRange += SELECTION[j].text;
        }
    } else {
        options.artboardRange = SELECTION[artboardIndex].text;
    }
    
    return options;
}

function launchExport(){
    var i=0;
    
    if (DatasetExists) {
        for (i=0;i<DatasetToExportList.length && RegularExportData.length > 0;i++){
            currentDataset = docObj.dataSets.getByName(DatasetToExportList[i]);
            currentDataset.display();
            regularExport();
        }
    
        for (i=0;i<DatasetToExportList.length && VectorExportData.length > 0;i++){ 
            currentDataset = docObj.dataSets.getByName(DatasetToExportList[i]);
            currentDataset.display();
            vectorExport();
        }
    } else {
        if (RegularExportData.length > 0){
            regularExport();
        }
        if (VectorExportData.length > 0){
            vectorExport();
        }
   }
    
    // WHEN EXPORT IS FINISHED, TERMINATE THE SCRIPT
    scriptCompleted();
}

function vectorizeAllTexts(myDoc){
    var mylayer;
    var i=0;
    var j=0;
    var myInstance;
    
    try{
        app.executeMenuCommand("unlockAll");
    }
    catch(e){
        //Window.alert('nothing to unlock or unlock error : '+e);
    }  
    
    for (i=0;i<myDoc.layers.length;i++){
        mylayer = myDoc.layers[i];
        if (mylayer.visible){ 
            if (mylayer.locked){
                mylayer.locked = false;
            }             
            while (mylayer.symbolItems.length>0){
                myInstance = mylayer.symbolItems[0];
                if (myInstance.hidden == false){
                    myInstance.breakLink();
                }
            }
        }
    }

    for (j=myDoc.textFrames.length;j>0;j--){
        if(myDoc.textFrames[j-1].hidden == false && myDoc.textFrames[j-1].locked == false){
            try{
                myDoc.textFrames[j-1].createOutline();
            }
            catch(e){
            }
        }
    }
}

function regularExport(){
    var i=0;
    var fileName = '';
    var destFile;
    var myTempoFolder;
    var myDataFolder;
    var options;
    
    for (i=0; i<RegularExportData.length; i++){
        myTempoFolder = RegularExportData[i][4];
        if (myTempoFolder.charAt(0) == '\\') {
            myFolder = docFolder + myTempoFolder;
            myDataFolder = new Folder( myFolder );
            // make certain the folder exists
            myDataFolder.create();
        } else {
            myFolder = myTempoFolder;
        }

        fileName = createFileName(RegularExportData[i]);
        destFile = new File( myFolder +'\\'+ fileName );
        options = RegularExportData[i][1];
        docObj.saveAs(destFile, options);
    }
    
}

function vectorExport(){
    var i=0;
    var fileName = '';
    var destFile;
    var myTempoFolder;
    var myDataFolder;
    var options;
    var tempDoc;
    
    for (i=0; i<VectorExportData.length; i++){
        myTempoFolder = VectorExportData[i][4];
        if (myTempoFolder.charAt(0) == '\\') {
            myFolder = docFolder + myTempoFolder;
            myDataFolder = new Folder( myFolder );
            // make certain the folder exists
            myDataFolder.create();
        } else {
            myFolder = myTempoFolder;
        }
        fileName = createFileName(VectorExportData[i]);
        destFile = new File( myFolder +'\\'+ fileName );
        
        // prepare a temporary PDF file, with final name
        options = VectorExportData[i][1];
        docObj.saveAs(destFile, options);
        tempDoc = app.open(destFile);
        vectorizeAllTexts(tempDoc);
        options = VectorExportData[i][6];
        // save the tempo file using definitive settings
        tempDoc.saveAs(destFile, options);
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
}

function createFileName(exportData){
    var filename = '';
    var pref = '';
    var suff = '';
    var myStr = '';
    var baseName = '';
    var artboardName = exportData[5][0];
    var prefSuffData;
    
    if (exportData[2][0] == true){
        pref = exportData[2][1];
    } else {
        prefSuffData = exportData[2][1];
        pref = getPrefSuffForExport(prefSuffData,artboardName);
    }    
    
    if (exportData[3][0] == true){
        suff = exportData[3][1];
    } else {
        prefSuffData = exportData[3][1];
        suff = getPrefSuffForExport(prefSuffData,artboardName);
    }

    if (exportData[5].length == 1){
        baseName = artboardName;
    } else{
        baseName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'));
    }
    
    if (DatasetExists){
        datasetIndex = exportData[6];
        if (DataSetNamingGRP.children[1].selection.index == 0){
            filename = pref+baseName+DataSetNamingGRP.children[3].text+currentDataset.name+suff+'.pdf';
        } else if (DataSetNamingGRP.children[1].selection.index == 1){
            filename = pref+currentDataset.name+DataSetNamingGRP.children[3].text+baseName+suff+'.pdf';
        } else {
            filename = pref+currentDataset.name+suff+'.pdf';
        }
    } else {
        filename = pref + baseName + suff + '.pdf'; 
    }
        
    filename = cleanPrefSuff(filename,'cleanAll');
    
    return filename;
}

function getPrefSuffForExport(prefSuffData, artboardName){
    var myStr = '';
    var resultStr = '';
    var swapSymbolArtboardSearch = false;
    var i = 0;
    var myArtboard = docObj.artboards.getByName(artboardName);
    
    for(i = 0; i<prefSuffData.length; i++) {
        swapSymbolArtboardSearch = false;
        if(prefSuffData[i][1] == 'symbol'){
            myStr = searchTXTfromSymbol(prefSuffData[i][0]);
            if (myStr == '' && prefSuffData[i][0] != ''){
                swapSymbolArtboardSearch = true;
            }
        } else if(prefSuffData[i][1] == 'artboardText'){
            myStr = searchTextInArtboard(myArtboard, prefSuffData[i][0]);
            if (myStr == '' && prefSuffData[i][0] != ''){
                swapSymbolArtboardSearch = true;
            }
        } else{
            myStr = prefSuffData[i][0];
        }
        if(swapSymbolArtboardSearch){
             if(prefSuffData[i][1] == 'symbol'){
                 myStr = searchTextInArtboard(myArtboard, prefSuffData[i][0]);
             } else {
                 myStr = searchTXTfromSymbol(prefSuffData[i][0]);
             }
        }
        resultStr += myStr;
    }

    return resultStr;
}



///////////////////////////////// ERROR MANAGEMENT
function checkSettingsBeforeExport(){
    //'[/!\\ NO☢☣♻☠⦿◘⛔⛝⛞⭕🚫🚯🚳🛇❗⚠]'
    var i=0;
    var myExportGRP;
    var errorFound = 0;
    var errorTXT = '';
    var PUCE = '☠';
    var SEPAR = '> '
    
    // check there is PDF preset selected in all export config sub-panel
    for (i=0;i < exportConfigNumber; i++){
        myExportGRP = PDFExpMainGRP.children[i];
        if (myExportGRP.children[0].children[1].children[1].selection == null){
            errorFound ++;
            errorTXT += PUCE+errorFound+SEPAR+'Please choose a PDF Preset for export profile : '+(i+1)+'\n';
        }
    }
    
    var artboardSelection = selectListGRP.children[0].selection;
    if (artboardSelection == null){
            errorFound ++;
        errorTXT += PUCE+errorFound+SEPAR+'Please select at least 1 artboard'+'\n';
    } else {
        ArtboardToExportList = [];
        for  (i=0;i < artboardSelection.length; i++){
            ArtboardToExportList.push(artboardSelection[i].subItems[0].text);
        }
    }

    if (DatasetExists){
        var datasetSelection = selectListGRP.children[1].selection;
        if (datasetSelection == null){
            errorFound ++;
            errorTXT += PUCE+errorFound+SEPAR+'Please select at least 1 dataset';
        } else {
            DatasetToExportList = [];
            for  (i=0;i < datasetSelection.length; i++){
                DatasetToExportList.push(datasetSelection[i].subItems[0].text);
            }
        }
    }
    
    if (errorFound > 0){
        popAlertWindow(errorTXT,errorFound,null);
    } else {
        prepareExport();
    }
}

function popAlertWindow(myTxt,errorCount,myAction){    
    // Prepare the Alert Window
    var panelTitle = '[ WARNING ]';
    var panelColor = [0.2,0,0,0.3];
    if (myAction != 'save'){
        panelTitle = '[ ERROR ]';
        panelColor = [1,0,0,0.5];
    }
    this.AlertWindow = new Window('dialog',"Corius' MultiPDF-Exporter  -  ["+ExportScriptVersion+"]");
    this.AlertWindow.preferredSize = [350,60+(15*(errorCount-1))];
    this.AlertWindow.orientation = 'column';
    this.AlertWindow.alignment = [ScriptUI.Alignment.LEFT, ScriptUI.Alignment.FILL];
    
    var messagePanel = this.AlertWindow.add('panel', undefined, panelTitle);
    var g = messagePanel.graphics;
    g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, panelColor);
    var alertMessage = messagePanel.add('statictext', undefined, '', {multiline: true}); 
    alertMessage.size = [ 340,15*errorCount];
    alertMessage.text = myTxt;
    
    this.alertBtnGRP = this.AlertWindow.add('group', undefined, '');
    this.alertBtnGRP.orientation = 'row';
    var cancelScriptBtn = this.alertBtnGRP.add('button', undefined, 'Cancel', {name:'cancel'});
    cancelScriptBtn.onClick = function() { AlertWindow.close() };
    if (myAction == 'save'){
        var saveFileBtn = this.alertBtnGRP.add('button', undefined, 'Save', {name:'save'});
        saveFileBtn.onClick = function() { docObj.save(); AlertWindow.close(); startPreparation();};
    }
        
    // Display the Alert Window
    AlertWindow.spacing = 10;
    this.AlertWindow.show();
}


///////////////////////////////// TERMINATE THE SCRIPT
function revertAndAbort(){
    this.ExporterWindow.close();
}

function scriptCompleted(){
    var undoCount = 1;
    var Regular = 0;
    var Vector = 0;
    var DatasetQty = DatasetToExportList.length;
    
    if (RegularExportData.length > 0){
        Regular = 1;
    }
    if (VectorExportData.length > 0){
        Vector = 1;
        undoCount += Regular;
    }
    if (DatasetExists){
        undoCount += (DatasetQty - 1) * Regular + (DatasetQty - 1) * Vector;
    }

    while(undoCount > 0 && docObj.saved == false){
            undoCount--;
            app.undo();
    }

    revertAndAbort();
}

///////////////////////////////// LOGO FILE CREATION
function myFileMaker() {
    // the (new String(....) in the line below is the output as pasted from myFileEnstringer
    // it is in JSON notation from the toSource() call above - when this line is executed, myBinary will contain an accurate binary representation of the image
    // the binary below is a PNG image.
    var myBinary = (new String("\u0089PNG\r\n\x1A\n\x00\x00\x00\rIHDR\x00\x00\x00\u00CE\x00\x00\x00B\b\x06\x00\x00\x00\x1Fm\u00E9g\x00\x00\x00\tpHYs\x00\x00\x0B\x12\x00\x00\x0B\x12\x01\u00D2\u00DD~\u00FC\x00\x00\x1C\u00C8IDATx\u009C\u00ED]\x0BtT\u00D5\u00B9\u00DE\u00F3d&\u0089\u00C0H\b\u00A0\u00B5\u0089\tX!b\u00B4\u0089\u008F*m\u0083\u0086\u00C2\x02lk\u00DB\u0080\u00D5U\u00B0\u00D4&\u00ED\u00D5\u00DE\u008B\u00ED\u00D5\tmea\u00EC\u0082\u00C4z\u0085zi+iK\u0081\u00D6V\x1E\u00B7\u00DEZ\u00B0HB\x1BZ\x14\u00A9I\x0B\u00CAC+\x19\u0088%\x17\u0090GHH2\u0093y\u009D\u00BB\u00F6q\u00FF\u00E3?\x7F\u00F6\u009Esf\u0092 \u0084\u00F9\u00D7\u009A\u00B5\u00B3\u00F7\u00D9\u00AF\u00F3\u00D8\u00FB\u00DB\u00FF3\x16M\u00D3X\u009AR#\u00FF\u0093\u00CBs-\u00D9\u00A3\x1E\u00B1\\\u0096U\u00AA\u009D\u00ED\u00B0\u0087\u00DF\u00DCw\u00B5\u00F3\u00AE\u0099\x07#o\u00BF\u00F3\x16\u00EB\u00F1{\u00DD\u008F>\u00DC\u009A~\u00B4C\u0093\u00D2\x0B'E\u00D2\x17MN\u00F6_\u00B5\u008E\u00CEw\u00B5\u00D3g61\u00BB\u00FD\u0098\u00C5\u00E5z-z\u00B6\u00FD\x1B\u00CE\u00B2;\u00CAy\u00AF\u00E1\u00A6\x7F\u00DC9\u00C4\x17\u00CF\n\u00C6\u00D8\r\u008C\u00B1 c\u00CCy\t\u00A5k\x18_8\u00E9_\u00F2\u00BF\u009E\x15+w\x06\u00D6oZ,k\u00DBS\u00FBtnp\u00C7_\u008F\u00FBW>[3\u00C4\u009Fm\u00A3\u00F6>\x1D\u00B9\u00C4\u00D2%i\u00C4I\u00818\u00DAD\x0E\x1F^\u009B\u00F5\u00D3gJU\u00AD{7\u00FEn\u0093\u00C55\u00ECS\u00CE\u00BBf\u00E5\\\u00C8\u00F7\u00D2O\u00DA\u00C6\x18\u00BB\u00E6\u00E0\u0081\x03\u00DA\u00EB\x7F{\u00DD\u00E2r\u00B9\u00B4@ `:\u009D5{\u00966*;\u00DB\u00C2\x18\u00E3\x1Fa,\u0085\u00FE\u00E8uZ\u009Ej\u00BD\x01\u0098\u00EF\u00F3i\u00B4I\u00E1\u00D7\u00F5\u009D\u00AA\u009Fw/}rJ\u00A2\u00B6\u00DD\u008FU/\u008Cvth\u00DD\u008F/\u009D;\u00D4\x11g\u00ED/\u00D7\x1C)\u00C8\u00CD\u00D3\nr\u00F3\u0092J\u008F\x1C>\"E*\u00E8\u008F^\u00A7\u00E5\u00A9\u00D6\x1B\u0080\u00F9.\u00B1_\u00D4\u00FB\u00DD\u0087D\u00D6\u009C\u009C'\u0080w\u00F1\u00FF\u00F7O\x1F\u00B0}\u00F4\u00AA/\u00B1\f\u00F7X\x16\nO\u00D0g\u00E4\u00B0\u00BFc\u00BF\u00B1h\u0084\u00D6y\u008EY\u00B3Gy{\x16?1\u008E\u00B9]M\x19\u008B\x1E\u00D99\u00C4\x1E\x05?\u00F3\u00B7\u00F2\x1D\u0099\u00A7O\u00FFh\u0085VTtCk(\x14\u00D4\x1C\x0E\u00A72]\u00B7v\r\u00FFx\u00F5<o'\u0090\"\u0096\u00F2\u009D\u00FD\u00D3\u00A5\u00A5\u00AD\u00D9\u00A3\u00B3\u00F5\u00BC@\u0088\u00D6\u00C2\u00C9\u00D7i\u00DB\x1B\x1Bc\u00E5P\u00AF\u00A5\u00E5\u0090\u00B6\u00A3\u00B1\u00B1U C\u00EB\u008E\u00C6F\u00FD:\x1F'77/\u0096\u0087\u00F1\x06`\u00BE\u00AE\u00F4\u00C2I\u0081\u00F8\u00A2\u00E9\u00A9^6\u00D7vu\u00DE\u008F]\u00F3\u00EF\x1B\u00C5\x17H\u00F8\u008D}=|\u00C1\u00C4z\u00CB\u00CA\u00EC\u00E0\u0089\u00ED\u00BAIWX\x0B\u00AE^\u00CE\u00EB\x04\u00FF\u00B0\u00E5d\u00F4\u00F8\u0089\u00DD\u00D1S\u00A7k\u0087\u00C8\"\u00E2\u008Cr.\u00FFXy\u00CA?\u00C2\u00DC\u00BC\u00DC\\\u00F1q)S\u00FE1\u00F3\u0094\x7F\u0094\u00B2z\u00A3\u00B2\u00B3[Geg\u00C7\u00F2|\u00D1T/Y\u0092\u00CB\x17\r\u00EE\x1F\u00EA\u00F1E\u00C3\u00AF\u00D3~\u00F8\u00C7\u008E\u00F30\u00DE\x00\u00CC7\u0090^8)\x10\u00E7q\u00EC%\x1F_\u00ED\u0098\u00F2\u0089\u008C\u00E0\u00CB\r\u00FB\u00A3\u00C7O\u00ACp\x7F\u00EB\u009B?\u00C7=u\x7Fw\u00F12\u00C7-7\x15E\u00DE\u00DC\u00BF\u00C6\u00F5`e\x15/\u00EBY\u00FC\u00C4BK\u00CE\u00E8[-\x19\x19\u00EB{\u00AA\u0097\u00B51\u0087\u00FD\u00DB\x17\u00F9\x02\u008AC\x1C\u00D8\u0091\x01!\x00\x01 \x15<\u0082\u00E9\u00FA\x14A(BQ$\x02d\x00$\u00A1y@\u00AA\u00FE\u00CE7\u008D8)\x12\u00D7\u00DD8gN\u00CF\b\u00FC\u00A4\u00EE\x1F\u00AE\x7F\u00AB\u00F88\u00ED\u0085/,\u00C7M%\u00F7\u00F0\u00BF\u00A3\u00A7\u00CF\u00FC\x03\u00CA3\u00AA\x1F[\x01\x7F\u00F7,\u00FB\u00E1\x14\x16\n?\u00DD\u00F3X\u00F5\u00A1\u008C'\x16\u00DF{A\u00DF\u00B0\u009A\u00E2\x10\x07vd@\b\u00BAs\u00F3c\x15G\b\u00B3\u00F5iJ\x11\u008A\"\x11 \x03 \t\u00CDC\u00BB\u00FE\u00CE7\u008D8)\u0092e\u00983\u008B\u00B7\u008C\x1E;\u00FE\x16\u00ED\u0081\u00F3<\u00F6\u0092\x1B\u00AB\x1Cw\u0094\u00E6\u00F9\u009F\\^\u009F\u00B1x\u00D1z\u00D9(\x02in\u00E6\x0B\u0088\u00FF.R\u00E4\u0091\"\x0E\u00E4)\u00EFB\u00AFC\x1Ex\x15\x15o\x01\u00ED\u0081g\x01$\u00A0\u00BC\x10\u00E5y\x00\u0089 \u00AFB\u0090d\u00E7\u009BF\u009C\x14)\u00DAv\u00EC\u0090\u00D6\u00D9\u00C9\\\u0095_\u00FBr\u00F0\u0085\x17?\x11\u00ED\u00EC<lq82-nw\u008Ek\u00FE}y\u00BC\u00D7\u00C0\u00DA_\u00FF\u00D9\u00FD\u00E8\u00C3\u009F1\x1A\u00E1\"?\u00AAI\x11\x07\u00F2*\u00C4\u00A0\u00F5)OCy\x0B\u00CA\u00B3\x00\x12\u00D0v\u0094\u00E7\x01$\u0082\u00BC\nA\u0092\u009Do\x1AqR$\u00F7\u00A3\x0F\u00FF\u00C0\u00FF\u00E4rf\u00C9\u00C9\u00AEp\u00DCT\u0092\u00E7,\u009C\u0098\x17=\u00DA\u00C6\"\u00FF:\u00CA\u00C2M\x7Fo\t\u00EF?XCy\u009E!J\x03\u00828*\x1E\x03\x10\u0083\"\x10E\x1E\u00CA\u00F3P\u009E\u0088\"S\x1Aq>D\u00E2\u008B\u00871\u00F6\x03\u009DWY\u00BF\u00A9\u0084\u009B\u00DC0\u0087\u00BD\u008D#\u0088\u00E3\x0E\u00A5^t\u00A8\u00D1\u0080 \u008E\u008A\u00C7\u00A0R4@ \u008A<\u0094\u00E7\u00A1\u00E3QdJ#\u00CE\x05@\u00E2\u00A8\u00A5\x1F\u00B7\u00B8\u00D4\f\u00FE\u00BEDH\u008A8*\u009E\u00C5h\u00C7\u00A7z\x15\u0095\x1E\u0087\"\x0F\u0095\u00BEA?\u0094\u00B7\x01$\u00E9\u00EF|\u00D3\u00883\b\u00C4\u00F5;*\u0081\u00C0\x10$)\u00E2\u00A8x\x16\u00A3\x1D_\u00A5WQI\u00CF\u00A8\u009E\x06R\u00E8\u0087\u00F26\u00AA\u00F1\u0093\u009Do\x1Aq\x06\u0083\u00C2\u00E1q\u00FD\u00E9\u0095\u008B\u00B2\u00B9\u0085u\u00E6\u00D2\u00EAE\x17\u00EC=~@q\u0088C\x11\u00C0\u00AC^F\u00C5\u00F3\x18I\u00C1T\u00FA\x1A\u00B3R\u00B5T\u00E7\u009BF\u009CA \u00DB\u00C7&\u00DC\u0090j\u00AF|\u00D1\u00B0a\u00CE\u00E7\u00AC#=\u00F7]\u00C8\u00F7\u0088(\x0EqT\b`\u00A4\u0097Q\u00F1<FR0\u0095\u00BE\u00C6\u00ACT-\u00D5\u00F9\u00A6\x11g0\u00C8\u00F9\u00BE\u008EG\u00F0;\u00CC:&g\u00ACe\u00E4\u0088\u00B1\u0091\u00B7\u00DF\u00D9\u00A3\u008F\u0096\u00C0fM\x0B\x046\u00B2@\u00E0\u00BF2\x16/j\u00BDH\u00EE6\u0098\u0088G0J\u00A9\u00FEE\u0095\u00A7\b\u00A1\u0092\u009A\u00A9\u0090FU?\u00D5\u00F9r\u00C4\u00B9\u00E4-\u009D\x07\u00EA\u00C7\u00AD\u00A5{j\u009F\u00DE\x16|\u00F55\u008D\u00FF\u00B8e4\u00FC\u008D\x7F\u00BD[\u00B7i\u00DCW\u00A7\u00F7\x7F\u00FF\u00B0\u0095[P\u00C3\u00F8=\u00CF\u00FC\u00E4\x01\u00FF\u009A_\u00FD\u00E9b\u00B4\u008E\x1El\x7F\x1C\u00955\u00B3Y\u00EBi#k\u00EC\x14\u00D2\u00B4?\u008E\x11\x19i\u00F5u\u00EB\u00E8k\u00AFy\u00C8>\u00F1\u00DA\"\u00AD\u00A3S/\x0B\u00EF\u00D9\u00BB\u0096#\u008C%+\u00EB\x05\u00EA\u00B7\u00A3\u008B\u00AF\u00FD\x01.\u00BE\u00BE\u00979\u00EC\u00D7X\u00B3\u00B3\u008Fk\u0081\u00C0S\u00D6\u0091#k\"G\u008F~\u00FE<)D\u00F3\x19c\u00C5\u008C\u00B1\x06\u00C6X\u00BB(\u00F3\u00A0\u00BF\u00CD\u0092\u00EE\u008F\x03~0\u00A7O\u009D\u00D6\u00BA\u00BA\u00BA,|g\u00CE\u00CC\u00CC\u00B4\u00A8\u00CA\u008D\u00F2\u00D4?\u00E7\u00C0\u0081\x03Z\u0093\u00C4\x7F\u00E6\u00BA\u00C9\u00D7i\u00A3Fe[\x04\"X\x04\u00AFc\u00F9c\u00FD6m\u00C2\u0084\t}\u00DAA}:\x0Fh\u00AF\u009A\x1F\u009D\u008F\u00CC\x1F'_\u00D34\u00AF\u00A6i\x1B4M\u00AB\x17\x7F\u0097\u0089\u00B4\u0082\u00D4\u00F5\u0088\u00F2z\u00F4[%\u00EA\u00E3:5\u00A2\u009E\u00EAW&\u00EA\u00E1\u00BE\u00BD\u008AvP\u00D6D\u00C6\u0085\u009FG\u00FC\u00EAI\u00BB\rh^\x1Eq/\x1BP;\u009E/\x17\u00FD\u00C6\u00CD\u0085\u00FB\u00D3tW/{\u0086\u00EE\u00B6\x1Caz\u00B75\u00EC\u00E1\u00C8\x12\u00F9\u00D7Q\u008D\u00A3\x05\u00F7\u00D3\u00E1\u00A8\u0082\u00EBu}\u00A7j\u008Fl\u00B7\u00E6u5\u00E1\u00B7\u00D3\u00FB\u00E2\u00E6\u00F78\n%\u00B9\u00DB\u00E7\u00A3_\u0085\u00B8\u00CFU\u00E2\x1E\u00CCx\u009En\x10u!_\u008F\u00DE\u00A7Y\u00CF\u00D5\u00C6D;~\u00B2\u00FE1\u00C9\"\u0082Y$J\x16\u0091L\u00CC'\u00CE\x1F\u00C7\u00CB\x18\u00ABAy\u00BE\u00FB\u0094\u00A1|\x15\u00FA\u009B\u00EFV\x1B\u00C4\u00CE\u00C5\u00A9Y\u00A4\x15\u00E2\u00C7w\u00B29\u008C\u00B1r\u00D1\u00AF\x11\u00B5\u008B\u00FA\rh\u00F73\u00D3NEed\u00EE\u009C6\u008A>\u00CB\u00C5\u00FC\u00F1N\x0Bu}\u00E2\u009E\u00E0~\x18\x17-w=\u00F4\u00F0R\u00C6\u00D8\u00BFCY`\u00CD\u00AF\x1Ev\u00CE\u009EYm\u00CB\u00BD*+\u00B4s\u00D7+\u00E1\u00BDoTq\u00A4\u00E0\u00A2h\u00C6\u00D8\u00D78\u0093o\x14k@\u00EB\u00EE\x1E\u00CF\u0084\u00E1g\u00F0\u00F7\u009BgD\u00DB\u00FE\u00EFS\u00C1\u00DFo\u00DE\u00EA\u00FC\u00DC\u00EC\x19&\u00EE\u008F\u00CF\u00B9^\u00DC\u00C3F\u0091\u00F29\u00D7\u0089wX!\u00CA\u009B\x13\u00F4Q'\u009EG\u00B1\u00A8\u00E7\x13\u00FD\u00D6\u00A1\u00F7`DA\u00ACi/\u00B9\u00F9&m\u00F1\u0092%\u00AD\u00A7O\u009F\u00D2\u00F3b\u00C7\u008E\u0095geeJ\u00EB\u00B5\u00B6\x1E\u00D1\u00D3\u00BD{\u00F7\u00C4x\x11\u00E0)x{\u0081\x00}\u00CA\u00A1\x1F,\u00FD\u00E2<\r\u00F4\x03\u00F5T\u00F3\u0083q\u00E1zP\u00F0@\u00B4\x1E\x1D\x17\u00F38^\u00ED\x03ZEv\u00DDzq\u00C5+\u00F2\u00C5\u009A\u00A6\u009DA\u00F5\x01\u0089<b\x07\x03j\x12u\u00CBE\u009F\u0098\u00BCdL \u008C\n\u00F9\u00E4Z\x0B\u009AW\x19\u0099\u00C3\x19\u0082tedL\u0098{\u008Bd~\u00F9h\u00DEgd;\u00AB\u00BFn\u00F5[\u00E0\u00F1\u00E9\u00FF\u00C5\u00DA\u0095\x1Ca\"\u00AD\u00EFv\u00CAb\np\u00C4\x01\u00DE\u0085\u00C7\x1E\u00E8\u00F9\u00E1\u00F2C\u00B2>\u00CF\u00FD\u00FBw^\u0083\u00BF{\u00B7l\u00ED\u00E6mx\u00DF\u00DD\u00DF\x7F\u00FC7&w\u00FB\n1\u00EFU\u00E2\u00BEZ\u00D0\u00BD$B\f\x0FBZ\u00FCL7 \u00D4\u00CA\u00EF\x0F\u008Fc\x16A\u008C<1\u0093\u00E5UR\u00ADg\u00C43\u00A9<@\u008B\t\u00D2\u00D4\u0092\u00B3n\u00AD\u00D8\u0091\u008B\u00C5\u00AF^\u00ECLL\u00ECLu\u00E2\u00EFv\u0081J\u00F5\"_,v\u00B49\b\u0099\u0080\u009AE[\x18\u00D7G\u00EA\u00B4\u00A39\u00C05\x1F*k\x10}\u0094\u0091\u00FE\x18\u00BA^L\u00EE\u0081\u00A11|\u00A2\fv\u00E4*\u00B4\u0083\u00F7!K\u0086\u00FB5\u00DB\u00F8\u0082\u0085\u0081\u00B5\u00BF\u00AEv-\u00987\u0095_\x0F\u00ED\u00DA}\u0099%g\u00F4l\u00FFS+F\u00B3\u00A8V-C\x18\u00AD\u00AB\u00EBn\u00FB\u00AD7\x1F\u0097\u00F5i\x1D1b,\x13\nS\u00FB-%\x19\u00E0r\u00D0\u00F3\u00FD\u00C7w\u0098A,\u0081(\u00FC\u00F9U\u008Aw\u00D0.\u00D04\x1F\u00DD\u00AF\u008C<\bY\u00CA\u00C5\u00AFN<\x13&\u00F2>\u0083>\u0080t\u00C4\x01D\u0080\x1D\x19\u00A4`\u00B0\u0083S\u00C4\u0080\u0094\u00EE\u00F8\u0090\u00CE\u00FF\u00EA\u00FD\u00BA\u00F4\fv\u00FC\u008F\\\u00F5\u00918d\u00A1H@\u00FB\u0083\u00F6\u0080pt|\u00E8\x0F\x10\u008A\u008E\u00A7\u00AA\x17$z\u009Cr\u00F1\x10\u00E0c\u00F5\u0091\u0087\x03\fd\u00BE\u00F8P=\u00E8\u00DAFE]\u00A8S\u00CE\u00D4\u0084\u00FB\u00A1\x0B\x0B\x13]T\u00C9\x12\u00DC\x0F\u008C\x07\u00FDm\x10\u00F3\u00AFB\x0B\u00C8\u0083\u008E.1\u00E2\u00D6\u00D0\u00B6\u00EB&-rN/\u00CB\b\u00BF\u00B9\u00BFK\u00EB\u00E8\u00B0\u00D8\u00AF\u00BF.3:|x\u00A1\u00AD\u00FC\x0B\u0085\u00A1\u009D\u00AF\u00CE\r\u00AC^7\u00C7\u00B5`\u00DEK\u00BC\u008D5/W?\u0086qQ\u00B4\u00D6\u00DE~(\u00D1|\u00AD\u00A3.\u00BF\u00D1\u0092\u0099\u00D9\ry\u00CB\u00F0\u00E1\u00F3\"mm|\x11\u00DDmp\u009F\u00ED\u00E8h\u00E6\x11\u00F7\u00D5,6*\u00D9\x06\u0080\u00DF\x1F\x1C\u00C3\u009B\u00D1\u00BB\u00AD\x15\u00F7nv\u00D10\u00D0\u00E3l\u00D9\u00BC9\u00A67\u00C9\u00CC\u00CC\u008C\u00E9I\u009CB\u00EF\x01\u00D7\u00CD\u00EAK\u00E6\u00CD\u00BF_\u00D7\u00BF\u00AC[\u00B36\u00D6\u00EF\x1Dw\u00DE\x19\u00BB\x0E\u00E5F\u00ED!O\u00C7\u00C7\u00FD\u00EDhl\u00EC3\u009E\u00AA\u009E\u0093\u00E8q`g\u00F6$8\x13\u00CFA;\x1A&\u00BA\u00C8\x18A\x02Fv~\u00A02\u0082r,\u00C1\u00D9\u00DA#)3\u009A\u0083\u00EC:\u00F0\x03\u00F8\x1Ex\u00DF-b\u00DCZ\u00C5\u00C6\u00C1u/\u00C7\u00B9\u00B7\u00A7v\u00B6\u00E3\\\u00F8\u00F5\u00E6{\u00ACc\u00C7|W\u00EB<w{\u00EF\u00FAM\x0F\u00DB\u00AE\u00BD\u00E6\u00CB\u00C3\u00EE\u009D{\u00B3\u00D6\u00DE\u00BE\u00861\u00A6G\u00B4\x01\x7F\x1D\u00AE\u00BFQM\u00CA2rD.\u00D4\u00D1\u00FC\u00FE\u00D8\u00E2\u00D2\u00DD\u00B2\x17?\u00B1\u00C3\u00E0\u009E\u0098@s@\u008A\x16\u00F1\u00DC\x13}\u00F0\u00C0\u00DBm\x14(U&\u00DA\u00F3g\u00D0\u0084\u00F8\u009C\u00BA\x04}P\u008A\u00B3\x1C\b\x12\u00CD>\u00CD\u00C3\u00CENy\x12Zn\u00C4\u00FB\u00A8\u0090\nR\u00CA\u00E3P\u00EBg:/\u00A8O\x11\u008B\u00CE\u0083\"\u008E\x19\u0082\x0F\u009A2\u00DC2$\u00A0\x1F\u00BA\u00EC\u00C3\u00C7\u008C?\u00FFXK\x12,\x00\x15\u00DA\u0094\u00A1\u00EB\u00AA\x05L\u00A9\x1D\u00FD\u00F2Q\x1Fe\u00E8C*\u00A3Hj/\u009CXe\x19>\u009C\u0085\u00FE\u00FC\u0097W9\u00AA\u00F8W>;\u00CF9s\u00BA~m\u00D8\u00BDso\t\u00FDe\u00E7q[~\u00FE\x18.jv\u0094~*\u00A6\x04\u00B5x<\u00E3\u008D\x10GF\u00D8S4\u00C13\u0081g\u00B8\x11!\b\x1C\u00A7\u00E7H\u00EE\x1F\x16U=\u00DA\u00DC\u00F2\x05\u00E2\u00C2\u0086\u00D1\x07m\r(\u00CEr\u00C0I4\u00FB4O\u0091\u0080\u00EE\u00F8*D\u00F9\u00F6\x7F,LJ\u00C3\x0F\u00F5)\x02\u00C2u:/\u00DA\u00BF\n\u0081(\u00E2`XO\u00F6H$[\x14\u00F4\u0098\u00E0S\u00A0\x0E\u00EE\u00A3\u0082H\u00ED\u00CC\u00B4o7\u0081F\f-\x0E\u00AF\u00E8\u00AB@\u00EC\u00BE5\u00A8=\u00F4UN\x17a\u00E0\u00C7\u00ABj\\\x0FV\x16\u0084v\u00EDf,\x1A\u00FD\x18/\u00D3\u00DE;\u00F9\x1Acl.\x1C\u00C9\u00B4\u00DE\u00DE\u00E3\u0096\x11\u00C3\u00C7\u00B0P\u00F8J\u00F6>\u008A\u00E8\u00E5\u00CCa\u00CFRM\u008A/D\u00CE\u00CB\u00D8n\u00BC>\x15\x13\u009D\x1A\u00B1\u0099\u0095#\x04\u0085\u00A3Z\u0095\x01\n\u00FB\u00D0\u00E9\x02\u00A3\u00B1O\u0089\u00B8j\nb)\x14\u00F0\x14F\b\x04<\x07\u00F0\x10T\nG\x11J\u0085LPn\u00C4#\u00A9x\x1F\x15\x12R\u00E9\u009F\u008C\u00C7\u00B1\u0092\u00E3Q\u00BE\u00E2c\u00AC\x10\u00C7\x013\u0084y\t\u009F\u00E2E\u00D4\u0092\u009D\u00CD+\u00C6\u0090\u0091j17\u00A3E\u009Ahab\u00A6\u00B7F\u00D4\u00E5\u00C8r\u00B9@\u00BAf\u00C9\x02\u008A\u0091mr\u00E1\u00FD\u00DC\u00DB3\u00FC\u00CA\u00AE\x16\u00DB5\x13\u00F2t{2\u00B7\u00AB\u0089;\u00AE\u00F1#\u0099\x1E\n\u00D7\u00ED\x1E\x1F\u00DE\x7F@\x17]k]]{\u00A1\u00ADe\u00D8\u00B0\u00B1Zo\u00B0K6)[\u00E1D]x\u00A0\u00D7\u00B3\u00D9\u0094\x0BLB^$\u00EE/A\u00F7`\u00E6Y0\u00B1\u0091\u00F8\u00C8}{\x12\x1C\u00A9\x13\u0091\u008E8\u0093&M\u00B2\u00CC\u00BB\x7F~\u00AE\b\u00D6\u00C7wr\u008B\u00D8\u00A1\u00A5\u00F9-\u009B7[\u00EE,-\u00CD-(\x18\u00AF\u00B7\u00DB\u00F7\u00E6>=\x7F\u00EA\u00E4\u00A9\u00B8\u00FA\u00F3\u00E6\u00DF\u00AF_\u00FFtii\\~\u00CE=s\u00E3\u00CA\u0085R2\u0096B=\u00E8\u0097+Eq9\u009D'\u00ED\x17\u00AE\u00C3}\u00C180\x7F@\u009C\u008D\u0088\u00DF\u00F0\u0088\u00BF\u00AB\u00D0\u008B('\fh-:&x\u0091D\u0087\u00BE\x00\u00FE\u00C1OS<\u00F4vq\u00CEnBe5\u00E2e\u00D2\u00A3\x02|\u00CC2\u00C9\x1C\u00BCX\u008F\u00E4\u00C8\u0086u3\f\x1D\u00D1V\u0089\u008F\u00AD\x1D\u00F1\x05\x1B\u00D0\u009Cc\u0088\u00C9%^\u00EE\u0085\x0F\u008E\t\u00ED~}\u00AF\u00D6yn\u008D\u00ADp\u00E2r\u00B6e\u00EBW2\x16=\u00F2\x03\x1D\u0081,\u0096\u0089\u0096\u00CC\u008Co:\u00A6\u00DC\u0096\x19|\u00E9\u00E5W\u00A0]l!8\x1Cc\u00A2GZ\x13\x1F\u00D5z\u00FC\u00C7Y\u00CE\u00E8\u00E9\t\u00EB\u00C4\x13\u0096\x10\u00D6\u00A0\u00E7\u00EDC\u008C?\x15\u00DA\x00yP[\u0099\u00D0\x05\u00D0\x0B\u00DEA\u00B1\u0081N'N\u008FC\u00F5%*\u00BD\rE\x1A\u00CA\u00CBP\u009E\u00C2\b\u00B1\u00A853\x1D\u0097\u0096\x1B!\u009E\u0091\u0094\x10\u00EBq*\u0090.@E\u00D8r`\x03\u00AAS/t5\u00E5H\u00E7\u00A3\u00A1\u00FA^R\x0Em\u008A\u0085\u00BE\u0081\x12\u00B6\x12\u00D0\u0090\u00BE\u00E6\u008C(\x03\x1DC9\u00BA\u00D6\u0082,\x00\u00BC\u00A4\u00DFz\u00A4\u008F:C\u00C6\u00A7:\u00AC3H\u00E7\u00C3\x02\x1B\u00FEg\x13/\u00E4\u00FA\x1A\u00AE\u00C7\u00E1\u00FA\x1BQ\u00C6t{\u00B3m\r{\u00B8\u00D6\u009F[\x0Fp[3]'\u00B3\u00ADa\x0FX\x0F`\u009D\u008EJ\u00DF\x13\u00F8\u00F5o\u00D7\u0084\u00F7\x1D\u00D0\u00B8\u00CE'I\u00CB\x01\u00AC\u00EFZE\u00DE\x0F\u00B6*\u00C0\u00F5\u00CB\u0091\u00B5@\u0085h\u00C7\u0090N\u00CBC\u00FAnA\u00CF\u00C9\u00B4\x1E\u00C7Hoc6\u00A2\u00E6@[\x1A\u00A4j\u00C9 \x19/f9\x00\u0092\x14\u0099\u00B6\x1E\u00F43X\u00DAR)v\u00AC\n\u0089\u0096\x1E\u00D7\u00CF\u0097H\u00CF\u0098\u00A8\u00BFJ R9\u00D9\u00F9j\u00D0\u00EEIwK\u00AF\u00E8\u00BFV\u00EC\u00AA>\u00C1\u00EC\u00E6\x13f\x1F\u00CFS\u00B6\u00FB\u0096\t\u00B4\u00A3G\u00B3J\u00BC\u00C3r^\u0085\x1F\u00C9\u00A2\u009D\u009D\u009B\u00B9e\x00G\x19\u00E0_8_\u00A3\u009D\u00ED(\x1AV\u00FE\x05\x16\u00AC\u00DF\u00BE\u0097\u00C6\x18\u00E0G8\u00E7\u00AC\u00C4F\x00\u00DC\x05!\u00DA~\u00D6\u00A5\x1F\u00DB\u00D6o\u00BA[D\u00FFOD\u00F9\x04\u00D5\x01i\u008B\u00C5=\u00C1\u00FB\u00C3\u00CF\u00B3\u008AH\u00DB6\u00A2\u00BA\r\u00A8l\x15B\u009A|$\u009ANd\u00BF&E\x1C\u0095\u00A6_%\u00C5\u00A2\u00BC\u0089\n9T<\x14E\b\u00D5|T\bB\u00F3F\u00E3Q\u00A9Z\x1D\u0092*aj\u0096<\u00BCv\u00C2\u0084z\x10c\u0089\u00EB\u00FB\u0088\u009E\u0084\u00A1\u00FAp,\u0082\u00C5\u00D3L\u00AE\u00C9\u00A4;t15\u008B\u008F\u00BD\u0098\u009C\u00D1A\u009F\u0084\x17;^p \u009A\u00C5\u00FCX\x1D\x1D\u00D3\u0092\u0095U\u00C4#t\u0082\u00E1%\u00E7_x\x19\\\u00E7\u008B\u0086\u00F3?\u0091\u00B7\u00FE\u00B9\u0092M\u00BB3nb\u009C\x7F\u00E1\x0B\u00C2\u0088\u00A8\u00A0\u0081\tCP\u00ED\u00DC\u00B9Y\u00D6\u0091\u009Eg\u0089\"\u00B4\x06}\u00D0>\u00F4\u00BC7&`\u00EA\u00F1\u00BB\u00C3\u009B\b\x16\u00FF\x03\u00CFY,\u009E=(\u00B6UG> \u00A7L\u00AA%\u00D3\u009B`\u00E9\u0096JJ&\u0093\u00BA\t^\u00A3u\u00D2\u00A4I}<2\u00A9\u009E\b\u00C6Q\u00CDG\u00A5g\u00A2y\u00A3\u00F1T\u00FE8f\u00ED\u0094\u0098I\u0099\u00BF\u00912-\x19\u0085\u009B\u008C6\u009Ax\u00C1t.\u00E6\u00C7Cam\u00F9?\u008F\u00B2]q\u0085\u00FEa\u00DB'M\u00E4\x16\x01E\u00E1=o\u009C\u00A0h\u00E3\u00F8\u00C4-\u00ACw\u00DDs\u009F7\u00D3\u00BDnA\u00BD\u00FF\u00E0r\u00EB\u00D81\u00BAt\u00CD\u00FF\u00D4\u008A_8g\u00CF\u009Cc\u009F\\\u0098\x15\u00F8I\x1D\u00E7}p\u00C0\u00C39\u00A6\u00E7mL\u00F4\x194$\u00F9\u00EE\x19\u00B5\x1CPi\u00F0U\x16\x00T*fdi@5\u00FA\x14\x11\u00A8\u00BE\u00C7l}#^\u00A7?z\u009C\u00F3I\u00A9\u0098\u00B7\u009B\u00A1\u00FE\x7Ft\u00C1`\u0093\u00ADpb!\u008Fj\u00A3uu\u00E9\u00A64\u00D1\x7F\x1D\u00DD\u008A\u00ABh\u00EF\u009D\u00E4\x0EkE\u00CA>\bqD\u00E11\u00A5\u00ADyyE\u00BD\x1B\x7F\u00B7\u00CF\u00FD\u009F\x0B\x0B#\u00FB\x0F2~D\u00B4N(\u00B0*\u009AU \u00FDM\x7F\u00A9?\u00CF\u00DB)C\x18\u00B3\x16\x00\u00A0\u00C7\u00A1z\x12#K\x03\u00DA\u008F\x11\u0092\x19\u00D5\u0087\u00EBx\u009E\x1C\u0099\u00E8<\u008C\x10g ^\u0084\u00D7\u00C4\u00CB\u0080\x17Om\u00C4\u008A\u0089\u00A5\u00F4\u0087JRQ\u00B1\u00C3\u00D1\x1F\x13\u00A0>\u00C4\x03\u00B1;>\u00FD\u00C9\u00D9\u00B6\u00DC\u00AB\nC\x7Fjl\t\u00ED\u00DC\u00F5=\u00C7\u00B4;\u009EO\u00D0\x04\u0094\u009D\u0095\u00A8,\x1F\u00F1&\u00C0\x0BUJ\u00DE\x03\u00E8\u00AA\u00CA\x10_\u00D8@\u00FA*\x17\u00E5\u0095,1\u00C5Y\x0EP\u00DEB\u00C53\u00A8\u00A4b*d\u00A2\b\u00A5\u0092\u00AEQ}\x0F\u00B5\u0096\u00A6\bd\u00C4K%\u00D2\u00E3\f\x16\u00E2\x18\u00B9\x04\u00F8D\u009Di\x12\u00DDQ*G\x06J``\n\u00BA\n\u0095r51\u0085B'\u0098\u00DB]\x00u\u00B8\u00CD\x1AOm\u00F9y\u00C3\u00FB9\u00BF\u00F7IX\x18\u00B8\u00BE\u00BE\u00E0.p\u00B5\u00E6V\x03\u00A1\u009D\u00BB\u00E6\u00DA\x0B'\u00B2\u00D0\u00EE\u00D7U-\u00B1\x0E\x06\\%<\u00E2\u00B9\u00E5#~N\u00B6y\u00C1\u0082\u00C1G\\\u0099X\u00BA\u00D8\x04\x1A\u00C5Y\x0E\u00D0\u009D^\u00C53`\u009ECf\u0083&C&\u008CP2\u00DEFf{\x06y\u008AD\u00AA\u00FE\u00E9<q^f\u00AB6\u00D0\u00D4\u008E|A\u00B0Dm\u009A\u00B8\x06:\x13@\u009B\x02!\u00D1Id\x10\u009A,\u00E5\u00A3]\u00D3\u0083>\u00A4\u00A4H\u00EB\u00EA>\u00E4\u009C9}\f\u00FF\u00A8\u00F9\x07\u00CD\u00B77\u00DA\u00DE\u00923Z\u00AE\u00F9w:]Fc\u00C5,\f\u00A8\u0099M8<\u008E[\x16hg;T: l=\x0E\x0B\x04\u009E_C\x02\u00E1\n\u00D4\u00AFGHSL6\u0096r\u00A4\u00833:5\u0098B\x1C*\u00DDR\u00F1\x14*d\u00A2H\u00A1\u00D2\x07\u00A9\u00F2\x14\u00B9T\u00D6\u00D5\u00C9\u00E8qTg\u00E8\u00FE\u0092\u00CCV\n\u00EC\u00A0\x1A\u00D0\x0E\x07\u00E2N\u00D5K6:\x12Qi\x1A\x13;0\u00B4\u0083\x17\u00DF\u00ACX\u0098\u00B2\u00F61\u008A\x1E?\u00A1\x07\u00D8\u00B0M.\u009CB\u00AFY\u00B2\u00B2\u00C6r\u00F1\u00B4\u00ED\u008A+\u008Atk\x02B\u00E06\u0090\u0088\x04?\u00D4\u0087b\u0091r\u0082r\u00AB\x03A2Q\u00BF\u008F\u00A42\u00F2!s\u009Br\u00F4N\u0098\u00E83\u0091\u00F9\x13%@\x1C\u00A9\x06\x1Fi\u00DAMY\fP\x0B\x04\u00A8?c\u00DAg\u00A4\u00ED\u00CC\u00E6\u00C1\"\x00R\u00B8\x0E\u00FDB\n\u00F3\u0080q\u00A9%\x03\u00B5\x1C8\u00DF\u0084-\x0BT\u00D6\u00D0\u00AB\u00C8\u00F1\u0083\u008AFe\u00DE\u00AAu\u00C8w\u00A8\f\u00953\u00E4#4\x07y\u0082\u00AA\u00DA\u00C7v\u00D9\u00E8\u0089\x13\u00CFG\u008F\u00B6=h\x1D\u0093\x13\u00B7p\u00B8\u00F1&\u00B7C\x0B\u00FDq[kF\u00F5c\u00B9\u0096\u00C6\u00BF|S\u00F5\u00B1a1\u00B3Y\u00E2\u008B\u0092W\u008DE\u00C6\u00F9\u0080\u00B0\u00DE\u00AC\u0099H5\u00A9\u00A8\x1E\u00F8\x1D\u0099\u00C4\u00B1\n\x1De}\u00C8\u00F6\u00CD#\u0090\u00C6+N\x06\x05f\u00F48TCOy\x07\x1AAS%\u00CDR\u00E9}\u00A8\u008D\u0099J\u00CF\u00A2\u00B2<PY6\u00A8x\u00A7d\u00F58\u0083M\u00D4D\u0084)\u0090\u0086\u00A2\x00|(\u0097#\u0093\x19\u00B0k\u00ABC/\u00DE+\u00EA\u00C1\"\u00A8 \x0Ew\u00F0\u0091\x18\u00B5\u00AF\u0084yq\u00FD\r7\u00A5q\u00CE\u009C~;7\u00F6d\u008C\u00C59\u00A5i\u00E1\u00D0o\u00C3o\u00EE\x7F\u00C8zu^la\u0089\u008F}>\u00E4\u00C1\u00C5\x00\x13(G\u00A3\u008A\x07n\u00B9,k<\u0097\u00ACqQ5\u00B9\u00B4\x11\x19w2\u00A4\f\x05\u0084\u00C5\u00CF\u00AE\x02\u00990\u00F9\x10\x12\u0083#\u009B\u008F\u00FC\u00DD\u008E\x16Oy\x02\u00BF\x1ELN\x19o\u00A0J\u008D\u00A4Y*\u00BD\x0F\u00D5\u00EF\u00A8\u00F4,\u0094\u0097A<IB^\u0086\u00F20\u00A9\u00EAq\x06\u008B\u00A8\x11gm\u0082#\u00C14\u00F1\x027\u00A02\u0090\x00\u00E1~\u00AA\u00D0\u00F9\u00BE\x18\u009D\u00D7\u009B%\x0B\u00A7VbL*k_\u0086\x17t\u00A4\u00C5\u00B7F\u00EB\u00EC\u00BC\u009D\x1B{F\u00DA\u00DA~\u00C9\u00CB\u00B8\u00AB@\u00E4\u00EDwJ\u00B9\u00822\u00EA;|\u00A3%g\u00F4t\u0093^\u009B:\u0081r4\u00BCg\u00AF\u00BC\u0082\u00DB]\x10=}\u00FA\u0084\u00A4?\u00CC\u00DB\x00\u009AT\u00A0\u00CD\x02\x16\x15U\u00FE2Q\x17+}\u00F3\u0091\x04\x14\x1B\u00DD\u0082\u00AD\u00A2\x19\x01\u008D)\u00CB\x01\u00B3\u00D2,\u0095t\u00CD\u0088'\u00A1\u00D2/\u0095^\u0088Z\x16\x18\u00F1\\*\u00DB\u00BB\u00C1\u008E\u00ABVFl\u00D0\u009AP\x04\u00952R\u0097\u00C6\x1F\x00[)Z\u0086\u00EB\u00D5\u00A3\u00F6\u00F5(n\u0080\x17\u008D}\x06\u0095%\u00D3>\u00EE^z\u00B7l\u00DD\u00C9/t?\u00BE\u0094G\u00A2\u00D1\u00B8}Y\u00F7\u00D2'\u00CFj\"\u00DA\u008D\x1E\u00E5F\u00C4\x1F\x10\u00B6i\u00DC\x0E\u00ED\b\u00D4\u00A5\u00FDA\u009DD\u00D7\u00C4\u0098\u00B2\u00E7Z#\u00EE\x0Bl\u00D3 \u00CA\rD\u00EA\u00F1\n\x1B\u00B3&E{\u0088\x01Q,~-\u00C8\u00A6\r\u009E]\u00B1ho\x14{ a\u00CC\x01\u00B3\u00D1gR\u00B5i3\x1B#\u00A0\u00BF\u00FF\u00C5\u00DA(\u00CA\u00CD`SI\u0092\u00FD'\u00AB\u0094\u00931\u00F9f\u00FCu\x12\u00B5\u00D7)\u00B2\u00EF\u00C0}\u00E1+\u00C7\u00EDr-\u0098\x17\u008B\x0B\u00AD\u00BDwRg\u00DC\u00E18g\u00BDf\u00FC\f\u00B3Lu\"\u00BE\x07\x04\x11\u00D1\u00C3Gh|5\u00F0\u009A\u00C5b\u00E4r\u00A4\u00D3\x01\u0081\x00 \u008FJ@\u00D0 P\x06\"\u00DC\u00D4\u00A2\u00BA y\u00F4J\u00D4\x042\u0092Z\x0E\u00A8\u00F4**i\u0099\nAT\u0088\u00A4\u0092\u008E\x19yp\u00AA<IU\u00BC\u0091\u00CA\u00E6\u00EE|\u00F38\u00C9z\x17R\u00A26tF~'\u0098<\u0092\u008F\u00C0t{~d\n\u00AC^\u00F7\u0080\u00B3l*WJ^\x16m;v\x05s8\u00F6\u00C1u~\u009Cs\u00DEQ\u00FA3=\u00D8 !]\u0090@\u00CB\x04\u00DF#a\u00FEy\f\u0082\x1B\u00C0\u00B0\u0094\\\u00F2\u0089\x05S+\u00F84F\\+\u00AA\u0090T\u00AC8\u0081KG\u0085\u00E0_\u00B0K\u0086\x17-\x1A\x1Cd\u00C5\u00E8}I-\x07\x12\u00E9U\x12Y\x06\x00o\u0093H\u009F\u0082\u00EB\u00ABb\x0F\u00A8<8U\u0096\x05*\u00DEHes7\u0098<N\u008D\u00E4C\u00C7\u00C110\u0095K\u00EA\u00CA\u00E2\u00A2\u0095\u008B\u00B6X9\u00E7EZp\u0086\u0094\u00A7X,\u009B\u008F\x18\u00FFfd\u00CC\u0098\u00A8}\x1F\u00E2.\u00D3\u0081\u009F\u00AD\u00DE\u00E1\u00FA\u00FA\u0082\u00D9\u00F6O\u00DE6-\u00FC\u00CA\u00AE\u00D5P\u0087\u00DB\u00AA\x05_x\u00F1{\u00D6\u00E1\u00C3gG\u00FD\x01]\u0080`\x19u\u00F9\u00D1\u00C8\u00FE\u0083\u00B92/P\u00AE\u00FB\u00E1\u00C6\u00A1\u00FA?\u00A3B\u00C4\x17\u009E\u00EB+\u00F7\x16D\u00FE\u00F9N\u008B$\u00A2'\u00B5\u00E9\u00C3\u00FEGp\u008F\u0095\"U!\u0085\x17\u00E9o\u00B0~k\x03\x12Mo\x10e\u00A6y\x1C\u0095\u00DF\u008B\u008470%\u00D5R\u00C5\x12P\u00D9\u00C4\u00A9\u00F4GT\u00BFD\u00AD\u00AFU\u00B1\x0F(\u008F%\u00B3\x1C\x18\f\u00DE\u0086\u00C6C\u00A3D\u00F9\x1B\u00EA\u00AB\u00A3\u0091\u0098i\u0098VI\u00E2\u00B7\x01\u00D1(\u009C\x15\u0092:\u00DE$\u00DA\u00F7\u00F9q\u009E\u0084\u00F33\u00DC\x7FF\u00F7\u00CDA<\n\u00E7qP<h\u00FD\u009A\x1E+z[C\u009FH\u009E\u00C1\u00BF\u00BE\u00D2\u00C5\x7F}\u00FAG\u00FE?\u00C0?u-zl\u0099\u0082\u00C7\u00C1\u0091J=(b)\u0094A\x1D\u00CC\u00A7xQ,\u00B93\u00C8\x1F\u0087\u00A1\u0098l4bk\u00D2<N\x7F\u00FDcR\u008D\u00C3\u00D6\u00DFxl\u00A9\u00F8\u00E3\f$\u00F9\x12\u00882e\u0091t\u00EA\u0088\x17)\u00AEKm\u00D8\u00C0\u00BE\r\\\t\u00DA\u00918\u0095\u00FA\u008E\u00D4!IR;\u00B9n\u00A6}_r:\u00B3\u00B8\u008Btxw\u00D3=\u00CE\u00CF\u00CDZ\u00EE\u00BCk\u00E6\u00FC`\u00FD\u00F6\x1B\u00C2M\x7F\x7F\u00C8b\u00B3\u00FF4\u00F2\u00CE\u00A1\u00ED`\u00ED\u00AC\u00B7\u008DD\u00BA\u00B1\x1B\x02\u00A7\u00EE\u00EF.^\u0096\u00B9\u00B4:\u00CEc\x14\u00C8:n\u00EC\f\x1E~\u008A\x1F\u00D3\u00BA\x1F\u00FD\u00DEV\u00AD\u00B7\u00D7\u009F\u00F5\u00A3\u00A7h\u0098\u00A8*t\u00CC\u00A4\ne\x1A)\b\u0090\u00A8\n\u00A1\x0B\u00EE\x07\u00EAB\x04\u00D6\u00CA$\u008F\u00D3R\x1E'\u0099\u00E8\u00FFf,\x06TR9*ES\u00F16FR\u00B6T\u00FCq\x06S\u00AA6\u00E4~\u00BA\u00C7\u00A7@\x10\u00EE\u00B1\u00C9%_8vt\u00E7\u00BC\x07\u00DE\u00D5\u0084w'G\u009C\u00C0s\u00CF\u00EF\u00E6\u00D7\x01\u0099x\u009B\u00E0\u00F6?\x1F\u00E2e4\n\u00A8\u00C8k\\Zw\u00EE\u00C1\u0085-*\u00CF\u00D1A\u00F8y\u0085\u00A4-\u0095\u00BE\x1B\u0093A\u0082dc<\u00F7W\u008A6P\x1E\u00A5\u00E7\x0Bq\u0086,\u00F1\u00A0\x1C\u00D1\u00D6w71!0`\u008CM\u0081\u00FFV\u00E0\u00BCs\u00EA\u00D4aw\x7F\u0096\u00C7\u0095~\u00D7qS\t\x0B\u009F:=\u00DEr\u00D9e\u008Fk\u009D\u00E7\u00B6\u00D8o(\u009A\x1F\u00F8\u00D9/o\u00B5\u00DFv\u00CBG\u00F4\u00F8\x04/\u00BCx\u00C4\u00F5`\u00A5\u00CE\u00EBq;8.e\u00B3\x15M\u00BE?\u00B4\u00F3\u00D5n\u00EB\u00B8\u00B1\u008Fg,zd\u00F9y|\u0086\u00FD\u00F1\u0085\u00D2\u00F9\u00B7\u008C\u008C\f7O\u00AF/\u00BA\u00DE\u00EDvg\u00B0\x11#\u0086\u00BB;::\u0095\u00E9\u00B1c\u00C7\u00DC\u00EF\u00B6\u00B62\u009B\u00DD\u00E6\x16\u00FD\u00C4\u00A5\u00D0\x1F\u00E4\u00A1_\u00C64=\x1F\u0089D\u00DC6\u009B\u008D}4\u00F7\u00A3\u00EE[n\u00BD5\u00D6/\\\x0F\u0085\u0082\u00D2~a<h\u0097\u0091\u00E1\u008E\x1B/\x10\b\u00C4\u00F5\x1F\n\u0085\u00DC\x0E\u0087#\u00D6\x1F\u009A\u00EF\u00C8\u00F4\u00C21I\u0082q\u00CF\fm\u00DB\x1E'\t\x13Nl?\u00D7%j\u00A1\u00F0\u00D3\u00F6\u00C9\u0085\u00A3\u00ADW]\u00C9\\_\u00FD\u00CA\u00ED\u00E1\u00FD\x07\u00B7\u00F0\u00A3\u009Ds\u00FA4\x1E\u00D5F\x0F-\u00C5-\x028\x05\u00FF\u00B0\u00E5\u00BD\u00E8\u0089\u0093\u00A3\x1D\u00A5\u009Fl\t\u00BF\u00B9\u00DF\u00CD\x17T`\u00F5\u00BA\x1F\u00BB\x16\u00CC;\u009F\u008B\u00A6\u00BF\u00A4\u008B\u00E4{zz\u00FC<]\u00FE\u00A3g\u00FC\u00B9y\u00BA\u00D9\u009E_\u00F4+M\u00D7\u00ADY\u00EB\u00AF^\u00B2\u0084E\u00C2\x11\u00E9u\u00E8\x0F\u00F2\u00D0\u00EF\u00FB\u00ED\u00A6\u00B2\u00ED\u008D\u008Dz\u00BEt\u00EAT\x7F\u00E9\u00D4\u00A9\u00A4\u00DF\u00A9\u00CAqa<\u00DA\x0E\u00C6\u009B5c\u0086\u009EB\u00FF\u00BF}\u00EE7\u00FA<i{\u00C6\u00D8\u00D9\u00F4\u00C21I\x16\u009B\u00BD4r\u00A4\u00B5[\x15,PH\u00C1nf\u00EF\u00C7\u0095\u00D6\u00C2\x7Fkz\u0085\u0085Bo[\u00C7\u008D\u00B5\u00F5>\x17\u00EFZ\u00C3CF\u00F1\u00E87\u009C\x1F\n\u009D<y\u0095\u00EB\u00BE{^\u00E2q\x0B\\\x0B\u00E6=t\u00C1\u00DC\u00B09\u008AC\x1C\u00D8\u0091a\u00A7\u00F6\u00FB\u00FDn\u00B7\u00DB\x1DKa'7\u00AA\x0F;<E\x1A\u00BA\u00F3C;\u00DA/\u00B4\x03d\x03\u00C4\x01D\u00A1\u00F3\u00A2\u00E3\x01\x12QD\u0083\u00F24\u00E2$A\u00D6+\u00C7\u008D\u008F\u00B4\u00F8\u00FE\u00E6\u0098r\u009B\u00A9F\u00D6\u00ECQ\u0087\u0086\u00DDw\u00CF\u00D7\u008C\u00EA\x05_z\u00F99\u009Er\x01\u0083\u0093\u00C4-\u00B8\b(\x0Eq`G\u00A6;5\u00A4\u00B0\u0093\u009B\u00ADO\u0091\u0086\u00EE\u00FC\u00D0\u008E\u00F6\x1B\u00DF\u00EE\u0083~\x01QT\u0088D\x11\u0093\"\x13J\u00D3\u0088c\u009A\u009C\u00CE\u0092h\u00CB\u00E1'\x06\u00B2\u00CB\u00C0\u00EAu+\u0087}\u00E9\u00F3\u00B7\u00F7\u00FE\u00EE\u00F7\u00DF>O\u00FF\u0089m\u00A0I\u008A8\u0094G\u00A1<\x06\u00AD\x0F;; \u0084\u00D1\u00CE\x0F\b\u00E4\u00B9\u00DC\x13\u00C7\u00AB@=z\u00DD\u0088\u00D7\u00A2\u00E3\u00C1|)\u00A2\u00A1\u00FB\x19y\u00C9K\u00CA\u00CC\u00FE\u00C06\u00CD\u00CC\x0F\u00A4j\u0089\u00EA\u00FA\x7F\u00B1vf\u00E8\u008D}\u00E7\x02\u00EB7-\u00BE\u0088\u009FKc\")\u0095Y\u00A9V\u00AA\u00FA\u009C$\u00A4`\x03\u00DD\u00CF\u0092\u00C1rd\x1BR\u00E4\x7Fr\u00F9\u00F7\u00AD\u00C3\u0087?\u009B\u00CC=\u0085\u00F7\x1F8\u00A6\u00BA\u00C6-\u00A9\u00AD\u00A3G\u00AD\u0089\u00BC\u00F5\u00F6\u00CB\u00C3\u00E6|\u00B1\u00FA\"~V\x03\u00828T\u009A\u0086y\x11\u009C\u00F2\u009D\u009FI\u0090\x05\u00CA\u00A1\x1Fz\u00DD\u00A8\x1F@\x1A\u008E0\u00B8\x1E\u00F0>X\u00DA&\u00C6I\u00F38f(\u00D2\u00D6V,QD&$\u008B\u00DDqBv\u009D/\x1A{\u00C9\u008D\u00DB\u00A3\u00A7N\u00AF\u00BC\u00C8\x17\rS\u00F18T*Fy\x13Z_%\u00DDR\u00F1H\u00C0{P)\x1B\u00E5MTR6\u0095TN\u00C5kQi[\u009A\u00C71I\u00C9.\u009AD\u00A4E#\u009B\u00A3\u00A7N\u00AF\x1B\x02\u008B\u0086\u00A9\x10\u0087\u00F2$T*\u00A5\u0092\u00AA\x01OB\u00A5a*=\u008EY\u00E9\x1C\u00ED\x07#L2R\u00B9\u00B4\x1E\u00E7C\"nFc\u00BB\u00F6c\u00DEas\u00BE\u00F8\u00D2\x10\u00B9%)\u00E2$\u0090FI\x11\x07K\u00C7\u00EE\u00FA\u00ECg\u00FBH\u00C3Tz\x1C\u00B3\u00D29\u0095\u00D4,Y\u00A9\\Z\u008F3\u00C8$b\x11\u00C4\u00E9{t\u00BE&;\u00BB\u00D2\u00B5`\u009E)/\u00D1\u008B\u0084\u00E2\x10G\u00A5'Q\u00E9M(BQi\x18\u00B50P\u00F1R*\u00DE\u008A\"\x07\u00E5\u00A9T\u00E3Rd\u0094H\u00DD\u00D2\u0088s\u00BE\u00C8\u00AC[\u00F5EFq\u0088c\u00A4'1\u00D2\u00E4cD\u00C1\u00C8\u00A3\u00E2\u008D\u008Cx+\u008A\x1C\u0094\u00A7R\u008DK\u0091Qb\x11\u0091F\u009C\u00C1\u00A0\u00E8\u00993\u0091\u00A1wWR\u00D2\x11G\u00C5\u00D3\x18\u00A5ToB\u00A5\\\u00A9\u00E6i\u00FF\u00B4\x1F*\u00EDS]O\u00D0.\u00AD\u00C7\x19\u008C\u00DF\u00B9o|\u00AB\u00F1\x12\u00B9\u00D7\x01\u00F1\u00C7\u00B9\b\u00D3\u00B4\x1EgP\u00C8\u00E9\x1C\u008C\u00A0\u00F1i\u00BA\u0080(}T\x1B\x04\u00B2\u008E\x18a\u00E6\u00DF\u00AD\x0F\x05*\u00BDD\u00EE\u00B3\x0F\u00A5\x11g\x10H\x12H0MC\u0089\x18c\u00FF\x0F\u00FE-\u00DFU$*\u00B5\u00C5\x00\x00\x00\x00IEND\u00AEB`\u0082"));
    // create a data folder for your script in the user's application data
    var myDataFolder = new Folder( Folder.userData.absoluteURI + "/CoriusScripts" );
    // make certain the folder exists
    myDataFolder.create();
    // write the image file
    var myFile = new File( myDataFolder.absoluteURI + "/CoriusLogo.png" );
    myFile.encoding = "BINARY";
    myFile.open( "w" );
    myFile.write( myBinary );
    myFile.close();
    // there is now a valid png image in your script's user data folder
}