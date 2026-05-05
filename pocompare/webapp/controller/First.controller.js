sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/export/Spreadsheet",
    "com/sap/pocompare/model/formatter",
    "sap/ui/core/Fragment"
], (Controller,JSONModel,MessageToast,MessageBox,Spreadsheet,formatter,Fragment) => {
    "use strict";
    var that=this;

    return Controller.extend("com.sap.pocompare.controller.First", {
        formattter:formatter,
        onInit() {
            // Initialize the model that will hold our Excel data
            var oModel = new JSONModel({
                data: []
            });
            this.getOwnerComponent().setModel(oModel, "excelModel");
            this.lineItemFlag=true
        },

        // Triggered when a file is selected via the FileUploader
        onFileChange: function (oEvent) {
            var aFiles = oEvent.getParameter("files");
            if (aFiles && aFiles.length > 0) {
                var oFile = aFiles[0];
                this._loadExternalLibrary("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js").then(function() {
                    this._readExcel(oFile);
                }.bind(this))
                .catch(function() {
                    sap.m.MessageToast.show("Failed to load the Excel library from CDN.");
                });
            }
        },
        onClearFile: function () {
            this.byId("excelUploader").clear();
            this.getView().getModel("excelModel").setProperty("/data", []);
            // this._headerFB.setVisible(false)
        },

        // Helper function using SheetJS (XLSX)
        _readExcel: function (file) {
            var that = this;
            var reader = new FileReader();

            reader.onload = function (e) {
                var data = e.target.result;
                
                // Parse the workbook
                var workbook = XLSX.read(data, { type: 'binary', cellDates: true });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];
                
                // Get headers only (first row)
                var headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0];

                const REQUIRED_HEADERS = [
                    "Vendor",
                    "Material",
                    "Material Desc",
                    "PO Number", 
                    "Line Item", 
                    "Quantity", 
                    "Delivery Date",
                    "Schedule Line Category"
                ];
                // Convert to JSON
                var jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    raw: false,
                    dateNF: 'yyyy-mm-dd'
                });

                // Map columns
                var formattedData = jsonData.map(function(row) {
                    return {
                        //Level 1 Start
                        VendorCode: row["Vendor Code"],
                        VendorName: row["Vendor Name"]||row["Vendor name"],
                        PONumber: row["PO Number"] || row["PONumber"]||row["PO/PR No."],
                        PODate: new Date(row["PO date"])?.toISOString()?.split('T')[0],
                        PanelVisible:false,
                        //Level 2 Start
                        // LineItem: row["Line Item"] || row["LineItem"],
                        POLineItem: row["PO Line Item"] || row["LineItem"],
                        Material: row["Material"],
                        MaterialDesc: row["Material Description"],
                        POQuantity: row["PO Quantity"],
                        UOM: row["Unit of Measure"],
                        DeliveryDate: new Date(row["Delivery Date"])?.toISOString()?.split('T')[0],
                        NetPrice: row["Net Price"],
                        Currency: row["Currency"],
                        Per: row["Per"],
                        MaterialGroup: row["Material Group"],
                        Plant : row["Plant"],
                        StorageLocation : row["Storage Location"],
                        //Level 3 Start
                        ConfirmationCategory:row["Confirmation category"]||row["Confirmation Category"],
                        FDDCategory:row["Fdelivery Date category"]||row["FDelivery Date Category"],
                        Quantity: row["Quantity"],
                        Reference: row["Reference"],
                        CreationDate: row["Created on Date"],
                        InboundDelivery: row["Inbound Delivery"],
                        Item: row["Item"],
                        HLItem: row["Higher Level Item"],
                        Batch: row["Batch"],
                        QtyReduced: row["Quantity Reduced"],
                        MRPRelevant: row["MRP relevant"]||row["MRP Relevent"],
                        MRPMaterial: row["MPN Material"]||row["MPN material"],
                        CreationIndicator: row["Creation Indicator"]||row["Creation Indicator"],
                        SequenceNumber: row["Sequence Number"]||row["Sequence number"],
                        //Status 4
                        StatusCode:"1",
                        Status:"New",
                        StatusState:formatter.stateFormatter("1"),
                        StatusMsg:formatter.statusDescription("1"),
                    };
                });

                let newData= that.transformDataForTreeTable(formattedData)

                that.getOwnerComponent().getModel("excelModel").setProperty("/data", newData);
                // that._headerFB.setVisible(true)
                MessageToast.show("Excel loaded for preview.");
            };
            reader.onerror = function (ex) {
                MessageBox.error("Error reading the Excel file.");
            };
            reader.readAsBinaryString(file);
        },
        transformDataForTreeTable: function(rawJsonString) {
            const flatData = rawJsonString;
            const groupedData = {};
            flatData.forEach((item) => {
                // ==========================================
                // LEVEL 1: Header Level Grouping
                // Key based on PONumber, VendorCode, VendorName, PODate
                // ==========================================
                const level1Key = `${item.PONumber}_${item.VendorCode}_${item.VendorName}_${item.PODate}`;
                // Initialize the Level 1 group if it doesn't exist
                if (!groupedData[level1Key]) {
                    groupedData[level1Key] = {
                        PONumber: item.PONumber,
                        VendorCode: item.VendorCode,
                        VendorName: item.VendorName,
                        PODate: item.PODate,
                        PanelVisible: item.PanelVisible,
                        DocumentDate: item.DocumentDate,
                        // We use a temporary map to easily group Level 2 items without duplicating them
                        _level2ItemsMap: {}, 
                        children: [] 
                    };
                }
                // ==========================================
                // LEVEL 2: Line Item Level Grouping
                // Key based on POLineItem (to group multiple sequences under one line)
                // ==========================================
                const level2Key = `${item.POLineItem}`;
                // Initialize Level 2 if it doesn't exist under this specific Level 1 node
                if (!groupedData[level1Key]._level2ItemsMap[level2Key]) {
                    groupedData[level1Key]._level2ItemsMap[level2Key] = {
                        // Showing upper-level fields + Line item info
                        PONumber: item.PONumber,
                        VendorCode: item.VendorCode,
                        VendorName: item.VendorName,
                        PODate: item.PODate,
                        POLineItem: item.POLineItem,
                        LineItemNumber: item.POLineItem, // Using actual PO Line Item instead of mock
                        Material: item.Material,
                        MaterialDesc: item.MaterialDesc,
                        POQuantity: item.POQuantity,
                        UOM: item.UOM,
                        NetPrice: item.NetPrice,
                        Currency: item.Currency,
                        Per: item.Per,
                        MaterialGroup: item.MaterialGroup,
                        Plant : item.Plant,
                        StorageLocation : item.StorageLocation,
                        NextPanelVisible: false,
                        DeliveryDate: item.DeliveryDate,
                        children: [] // This will hold the Level 3 sequences

                    };
                }
                // ==========================================
                // LEVEL 3: Sequence Level Details
                // ==========================================
                const level3Node = {
                    PONumber: item.PONumber, 
                    VendorCode: item.VendorCode,
                    VendorName: item.VendorName,
                    PODate: item.PODate,
                    POLineItem: item.POLineItem,
                    LineItemNumber: item.POLineItem,
                    SequenceNumber: item.SequenceNumber, // The differentiator for Level 3
                    DeliveryDate: item.DeliveryDate,
                    ConfirmationCategory: item.ConfirmationCategory,
                    FDDCategory: item.FDDCategory,
                    Quantity: item.Quantity,
                    Reference: item.Reference,
                    CreationDate: item.CreationDate,
                    InboundDelivery: item.InboundDelivery,
                    Item: item.Item,
                    HLItem: item.HLItem,
                    Batch: item.Batch,
                    QtyReduced: item.QtyReduced,
                    MRPRelevant: item.MRPRelevant,
                    MRPMaterial: item.MRPMaterial,
                    CreationIndicator: item.CreationIndicator,
                    Status: 1,
                    newRecFlag:false,
                    StatusState: formatter.stateFormatter("1"),
                    StatusMsg: formatter.statusDescription("1")
                };
                // Push the Level 3 detail into the correct Level 2 node's children array
                groupedData[level1Key]._level2ItemsMap[level2Key].children.push(level3Node);
            });
            // ==========================================
            // FINAL CLEANUP: Convert Maps back to Arrays
            // ==========================================
            // UI5 JSONModels need standard arrays for "children", not object maps.
            const treeData = Object.values(groupedData).map(level1Node => {
                // Extract Level 2 items from the temporary map into the children array
                level1Node.children = Object.values(level1Node._level2ItemsMap);
                // Remove the temporary map so it doesn't clutter your UI5 model
                delete level1Node._level2ItemsMap; 
                return level1Node;
            });

            return treeData;
        },
        onDownloadTemplate: function () {
            // Fetch the specific column configuration for the PO template
            var aCols = this._createColumnConfig();

            // Configure the export settings
            var oSettings = {
                workbook: {
                    columns: aCols
                },
                dataSource: [], 
                fileName: 'PO_Template.xlsx',
                worker: false 
            };

            // Generate and trigger the download of the spreadsheet
            var oSheet = new Spreadsheet(oSettings);
            oSheet.build().finally(function() {
                oSheet.destroy();
            });
        },
        onSaveTemplate: function (oEvent) {
            var treeData = this.getView().getModel("excelModel").getProperty("/data");
            var flatExcelData = this.transformTreeToFlatData(treeData);
            var aCols = [
                // Level 1
                { label: 'Vendor Code', property: 'VendorCode', type: 'string' },
                { label: 'Vendor name', property: 'VendorName', type: 'string' },
                { label: 'PO Number', property: 'PONumber', type: 'string' },
                { label: 'PO date', property: 'PODate', type: 'string' },
                // Level 2
                { label: 'PO Line Item', property: 'POLineItem', type: 'string' },
                { label: 'Material', property: 'Material', type: 'string' },
                { label: 'Material Description', property: 'MaterialDesc', type: 'string' },
                { label: 'PO Quantity', property: 'POQuantity', type: 'string' },
                { label: 'Unit of Measure', property: 'UOM', type: 'string' },
                { label: 'Delivery Date', property: 'DeliveryDate', type: 'string' },
                { label: 'Net Price', property: 'NetPrice', type: 'string' },
                { label: 'Currency', property: 'Currency', type: 'string' },
                { label: 'Per', property: 'Per', type: 'string' },
                { label: 'Material Group', property: 'MaterialGroup', type: 'string' },
                { label: 'Plant', property: 'Plant', type: 'string' },
                { label: 'Storage Location', property: 'StorageLocation', type: 'string' },
                // Level 3
                { label: 'Confirmation category', property: 'ConfirmationCategory', type: 'string' },
                { label: 'Fdelivery Date Category', property: 'FDDCategory', type: 'string' },
                { label: 'Quantity', property: 'Quantity', type: 'string' },
                { label: 'Reference', property: 'Reference', type: 'string' },
                { label: 'Created on Date', property: 'CreationDate', type: 'string' },
                { label: 'Inbound Delivery', property: 'InboundDelivery', type: 'string' },
                { label: 'Item', property: 'Item', type: 'string' },
                { label: 'Higher Level Item', property: 'HLItem', type: 'string' },
                { label: 'Batch', property: 'Batch', type: 'string' },
                { label: 'Quantity Reduced', property: 'QtyReduced', type: 'string' },
                { label: 'MRP relevant', property: 'MRPRelevant', type: 'string' },
                { label: 'MPN Material', property: 'MRPMaterial', type: 'string' },
                { label: 'Creation Indicator', property: 'CreationIndicator', type: 'string' },
                { label: 'Sequence Number', property: 'SequenceNumber', type: 'string' },
                // Level 4
                { label: 'Status', property: 'Status', type: 'string' },
                { label: 'StatusMsg', property: 'StatusMsg', type: 'string' },
            ];
            // 3. Configure and start the export
            var oSettings = {
                workbook: { columns: aCols },
                dataSource: flatExcelData,
                fileName: 'BTP_Harman_POC_Template.xlsx'
            };

            var oSheet = new Spreadsheet(oSettings);
            oSheet.build().finally(function() {
                oSheet.destroy();
            });
        },
        transformTreeToFlatData:function(treeData) {
            const flatData = [];
            // Loop through Level 1 (The PO / Vendor groups)
            treeData.forEach(groupNode => {
                // Extract the parent-level data
                const poNumber = groupNode.PONumber;
                const vendorCode = groupNode.VendorCode;
                const vendorName = groupNode.VendorName;
                const poDate = groupNode.PODate;
                // Check if this group has children (Level 2 Line Items)
                if (groupNode.children && groupNode.children.length > 0) {
                    // Loop through Level 2
                    groupNode.children.forEach(itemNode => {
                        if (itemNode.children && itemNode.children.length > 0) {
                            // Loop through Level 3
                            itemNode.children.forEach(subItemNode => {
                                const flatRow = {
                                    PONumber: poNumber,
                                    VendorCode: vendorCode,
                                    VendorName: vendorName,
                                    PODate:poDate,
                                    //Level 2 Start
                                    POLineItem: itemNode.POLineItem,
                                    Material: itemNode.Material,
                                    MaterialDesc: itemNode.MaterialDesc,
                                    POQuantity: itemNode.POQuantity,
                                    UOM: itemNode.UOM,
                                    DeliveryDate: itemNode.DeliveryDate,
                                    NetPrice: itemNode.NetPrice,
                                    Currency: itemNode.Currency,
                                    Per: itemNode.Per,
                                    MaterialGroup:itemNode.MaterialGroup,
                                    Plant : itemNode.Plant,
                                    StorageLocation : itemNode.StorageLocation,
                                    //Level 3 Start
                                    ConfirmationCategory:subItemNode.ConfirmationCategory,
                                    FDDCategory:subItemNode.FDDCategory,
                                    Quantity: subItemNode.Quantity,
                                    Reference: subItemNode.Reference,
                                    CreationDate: subItemNode.CreationDate,
                                    InboundDelivery: subItemNode.InboundDelivery,
                                    Item: subItemNode.Item,
                                    HLItem: subItemNode.HLItem,
                                    Batch: subItemNode.Batch,
                                    QtyReduced: subItemNode.QtyReduced,
                                    MRPRelevant: subItemNode.MRPRelevant,
                                    MRPMaterial: subItemNode.MRPMaterial,
                                    CreationIndicator: subItemNode.CreationIndicator,
                                    SequenceNumber: subItemNode.SequenceNumber,
                                    Status:subItemNode.Status,
                                    StatusMsg:subItemNode.StatusMsg
                                };
                                flatData.push(flatRow);
                            });
                        }
                    });
                }
            });

            return flatData;
        },

        _createColumnConfig: function() {
            return [
                {
                    label: 'PO Number',
                    property: 'poNumber',
                    type: 'string', 
                    width: 20
                },
                {
                    label: 'Line Item',
                    property: 'lineItem',
                    type: 'string',
                    width: 20
                },
                {
                    label: 'Quantity',
                    property: 'quantity',
                    type: 'number',
                    width: 15
                },
                {
                    label: 'Delivery Date',
                    property: 'deliveryDate',
                    type: 'date',
                    format: 'yyyy-MM-dd', 
                    width: 20
                }
            ];
        },
        onShowExpanded:function(oEvent){

            if(this.lineItemFlag){
                let oSource=oEvent.getSource()
                let oBindingContext=oSource.getBindingContext("excelModel")
                let oPath=oBindingContext.getPath()
                let oPanelVisiblePath=oPath+"/PanelVisible"

                let bVisiblePath=this.getOwnerComponent().getModel("excelModel").getProperty(oPanelVisiblePath)
                if(bVisiblePath){
                    this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath,false)
                    // oSource.setIcon("sap-icon://dropdown")
                    
                }else{
                    this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath,true)
                    // oSource.setIcon("sap-icon://slim-arrow-up")
                }
            }
            this.convertLIFlag(true)
        },
        convertLIFlag:function(bFlag){
            this.lineItemFlag=bFlag
        },
        onSubShowExpanded:function(oEvent){
            let oSource=oEvent.getSource()
            let oBindingContext=oSource.getBindingContext("excelModel")
            let oPath=oBindingContext.getPath()
            let oPanelVisiblePath=oPath+"/NextPanelVisible"

            let bVisiblePath=this.getOwnerComponent().getModel("excelModel").getProperty(oPanelVisiblePath)
            if(bVisiblePath){
                this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath,false)
                oSource.setIcon("sap-icon://dropdown")
            }else{
             this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath,true)
             oSource.setIcon("sap-icon://slim-arrow-up")
            }
             
        },
        onAddVendorRow: function (oEvent) {
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("excelModel");
            var sOuterRowPath = oContext.getPath(); // e.g., "/data/0"
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var aVendorInputTable = oModel.getProperty(sOuterRowPath + "/children");
            var iLINumber;
            if(aVendorInputTable.length!=0){
                var iMaxTabLength=aVendorInputTable.length
                var iMaxLINumber=aVendorInputTable[iMaxTabLength-1].LineItemNumber
                iLINumber=(Number(iMaxLINumber)+10).toString()
            }else{
                iLINumber=10
            }
            aVendorInputTable.push({
                LineItemNumber: iLINumber,
                Quantity: 0,
                DeliveryDate: "",
                Status:1,
                StatusMsg:formatter.stateFormatter("1"),
                StatusState:formatter.statusDescription("1")
            });
            oModel.setProperty(sOuterRowPath + "/children", aVendorInputTable);
        },
        onAddVendorRowSP: function (oEvent) {
            this.convertLIFlag(false)
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("alSidePanel");
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var oSPModel = this.getOwnerComponent().getModel("alSidePanel");
            var aVendorInputTable=oModel.getProperty(this.sCurrentPath)
            var iSNum;
            var oVendorObject;
            if(aVendorInputTable.length!=0){
                oVendorObject=aVendorInputTable[0]
                var iMaxTabLength=aVendorInputTable.length
                var iMaxSN=aVendorInputTable[iMaxTabLength-1].SequenceNumber
                iSNum=(Number(iMaxSN)+1).toString()
            }else{
                oVendorObject={}
                iSNum=1
            }
            aVendorInputTable.push({
                // LineItemNumber: iLINumber,
                ConfirmationCategory:oVendorObject?.ConfirmationCategory,
                FDDCategory:oVendorObject?.FDDCategory,
                Quantity: oVendorObject?.Quantity,
                Reference: oVendorObject?.Reference,
                CreationDate: oVendorObject?.CreationDate,
                InboundDelivery: oVendorObject?.InboundDelivery,
                Item:oVendorObject?.Item,
                HLItem: oVendorObject?.HLItem,
                Batch: oVendorObject?.Batch,
                QtyReduced: oVendorObject?.QtyReduced,
                MRPRelevant: oVendorObject?.MRPRelevant,
                MRPMaterial: oVendorObject?.MRPMaterial,
                CreationIndicator: oVendorObject?.CreationIndicator,
                SequenceNumber: iSNum,
                newRecFlag:true,
                StatusMsg:formatter.statusDescription("1"),
                StatusState:formatter.stateFormatter("1")
            });
            oSPModel.refresh(true);
            oSPModel.setProperty("/",aVendorInputTable)
            oModel.setProperty(this.sCurrentPath, aVendorInputTable);
        },
        onDeleteVendorRow: function (oEvent) {
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("excelModel");
            var sInnerRowPath = oContext.getPath(); 
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var sParentArrayPath = sInnerRowPath.substring(0, sInnerRowPath.lastIndexOf("/")); 
            var sRowIndex = sInnerRowPath.substring(sInnerRowPath.lastIndexOf("/") + 1); 
            var aVendorInputTable = oModel.getProperty(sParentArrayPath);
            aVendorInputTable.splice(parseInt(sRowIndex, 10), 1); 
            
            oModel.setProperty(sParentArrayPath, aVendorInputTable);
        },
        onDeleteVendorRowSP: function (oEvent) {
            this.convertLIFlag(false)
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("alSidePanel");
            var sInnerRowPath = oContext.getPath(); 
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var oModelData = this.getOwnerComponent().getModel("excelModel").getProperty(this.sCurrentPath);
            var oSPModel = this.getOwnerComponent().getModel("alSidePanel");
            var oSPModelData = this.getOwnerComponent().getModel("alSidePanel").getProperty("/");
            var aVendorInputTable = oSPModel.getProperty(sInnerRowPath);
            var iLINumber=aVendorInputTable?.LineItemNumber

            let iIndex = oModelData.findIndex(item => item.LineItemNumber == iLINumber);
            let iNewIndex = oSPModelData.findIndex(item => item.LineItemNumber == iLINumber);
            if (iIndex > -1) {
                oModelData.splice(iIndex, 1);
                oModel.refresh(); 
                oSPModel.refresh(); 
            }
        },
        onApprovePress: function (oEvent) {
            MessageToast.show("Line Item Approved!");
            this._closeDialog();
        },
        onRejectPress: function (oEvent) {
            MessageToast.show("Line Item Rejected!");
            this._closeDialog();
        },
        onActionTaken:function(oEvent){
            let sButtonText=oEvent.getSource().getText()
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("alSidePanel");
            var sInnerRowPath = oContext.getPath();
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var oSPModel = this.getOwnerComponent().getModel("alSidePanel");
            let sChangePropertyStatus=sInnerRowPath+"/Status"
            let sChangePropertyStatusMsg=sInnerRowPath+"/StatusMsg"
            let sChangePropertyStatusState=sInnerRowPath+"/StatusState"
            if(sButtonText=="Approve"){
                oSPModel.setProperty(sChangePropertyStatus,2);
                oSPModel.setProperty(sChangePropertyStatusMsg,formatter.statusDescription("2"));
                oSPModel.setProperty(sChangePropertyStatusState,formatter.stateFormatter("2"));
            }else if(sButtonText=="Reject"){
                oSPModel.setProperty(sChangePropertyStatus,3);
                oSPModel.setProperty(sChangePropertyStatusMsg,formatter.statusDescription("3"));
                oSPModel.setProperty(sChangePropertyStatusState,formatter.stateFormatter("3"));
            }
        },
        _closeDialog: function () {
            if (this._pCompareDialog) {
                this._pCompareDialog.then(function (oDialog) {
                    oDialog.close();
                });
            }
        },
        // onConfirmationRowPress: function (oEvent) {   
        //     var oItem = oEvent.getSource();
        //     var oCtx  = oItem.getBindingContext("excelModel");
        //     var oModel=this.getOwnerComponent().getModel("excelModel")
        //     var oCPath=oCtx.getPath()+"/children"
        //     this.sCurrentPath=oCPath;
        //     var oExcelTabData=oCtx.getModel("excelData").getProperty(oCPath)
        //     var oSPJSONModel=new JSONModel(oExcelTabData)
        //     // this.getView().byId("idPOLIDataTable").bindItems(oCPath)
        //     this.getOwnerComponent().setModel(oSPJSONModel,"alSidePanel")
        // },
        onConfirmationRowPress: function (oEvent) {   
            // oEvent.cancelBubble()
            // oEvent.getParameter("event").stopPropagation();
            // 1. Get the current pressed row
            this.convertLIFlag(false)
            var oCurrentItem = oEvent.getSource();

            // 2. Manage the highlight logic
            if (this.prevRecord) {
                this.prevRecord.removeStyleClass("myCustomHighlight");
            }
            
            oCurrentItem.addStyleClass("myCustomHighlight");
            
            // 3. Store this item as the "previous" for the next time a row is pressed
            this.prevRecord = oCurrentItem;
            var oItem = oEvent.getSource();
            var oCtx  = oItem.getBindingContext("excelModel");
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var oCPath = oCtx.getPath() + "/children";
            this.sCurrentPath = oCPath;
            var oExcelTabData = oModel.getProperty(oCPath);
            oExcelTabData.newRecFlag=false;
            var aCopiedData = JSON.parse(JSON.stringify(oExcelTabData));
            var oSPJSONModel = new JSONModel(aCopiedData);
            this.getOwnerComponent().setModel(oSPJSONModel, "alSidePanel");
        },
        onRowSelect: function(oEvent) {
            
            // var oItem = oEvent.getSource();
            // var oCtx  = oItem.getBindingContext("excelModel");
            // var oModel = this.getOwnerComponent().getModel("excelModel");
            // var oCPath = oCtx.getPath() + "/children";
            // this.sCurrentPath = oCPath;
            // var oExcelTabData = oModel.getProperty(oCPath);
            // var aCopiedData = JSON.parse(JSON.stringify(oExcelTabData));
            // var oSPJSONModel = new JSONModel(aCopiedData);
            // this.getOwnerComponent().setModel(oSPJSONModel, "alSidePanel");
        },
        onCloseDialog: function () {
            if (this._oDialog) {
                this._oDialog.close();
            }
        },
        _loadExternalLibrary: function (sUrl) {
            return new Promise(function (resolve, reject) {
                // If already loaded, resolve immediately
                if (window.XLSX) {
                    resolve();
                    return;
                }

                var script = document.createElement('script');
                script.type = 'text/javascript';
                script.src = sUrl;
                script.onload = resolve;
                script.onerror = reject;
                document.head.appendChild(script);
            });
        }
    });
});