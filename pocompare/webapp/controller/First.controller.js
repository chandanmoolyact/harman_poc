sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/export/Spreadsheet",
    "com/sap/pocompare/model/formatter",
    "sap/ui/core/Fragment"
], (Controller, JSONModel, MessageToast, MessageBox, Spreadsheet, formatter, Fragment) => {
    "use strict";
    var that = this;

    return Controller.extend("com.sap.pocompare.controller.First", {
        formattter: formatter,
        onInit() {
            // Initialize the model that will hold our Excel data
            var oModel = new JSONModel({
                data: []
            });
            this.getOwnerComponent().setModel(oModel, "excelModel");
            this.lineItemFlag = true
        },

        // Triggered when a file is selected via the FileUploader
        onFileChange: function (oEvent) {
            var aFiles = oEvent.getParameter("files");
            if (aFiles && aFiles.length > 0) {
                var oFile = aFiles[0];
                this._loadExternalLibrary("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js").then(function () {
                    this._readExcel(oFile);
                }.bind(this))
                    .catch(function () {
                        sap.m.MessageToast.show("Failed to load the Excel library from CDN.");
                    });
            }
        },
        onClearFile: function () {
            this.byId("excelUploader").clear();
            this.getView().getModel("excelModel").setProperty("/data", []);
        },
        onCancelTemplate: function () {
            this.byId("excelUploader").clear();
            this.getView().getModel("excelModel").setProperty("/data", []);
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
                // var formattedData = jsonData.map(function(row) {
                //     return {
                //         //Level 1 Start
                //         VendorCode: row["Vendor Code"],
                //         VendorName: row["Vendor Name"]||row["Vendor name"],
                //         PONumber: row["PO Number"] || row["PONumber"]||row["PO/PR No."],
                //         PODate: new Date(row["PO date"])?.toISOString()?.split('T')[0],
                //         PanelVisible:false,
                //         //Level 2 Start
                //         POLineItem: row["PO Line Item"] || row["LineItem"],
                //         Material: row["Material"],
                //         MaterialDesc: row["Material Description"],
                //         POQuantity: row["PO Quantity"],
                //         UOM: row["Unit of Measure"],
                //         DeliveryDate: new Date(row["Delivery Date"])?.toISOString()?.split('T')[0],
                //         NetPrice: row["Net Price"],
                //         Currency: row["Currency"],
                //         Per: row["Per"],
                //         MaterialGroup: row["Material Group"],
                //         Plant : row["Plant"],
                //         StorageLocation : row["Storage Location"],
                //         //Level 3 Start
                //         ConfirmationCategory:row["Confirmation category"]||row["Confirmation Category"],
                //         FDDCategory:row["Fdelivery Date category"]||row["FDelivery Date Category"]||row["Delivery Date Category"],
                //         Quantity: row["Quantity"],
                //         Reference: row["Reference"],
                //         CreationDate: new Date(row["Created on Date"])?.toISOString()?.split('T')[0],
                //         InboundDelivery: row["Inbound Delivery"],
                //         Item: row["Item"],
                //         HLItem: row["Higher Level Item"],
                //         Batch: row["Batch"],
                //         QtyReduced: row["Quantity Reduced"],
                //         MRPRelevant: row["MRP relevant"]||row["MRP Relevent"],
                //         MRPMaterial: row["MPN Material"]||row["MPN material"],
                //         CreationIndicator: row["Creation Indicator"]||row["Creation Indicator"],
                //         SequenceNumber: row["Sequence Number"]||row["Sequence number"],
                //         RejectFlag: row["Rejected"],
                //         RejectReason: row["Comment"],
                //         RejectDate: row["Rejection Date"],   
                //         RejectDate: new Date(row["Rejection Date"]?.toISOString()?.split('T')[0], 

                //         QlikQty: row["Qlik Qty"]||row["QLIK Qty"],  
                //         QlikDate: new Date(row["Qlik Date"]||row["QLIK Date"])?.toISOString()?.split('T')[0],
                //         //Status 4
                //         StatusCode:"1",
                //         Status:"New",
                //         StatusState:formatter.stateFormatter("1"),
                //         StatusMsg:formatter.statusDescription("1"),
                //     };
                // });


                var formattedData = jsonData.map(function (row) {
                    // Helper function to safely format dates to YYYY-MM-DD
                    var formatDate = function (dateVal) {
                        if (!dateVal) return "";
                        var d = new Date(dateVal);
                        if (isNaN(d.getTime())) return "";

                        var month = (d.getUTCMonth() + 1).toString().padStart(2, '0');
                        var day = d.getUTCDate().toString().padStart(2, '0');
                        var year = d.getUTCFullYear();
                        return month + "/" + day + "/" + year;
                    };

                    return {
                        // Level 1
                        VendorCode: row["Vendor Code"],
                        VendorName: row["Vendor Name"] || row["Vendor name"],
                        PONumber: row["PO Number"] || row["PONumber"] || row["PO/PR No."],
                        PODate: formatDate(row["PO date"]),
                        PanelVisible: false,

                        // Level 2
                        POLineItem: row["PO Line Item"] || row["LineItem"],
                        Material: row["Material"],
                        MaterialDesc: row["Material Description"],
                        POQuantity: row["PO Quantity"],
                        UOM: row["Unit of Measure"],
                        DeliveryDate: formatDate(row["Delivery Date"]),
                        NetPrice: row["Net Price"],
                        Currency: row["Currency"],
                        Per: row["Per"],
                        MaterialGroup: row["Material Group"],
                        Plant: row["Plant"],
                        StorageLocation: row["Storage Location"],

                        // Level 3
                        ConfirmationCategory: row["Confirmation category"] || row["Confirmation Category"],
                        FDDCategory: row["Fdelivery Date category"] || row["FDelivery Date Category"] || row["Delivery Date Category"],
                        Quantity: row["Quantity"],
                        Reference: row["Reference"],
                        CreationDate: formatDate(row["Created on Date"]),
                        InboundDelivery: row["Inbound Delivery"],
                        Item: row["Item"],
                        HLItem: row["Higher Level Item"],
                        Batch: row["Batch"],
                        QtyReduced: row["Quantity Reduced"],
                        MRPRelevant: row["MRP relevant"] || row["MRP Relevent"],
                        MRPMaterial: row["MPN Material"] || row["MPN material"],
                        CreationIndicator: row["Creation Indicator"],
                        SequenceNumber: row["Sequence Number"] || row["Sequence number"],
                        RejectFlag: row["Rejected"],
                        RejectReason: row["Comment"],
                        RejectDate: formatDate(row["Rejection Date"]),

                        QlikQty: row["Qlik Qty"] || row["QLIK Qty"],
                        QlikDate: formatDate(row["Qlik Date"] || row["QLIK Date"]),

                        // Status Initial States
                        StatusCode: "1",
                        Status: "New",
                        DateState: "None",
                        DateMsg: "",
                        // These are the properties our validation logic uses!
                        StatusState: formatter.stateFormatter("1"), // e.g., "None" or "Information"
                        StatusMsg: formatter.statusDescription("1")
                    };
                });

                let newData = that.transformDataForTreeTable(formattedData)

                let aDateNewData=that._processData(newData)
                that.aOldData = JSON.parse(JSON.stringify(aDateNewData));

                that.getOwnerComponent().getModel("excelModel").setProperty("/data", aDateNewData);
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
                        POLineItem: item.POLineItem, // Using actual PO Line Item instead of mock
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
                    RejectFlag:item.RejectFlag,
                    RejectReason:item.RejectReason,
                    QlikQty: item.QlikQty,
                    QlikDate: item.QlikDate,
                    RejectDate:item.RejectDate,   
                    StatusState: formatter.stateFormatter("1"),
                    StatusMsg: formatter.statusDescription("1"),
                    DateState: "None",
                    DateMsg: "",
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
        _processData: function (aData) {
            aData.forEach(level1 => {
                if (level1.children) {
                    level1.children.forEach(level2 => {
                        if (level2.children) {
                            level2.children.forEach(level3 => {
                                // 1. Get the dates
                                const oDelDate = new Date(level3.DeliveryDate);
                                const oCreDate = new Date(level3.CreationDate);

                                // 2. Calculate the difference in time
                                const iDiffInTime = oCreDate.getTime() - oDelDate.getTime();
                                
                                // 3. Convert time to days
                                const iDiffInDays = iDiffInTime / (1000 * 3600 * 24);

                                // 4. Check logic: +- 5 days
                                if (Math.abs(iDiffInDays) <= 5) {
                                    level3.DateState = "Success"; // Green
                                    level3.DateMsg = "Within 5-day range";
                                } else {
                                    level3.DateState = "Error";   // Red
                                    level3.DateMsg = "Outside 5-day range";
                                }
                            });
                        }
                    });
                }
            });
            return aData;
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
            oSheet.build().finally(function () {
                oSheet.destroy();
            });
        },
        onSaveTemplate: async function (oEvent) {
            var treeData = this.getView().getModel("excelModel").getProperty("/data");



            var oCAPModel = this.getOwnerComponent().getModel("capService");
            var oActionContext = oCAPModel.bindContext("/sendMailContent(...)");
            // oActionContext.setParameter("poHeader", JSON.stringify(treeData));
            oActionContext.setParameter("poHeader", treeData);
            try {
                var oResults = await oActionContext.execute()
            } catch (oError) {
                console.log(oError)
            }
            var flatExcelData = this.transformTreeToFlatData(treeData);
            var sGeneratedMsg = this.getChangeSummary(this.aOldData, treeData)

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
                { label: 'Delivery Date Category', property: 'FDDCategory', type: 'string' },
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
                { label: 'Rejected', property: 'RejectFlag', type: 'string' },
                { label: 'Comment', property: 'RejectReason', type: 'string' },
                { label: 'Qlik Qty', property: 'QlikQty', type: 'string' },
                { label: 'Qlik Date', property: 'QlikDate', type: 'string' },
                { label: 'Rejection Date', property: 'RejectDate', type: 'string' },
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
            oSheet.build().finally(function () {
                oSheet.destroy();
            });

            // this.aOldData=JSON.parse(JSON.stringify(treeData));
            this.onGenSaveMessage(sGeneratedMsg, treeData)
        },
        onGenSaveMessage: function (sMsg, treeData) {
            var that = this
            MessageBox.success(sMsg, {
                actions: [MessageBox.Action.OK, MessageBox.Action.CANCEL],
                emphasizedAction: MessageBox.Action.OK,
                onClose: function (sAction) {
                    MessageToast.show(sMsg);


                    const resetNewRecFlag = (data) => {
                        data.forEach(parent => {
                            parent.children?.forEach(lineItem => {
                                lineItem.children?.forEach(record => {
                                    record.newRecFlag = false;
                                });
                            });
                        });
                        return data;
                    };
                    var aOldExcelList = resetNewRecFlag(treeData);
                    that.aOldData = JSON.parse(JSON.stringify(aOldExcelList));
                    this.getView().getModel("excelModel").setProperty("/data", that.aOldData)
                },
                dependentOn: this.getView()
            });
        },
        // This function is called on the 'change' or 'liveChange' event of the input
        onQuantityLiveChange: function (oEvent) {
            const oInput = oEvent.getSource();
            const sNewValue = oEvent.getParameter("newValue");

            // 1. Get the Binding Context for the current line (Third Level)
            const oContext = oInput.getBindingContext("excelModel");
            const sPath = oContext.getPath();

            // 2. Identify the Parent Path (Second Level - PO Line Item)
            // Example path: /0/children/1/children/2 -> Parent: /0/children/1
            const aPathParts = sPath.split("/");
            aPathParts.pop(); // Remove current index
            aPathParts.pop(); // Remove "children" literal
            const sParentPath = aPathParts.join("/");

            const oModel = oContext.getModel("excelModel");
            const oParentData = oModel.getProperty(sParentPath);

            // 3. Get PO Quantity (Total allowed)
            const fPOQuantity = parseFloat(oParentData.POQuantity || 0);

            // 4. Calculate sum of all AB lines under this parent
            const aConfirmations = oParentData.children || [];
            let fTotalConfirmedQty = 0;

            aConfirmations.forEach((item, index) => {
                // Use the live value for the row being edited, otherwise use the model value
                if (oContext.getPath() === sParentPath + "/children/" + index) {
                    fTotalConfirmedQty += parseFloat(sNewValue || 0);
                } else {
                    // Only sum items with Category "AB"
                    if (item.ConfirmationCategory === "AB") {
                        fTotalConfirmedQty += parseFloat(item.Quantity || 0);
                    }
                }
            });

            // 5. Validation Check
            if (fTotalConfirmedQty > fPOQuantity) {
                oInput.setValueState("Error");
                oInput.setValueStateText("The sum of Confirmation quantity exceeds the PO line quantity");
            } else {
                oInput.setValueState("None");
                oInput.setValueStateText("");
                const oTable = oEvent.getSource().getParent().getParent() // Ensure you have the ID of your TreeTable
                const aRows = oTable.getItems();

                aRows.forEach(oRow => {
                    const oRowContext = oRow.getBindingContext("excelModel");
                    if (oRowContext) {
                        const sRowPath = oRowContext.getPath();

                        // Check if this row belongs to the same parent PO Line
                        if (sRowPath.startsWith(sParentPath + "/children/")) {
                            // Find the Input control within the row (usually inside a template/cell)
                            // Adjust the index [n] based on which column your Quantity input is in
                            const oRowInput = oRow.getCells()[10];
                            // const oRowInput = oRow.getCells().find(oCell => oCell.getMetadata().getName() === "sap.m.Input");
                            oRowInput.setValueState("None");
                            oRowInput.setValueStateText("");
                        }
                    }
                });

            }
        },
        onDateLiveChange: function (oEvent) {
            const oDP = oEvent.getSource();
            const oModel = this.getView().getModel("excelModel");
            const oContext = oDP.getBindingContext("excelModel");
            const sNewDateValue = oEvent.getParameter("value");

            if (!sNewDateValue) return;

            // 1. Navigate to Parent Delivery Date
            const sPath = oContext.getPath();
            const sParentPath = sPath.substring(0, sPath.lastIndexOf("/children/"));
            const sParentDateStr = oModel.getProperty(sParentPath + "/DeliveryDate");

            const dParent = new Date(sParentDateStr);
            const dChild = new Date(sNewDateValue);

            // 2. Calculate Day Difference
            const iDiffInMs = Math.abs(dChild - dParent);
            const iDiffInDays = iDiffInMs / (1000 * 60 * 60 * 24);

            // 3. Determine State
            const sState = (iDiffInDays <= 5) ? "Success" : "Error";
            const sMsg = (iDiffInDays <= 5) ? "Within ±5 day range" : "Beyond 5 days of PO Delivery Date";

            // 4. Update the Model
            oModel.setProperty(sPath + "/DateState", sState);
            oModel.setProperty(sPath + "/DateMsg", sMsg);
        },

        // onQuantityLiveChange: function (oEvent) {
        //     const oCurrentInput = oEvent.getSource();
        //     const oContext = oCurrentInput.getBindingContext("excelModel");
        //     const oModel = oContext.getModel("excelModel");

        //     // 1. Navigate to the Parent (Level 2)
        //     const sPath = oContext.getPath();
        //     const aPathParts = sPath.split("/");
        //     aPathParts.splice(-2); // Remove index and "children" to get parent path
        //     const sParentPath = aPathParts.join("/");

        //     const oParentData = oModel.getProperty(sParentPath);
        //     const fPOQuantity = parseFloat(oParentData.POQuantity || 0);
        //     const aConfirmations = oParentData.children || [];

        //     // 2. Calculate the current total of all AB lines
        //     // Note: We use the live value for the row being edited, model value for others
        //     let fTotalConfirmedQty = 0;
        //     aConfirmations.forEach((item, index) => {
        //         if (item.ConfirmationCategory === "AB") {
        //             const sItemPath = sParentPath + "/children/" + index;
        //             const fQty = (sItemPath === sPath) 
        //                 ? parseFloat(oEvent.getParameter("newValue") || 0) 
        //                 : parseFloat(item.Quantity || 0);
        //             fTotalConfirmedQty += fQty;
        //         }
        //     });

        //     // 3. Find all rendered Input controls for this specific PO Line
        //     // We use the Table's rows to find siblings and update their states
        //     const bIsInvalid = fTotalConfirmedQty > fPOQuantity;
        //     const oTable = oEvent.getSource().getParent().getParent() // Ensure you have the ID of your TreeTable
        //     const aRows = oTable.getItems();

        //     aRows.forEach(oRow => {
        //         const oRowContext = oRow.getBindingContext("excelModel");
        //         if (oRowContext) {
        //             const sRowPath = oRowContext.getPath();

        //             // Check if this row belongs to the same parent PO Line
        //             if (sRowPath.startsWith(sParentPath + "/children/")) {
        //                 // Find the Input control within the row (usually inside a template/cell)
        //                 // Adjust the index [n] based on which column your Quantity input is in
        //                 const oRowInput = oRow.getCells().find(oCell => oCell.getMetadata().getName() === "sap.m.Input");

        //                 if (oRowInput) {
        //                     if (bIsInvalid) {
        //                         oRowInput.setValueState("Error");
        //                         oRowInput.setValueStateText("The sum of Confirmation quantity exceeds the PO line quantity");
        //                     } else {
        //                         oRowInput.setValueState("None");
        //                         oRowInput.setValueStateText("");
        //                     }
        //                 }
        //             }
        //         }
        //     });
        // },

        transformTreeToFlatData: function (treeData) {
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
                                    PODate: poDate,
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
                                    MaterialGroup: itemNode.MaterialGroup,
                                    Plant: itemNode.Plant,
                                    StorageLocation: itemNode.StorageLocation,
                                    //Level 3 Start
                                    ConfirmationCategory: subItemNode.ConfirmationCategory,
                                    FDDCategory: subItemNode.FDDCategory,
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

                                    RejectFlag: subItemNode.RejectFlag,
                                    RejectReason: subItemNode.RejectReason,
                                    QlikQty: subItemNode.QlikQty,
                                    QlikDate: subItemNode.QlikDate,
                                    RejectDate: subItemNode.RejectDate,


                                    Status: subItemNode.Status,
                                    StatusMsg: subItemNode.StatusMsg
                                };
                                flatData.push(flatRow);
                            });
                        }
                    });
                }
            });

            return flatData;
        },

        _createColumnConfig: function () {
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
        onShowExpanded: function (oEvent) {

            if (this.lineItemFlag) {
                let oSource = oEvent.getSource()
                let oBindingContext = oSource.getBindingContext("excelModel")
                let oPath = oBindingContext.getPath()
                let oPanelVisiblePath = oPath + "/PanelVisible"

                let bVisiblePath = this.getOwnerComponent().getModel("excelModel").getProperty(oPanelVisiblePath)
                if (bVisiblePath) {
                    this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath, false)
                    // oSource.setIcon("sap-icon://dropdown")

                } else {
                    this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath, true)
                    // oSource.setIcon("sap-icon://slim-arrow-up")
                }
            }
            this.convertLIFlag(true)
        },
        convertLIFlag: function (bFlag) {
            this.lineItemFlag = bFlag
        },
        convertSubLIFlag: function (bFlag) {
            this.sublineItemFlag = bFlag
        },
        onSubShowExpanded: function (oEvent) {
            if (this.sublineItemFlag) {
                this.convertLIFlag(false)
                let oSource = oEvent.getSource()
                let oBindingContext = oSource.getBindingContext("excelModel")
                let oPath = oBindingContext.getPath()
                let oPanelVisiblePath = oPath + "/NextPanelVisible"

                let bVisiblePath = this.getOwnerComponent().getModel("excelModel").getProperty(oPanelVisiblePath)
                if (bVisiblePath) {
                    this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath, false)
                    // oSource.setIcon("sap-icon://dropdown")
                } else {
                    this.getOwnerComponent().getModel("excelModel").setProperty(oPanelVisiblePath, true)
                    //  oSource.setIcon("sap-icon://slim-arrow-up")
                }


                var oCurrentItem = oEvent.getSource();

                // 2. Manage the highlight logic
                if (this.prevRecord) {
                    this.prevRecord.removeStyleClass("myCustomHighlight");
                }

                oCurrentItem.addStyleClass("myCustomHighlight");

                // 3. Store this item as the "previous" for the next time a row is pressed
                this.prevRecord = oCurrentItem;
            }
            this.convertSubLIFlag(true)
            
            // var oItem = oEvent.getSource();
            // var oCtx  = oItem.getBindingContext("excelModel");
            // var oModel = this.getOwnerComponent().getModel("excelModel");
            // var oExcelTabData = oModel.getProperty(oCPath);
            // oExcelTabData.newRecFlag=false;
            // var aCopiedData = JSON.parse(JSON.stringify(oExcelTabData));
            // var oSPJSONModel = new JSONModel(aCopiedData);
            // this.getOwnerComponent().setModel(oSPJSONModel, "alSidePanel");

        },
        onAddVendorRowSP: function (oEvent) {
            this.convertLIFlag(false)
            this.convertSubLIFlag(false)
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("excelModel");
            var oModel = this.getOwnerComponent().getModel("excelModel");
            // var oSPModel = this.getOwnerComponent().getModel("alSidePanel");
            this.sCurrentPath = oContext.getPath() + "/children"
            var aVendorInputTable = oModel.getProperty(this.sCurrentPath)
            var iSNum;
            var oVendorObject;
            if (aVendorInputTable.length != 0) {
                oVendorObject = aVendorInputTable[0]
                var iMaxTabLength = aVendorInputTable.length
                var iMaxSN = aVendorInputTable[iMaxTabLength - 1].SequenceNumber
                iSNum = (Number(iMaxSN) + 1).toString()
            } else {
                oVendorObject = {}
                iSNum = 1
            }
            aVendorInputTable.push({
                // LineItemNumber: iLINumber,
                ConfirmationCategory: oVendorObject?.ConfirmationCategory,
                FDDCategory: oVendorObject?.FDDCategory,
                Quantity: 0,
                Reference: oVendorObject?.Reference,
                CreationDate: oVendorObject?.CreationDate,
                InboundDelivery: oVendorObject?.InboundDelivery,
                Item: oVendorObject?.Item,
                HLItem: oVendorObject?.HLItem,
                Batch: oVendorObject?.Batch,
                QtyReduced: oVendorObject?.QtyReduced,
                MRPRelevant: oVendorObject?.MRPRelevant,
                MRPMaterial: oVendorObject?.MRPMaterial,
                CreationIndicator: oVendorObject?.CreationIndicator,
                SequenceNumber: iSNum,
                newRecFlag: true,
                StatusMsg: formatter.statusDescription("1"),
                StatusState: formatter.stateFormatter("1")
            });
            oModel.setProperty(this.sCurrentPath, aVendorInputTable);
        },
        onDeleteVendorTreeRow: function (oEvent) {
            this.convertLIFlag(false);
            this.convertSubLIFlag(false)

            var oModel = this.getOwnerComponent().getModel("excelModel");
            var oContext = oEvent.getSource().getBindingContext("excelModel");
            var sPath = oContext.getPath(); // e.g., "/nodes/0/nodes/1/nodes/5"

            // 1. Find the last slash to separate the index from the parent path
            var iLastSlashIndex = sPath.lastIndexOf("/");
            var sParentPath = sPath.substring(0, iLastSlashIndex);
            var iIndex = sPath.substring(iLastSlashIndex + 1);

            // 2. Get the actual array from the model (the parent collection)
            var aParentCollection = oModel.getProperty(sParentPath);

            // 3. Remove the item directly from the model's data
            if (Array.isArray(aParentCollection)) {
                aParentCollection.splice(iIndex, 1);
                oModel.refresh(true);
            }
        },
        onRejectVendorTreeRow: function (oEvent) {
            this.convertLIFlag(false);
            this.convertSubLIFlag(false)
            var oResourceBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
            var oModel = this.getOwnerComponent().getModel("excelModel");

            // Get the binding context of the row where the button was clicked
            var oContext = oEvent.getSource().getBindingContext("excelModel");
            var sRejPath = oContext.getPath();

            // Create a Dialog dynamically
            if (!this.oRejectDialog) {
                this.oRejectDialog = new sap.m.Dialog({
                    title: "Reject Record",
                    type: "Message",
                    content: [
                        new sap.m.Label({
                            text: "Please provide a reason for rejection:",
                            labelFor: "rejectionTextArea"
                        }),
                        new sap.m.TextArea("rejectionTextArea", {
                            width: "100%",
                            placeholder: "Enter reason here...",
                            rows: 4
                        })
                    ],
                    beginButton: new sap.m.Button({
                        type: "Emphasized",
                        text: "Confirm",
                        press: function () {
                            var sReason = sap.ui.getCore().byId("rejectionTextArea").getValue();

                            if (!sReason) {
                                sap.m.MessageToast.show("Please enter a reason before submitting.");
                                return;
                            }

                            var sActivePath = this.oRejectDialog.data("activePath");

                            oModel.setProperty(sActivePath + "/RejectFlag", "X");
                            oModel.setProperty(sActivePath + "/RejectReason", sReason);
                            var oToday = new Date().toLocaleDateString();
                            oModel.setProperty(sActivePath + "/RejectDate", oToday);
                            oModel.refresh(true);
                            sap.m.MessageToast.show("Record rejected successfully.");

                            // Close and clean up
                            this.oRejectDialog.close();
                            sap.ui.getCore().byId("rejectionTextArea").setValue(""); // Clear for next time
                        }.bind(this)
                    }),
                    endButton: new sap.m.Button({
                        text: "Cancel",
                        press: function () {
                            this.oRejectDialog.close();
                        }.bind(this)
                    })
                });
            }
            this.oRejectDialog.data("activePath", sRejPath);
            this.oRejectDialog.open();
        },
        onApprovePress: function (oEvent) {
            MessageToast.show("Line Item Approved!");
            this._closeDialog();
        },
        onRejectPress: function (oEvent) {
            MessageToast.show("Line Item Rejected!");
            this._closeDialog();
        },
        onActionTaken: function (oEvent) {
            let sButtonText = oEvent.getSource().getText()
            var oButton = oEvent.getSource();
            var oContext = oButton.getBindingContext("alSidePanel");
            var sInnerRowPath = oContext.getPath();
            var oModel = this.getOwnerComponent().getModel("excelModel");
            var oSPModel = this.getOwnerComponent().getModel("alSidePanel");
            let sChangePropertyStatus = sInnerRowPath + "/Status"
            let sChangePropertyStatusMsg = sInnerRowPath + "/StatusMsg"
            let sChangePropertyStatusState = sInnerRowPath + "/StatusState"
            if (sButtonText == "Approve") {
                oSPModel.setProperty(sChangePropertyStatus, 2);
                oSPModel.setProperty(sChangePropertyStatusMsg, formatter.statusDescription("2"));
                oSPModel.setProperty(sChangePropertyStatusState, formatter.stateFormatter("2"));
            } else if (sButtonText == "Reject") {
                oSPModel.setProperty(sChangePropertyStatus, 3);
                oSPModel.setProperty(sChangePropertyStatusMsg, formatter.statusDescription("3"));
                oSPModel.setProperty(sChangePropertyStatusState, formatter.stateFormatter("3"));
            }
        },
        _closeDialog: function () {
            if (this._pCompareDialog) {
                this._pCompareDialog.then(function (oDialog) {
                    oDialog.close();
                });
            }
        },
        onSubSubRowPress: function (oEvent) {
            this.convertLIFlag(false);
            this.convertSubLIFlag(false)
        },
        onCloseDialog: function () {
            if (this._oDialog) {
                this._oDialog.close();
            }
        },
        getChangeSummary: function (originalData, currentData) {
            let updatedCount = 0;
            let addedCount = 0;

            // Helper to flatten the nested children into a simple array of records
            const flatten = (data) => {
                let results = [];
                data.forEach(parent => {
                    parent.children.forEach(lineItem => {
                        lineItem.children.forEach(record => {
                            results.push(record);
                        });
                    });
                });
                return results;
            };


            const flattenReset = (data) => {
                let results = [];
                data.forEach(parent => {
                    // Ensure children exist to avoid errors
                    parent.children?.forEach(lineItem => {
                        lineItem.children?.forEach(record => {
                            // Set the property to false
                            record.newRecFlag = false;

                            results.push(record);
                        });
                    });
                });
                return results;
            };

            const oldRecords = flattenReset(originalData);
            const newRecords = flatten(currentData);

            newRecords.forEach(newRec => {
                // 1. Check if it's a brand new record
                if (newRec.newRecFlag === true) {
                    addedCount++;
                } else {
                    // 2. Check if an existing record was modified
                    // Find the matching record in the original snapshot by SequenceNumber
                    // const oldRec = oldRecords.find(r => r.SequenceNumber === newRec.SequenceNumber);
                    const oldRec = oldRecords.find(r =>
                        r.SequenceNumber == newRec.SequenceNumber &&
                        r.POLineItem == newRec.POLineItem &&
                        r.VendorCode == newRec.VendorCode &&
                        r.PONumber == newRec.PONumber
                    )

                    if (oldRec) {
                        // Compare relevant fields (Quantity, DeliveryDate, etc.)
                        // We stringify to do a quick "dirty" deep comparison
                        if (JSON.stringify(oldRec) !== JSON.stringify(newRec)) {
                            updatedCount++;
                        }
                    }
                }
            });


            let sResMessage = this.generateMessage(addedCount, updatedCount);
            return sResMessage;
        },

        generateMessage: function (added, updated) {
            if (added === 0 && updated === 0) return "No changes to save.";

            let msg = "Your data has been saved";
            let details = [];

            if (updated > 0) details.push(`${updated} record${updated > 1 ? 's' : ''} updated`);
            if (added > 0) details.push(`${added} record${added > 1 ? 's' : ''} added`);

            return `${msg}, ${details.join(", ")}`;
        },
        handlePopoverPress: function (oEvent) {
             this.convertLIFlag(false);
            this.convertSubLIFlag(false)
            var oButton = oEvent.getSource(),
                oView = this.getView(),
                // Capture the specific row context from the clicked button
                oContext = oButton.getBindingContext("excelModel");

            // Create popover if it doesn't exist
            if (!this._pPopover) {
                this._pPopover = sap.ui.core.Fragment.load({
                    id: oView.getId(),
                    name: "com.sap.pocompare.view.fragments.CommentPopover",
                    controller: this
                }).then(function (oPopover) {
                    oView.addDependent(oPopover);
                    return oPopover;
                });
            }

            this._pPopover.then(function (oPopover) {
                // Bind the popover to the specific row's context
                oPopover.setBindingContext(oContext, "excelModel");
                oPopover.openBy(oButton);
            });
        },
        handleQlikPopoverPress: function (oEvent) {
             this.convertLIFlag(false);
            this.convertSubLIFlag(false)
            var oButton = oEvent.getSource(),
                oView = this.getView(),
                // Capture the specific row context from the clicked button
                oContext = oButton.getBindingContext("excelModel");

            // Create popover if it doesn't exist
            if (!this._pqPopover) {
                this._pqPopover = sap.ui.core.Fragment.load({
                    id: oView.getId(),
                    name: "com.sap.pocompare.view.fragments.QlikCommentPopover",
                    controller: this
                }).then(function (oPopover) {
                    oView.addDependent(oPopover);
                    return oPopover;
                });
            }

            this._pqPopover.then(function (oPopover) {
                // Bind the popover to the specific row's context
                oPopover.setBindingContext(oContext, "excelModel");
                oPopover.openBy(oButton);
            });
        },

        onClosePopover: function () {
            this._pPopover.then(function (oPopover) {
                oPopover.close();
            });
        },
        onCloseQlikPopover: function () {
            this._pqPopover.then(function (oPopover) {
                oPopover.close();   
            });
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