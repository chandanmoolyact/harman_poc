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

    return Controller.extend("com.sap.pocompare.controller.Main", {
        formattter:formatter,
        onInit() {
            // Initialize the model that will hold our Excel data
            var oModel = new JSONModel({
                data: []
            });
            this.getOwnerComponent().setModel(oModel, "excelModel");
            // this._headerFB=this.getView().byId("idHeaderFB")
            // this._headerFB.setVisible(false)
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

        // Allows the user to reset the view
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

                // VALIDATION: Check if every required header exists in the uploaded file
                // var bIsValid = REQUIRED_HEADERS.every(function (header) {
                //     return headers.includes(header);
                // });

                // if (!bIsValid) {
                //     that.onClearFile();
                //     MessageBox.error("Incorrect Template. Please use the official template with the correct columns");
                //     return;  
                // }

                // Convert to JSON
                var jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    raw: false,
                    dateNF: 'yyyy-mm-dd'
                });

                // Map columns
                var formattedData = jsonData.map(function(row) {
                    return {
                        Vendor: row["Vendor"],
                        Material: row["Material"],
                        MaterialDesc: row["Material Desc"],
                        PONumber: row["PO Number"] || row["PONumber"],
                        LineItem: row["Line Item"] || row["LineItem"],
                        Quantity: row["Quantity"],
                        DeliveryDate: new Date(row["Delivery Date"] || row["Date"])?.toISOString()?.split('T')[0],
                        ScheduleLineCategory:row["Schedule Line Category"],
                        StatusCode:"1",
                        Status:"New",
                        StatusState:formatter.stateFormatter("1"),
                        PanelVisible:false,
                        VendorInputTable:[{
                            LineItem: row["Line Item"] || row["LineItem"],
                            Quantity: row["Quantity"],
                            DeliveryDate: new Date(row["Delivery Date"] || row["Date"])?.toISOString()?.split('T')[0]
                        }]
                    };
                });

                that.getOwnerComponent().getModel("excelModel").setProperty("/data", formattedData);
                // that._headerFB.setVisible(true)
                MessageToast.show("Excel loaded for preview.");
            };

            reader.onerror = function (ex) {
                MessageBox.error("Error reading the Excel file.");
            };

            reader.readAsBinaryString(file);
        },

        // Triggered on Submit
        onSubmitData: function () {
            var that=this
            const oExcelModel = this.getOwnerComponent().getModel("excelModel");
            var aData = oExcelModel.getProperty("/data");

            if (!aData || aData.length === 0) {
                MessageBox.warning("No data to submit.");
                return;
            }

            var oPayload = { items: aData };

            MessageBox.success("Data submitted. Any mismatches will be routed for manager approval.", {
				actions: [MessageBox.Action.OK, MessageBox.Action.CANCEL],
				emphasizedAction: MessageBox.Action.OK,
				onClose: function (sAction) {
					// MessageToast.show("Action selected: " + sAction);
                    that.onClearFile();

				},
				dependentOn: this.getView()
			});
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
        // --- NEW: Handle Row Click & Open Dialog ---
        onRowPress: async function (oEvent) {
            // var oView = this.getView();
            // // Get the specific data context of the row that was clicked
            // var oContext = oEvent.getSource().getBindingContext("excelModel");

            // // If the dialog hasn't been created yet, load it
            // if (!this._pCompareDialog) {
            //     this._pCompareDialog = await this.loadFragment({
            //         // id: oView.getId(),
            //         // Replace "com.sap.pocompare.view.fragments" with your actual path
            //         name: "com.sap.pocompare.view.fragments.sections.CompareDialog"
            //         // controller: this
            //     });
            //     oView.addDependent(this._pCompareDialog);
            // }

            // // Once the dialog is ready, bind the row context to it and open
            // this._pCompareDialog.setBindingContext(oContext, "excelModel");
            // this._pCompareDialog.open();
        },
        onShowExpanded:function(oEvent){
            let oSource=oEvent.getSource()
            let oBindingContext=oSource.getBindingContext("excelModel")
            let oPath=oBindingContext.getPath()
            let oPanelVisiblePath=oPath+"/PanelVisible"

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
            // 1. Get the button that was clicked
            var oButton = oEvent.getSource();
            
            // 2. Get the binding context of the outer row 
            // (since the Add button is outside the inner table, but inside the outer row)
            var oContext = oButton.getBindingContext("excelModel");
            var sOuterRowPath = oContext.getPath(); // e.g., "/data/0"
            
            // 3. Get the model
            var oModel = this.getOwnerComponent().getModel("excelModel");
            
            // 4. Get the current array for this specific inner table
            var aVendorInputTable = oModel.getProperty(sOuterRowPath + "/VendorInputTable");
            
            // 5. Push a new blank object. 
            // Make sure the keys match exactly what you defined in your map function
            aVendorInputTable.push({
                LineItem: oContext.getProperty("LineItem"), // Optional: carry over the line item
                Quantity: "",
                DeliveryDate: "" 
            });
            
            // 6. Update the model to trigger the UI refresh
            oModel.setProperty(sOuterRowPath + "/VendorInputTable", aVendorInputTable);
        },
        onDeleteVendorRow: function (oEvent) {
            // 1. Get the delete button that was clicked inside the inner table row
            var oButton = oEvent.getSource();
            
            // 2. Get the binding context for this specific INNER row
            var oContext = oButton.getBindingContext("excelModel");
            var sInnerRowPath = oContext.getPath(); // e.g., "/data/0/VendorInputTable/1"
            
            // 3. Get the model
            var oModel = this.getOwnerComponent().getModel("excelModel");
            
            // 4. Split the path to figure out the parent array and the index of the row to delete
            // sInnerRowPath.lastIndexOf("/") finds the last slash to separate array path from the index
            var sParentArrayPath = sInnerRowPath.substring(0, sInnerRowPath.lastIndexOf("/")); // "/data/0/VendorInputTable"
            var sRowIndex = sInnerRowPath.substring(sInnerRowPath.lastIndexOf("/") + 1); // "1"
            
            // 5. Get the array, splice out the deleted row, and update the model
            var aVendorInputTable = oModel.getProperty(sParentArrayPath);
            aVendorInputTable.splice(parseInt(sRowIndex, 10), 1); 
            
            oModel.setProperty(sParentArrayPath, aVendorInputTable);
        },

        // --- Dialog Actions ---
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
            var oContext = oButton.getBindingContext("excelModel");
            var sInnerRowPath = oContext.getPath();
        
            // 3. Get the model
            var oModel = this.getOwnerComponent().getModel("excelModel");
            let sChangeProperty=sInnerRowPath+"/StatusCode"
            if(sButtonText=="Approve"){
                oModel.setProperty(sInnerRowPath,2);
            }else if(sButtonText=="Reject"){
                oModel.setProperty(sInnerRowPath,3);
            }
        },

        _closeDialog: function () {
            if (this._pCompareDialog) {
                this._pCompareDialog.then(function (oDialog) {
                    oDialog.close();
                });
            }
        },

        onConfirmationRowPress: function (oEvent) {   
            var oItem = oEvent.getSource();
            var oCtx  = oItem.getBindingContext("excelModel");
            var oModel = this.getView().getModel("excelModel");

            // Store selected row info so dialog title binds correctly
            oModel.setProperty("/SelectedLineItem", oCtx.getProperty("LineItem"));

            // Optionally filter SAPPOTable based on selected row key here
            // oModel.setProperty("/SAPPOTable", this._getFilteredPORows(oCtx));

            if (!this._oDialog) {
                this._oDialog = this.byId("poDetailDialog");
            }
            this._oDialog.open();
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