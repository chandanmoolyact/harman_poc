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

    return Controller.extend("com.sap.pocompare.controller.Master", {
        formattter:formatter,
        onInit() {
            // Initialize the model that will hold our Excel data
            var oModel = new JSONModel({
                data: []
            });
            this.getOwnerComponent().setModel(oModel, "excelModel");
            this._headerFB=this.getView().byId("idHeaderFB")
            this._headerFB.setVisible(false)
        },

        // Triggered when a file is selected via the FileUploader
        onFileChange: function (oEvent) {
            var aFiles = oEvent.getParameter("files");
            
            if (aFiles && aFiles.length > 0) {
                var oFile = aFiles[0];
                this._readExcel(oFile);
            }
        },

        // Allows the user to reset the view
        onClearFile: function () {
            this.byId("excelUploader").clear();
            this.getView().getModel("excelModel").setProperty("/data", []);
            this._headerFB.setVisible(false)
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
                var bIsValid = REQUIRED_HEADERS.every(function (header) {
                    return headers.includes(header);
                });

                if (!bIsValid) {
                    that.onClearFile();
                    MessageBox.error("Incorrect Template. Please use the official template with the correct columns");
                    return;  
                }

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
                        VisiblePanel:false
                    };
                });

                that.getOwnerComponent().getModel("excelModel").setProperty("/data", formattedData);
                that._headerFB.setVisible(true)
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
        onRowPress: function (oEvent) {
            var oView = this.getView();
            // Get the specific data context of the row that was clicked
            var oContext = oEvent.getSource().getBindingContext("excelModel");

            // If the dialog hasn't been created yet, load it
            if (!this._pCompareDialog) {
                this._pCompareDialog = Fragment.load({
                    id: oView.getId(),
                    // Replace "com.sap.pocompare.view.fragments" with your actual path
                    name: "com.sap.pocompare.view.fragments.CompareDialog", 
                    controller: this
                }).then(function (oDialog) {
                    oView.addDependent(oDialog);
                    return oDialog;
                });
            }

            // Once the dialog is ready, bind the row context to it and open
            this._pCompareDialog.then(function (oDialog) {
                oDialog.setBindingContext(oContext, "excelModel");
                oDialog.open();
            });
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

        _closeDialog: function () {
            if (this._pCompareDialog) {
                this._pCompareDialog.then(function (oDialog) {
                    oDialog.close();
                });
            }
        }
    });
});