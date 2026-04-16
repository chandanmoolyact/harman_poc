sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/export/Spreadsheet",
], (Controller,JSONModel,MessageToast,MessageBox,Spreadsheet) => {
    "use strict";
    var that=this;

    return Controller.extend("com.sap.pocompare.controller.Master", {
        onInit() {
            // Initialize the model that will hold our Excel data
            var oModel = new JSONModel({
                data: []
            });
            this.getOwnerComponent().setModel(oModel, "excelModel");
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
                    "PO Number", 
                    "Line Item", 
                    "Quantity", 
                    "Delivery Date"
                ];

                // VALIDATION: Check if every required header exists in the uploaded file
                var bIsValid = REQUIRED_HEADERS.every(function (header) {
                    return headers.includes(header);
                });

                if (!bIsValid) {
                    // Punish user: Clear UI and show error
                    that.onClearFile();
                    // oErrorStrip.setVisible(true);
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
                        PONumber: row["PO Number"] || row["PONumber"],
                        LineItem: row["Line Item"] || row["LineItem"],
                        Quantity: row["Quantity"],
                        DeliveryDate: new Date(row["Delivery Date"] || row["DeliveryDate"])?.toISOString()?.split('T')[0]
                    };
                });

                that.getOwnerComponent().getModel("excelModel").setProperty("/data", formattedData);
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

            // sap.ui.core.BusyIndicator.show(0);

            // Bind to your CAP OData V4 action
            // var oODataModel = this.getOwnerComponent().getModel(); 
            MessageBox.success("Data submitted. Any mismatches will be routed for manager approval.", {
				actions: [MessageBox.Action.OK, MessageBox.Action.CANCEL],
				emphasizedAction: MessageBox.Action.OK,
				onClose: function (sAction) {
					// MessageToast.show("Action selected: " + sAction);
                    that.onClearFile();

				},
				dependentOn: this.getView()
			});



            // oAction.setParameter("payload", JSON.stringify(oPayload));
            
            // oAction.execute().then(function () {
            //     sap.ui.core.BusyIndicator.hide();
            //     MessageBox.success("Data submitted. Any mismatches will be routed for manager approval.", {
            //         onClose: function() {
            //             this.onClearFile(); // Reset the UI
            //         }.bind(this)
            //     });
            // }.bind(this)).catch(function (oError) {
            //     sap.ui.core.BusyIndicator.hide();
            //     MessageBox.error("Failed to submit data: " + oError.message);
            // });
        },
        onDownloadTemplate: function () {
            // Fetch the specific column configuration for the PO template
            var aCols = this._createColumnConfig();

            // Configure the export settings
            var oSettings = {
                workbook: {
                    columns: aCols
                },
                // An empty array guarantees no rows are added, keeping it a pure template
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
            // Defined based on your provided file structure
            return [
                {
                    label: 'PO Number',
                    property: 'poNumber',
                    type: 'string', // Stored as a string to preserve leading zeros if any
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
                    format: 'yyyy-MM-dd', // Common standard for SAP/S4HANA date formats
                    width: 20
                }
            ];
        }
    });
});