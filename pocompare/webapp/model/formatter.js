sap.ui.define([
    "sap/ui/model/json/JSONModel",
    "sap/ui/Device"
], 
function (JSONModel, Device) {
    "use strict";

    return {
        /**
         * Provides runtime information for the device the UI5 app is running on as a JSONModel.
         * @returns {sap.ui.model.json.JSONModel} The device model.
         */
        stateFormatter: function (sStatus) {
            if(sStatus=="1"){
                return "Information"
            }else if(sStatus=="2"){
                return "Success"
            }else if(sStatus=="3"){
                return "Error"
            }
        }
    };

});