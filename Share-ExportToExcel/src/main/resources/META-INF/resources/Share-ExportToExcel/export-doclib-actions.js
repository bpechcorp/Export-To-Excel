(function() {
	YAHOO.Bubbling.fire("registerAction", {
		actionName : "onActionExportData",
		fn : function com_pdf_flipview_onActionExportData(file) {
			window.location.href = Alfresco.constants.PROXY_URI+"export/excel?nodeRef="+file.nodeRef;
		}
	});
})();