Components.utils.import("resource://gre/modules/XPCOMUtils.jsm");
Components.utils.import("resource://gre/modules/Services.jsm");

var timer, Zotero;
function ZoteroMacWordStartupListener() {
    function loadObserver() {
        // We'll get this notification before the service is actually
        // done initializing, so wait for it actually complete
        timer = Components.classes["@mozilla.org/timer;1"].
            createInstance(Components.interfaces.nsITimer);
        var timerCallback = {"notify":function() {
            timer = undefined;
            Zotero = Components.classes["@zotero.org/Zotero;1"]
                     .getService(Components.interfaces.nsISupports)
                     .wrappedJSObject;
            initWord2016Pipe();
        }}
        timer.initWithCallback(timerCallback, 0, Components.interfaces.nsITimer.TYPE_ONE_SHOT);
    }
    Services.obs.addObserver(loadObserver, "zotero-loaded", false);
}
ZoteroMacWordStartupListener.prototype = {
    observe: function() {},
    contractID: "@zotero.org/Zotero/integration/startupListener?agent=MacWord2016;1",
    classDescription: "Zotero Mac Word 2016 Startup Listener",
    classID: Components.ID("{26522064-b955-4bb0-9ccb-37a5c8c96fa0}"),
    service: true,
    _xpcom_categories: [{category:"profile-after-change", entry:"ZoteroMacWordStartupListener"}],
    QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports, Components.interfaces.nsIObserver])
};

const NSGetFactory = XPCOMUtils.generateNSGetFactory([ZoteroMacWordStartupListener]);

function initWord2016Pipe() {
    // We only use an integration pipe on OS X.
    // On Linux, we use the alternative communication method in the OOo plug-in
    // On Windows, we use a command line handler for integration. See
    // components/zotero-integration-service.js for this implementation.
    if(!Zotero.isMac || Zotero.isConnector) return;

    // Determine where to put the pipe
    // on OS X, first try /Users/Shared for those who can't put pipes in their home
    // directories
    var office2016Container = Components.classes["@mozilla.org/file/directory_service;1"].
        getService(Components.interfaces.nsIProperties).
        get("Home", Components.interfaces.nsIFile);
    office2016Container.append("Library");
    office2016Container.append("Containers");
    office2016Container.append("com.microsoft.Word");
    office2016Container.append("Data");
    
    if(!office2016Container.exists() || !office2016Container.isDirectory() || !office2016Container.isWritable()) return;

    var pipe = office2016Container.clone();
    pipe.append(".zoteroIntegrationPipe");

    if(pipe.exists()) {
        if(!Zotero.Integration.deletePipe(pipe)) return;
    }
    
    // try to initialize pipe
    Zotero.Integration.initPipe(pipe);
}
