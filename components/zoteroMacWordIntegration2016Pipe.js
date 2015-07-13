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

/**
 * Runs an AppleScript on OS X
 *
 * @param script {String}
 * @param block {Boolean} Whether the script should block until the process is finished.
 */
var _osascriptFile;
function _executeAppleScript(script, block) {
    if(_osascriptFile === undefined) {
        _osascriptFile = Components.classes["@mozilla.org/file/local;1"].
            createInstance(Components.interfaces.nsILocalFile);
        _osascriptFile.initWithPath("/usr/bin/osascript");
        if(!_osascriptFile.exists()) _osascriptFile = false;
    }
    if(_osascriptFile) {
        var proc = Components.classes["@mozilla.org/process/util;1"].
                createInstance(Components.interfaces.nsIProcess);
        proc.init(_osascriptFile);
        try {
            proc.run(!!block, ['-e', script], 2);
        } catch(e) {}
    }
}

/**
 * Deletes a defunct pipe on OS X
 */
function _deletePipe(pipe) {
    try {
        if(pipe.exists()) {
            Zotero.IPC.safePipeWrite(pipe, "Zotero shutdown\n");
            pipe.remove(false);
        }
        return true;
    } catch (e) {
        // if pipe can't be deleted, log an error
        Zotero.debug("Error removing old integration pipe "+pipe.path, 1);
        Zotero.logError(e);
        Components.utils.reportError(
            "Zotero word processor integration initialization failed. "
                + "See http://forums.zotero.org/discussion/12054/#Item_10 "
                + "for instructions on correcting this problem."
        );
        
        // can attempt to delete on OS X
        try {
            var promptService = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
                .getService(Components.interfaces.nsIPromptService);
            var deletePipe = promptService.confirm(null, Zotero.getString("integration.error.title"), Zotero.getString("integration.error.deletePipe"));
            if(!deletePipe) return false;
            let escapedFifoFile = pipe.path.replace("'", "'\\''");
            _executeAppleScript("do shell script \"rmdir '"+escapedFifoFile+"'; rm -f '"+escapedFifoFile+"'\" with administrator privileges", true);
            if(pipe.exists()) return false;
        } catch(e) {
            Zotero.logError(e);
            return false;
        }
    }
}

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
        if(!_deletePipe(pipe)) return;
    }
    
    // try to initialize pipe
    try {
        Zotero.IPC.Pipe.initPipeListener(pipe, function(string) {                       
            if(string != "") {
                // exec command if possible
                var parts = string.match(/^([^ \n]*) ([^ \n]*)(?: ([^\n]*))?\n?$/);
                if(parts) {
                    var agent = parts[1].toString();
                    var cmd = parts[2].toString();
                    var document = parts[3] ? parts[3].toString() : null;
                    Zotero.Integration.execCommand(agent, cmd, document);
                } else {
                    Components.utils.reportError("Zotero: Invalid integration input received: "+string);
                }
            }
        });
    } catch(e) {
        Zotero.logError(e);
    }
}