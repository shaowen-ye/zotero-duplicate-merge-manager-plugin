var DuplicateMergeManagerPlugin;

function log(msg) {
  Zotero.debug("Duplicate Merge Manager: " + msg);
}

function install() {
  log("Installed");
}

async function startup({ id, version, rootURI }) {
  log(`Starting ${version}`);
  Services.scriptloader.loadSubScript(rootURI + "duplicate-merge-manager-plugin.js");
  DuplicateMergeManagerPlugin.init({ id, version, rootURI });
  DuplicateMergeManagerPlugin.addToAllWindows();
}

function onMainWindowLoad({ window }) {
  DuplicateMergeManagerPlugin.addToWindow(window);
}

function onMainWindowUnload({ window }) {
  DuplicateMergeManagerPlugin.removeFromWindow(window);
}

function shutdown() {
  log("Shutting down");
  if (DuplicateMergeManagerPlugin) {
    DuplicateMergeManagerPlugin.removeFromAllWindows();
    DuplicateMergeManagerPlugin = undefined;
  }
}

function uninstall() {
  log("Uninstalled");
}
