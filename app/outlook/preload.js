"use strict";
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("workspaceTabs", {
  select: (tab, keepFocus = false) => {
    if (["teams", "outlook", "reload"].includes(tab)) {
      ipcRenderer.send("workspace-tab-action", { tab, keepFocus: keepFocus === true });
    }
  },
  onState: (callback) => {
    ipcRenderer.on("workspace-tab-state", (_event, state) => callback(state));
  },
});
