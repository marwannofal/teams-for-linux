const { app, BaseWindow, WebContentsView } = require("electron");
const { workspaceBounds } = require("./layout");

/** Native workspace window with a separately sized, persistent Teams renderer. */
function createWorkspaceWindow(options, config) {
  const { webPreferences, ...windowOptions } = options;
  const window = new BaseWindow(windowOptions);
  const teamsView = new WebContentsView({ webPreferences });
  const contents = teamsView.webContents;
  window.contentView.addChildView(teamsView);
  Object.defineProperties(window, {
    webContents: { value: contents },
    teamsView: { value: teamsView },
  });

  // Existing main-process services address the Teams renderer through these
  // BrowserWindow conveniences. BaseWindow provides the native window APIs.
  for (const method of ["loadURL", "loadFile", "getURL", "reload", "send", "openDevTools", "closeDevTools", "toggleDevTools", "isDevToolsOpened", "capturePage"]) {
    window[method] = (...args) => contents[method](...args);
  }
  const layout = () => teamsView.setBounds(workspaceBounds(window, config));
  window.on("resize", layout);
  layout();
  // BrowserWindow normally forwards these events from its primary renderer.
  contents.on("page-title-updated", (event, title) => {
    window.emit("page-title-updated", event, title);
    if (!event.defaultPrevented && !window.isDestroyed()) window.setTitle(title);
  });
  window.on("focus", (event) => {
    contents.focus();
    app.emit("browser-window-focus", event, window);
  });
  window.on("blur", (event) => app.emit("browser-window-blur", event, window));
  contents.once("destroyed", () => {
    if (!window.isDestroyed()) window.destroy();
  });
  // BaseWindow doesn't destroy the WebContentsViews it owns automatically.
  window.once("closed", () => {
    window.removeListener("resize", layout);
    if (!contents.isDestroyed()) contents.close();
  });
  return window;
}

module.exports = { createWorkspaceWindow };
