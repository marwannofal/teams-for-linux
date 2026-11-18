const { WebContentsView, ipcMain, shell } = require("electron");
const path = require("node:path");

const { DEFAULT_URL, OUTLOOK_HOSTS, AUTH_HOSTS, isOutlookUrl } = require("./urls");
const { HEADER_HEIGHT, workspaceBounds } = require("./layout");
const { isIntegratedTitleBar } = require("../mainAppWindow/titleBar");
const CHANNEL = "workspace-tab-action";

/** Keeps Teams alive underneath an independent, sandboxed Outlook tab. */
class OutlookTabs {
  #window;
  #config;
  #outlook;
  #bar;
  #active = "teams";
  #status = "idle";
  #disposed = false;
  #loading = false;
  #navigationFailed = false;
  #getTeamsContents;
  #popups = new Set();
  #resize = () => this.#layout();
  #focus = () => {
    if (this.#active === "outlook" && !this.#disposed) this.#outlook.webContents.focus();
  };
  #action = (event, action) => {
    // Pin to this exact app-owned view, including its main frame.
    if (event.sender !== this.#bar.webContents ||
        event.senderFrame !== this.#bar.webContents.mainFrame) return;
    if (!action || typeof action !== "object") return;
    if (action.tab === "teams" || action.tab === "outlook") this.select(action.tab, action.keepFocus === true);
    if (action.tab === "reload" && this.#active === "outlook") this.#load();
  };

  constructor(window, config, getTeamsContents = () => window.webContents) {
    this.#window = window;
    this.#config = config;
    this.#getTeamsContents = getTeamsContents;
  }

  initialize() {
    const url = this.#config.outlook?.url ?? DEFAULT_URL;
    if (!isOutlookUrl(url)) throw new Error("outlook.url must use an official Outlook HTTPS host");
    this.#outlook = new WebContentsView({
      webPreferences: {
        session: this.#window.webContents.session,
        contextIsolation: true,
        sandbox: true,
        nodeIntegration: false,
        webviewTag: false,
        spellcheck: true,
        focusOnNavigation: false,
      },
    });
    this.#outlook.setVisible(false);
    this.#window.contentView.addChildView(this.#outlook);
    this.#configureNavigation(this.#outlook.webContents);
    this.#outlook.webContents.on("did-start-navigation", (_event, _url, inPlace, mainFrame) => {
      if (!mainFrame || inPlace) return;
      this.#navigationFailed = false;
      this.#status = "loading";
      this.#push();
    });
    this.#outlook.webContents.on("did-finish-load", () => {
      // Chromium also finishes loading its error document after did-fail-load.
      // Keep the Retry state until a subsequent real navigation succeeds.
      if (!this.#navigationFailed) this.#status = "ready";
      this.#push();
    });
    this.#outlook.webContents.on("did-fail-load", (_event, code, _description, _url, mainFrame) => {
      if (mainFrame && code !== -3) {
        this.#navigationFailed = true;
        this.#status = "error";
        this.#push();
      }
    });
    this.#outlook.webContents.on("render-process-gone", () => {
      this.#status = "error";
      this.#push();
    });
    this.#bar = new WebContentsView({
      webPreferences: {
        preload: path.join(__dirname, "preload.js"),
        contextIsolation: true,
        sandbox: true,
        nodeIntegration: false,
        webviewTag: false,
      },
    });
    this.#window.contentView.addChildView(this.#bar);
    this.#bar.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    this.#bar.webContents.on("will-navigate", (event) => event.preventDefault());
    this.#bar.webContents.on("did-finish-load", () => this.#push());
    // Select a workspace tab or retry Outlook, only from the local tab switcher.
    ipcMain.on("workspace-tab-action", this.#action);
    this.#window.on("resize", this.#resize);
    this.#window.on("focus", this.#focus);
    this.#window.once("closed", () => this.dispose());
    this.#layout();
    this.#bar.webContents.loadFile(path.join(__dirname, "tabs.html")).catch(() => {
      console.error("[OUTLOOK] Could not load tab switcher");
    });
  }

  select(tab, keepFocus = false) {
    if (this.#disposed || !["teams", "outlook"].includes(tab)) return;
    this.#active = tab;
    this.#layout();
    this.#outlook.setVisible(tab === "outlook");
    if (tab === "outlook") {
      this.#window.contentView.addChildView(this.#outlook);
      if (this.#status === "idle") this.#load();
      if (!keepFocus) this.#outlook.webContents.focus();
    } else {
      if (!keepFocus) this.#getTeamsContents()?.focus();
    }
    this.raise();
    this.#push();
  }

  raise() {
    if (this.#disposed || !this.#bar) return;
    if (this.#active === "outlook") this.#window.contentView.addChildView(this.#outlook);
    this.#window.contentView.addChildView(this.#bar);
  }

  #layout() {
    if (this.#disposed) return;
    const [width] = this.#window.getContentSize();
    this.#bar.setBounds({ x: 0, y: 0, width, height: HEADER_HEIGHT });
    this.#outlook.setBounds(workspaceBounds(this.#window, this.#config));
  }

  async #load() {
    if (this.#loading || this.#disposed) return;
    this.#loading = true;
    this.#navigationFailed = false;
    this.#status = "loading";
    this.#push();
    try {
      await this.#outlook.webContents.loadURL(this.#config.outlook?.url ?? DEFAULT_URL);
    } catch (error) {
      if (!this.#disposed && error.code !== "ERR_ABORTED") {
        this.#status = "error";
        this.#push();
        console.warn("[OUTLOOK] Page load failed; use Retry in the tab switcher");
      }
    } finally {
      this.#loading = false;
    }
  }

  #configureNavigation(contents) {
    contents.on("will-navigate", (event, value) => {
      const url = new URL(value);
      if (url.protocol === "https:") return;
      event.preventDefault();
      this.#openExternal(value);
    });
    contents.setWindowOpenHandler(({ url }) => {
      let parsed;
      try { parsed = new URL(url); } catch { return { action: "deny" }; }
      // Microsoft and federated sign-in flows may open an empty window first
      // and navigate it afterwards. It still has no preload or Node access.
      if (url === "about:blank" || (parsed.protocol === "https:" &&
          (AUTH_HOSTS.has(parsed.hostname) || OUTLOOK_HOSTS.has(parsed.hostname)))) {
        return { action: "allow", overrideBrowserWindowOptions: {
          parent: this.#window,
          webPreferences: { contextIsolation: true, sandbox: true, nodeIntegration: false, webviewTag: false },
        } };
      }
      this.#openExternal(url);
      return { action: "deny" };
    });
    contents.on("did-create-window", (child) => {
      this.#popups.add(child);
      child.once("closed", () => this.#popups.delete(child));
      this.#configureNavigation(child.webContents);
    });
  }

  #openExternal(value) {
    try {
      if (!["https:", "http:", "mailto:", "tel:"].includes(new URL(value).protocol)) return;
      shell.openExternal(value).catch(() => console.warn("[OUTLOOK] Could not open external link"));
    } catch { /* Invalid or unsupported link. */ }
  }

  #push() {
    if (this.#disposed || !this.#bar || this.#bar.webContents.isDestroyed()) return;
    this.#bar.webContents.send("workspace-tab-state", {
      active: this.#active, status: this.#status,
      integratedTitleBar: isIntegratedTitleBar(this.#config),
    });
  }

  dispose() {
    if (this.#disposed) return;
    this.#disposed = true;
    ipcMain.removeListener(CHANNEL, this.#action);
    this.#window.removeListener("resize", this.#resize);
    this.#window.removeListener("focus", this.#focus);
    for (const popup of this.#popups) if (!popup.isDestroyed()) popup.destroy();
    this.#popups.clear();
    for (const view of [this.#bar, this.#outlook]) {
      if (view && !view.webContents.isDestroyed()) view.webContents.close();
    }
  }
}

module.exports = { OutlookTabs, isOutlookUrl, DEFAULT_URL };
