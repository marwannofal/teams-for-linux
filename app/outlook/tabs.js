const tabs = ["teams", "outlook"].map((id) => document.getElementById(id));
for (const tab of tabs) {
  tab.addEventListener("click", () => window.workspaceTabs.select(tab.id));
  tab.addEventListener("keydown", (event) => {
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
    event.preventDefault();
    const next = event.key === "Home" ? tabs[0] : event.key === "End" ? tabs[1] : tabs.find((item) => item !== tab);
    next.focus();
    window.workspaceTabs.select(next.id, true);
  });
}
document.getElementById("reload").addEventListener("click", () => window.workspaceTabs.select("reload"));
window.workspaceTabs.onState(({ active, status, integratedTitleBar }) => {
  document.body.dataset.integrated = String(integratedTitleBar === true);
  for (const tab of tabs) {
    tab.setAttribute("aria-selected", String(tab.id === active));
    tab.tabIndex = tab.id === active ? 0 : -1;
  }
  const reload = document.getElementById("reload");
  reload.hidden = active !== "outlook";
  reload.title = status === "error" ? "Retry Outlook" : "Reload Outlook";
  reload.setAttribute("aria-label", reload.title);
  document.getElementById("status").textContent = active !== "outlook" ? "" :
    status === "loading" ? "Loading…" : status === "error" ? "Load failed" : "";
});
