const { isOutlookUrl } = require("./urls");

const HEADER_HEIGHT = 52;
const hasOutlookHeader = (config) => config.outlook?.enabled === true && isOutlookUrl(config.outlook.url);

function workspaceBounds(window, config) {
  const [width, height] = window.getContentSize();
  const y = hasOutlookHeader(config) ? HEADER_HEIGHT : 0;
  return { x: 0, y, width, height: Math.max(0, height - y) };
}

module.exports = { HEADER_HEIGHT, hasOutlookHeader, workspaceBounds };
