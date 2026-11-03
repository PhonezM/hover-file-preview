/* The PowerPoint renderer is 685 KB of the extension. Loading it into every
   page costs parse/compile time even on pages without a deck, so it is
   injected on demand into the tab's isolated world (the content script's
   global) only when a .pptx link is actually hovered.

   A dynamic import would be smaller, but Firefox MV3 still cannot import an
   extension module from a content script (Bugzilla 1803950), and a page-world
   <script> would be subject to the page's CSP. scripting.executeScript runs in
   the isolated world under the extension CSP, which works in both browsers. */
(function () {
  "use strict";

  var api = typeof browser !== "undefined" ? browser : chrome;
  var RENDERER = ["dist/pptx.js"];

  api.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
    if (!msg || msg.type !== "hfp:load-renderer") return;
    var tabId = sender && sender.tab && sender.tab.id;
    if (typeof tabId !== "number") {
      sendResponse({ ok: false, error: "no tab" });
      return;
    }
    var target = { tabId: tabId };
    if (typeof sender.frameId === "number") target.frameIds = [sender.frameId];

    Promise.resolve(api.scripting.executeScript({ target: target, files: RENDERER }))
      .then(function () { sendResponse({ ok: true }); })
      .catch(function (e) {
        sendResponse({ ok: false, error: String((e && e.message) || e) });
      });
    return true;   // keep the message channel open for the async reply
  });
})();
