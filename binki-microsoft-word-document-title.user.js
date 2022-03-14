// ==UserScript==
// @name binki-microsoft-word-document-title
// @version 1.0
// @author Nathan Phillip Brink (binki) (@ohnobinki)
// @homepageURL https://github.com/binki/binki-microsoft-word-document-title/
// @match https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?*
// @require https://github.com/binki/binki-userscript-delay-async/raw/252c301cdbd21eb41fa0227c49cd53dc5a6d1e58/binki-userscript-delay-async.js
// @require https://github.com/binki/binki-userscript-when-element-changed-async/raw/88cf57674ab8fcaa0e86bdf5209342ec7780739a/binki-userscript-when-element-changed-async.js
// ==/UserScript==
(async () => {
  const titleParamName = 'x-title';
  // Extract the title from the URI before the document fully loads so we can set it more quickly
  // if it is already set. This also sets the title if the document is in a broken/unauthenticated
  // state (for Sharefile documents, consider using binki-sharefile-microsoft-word-recovery to automatically
  // refresh these).
  {
    const url = new URL(location.href);
    const title = url.searchParams.get(titleParamName);
    if (title) {
      document.title = title;
    }
  }
  while (true) {
    const throttlePromise = delayAsync(200);

    const documentTitleSpan = document.querySelector('span[data-unique-id=DocumentTitleContent]');
    if (documentTitleSpan) {
      const title = documentTitleSpan.textContent;
      // Put the title in a GET parameter in case if the page is loaded on a Greasemonkey-free
      // copy of Firefox such as Firefox for Android which will overwrite the user’s history
      // with an untitled enry, breaking address bar completion. Having it in the URI itself
      // will mitigate this somewhat.
      const url = new URL(location.href);
      url.searchParams.set(titleParamName, title);
      history.replaceState(history.state, title, url.toString());
      document.title = title;
      // Even though it is just text update, seems like it actually updates the text node itself.
      // This doesn’t trigger characterData but does trigger subtree. So no point in trying to make
      // this more specific—it is already specific to the node itself.
      await whenElementChangedAsync(documentTitleSpan);
    } else {
      // Wait for this element to appear.
      await whenElementChangedAsync(document.body);
    }

    // Don’t go too fast. If we received two mutation events in a row, wait a bit for the next
    // one. If nothing had happened for a while, the delay will already be expired so this won’t
    // hold anything up.
    await throttlePromise;
  }
})();
