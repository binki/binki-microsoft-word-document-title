Update the browser’s title to match the Microsoft Word document’s title.

## Problem

When editing files in Microsoft Word Online using one’s [personal OneDrive](https://onedrive.live.com/) or [business OneDrive](https://office.live.com/start/OneDrive.aspx), Word updates the browser’s title to match the document:

![Editing “Example Word Document.docx”.](https://i.imgur.com/C6jeCTt.png)

This way, you can search for the document by name in your address bar:

![Typing “example word” into the address bar.](https://i.imgur.com/0nSyjan.png)

However, if you launched a Microsoft Word document through the [special WOPI protocol](https://docs.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/), Word does not do this.
For example, a document named “Another Example.docx” just has a document title of “Word”:

![Editing “Another Example.docx” opened via ShareFile](https://i.imgur.com/BVzFMrg.png)

This makes searching for the document nearly impossible. For example, I have no results in my address bar:

![Typing “another example” into the address bar.](https://i.imgur.com/erk5Z3r.png)

## Solution

This user script detects this situation and monitors the document’s title shown in the in-page blue header and copies the text to the document’s title.
This updates the displayed value shown by the browser, making the tab searchable.

However, this does not solve the problem completely.
Every time you visit this document on a synced Firefox instance, Firefox will update the document title for that URI to the latest one it saw.
So if you launch the document via Firefox for Android, which cannot run user scripts, you will overwrite the document’s history entry title to just be “Word”.
To mitigate this issue, this script transparently adds an `x-title` GET parameter to the URI of the document.
This provides two functionalities.
During future loads of the document, the document’s title can be updated immediately upon navigating to the document—even before the document loads enough to populate the header within the document with the title and even if there is an issue with loading the document itself.
Also, the address bar search can match the document title keywords in the URI itself, mitigating the fact that the title may have been overwritten by a device participating in Firefox Sync that does not have this userscript installed.

## Installation

[Install](binki-microsoft-word-document-title.user.js?raw=1)

## See Also

* [binki-sharefile-microsoft-word-recovery](https://github.com/binki/binki-sharefile-microsoft-word-recovery): Enable direct navigation to the WOPI editor for Word documents stored in ShareFile, recovering from the “Sorry, we ran into a problem.”. This is likely necessary if you using ShareFile and search for documents by title.
