function onOpen() {
  const ui = DocumentApp.getUi();

  // Bullet Tools Submenu
  const bulletMenu = ui.createMenu("Bullet Tools")
    .addItem("Remove Bullet Break Lines", "promptcleanBulletsAndJoinupdated");

  // Link Tools Submenu
  const linkMenu = ui.createMenu("Link Tools")
    .addItem("Add URL to Paragraph Text", "promptaddLinkToParagraphsOnlyupdated")
    .addItem("Change URL and Text", "promptchangeTextWithUrlIfLinkedupdated");

  // Utility / Experimental Submenu
  const utilityMenu = ui.createMenu("Articles")
    .addItem("Total No. of Articles", "countArticles")
    .addItem("Check links", "validateLinksPerArticle")
    .addItem("Check Bullet points", "validateBulletGroupsWithTitle");

  // Main ASVSI Menu
  ui.createMenu("ASVSI")
    .addSubMenu(bulletMenu)
    .addSubMenu(linkMenu)
    .addSeparator()
    .addSubMenu(utilityMenu)
    .addItem("OMG", "omg")
    .addToUi();
}


function promptchangeTextWithUrlIfLinkedupdated() {
  const ui = DocumentApp.getUi();

  const targetKeywordResponse = ui.prompt(
    "ASVSI SmartLink Engine",
    "Enter the Target keyword:",
    ui.ButtonSet.OK_CANCEL
  );

  if (targetKeywordResponse.getSelectedButton() !== ui.Button.OK) return;

  const textResponse = ui.prompt(
    "ASVSI SmartLink Engine",
    "Enter the Text to Replace with:",
    ui.ButtonSet.OK_CANCEL
  );

  if (textResponse.getSelectedButton() !== ui.Button.OK) return;

  const urlResponse = ui.prompt(
    "ASVSI SmartLink Engine",
    "Enter the URL to apply:",
    ui.ButtonSet.OK_CANCEL
  );

  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;

  const targetkeyword = targetKeywordResponse.getResponseText().trim();
  const replacekeyword = textResponse.getResponseText().trim();
  const url = urlResponse.getResponseText().trim();

  if (!targetkeyword || !replacekeyword || !url) {
    ui.alert("Keyword and URL cannot be empty.");
    return;
  }

  changeTextWithUrlIfLinkedupdated(targetkeyword, replacekeyword, url);

  ui.alert("Links added successfully.");
}

function promptaddLinkToParagraphsOnlyupdated() {
  const ui = DocumentApp.getUi();

  const keywordResponse = ui.prompt(
    "ASVSI SmartLink Engine",
    "Enter The keyword:",
    ui.ButtonSet.OK_CANCEL
  );

  if (keywordResponse.getSelectedButton() !== ui.Button.OK) return;

  const urlResponse = ui.prompt(
    "ASVSI SmartLink Engine",
    "Enter the URL for Keyword:",
    ui.ButtonSet.OK_CANCEL
  );

  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;

  const targetkeyword = keywordResponse.getResponseText().trim();
  const url = urlResponse.getResponseText().trim();

  if (!targetkeyword || !url) {
    ui.alert("Keyword and URL cannot be empty.");
    return;
  }

  addLinkToParagraphsOnlyupdated(targetkeyword, url);

  ui.alert("Links added successfully.");
}

function promptcleanBulletsAndJoinupdated() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
      "ASVSI SmartLink Engine",
    "Remove break line between bullet points?",
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.NO) return;

  cleanBulletsAndJoinupdated();

  ui.alert("Bullet point's Break lines removed successfully.");
}

function cleanBulletsAndJoinupdated() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  for (let i = 0; i < paragraphs.length - 1; i++) {
    const current = paragraphs[i];
    const next = paragraphs[i + 1];

    // Only clean list items (bulleted or numbered)
    if (current.getType() === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = current.asListItem();

      // Store original attributes (bullet style, indentation, etc.)
      const attrs = listItem.getAttributes();

      // Trim trailing spaces safely
      const cleanText = listItem.getText().replace(/\s+$/g, "");
      if (cleanText !== listItem.getText()) {
        listItem.setText(cleanText);
        listItem.setAttributes(attrs); // restore bullet style
      }

      // Remove blank line after bullet (only if it's not a list item)
      if (
        next.getText().trim() === "" &&
        next.getType() !== DocumentApp.ElementType.LIST_ITEM
      ) {
        const parent = next.getParent();
        try {
          parent.removeChild(next);
          i--; // adjust index
        } catch (e) {
          Logger.log("Skipped a nested or invalid paragraph: " + e);
        }
      }
    }
  }

  Logger.log("Cleaned bullets and kept original styles intact!");
}

const addLinkToParagraphsOnlyupdated = (_searchText,_linkUrl) => {
  const searchText = _searchText;
  const linkUrl = _linkUrl;

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const total = body.getNumChildren();

  for (let i = 0; i < total; i++) {
    const element = body.getChild(i);

    // Only allow PARAGRAPH elements
    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph = element.asParagraph();

    // Strictly allow only NORMAL paragraphs (exclude title & headings)
    if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) continue;

    const text = paragraph.editAsText();
    let match = null;

    while ((match = text.findText(searchText, match))) {
      const start = match.getStartOffset();
      const end = match.getEndOffsetInclusive();

      // Skip if already linked
      if (text.getLinkUrl(start)) continue;

      text.setLinkUrl(start, end, linkUrl);
    }
  }

  doc.saveAndClose();
};


const changeTextWithUrlIfLinkedupdated = (_targetText,_replaceText,_replaceUrl) => {
  // const searchText = 'NCHM JEE';
  // const replaceText = 'NCHM JEE1';
  // const replaceUrl = 'https://asvsi.com/';
  const searchText = _targetText;
  const replaceText = _replaceText;
  const replaceUrl = _replaceUrl;

  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  let search = null;

  while ((search = body.findText(searchText, search))) {
    const element = search.getElement();
    const startIndex = search.getStartOffset();
    const endIndex = search.getEndOffsetInclusive();
    const textElement = element.asText();

    // ✅ Only replace if the matched text is already a hyperlink
    const currentLink = textElement.getLinkUrl(startIndex);
    if (!currentLink) {
      continue; // skip plain text
    }

    // Delete old text and insert new
    textElement.deleteText(startIndex, endIndex);
    textElement.insertText(startIndex, replaceText);
    textElement.setLinkUrl(startIndex, startIndex + replaceText.length - 1, replaceUrl);
  }

  document.saveAndClose();
};
function countArticles() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();

  let count = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    if (paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.TITLE) {
      count++;
    }
  }
  
  DocumentApp.getUi().alert("Total Articles:" + count);
}
function validateLinksPerArticle() {
  const body = DocumentApp.getActiveDocument().getBody();
  const total = body.getNumChildren();

  let results = [];
  let currentArticle = null;
  let linkCount = 0;

  for (let i = 0; i < total; i++) {
    const element = body.getChild(i);

    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph = element.asParagraph();
    const heading = paragraph.getHeading();

    // New article detected
    if (heading === DocumentApp.ParagraphHeading.TITLE) {

      if (currentArticle !== null) {
        results.push({ title: currentArticle, links: linkCount });
      }

      currentArticle = paragraph.getText();
      linkCount = 0;
    }

    // Count links in this paragraph safely
    linkCount += countLinksInParagraph(paragraph);
  }

  // Push last article
  if (currentArticle !== null) {
    results.push({ title: currentArticle, links: linkCount });
  }

  // Validate
  let message = "";
  let hasIssue = false;

  results.forEach(article => {
    if (article.links !== 3) {
      hasIssue = true;
      message += `"${article.title}" → ${article.links} links\n`;
    }
  });

  if (hasIssue) {
    DocumentApp.getUi().alert(
      "Articles with invalid link count (should be 3):\n\n" + message
    );
  } else {
    DocumentApp.getUi().alert("All articles have exactly 3 links ✅");
  }
}


function countLinksInParagraph(paragraph) {
  const text = paragraph.editAsText();
  const length = text.getText().length;

  let count = 0;
  let inLink = false;

  for (let i = 0; i < length; i++) {
    const url = text.getLinkUrl(i);

    if (url && !inLink) {
      count++;
      inLink = true;
    }

    if (!url) {
      inLink = false;
    }
  }

  return count;
}

function validateBulletGroupsWithTitle() {
  const body = DocumentApp.getActiveDocument().getBody();
  const total = body.getNumChildren();

  let currentTitle = null;
  let currentBullets = [];
  let results = [];

  for (let i = 0; i < total; i++) {
    const element = body.getChild(i);
    const type = element.getType();

    // Detect TITLE
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      if (paragraph.getHeading() === DocumentApp.ParagraphHeading.TITLE) {

        // Before switching title, validate previous bullet group
        if (currentBullets.length > 0) {
          if (currentBullets.length < 5 || currentBullets.length > 6) {
            results.push({
              title: currentTitle,
              bullets: [...currentBullets]
            });
          }
          currentBullets = [];
        }

        currentTitle = paragraph.getText();
      }
    }

    // Detect bullet items
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = element.asListItem();
      currentBullets.push(listItem.getText());
    } else {
      // If bullet group ends, validate it
      if (currentBullets.length > 0) {
        if (currentBullets.length < 5 || currentBullets.length > 6) {
          results.push({
            title: currentTitle,
            bullets: [...currentBullets]
          });
        }
        currentBullets = [];
      }
    }
  }

  // Final check at end of document
  if (currentBullets.length > 0) {
    if (currentBullets.length < 5 || currentBullets.length > 6) {
      results.push({
        title: currentTitle,
        bullets: [...currentBullets]
      });
    }
  }

  // Show result
  if (results.length === 0) {
    DocumentApp.getUi().alert("All bullet groups contain 5–6 items ✅");
    return;
  }

  let message = "Bullet Group Errors:\n\n" + results
  .map(group => 
    `• ${group.title || "Untitled"} — ${group.bullets.length} items`
  )
  .join("\n");

  DocumentApp.getUi().alert(message);
}


const omg=()=>{
    const ui = DocumentApp.getUi();
  ui.alert("Done!");
}
