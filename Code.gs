/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Format', 'formatMarkdown')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

function formatMarkdown() {
  processSourceCode();
  processBackquotes();
  processBold();
  processLinks();
  processItalics();
  processHeadings();
}

/**
 * Search for two lines starting with ``` in the doc,
 * and add all the lines appearing between them into a single-cell
 * table, set the font to a monospace font.
 *
 * TODO(sujeet): Syntax highliting (using hilite.me maybe?).
 */
function processSourceCode() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  var startingTripleTick = body.findText('```');
  if (!startingTripleTick) return;
  var endTripleTick = body.findText('```', startingTripleTick);
  if (!endTripleTick) return;
  
  var firstLine = startingTripleTick.getElement();
  var lastLine = endTripleTick.getElement();
  
  var rangeBuilder = doc.newRange();
  rangeBuilder.addElementsBetween(firstLine, lastLine);
  var range = rangeBuilder.build();
  var lineRanges = range.getRangeElements();
  var lines = [];
  
  var firstLineIndex = body.getChildIndex(lineRanges[0].getElement());
  var code = "";
  
  // Don't iterate over 0th and last line because they are the tripleticks
  lineRanges[0].getElement().removeFromParent();
  for (var i = 1; i < lineRanges.length - 1; ++i) {
    code += lineRanges[i].getElement().asText().getText() + '\n';
    lineRanges[i].getElement().removeFromParent();
  }
  lineRanges[lineRanges.length-1].getElement().removeFromParent();
  
  var cell = body.insertTable(firstLineIndex)
                 .setBorderWidth(0)
                 .appendTableRow()
                 .appendTableCell();
  cell.setText(code.trim());
  cell.setBackgroundColor('#f0f0f2');
  cell.setFontFamily('Consolas');
  
  processSourceCode();
}


/**
 * Search for `some text` and replace it with
 * its backtick-free version with a monospace font
 * (uses slack color theme presently)
 */
function processBackquotes() {
  var backquote = DocumentApp
                    .getActiveDocument()
                    .getBody()
                    .findText('`.*?`');
  if (backquote) {
    var start = backquote.getStartOffset();
    var end = backquote.getEndOffsetInclusive();
    var text = backquote.getElement().asText();
    text.setBackgroundColor(start, end, '#f0f0f2');
    text.setFontFamily(start, end, 'Consolas');
    text.setForegroundColor(start, end, '#cc2255');
    text.deleteText(end, end);
    text.deleteText(start, start);
    processBackquotes();
  }
}


/**
 * Search for **some text** and replace it with its
 * asterisk-free version with a bold face.
 */
function processBold() {
  var bold = DocumentApp
               .getActiveDocument()
               .getBody()
               .findText('\\*\\*.*?\\*\\*');
  if (bold) {
    var start = bold.getStartOffset();
    var end = bold.getEndOffsetInclusive();
    var text = bold.getElement().asText();
    text.setBold(start, end, true);
    text.deleteText(end-1, end);
    text.deleteText(start, start+1);
    processBold();
  }
}


/**
 * Search for _some text_ and replace it with its
 * underscore-free, italicized version.
 */
function processItalics() {
  var italics = DocumentApp
                  .getActiveDocument()
                  .getBody()
                  .findText(' _.*?_ ');
  if (italics) {
    var start = italics.getStartOffset();
    var end = italics.getEndOffsetInclusive();
    var text = italics.getElement().asText();
    text.setItalic(start, end, true);
    text.deleteText(end-1, end-1);
    text.deleteText(start+1, start+1);
    processItalics();
  }
}


/**
 * Convert patterns of the form [Link Name](http://example.com/address)
 * to hyperlinks where the link text is "Link Name" and
 * the link url is "http://example.com/address"
 */
function processLinks() {
  // Links are of the form "[Link Name](http://example.com/page/address)"
  var link = DocumentApp
               .getActiveDocument()
               .getBody()
               .findText('\\[.*?\\]\\(https?:\\/\\/.*?\\)');
  if (link) {
    var start = link.getStartOffset();
    var end = link.getEndOffsetInclusive();
    var text = link.getElement().asText();
    var linkName = text.getText().split('[')[1].split(']')[0];
    var url = text.getText().split(']')[1].split('(')[1].split(')')[0];
    text.deleteText(start, end);
    text.insertText(start, linkName);
    text.setLinkUrl(start, start + linkName.length - 1, url);
    processLinks();
  }
}


/**
 * Do the following conversions:
 * # my heading   -> "my heading" styled as Heading1
 * ## another one -> "another one" styled as Heading2
 * ### third      -> "third" styled as Heading3
 */
function processHeadings() {
  var headingStarts = ['# ', '## ', '### '];
  var headingFormats = [
    DocumentApp.ParagraphHeading.HEADING1,
    DocumentApp.ParagraphHeading.HEADING2,
    DocumentApp.ParagraphHeading.HEADING3
  ];
  for (var i = 0; i < headingStarts.length; ++i) {
    var headingStart = headingStarts[i];
    var heading = DocumentApp
                    .getActiveDocument()
                    .getBody()
                    .findText(headingStart + '.*');
    while (heading) {
      var start = heading.getStartOffset();
      if (start == 0) {
        // We want to style the text as a heading only if
        // the paragraph starts with the pounds.
        var elem = heading.getElement();
        elem.asText().deleteText(0, i+1);
        while (elem.getType() != DocumentApp.ElementType.PARAGRAPH) {
          elem = elem.getParent();
        }
        elem.setHeading(headingFormats[i]);
      }
      heading = DocumentApp
                  .getActiveDocument()
                  .getBody()
                  .findText(headingStart + '.*', heading);
    }
  }
}
