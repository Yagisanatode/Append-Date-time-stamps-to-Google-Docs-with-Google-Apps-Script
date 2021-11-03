/**
 * ########### APPEND GOOGLE DOC PARAGRAPHS, LIST ITEMS AND TABLE CELL ITEMS WITH A DATE-TIME STAMP ###########
 * @author Yagisanatode <yagisanatode@gmail.com>
 * [Check out the tutorial]{@link https://yagisanatode.com/2021/10/31/append-list-items-paragraphs-and-table-cell-items-with-a-date-time-stamp-in-google-docs-using-google-apps-script/}
 */

/**
 * Creates a menu item in the current document. 
 * Google Apps Script simple trigger 
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu("CheckBox DTS")
    .addItem("Add DTS", "appendElement")
    .addToUi();
}


/**
 * Creates a date-time stamp from the area the cursor is on.
 * Called when the "Add DTS" button is selected from the menu.
 */
function appendElement() {
  // Document variables. 
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const element = cursor.getElement();

  // Input variables.
  let date = new Date();
  let email = Session.getActiveUser().getEmail();
  const inputVal = ` -Complete: ${date.toDateString()} ${date.toLocaleTimeString()} | ${email} `;// Add a space at the end.
  
  // Element lenght before and after adding date-time stamp.
  const inputValLength = inputVal.length;
  const elementLength = element.asText().getText().length;
  const newElementLength = elementLength + inputValLength - 2; // -2 to not style the last space.

  // Alternative approach to text formatting. Good for 1 or two style items but slow for more.
  //  element.asText()
  //   .appendText(inputVal)
  //   .setItalic(elementLength, newElementLength, true)
  //   .setBold(elementLength, newElementLength, true)
  //   .setFontFamily(elementLength, newElementLength, "Merriweather")
  //   .setForegroundColor(elementLength, newElementLength, "#34a853")
  //   .setBackgroundColor(elementLength, newElementLength, "#eeeeee")

  // Create the styling for the text.
  const attr = DocumentApp.Attribute;
  const style = {
    [attr.FONT_FAMILY]: "Merriweather",
    [attr.ITALIC]: true,
    [attr.BOLD]: true,
    [attr.BACKGROUND_COLOR]: "#eeeeee",
    [attr.FOREGROUND_COLOR]: "#34a853",
  };

  // Append the text and add the styling to it. 
  element.asText()
    .appendText(inputVal)
    .setAttributes(elementLength, newElementLength, style);

};

