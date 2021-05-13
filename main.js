// !!!! ANY LINES BEGINNING WITH DOUBLE SLASH (//) ARE CODE COMMENTS DESCRIBING THIS FUNCTION !!!!
//The following function generates a new menu in the UI in order to run the script.
//This isn't completely necessary and is primarily used for testing as it provides an easy way to trigger functions.
//Button creation function runs when opening the spreadsheet. 
//The deck generation function will run automatically when a form is submitted.
function onOpen() {
 //API call to add new menu in the UI
 const ui = SpreadsheetApp.getUi();
 const menu = ui.createMenu('Generate Decks');

 //Populate new menu with a button that calls the deck generation function
 menu.addItem('Create New Post Fight Decks', 'createNewGoogleSlides');
 menu.addToUi();
}

//File and Folder IDs have been removed for code upload.
//File and Folder IDs need to be modified to apply to different sheets/slides/docs

function createNewGoogleSlides() {
  // Identify the template slide via its file ID. This is found in the slide URL
  const googleSlideTemplate = DriveApp.getFileById('Template ID Goes Here as String');

  // Identify destination folder for the newly generated deck. Again, can be found in the folder URL
  const destinationFolder = DriveApp.getFolderById('Destination Folder ID Goes Here as String')

  // Identify individual sheet to pull data from
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormResponses');

  // Split sheet data into rows
  const rows = sheet.getDataRange().getValues();
  Logger.log(rows)

  // Iterate through each row in the spreadsheet and check to see if a deck already exists.
  // If there is no link to the deck in the first cell of a row, a new one is generated.
  rows.forEach(function(row, index){

    //check first index for doc link
    if (index === 0) return;

    // If the link already exists, move to next row
    if (row[0]) return; 

    //create a copy of the template with new name and place it into the folder defined above
    const copy = googleSlideTemplate.makeCopy(`${row[5]}, ${row[2]} Post Fight Summary`, destinationFolder)


    //Open the newly created deck in order to begin edits.
    const deck = SlidesApp.openById(copy.getId());
    //const body = slide.getBody();
    //console.log('checkpoint2')

    //By default, the date will include way too much information. It will go down to the first second of the day
    //This line just refromats that into MM/DD/YYYY
    const betterDate = new Date(row[6]).toLocaleDateString();

    //This is where 'smart fields' come in. The program just executes a find and replace function.
    //Defined tags will be replaced with data from the spreadsheet row. Formatting is retained from deck template
    //When Applying this script to new files/templates, this is where the bulk of the editing is.  
    
    deck.replaceAllText('{{Client}}', row[2]);
    deck.replaceAllText('{{First Name}}', row[3]);
    deck.replaceAllText('{{Last Name}}', row[4]);
    deck.replaceAllText('{{Event}}', row[5]);
    deck.replaceAllText('{{Event Date}}', betterDate);
    deck.replaceAllText('{{FightNightPurse}}', row[8]);
    deck.replaceAllText('{{comP}}', row[9]);
    deck.replaceAllText('{{PSMCom}}', row[10]);
    deck.replaceAllText('{{ClientResults}}', row[11])
    deck.replaceAllText('{{fNum1}}', row[12]);
    deck.replaceAllText('{{fNum2}}', row[13]);
    deck.replaceAllText('{{Promotion}}', row[14]);
    deck.replaceAllText('{{NewFollowers}}', row[15]);
    deck.replaceAllText('{{GrowthDuring}}', row[16]);
    deck.replaceAllText('{{AccountsReached}}', row[17]);
    deck.replaceAllText('{{WeeklyImpressions}}', row[18]);
    deck.replaceAllText('{{CashTotal}}',row[19]);
    deck.replaceAllText('{{Sponsor1}}', row[20]);
    deck.replaceAllText('{{Fee1}}', row[22]);
    deck.replaceAllText('{{PSMcom1}}', row[23]);
    deck.replaceAllText('{{Term1}}', row[24]);
    deck.replaceAllText('{{Deliverables1}}', row[25]);
    deck.replaceAllText('{{Sponsor2}}', row[26]);
    deck.replaceAllText('{{Fee2}}', row[28]);
    deck.replaceAllText('{{PSMcom2}}', row[29]);
    deck.replaceAllText('{{Term2}}', row[30]);
    deck.replaceAllText('{{Deliverables2}}', row[31]);
    deck.replaceAllText('{{Sponsor3}}', row[32]);
    deck.replaceAllText('{{Fee3}}', row[34]);
    deck.replaceAllText('{{PSMcom3}}', row[35]);
    deck.replaceAllText('{{Term3}}', row[36]);
    deck.replaceAllText('{{Deliverables3}}', row[37]);
    deck.replaceAllText('{{cmLink}}', row[38]);
    deck.replaceAllText('{{oLink}}', row[39]);
    //deck.replaceAllText


    // I'm still having an issue with automatic image uploads - need to replace drive image uploads with 
    //hosted images - for instance, it works with stock images. Current workflow collects images, but requires drag and drop.

    //deck.replace('{{img}}',)
    //todo - figure out image permissions
    const slide1 = deck.getSlideById('p');
    const pageElement = slide1.getPageElementById('gd872a1c164_0_9');

    //gd872a1c164_0_9

    const image = pageElement.asImage();
    Logger.log(slide1.getImages());


    //replace this w url from row in sheets - change to true for cropping
    //image.replace(row[7], false);

    // Save changes to the deck and fetch it's URL
    deck.saveAndClose();
    const url = deck.getUrl();


    //Send an email confirmation to whoever submitted the form containing basic instructions as well as the URL
    //todo - add other emails and reformat
    
    MailApp.sendEmail({
      to: row[42],
      subject: (row[2] + " Post Fight Summary"),
      htmlBody: "Hello, <br> <br> Here is the post-fight deck you requested for " + row[2] 
      +": <br> <br>" + url + "<br> <br> Please review to ensure accuracy and remove unused sections. <br> <br> An Asana task will be generated in the next few minutes. Once the Asana task is marked as complete by yourself or Griffin, a record describing this deck will be uploaded to Insightly. <br> <br> Thanks, <br> Jacob"
    })

    //insert url value into spreadhsheet
    sheet.getRange(index + 1, 1).setValue(url)
  
  })

  
}
