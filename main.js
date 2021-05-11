function onOpen() {
 // executes on google sheet, so call google sheet api
 // function to generate additional menu item in the g-sheet UI to run custom command
 const ui = SpreadsheetApp.getUi();
 const menu = ui.createMenu('Generate Decks');
 menu.addItem('Create New Decks', 'createNewGoogleSlides');
 menu.addToUi();
}

//currently works on open - working on finding a way to eliminate that button press because
//that would be kind of cool 

function createNewGoogleSlides() {
  // Identify the template slide
  const googleSlideTemplate = DriveApp.getFileById('1FepW7lDnWqFRit7iELb-i8pnTYFKfknTpViE0pLSLhs');
  // identify destination folder
  const destinationFolder = DriveApp.getFolderById('1drI8EKQ5xgmVhYM8lLZkGzxxWJzcSS4e')
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const rows = sheet.getDataRange().getValues();
  Logger.log(rows)
  //Logger.log('checkpoint a')


  rows.forEach(function(row, index){
    if (index === 0) return;
    //check if deck already exists
    if (row[0]) return; //doc link index
    //Logger.log(row)
    const copy = googleSlideTemplate.makeCopy(`${row[1]}, ${row[2]} Post Fight Summary`, destinationFolder)
    //works up to here
    //make a copy of the deck template


    //need to get deck id for here
    const deck = SlidesApp.openById(copy.getId());
    //const body = slide.getBody();
    //console.log('checkpoint2')

    //date formatter
    const friendlyDate = new Date(row[6]).toLocaleDateString();


    //assign values here
    
    deck.replaceAllText('{{Client}}', row[2]);
    deck.replaceAllText('{{First Name}}', row[3]);
    deck.replaceAllText('{{Last Name}}', row[4]);
    deck.replaceAllText('{{Event}}', row[5]);
    deck.replaceAllText('{{Event Date}}', friendlyDate);

    //deck.replace('{{img}}',)

    const slide1 = deck.getSlideById('p')
    const pageElement = slide1.getPageElementById('gd872a1c164_0_9')

    //gd872a1c164_0_9

    const image = pageElement.asImage();
    Logger.log(slide1.getImages());


    //replace this w url from row in sheets
    image.replace(row[7], true)



    deck.saveAndClose();
    const url = deck.getUrl();

    MailApp.sendEmail(row[1], row[2], url)

    sheet.getRange(index + 1, 1).setValue(url)
  
  })

  
}

