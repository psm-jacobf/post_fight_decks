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

//function to get id from image url
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }


function createNewGoogleSlides() {
  // Identify the template slide via its file ID. This is found in the slide URL
  const googleSlideTemplate = DriveApp.getFileById('1FepW7lDnWqFRit7iELb-i8pnTYFKfknTpViE0pLSLhs');

  // Identify destination folder for the newly generated deck. Again, can be found in the folder URL
  const destinationFolder = DriveApp.getFolderById('1uc09uKRJoL6P4hgwwm7BTNxP2tiHxU-E')

  // Identify individual sheet to pull data from
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormResponses');

  // Split sheet data into rows
  const rows = sheet.getDataRange().getValues();
  //Logger.log(rows)

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

    //currency formatter
    var cur_formatter = new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0, // (this suffices for whole numbers, but will print 2500.10 as $2,500.1)
      maximumFractionDigits: 0, // (causes 2500.99 to be printed as $2,501)
    });

    if(row[43].length < 1){
      var bonus_val = 0
    } else {
      var bonus_val = parseInt(row[43]);
    }
    //This is where 'smart fields' come in. The program just executes a find and replace function.
    //Defined tags will be replaced with data from the spreadsheet row. Formatting is retained from deck template
    //When Applying this script to new files/templates, this is where the bulk of the editing is.  
    
    deck.replaceAllText('{{Client}}', row[2]);
    deck.replaceAllText('{{First Name}}', row[3]);
    deck.replaceAllText('{{Last Name}}', row[4]);
    deck.replaceAllText('{{Event}}', row[5]);
    deck.replaceAllText('{{Event Date}}', betterDate);
    deck.replaceAllText('{{FightNightPurse}}', cur_formatter.format(row[8]));
    deck.replaceAllText('{{FightNightBonus}}', cur_formatter.format(bonus_val))
    var fntotal = parseInt(row[8]) + parseInt(bonus_val)
    
    deck.replaceAllText('{{FightNightTotal}}', cur_formatter.format(fntotal))
    deck.replaceAllText('{{comP}}', row[9]);
    var comval = parseInt(row[10].value);
    //deck.replaceAllText('{{PSMCom}}', comval);
    deck.replaceAllText('{{PSMCom}}', cur_formatter.format(fntotal * .1))
    //deck.replaceAllText('{{PSMCom}}', cur_formatter.format((parseInt(row[8]))*.1));
    
    deck.replaceAllText('{{ClientResults}}', row[11])
    deck.replaceAllText('{{fNum1}}', row[12]);
    deck.replaceAllText('{{fNum2}}', row[13]);
    deck.replaceAllText('{{Promotion}}', row[14]);
    deck.replaceAllText('{{NewFollowers}}', row[15]);
    deck.replaceAllText('{{GrowthDuring}}', row[16]);
    deck.replaceAllText('{{ProfileVisits}}', row[17]);
    deck.replaceAllText('{{AccountsReached}}', row[18]);
    deck.replaceAllText('{{WeeklyImpressions}}', row[19]);
    deck.replaceAllText('{{CashTotal}}',cur_formatter.format(row[20]));
    deck.replaceAllText('{{CommTotal}}', cur_formatter.format((row[20]*.2))) //make sure this works
    deck.replaceAllText('{{Sponsor1}}', row[21]);
    deck.replaceAllText('{{Fee1}}', row[23]);
    deck.replaceAllText('{{PSMcom1}}', row[24]);
    deck.replaceAllText('{{Term1}}', row[25]);
    deck.replaceAllText('{{Deliverables1}}', row[26]);
    deck.replaceAllText('{{Sponsor2}}', row[27]);
    deck.replaceAllText('{{Fee2}}', row[29]);
    deck.replaceAllText('{{PSMcom2}}', row[30]);
    deck.replaceAllText('{{Term2}}', row[31]);
    deck.replaceAllText('{{Deliverables2}}', row[32]);
    deck.replaceAllText('{{Sponsor3}}', row[33]);
    deck.replaceAllText('{{Fee3}}', row[35]);
    deck.replaceAllText('{{PSMcom3}}', row[36]);
    deck.replaceAllText('{{Term3}}', row[37]);
    deck.replaceAllText('{{Deliverables3}}', row[38]);
    deck.replaceAllText('{{cmLink}}', row[39]);
    deck.replaceAllText('{{oLink}}', row[40]);
    //deck.replaceAllText

  

    

    //Cover Slide Image Replacement
    if(row[7].length < 1){
      //pass
    } else {
      const slide1 = deck.getSlideById('p');
      const pageElement = slide1.getPageElementById('gdb3f89f5fe_0_0');

      //gd872a1c164_0_9

      const image1 = pageElement.asImage();
      img_id1 = getIdFromUrl(row[7]);
      var img1 = DriveApp.getFileById(img_id1).getBlob();
      image1.replace(img1, true); //boolean dictates whether or not to crop 
        
    } 
    
    //Logo 1 image replacement

    if(row[22].length < 1){
      //pass
    } else {
      const slide4 = deck.getSlideById('gd872a1c164_0_41');
      const pageElement2 = slide4.getPageElementById('gd872a1c164_0_55');

      //gd872a1c164_0_9

      const logo1 = pageElement2.asImage();
      logo_id1 = getIdFromUrl(row[22]);
      var lg1 = DriveApp.getFileById(logo_id1).getBlob();
      logo1.replace(lg1, false); //boolean dictates whether or not to crop 
        
    } 


    //logo 2 replacement
    if(row[28].length < 1){
      //pass
    } else {
      const slide4 = deck.getSlideById('gd872a1c164_0_41');
      const pageElement3 = slide4.getPageElementById('gd872a1c164_0_58');

      //gd872a1c164_0_9

      const logo2 = pageElement3.asImage();
      logo_id2 = getIdFromUrl(row[28]);
      var lg2 = DriveApp.getFileById(logo_id2).getBlob();
      logo2.replace(lg2, false); //boolean dictates whether or not to crop 
        
    } 


    //logo 3 replacement
    if(row[34].length < 1){
      //pass
    } else {
      const slide4 = deck.getSlideById('gd872a1c164_0_41');
      const pageElement4 = slide4.getPageElementById('gd872a1c164_0_60');

      //gd872a1c164_0_9

      const logo3 = pageElement4.asImage();
      logo_id3 = getIdFromUrl(row[34]);
      var lg3 = DriveApp.getFileById(logo_id3).getBlob();
      logo3.replace(lg3, false); //boolean dictates whether or not to crop 
        
    } 






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
