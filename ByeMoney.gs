

function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu('💸')
  .addItem('Run Payroll', 'payroll')
  .addSeparator()
  .addItem('Delete All','delete3LetterTabs')
  .addItem('Mulligan', 'mulligan')
  .addSeparator()
  .addItem('Send Emails', 'sendOutEmails')
  .addToUi();

}

var activeSS = SpreadsheetApp.getActiveSpreadsheet();

var ssSheets = activeSS.getSheets();
var lengthSS  = ssSheets.length;  
var employees = getSheetObj('Employees');     // Employees is a local sheet on Nates Payroll imported range from Employees SpreadSheet
var jobs      = getSheetObj('Jobs');          // Jobs comes in every week - replace sheet when importing       
var mileage   = getSheetObj('Mileage');       // Mileage is from the forms Employees submit daily - Query of the results page is imported locally to show only the miles needed
var timecards = getSheetObj('Timecard');      // Timecards is the 2nd sheet along with jobs that has to be manually imported by hand from HCP
var rates     = getSheetObj('Rates');         // Rates is another local sheet on Nates Payroll that is imported from Rates SpreadSheet which updates when Min wage goes up every July 1st.. ugh
var template = activeSS.getSheetByName('Template').activate();                                  


function tabList(){
  var tabs = [];
  employees.forEach(function(employee){
    if (employee['fullName']){
      tabs.push(employee['fullName'].slice(0,3));
    } 
  });
  return tabs;
}








// copies the template to print the payroll to and renames it
function newTab(name) {
  var temp = template;
  var ss = activeSS;
  return temp.copyTo(ss).setName(name).setActiveSelection('D6');
  //return SpreadsheetApp.getActiveSpreadsheet().insertSheet(name).setActiveSelection('D6');
}

// Makes the dropdown list for Mulligan
var mulligan = function(){
  var html = HtmlService.createTemplateFromFile('mulligan');
  
  var tabhtml = '<select id="selection" name="Employees"><option value="Pickem" disabled selected>Pickem</option>'

  ssSheets.forEach(function(x){
    
    if(x.getName().length == 3){
      tabhtml += '<option value="' + x.getName() + '">' + x.getName() + '</option>'
    }
  });
 
  tabhtml += '</select>';
  html.tabs = tabhtml;
  
  // Make the SideBar
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setWidth(200));
}

// Returns an array of objects with 2 properties the name of the 3 letter sheet and the sheet itself
var threeLetterSheets = function(daSheets){
  var threeLettSheets = [];
  
  for(sheet in daSheets){
    var namy = daSheets[sheet].getName();                      // nonsensical variable names are on the house
    
    if(namy.length == 3){
      threeLettSheets.push({ name: namy, sheet: sheet});
    }
  }
  return threeLettSheets;
}



function deleteCreateSingleTab(name){
  Logger.log(name);
  var result = [];
  ssSheets.forEach(function(obj){
     // if the button has passed a 3 letter name
   
     if( obj.getName() === name ){
       activeSS.deleteSheet(obj);
       var newEmployees = employees.filter(function(employee){
         if(employee['fullName'].slice(0,3) === name){
           return employee;
         }
       })
       employees = newEmployees;
       payroll();
       result.push({ 'name' : name, 'arrayOfEmps' : newEmployees});
     }
  });
  Logger.log(result);   
  return name;
}


function delete3LetterTabs(){             // Perfect code - made out of laziness. Got tired of deleting all the sheets by hand while testing ;)
                
  ssSheets.forEach(function(obj){
    // if the tab is a 3 letter tab delete it
    if(obj.getName().length == 3 ){
      activeSS.deleteSheet(obj);
      // Didn't really realize I would have to get parent but it makes sense can't delete something if you are using it right?
     }
     });
  
}

function test(){
Logger.log(delete3LetterTabs()); 
//Logger.log(normalizeName( 'Tiffany', 'McCoy, Tiffany'));
}

function nameFL(fullName, displayName){                                //displayName is First name from Employee roster
                                                                       //fullName is last name comma first middle initial format Ashcroft, Dany G
  var comma = fullName.indexOf(',');                                   //HCP uses full names first name (space) last name format this creates that
  return displayName +' '+ fullName.slice(0, comma); 

}
function normalizeName( displayName, fullName ){
   
  var lowerCaseFirstLetter = displayName.slice ( 0, 1 ).toLowerCase();
  var camelCaseFirstLetterLastName = fullName.slice ( 0, 1 ).toUpperCase();
  var lowerCaseLastName = fullName.slice ( 0, fullName.indexOf ( ',' ) ).toLowerCase();
   
  return lowerCaseFirstLetter                                                 //quick normalizedName to use in var duration below -- timecards[card][normalizedName] <-var for dynamic obj property retrieval (couldn't
         + displayName.slice ( 1 )
         + camelCaseFirstLetterLastName      //think of another way to pull normalizedHeaders without tapping SpreadsheetApp service again.. needlessly 
         + lowerCaseLastName.slice ( 1 );                      
                                                                                                                            
}
// This is the monolith function to kill all functions. Its where the magic happens (in some 100 lines!)
// I have ideas on how to break it up but I need to get this working first.
function payroll() {                            
       
  var payDay    = new Date();                   // This will be a pretty little date in the corner, it doesn't have a place yet. 
  
//-----1-----FIRST LOOP EMPLOYEES----------------------------[1]
  for(emp in employees){ 
    
    var col = 0;       //counter
    var effWage = 0;   //counter
    var payStub;
    var employee = employees[emp];              //Gets used too much not to shorten it
    var fullName = employee['fullName'];
    var displayName = employee['displayName'];
    
    var nameFirstLast = nameFL ( fullName, displayName );               //function formats name to Greg Fashbaugh 
    var normalizedName = normalizeName ( displayName, fullName );       //format name to gregFashbaugh
    
    if(employee['currentlyEmployed']){                             // if a fullName exists,
      payStub = newTab(fullName.slice(0, 3)); // make a new tab/sheet with first 3 letters of last name as sheet name
    }else{                                    // otherwise break out of this loop
      break;
    }
    
    
    payStub.setValue(nameFirstLast);                                    //sets the name on the new tab (paystub)
    
    
   // payStub.offset(-1, col).setValue(new Date(setDate((timecards[0]['date']).getDate() + 11).toDateString()));       



    
//-----2-----SECOND LOOP (NESTED) TIMECARDS-------------------[2]
    for(card in timecards){
      var totalAmt = 0;                                       //running totals so we can have a total before we move on to the next day
      var totalSales = 0; 
         
      var date1 = timecards[card]['date'];
      var duration = timecards[card][normalizedName];
      
      if(typeof duration === 'object') {                      //getMinutes() and gethours() do not like empty objects.. or is it strings?  the jury is still out!

       var minutes = duration.getMinutes();                   //If it is not an obj run the datetime methods to pull min and hours from datetime obj
       var hours = duration.getHours();         
      
      }else{
        
        var minutes = '0';                                    //if it is an ---empty obj---(typeof string) assign string zeros to each one
        var hours = '0';
      };
      
      var min = minutes/60;                                   //divide by 60 to get decimal
      var hrs = min + hours;
      
      var decimal = parseFloat(hrs).toFixed(2);
                              
      var row = 6;
      
      payStub.offset(5, col).setValue(date1.toDateString());  //write the date to the sheet
      payStub.offset(5, col - 2).setValue('Job | Team');      //write the header to the sheet    
      payStub.offset(5, col - 1).setValue('Team Share');           //write the header to the sheet


      
//-----3-----THIRD LOOP (NESTED) JOBS ----------------------[3]
      for(job in jobs){
        
        var jobDate   = jobs[job]['date'];
        var jobEmp    = jobs[job]['employee'];
        var jobStatus = jobs[job]['jobStatus'];
        
        if(jobs[job]['date'] && jobDate.toDateString() == date1.toDateString() && jobEmp.indexOf(nameFirstLast) !== -1 && jobStatus){    // and keep ones that match employee AND day of the week
        
          col = col - 2;
          
          var empsOnJob = jobEmp.split(', ');                //the list of employees on each job is a string of names seperated by commas. split puts each name into its own spot in an array.
          var num = empsOnJob.length;
          var numOnTeam = num.toString(); 

          var shortTeam = [];                                //names are too long so incoming short jokes
          var dailyEff = 0;                                                             
          for(var i = 0; numOnTeam > i; i++){                     
            var shortName = empsOnJob[i].slice(0, 3);        //no shortName! Nooo!  This gets first 3 letters from name and makes new array. 
            shortTeam.push(' ' + shortName);                 //shortTeam PUSH!  lol
          };
          
          payStub.offset(row, col++)
            .setValue(jobs[job]['customer'].slice(0,3)       //write 3 first letters of last name of client on paystub so employee can recognize how much they made on a particular job
                      + ' | '                                //write the teammates first name so again can recognize who they work well with
                      + shortTeam.toString()
                     );
                                                                                                                                               
          var timeCleaning = parseFloat(jobs[job]['totalDuration']/3600);    //Time is given to me in seconds from HCP (housecleaning app) so to turn into hours you devide seconds by 3600.
          var sales = parseFloat(0.425 * jobs[job]['subtotal']);                     //Objects everywhere in this thing but I can't seem to OOP what the hell.
         
          if(numOnTeam == '2'){
            numOnTeam = 'two';
          }else if (numOnTeam == '3'){
            numOnTeam = 'three'; 
          }else if (numOnTeam == '4'){
            numOnTeam = 'four';
          }else if (numOnTeam == '5'){
            numOnTeam = 'five';
          }else if (numOnTeam == '6'){
            numOnTeam = 'six';
          }
            
          for(rate in rates){
            if(rates[rate]['position'] == employees[emp]['jobTitle']) {
              
              dailyEff = rates[rate][numOnTeam];
              sales = dailyEff * sales;
              
            }; 
          }
          
          payStub.offset(row, col++).setValue('$' + sales.toFixed(2));       //parseFloat to force numbers then toFixed to decimal places to 2 spots.  over and over want to know a better way.
          payStub.offset(row, col).setValue(timeCleaning.toFixed(2));
          //Logger.log(dailyEff);
          totalAmt   += timeCleaning;                                          //because down here if they revert back to strings you get a dumb long string of numbers like '1230.23.234.223.3.234.' when all I want is 3.52
          totalSales += sales;                                              //probably just me picking the wrong way to add up a total.   
          row += 1;          
        };
      };
      var minWage = parseFloat(decimal * 10.75).toFixed(2);
      
      payStub.offset(17,col - 1).setValue('$' + parseFloat(totalSales).toFixed(2));   //display total TS
      payStub.offset(14,col - 1).setValue('$' + parseFloat(totalSales/totalAmt).toFixed(2));
      payStub.offset(15,col - 2).setValue('Hourly Wage + Office Time');
      payStub.offset(14,col - 2).setValue('Hourly Wage Cleaning');
      payStub.offset(15,col - 1).setValue('$' + parseFloat(totalSales/decimal).toFixed(2));
      payStub.offset(14,col).setValue(parseFloat(totalAmt).toFixed(2));                // Hours spent on the Job
      payStub.offset(17,col).setValue('$' + parseFloat((decimal * 10.75)).toFixed(2)); 
      payStub.offset(18,col).setValue('Min Wage');
      payStub.offset(18,col - 1).setValue('Team Share');
     
      payStub.offset(15,col).setValue(decimal);// Hours clocked in   ----   Big disparagement
      var firstEntry = '' ;
      var secondEntry = '';
      effWage = totalAmt + effWage;
//      for(miles in mileage){
//        var mileageDate = mileage[miles]['timestamp'].toDateString();
//        var date1str = date1; //toDateString();
//        Logger.log((mileage[miles]['timestamp']));
//        if (mileageDate == date1str && employees[emp]['email'].toLowerCase().trim() == mileage[miles]['emailAddress'].toLowerCase().trim()){ //if date on mileage form matches && email matches then write
//         
//          //(!firstEntry) ? firstEntry = mileage[miles] : secondEntry = mileage[miles];
//          payStub.offset(15,col).setValue(mileage[miles]['mileage'] + ' Miles on ' + mileageDate);
//      };
//      };
      col += 3;
    };
    //payStub.offset(17,col).setValue((effWage).toFixed(2) + ' eff Total');
  Logger.log(effWage);
  };  
}

 
/*
 * @param(string) sheetName is the name of the sheet as a string
 */
function getSheetObj(sheetName) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  return objectSheet(sheet);
  
}

function objectSheet(data) {
  
  var headers = data.shift();
  var fixedHeaders = normalizeHeaders(headers);
  
  // convert 2d array into object array
  // https://stackoverflow.com/a/22917499/1027723
  var obj = data.map(function(values) {
    return fixedHeaders.reduce(function(o, k, i) {
      o[k] = values[i];
      return o;
    }, {});
  });
  return obj;
}

//formQuestions(form,obj);
// ScriptApp.newTrigger().forSpreadsheet(ss).onFormSubmit().create();

// Returns an Array of normalized Strings.
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
//
//
//[empData]
//[invoice, hcpId, createdAt, date, endTime, travelDuration, onJobDuration, totalDuration, customer, firstName, lastName, 
// email, company, mobilePhone, homePhone, customerTags, address, street, streetLine2, city, state, zip, description, lineItems, 
// amount, labor, materials, subtotal, paymentHistory, creditCardFee, paidAmount, due, discount, tax, taxableAmount, taxRate, 
// jobTags, notes, employee, jobStatus, finished, payment, invoiceSent, window, attachments, segments, hcJob, tipAmount, onlineBookingSource]

function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(normalizeHeader(headers[i]));
    //Logger.log("string: "+headers[i]);
  }
  return keys;
}
// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  
  //Logger.log("header: "+key);
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}