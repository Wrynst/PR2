function Employee( fName, lName, startDate, address, email, phone, ssNum, w4Claim ) { 
  this.fName     = fName; 
  this.lName     = lName;
  this.startDate = startDate;  
  this.address   = address;
  this.email     = email;
  this.phone     = phone;
  this.ssNum     = ssNum;
  this.w4Claim   = w4Claim;
  
  var canDrive   = false;
  var position   = 'greenHorn';
  
  var recSheetId = ''
  
  var hoursWorked,
      sickHoursEarned,
      daysWorked;

  
  this.work = function( teamNum, day ) {
    daysWorkedCount += 1;
    daysWorked += day.toDateString();
    team = teamNum;
  }

} 

function Team( tm, members, teamNum, day, jobs ) {
  this.tm      = tm;
  this.members = members;
  this.teamNum = teamNum;
  this.day     = day;
  this.jobs    = jobs;
  
  
}

function Job( teamNum, numEmpl, clientName, price, day, duration ) {
  this.teamNum    = teamNum;
  this.clientName = clientName;
  this.price      = price;
  this.day        = day;
  this.duration   = duration;
  
  this.teamShare = function(price){
    return price*0.425;
  }
  
}

function Paystub(name,payday,days,teams,jobs){
  this.name = name;
  this.payday = payday;
  this.days = days;
  this.teams = teams;
  this.jobs = jobs;
  
  this.effBonus = function() {
    
  }

}

function getEmployees(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log(ss);
}
