var url = 'https://www.google.com';

function testStatusLink() {
  
  // test a link with GET
  
  var response = UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true, 'followRedirects': false });
  var code = response.getResponseCode();
  
  // decode all the keys used in a headers response
  
  Logger.log("headers : " + JSON.stringify(response.getAllHeaders()));
  
  switch (code){
    case 200 : { Logger.log("OK      cookies: "+ response.getAllHeaders()['Set-Cookie'])  ; break;}
    case 301 :
    case 302 :
    case 307 :
    case 308 : { Logger.log("Code: "+ code +"  Redirection to: " + response.getAllHeaders()['Location']) ; break;}
    case 400 : { Logger.log("Code: "+ code +"  Error in Request") ; break;}
    case 401 : { Logger.log("Code: "+ code +"  Need Authorization with: " + response.getAllHeaders()['WWW-Authenticate']) ; break;}
    case 404 : { Logger.log("Code: "+ code +"  Dead link " ) ; break;}
    default  : { Logger.log("Error Number "+ code ) ; break;}
  }
  
}
