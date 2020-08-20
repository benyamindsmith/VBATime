function ExtractEmail(cell){
  
//First lets define our variables

//  Email Regex
var regExp = new RegExp( "([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", "gi"); // "g" is for global,
                                                                                     //"i" is for case insensitive
// Extract Emails
var Email = regExp.exec(cell)[0];

return Email;
  
}
