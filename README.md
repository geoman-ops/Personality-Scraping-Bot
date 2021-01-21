/*
# Personality-Scraping-Bot
Scrape LinkedIn Urls from Google using SerpWoW scraper and get Personality insights using Humantic AI. 

In this project I am using a Speadsheet, and Google's appscript in order to to retrieve LinkedIn Profile URLs and send them to Humantic AI using an API call. Humantic AI returns some personality insights for a specific LinikedIn profile which then are saved to my inital spreadsheet. Then we can store and use those insights in a CRM.
*/






/*-------------A RANDOM TEST FUNCTION-----------*/


function testGetBingResults(row=2) {
  
//  const jsonResults = BingSearchSerpWowScraping.resultsGoogle();
  
//  var url_string = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "F"+row).values[0][0];
  var company_name_1 = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].toUpperCase().split(' ').slice(0,1);
  var company_name_2 = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].toUpperCase().split(' ').slice(1,2);
//  const firstName = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].toUpperCase();
//  const lastName = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].toUpperCase();
  
  
  Logger.log(company_name_1);
   Logger.log(company_name_2);
  
//  var response = UrlFetchApp.fetch(url_string);
//  var content = response.getContentText();
//  var jsonResults = JSON.parse(resultsFromGoogle);
//  
//  let topFiveResults = [];
//  
//  let linkedinUrl = "";
//  let resultTitle = "";
//  let matchedKeywords = [];
//  
//  
//  for (let i =0; (i<jsonResults["organic_results"].length) && (i<=4); i++) {
//    
//    linkedinUrl = jsonResults["organic_results"][i]["link"];
//    resultTitle = jsonResults["organic_results"][i]["title"];
//    matchedKeywords = jsonResults["organic_results"][i]["snippet_matched"];
//    topFiveResults.push({
//      link: linkedinUrl,
//      title: resultTitle,
//      keywordsMatch: matchedKeywords
//    });
//  }
//  
//  
//  
//  filterResults(topFiveResults,company_name,firstName,lastName);
//  
//  return true;
  
}



// First 100 Results

/* ----------------CONSTRUCT SEARCH URL------------------ */



function create_Search_URL(row) {
  
  // Get Company Name Cell - Check if it is empty
  var company_name_check= SpreadsheetApp.getActiveSheet().getRange("E"+row).isBlank();
  
  
  // Get First Name Cell - Check if it is empty
  var first_name_check= SpreadsheetApp.getActiveSheet().getRange("C"+row).isBlank();
  
  // Get Last Name Cell - Check if it is empty
  var last_name_check= SpreadsheetApp.getActiveSheet().getRange("D"+row).isBlank();
  
  // Get Email Address Cell - Check if it is empty
  var email_address_check= SpreadsheetApp.getActiveSheet().getRange("B"+row).isBlank();
  
  // Get Initial URL Cell - Check if it is empty
  var search_url_check= SpreadsheetApp.getActiveSheet().getRange("F"+row).isBlank();
  
//  var company_name_dirty = "";
  var company_name = "";
  var first_name_dirty = "";
  var first_name = "";
  var last_name_dirty = "";
  var last_name = "";
  var search_url = "";
  var email = "";
  
  if (search_url_check === true) {
    
    if (email_address_check === true) {
      
      var search_url = "No Result";
      
      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
      
      return search_url;     // Returns True when one of the 3 main cells (First Name, Last Name, Email) is empty and false when ALL of them are not empty.
      
    } else if (first_name_check && last_name_check) {
      var search_url = "No Result";
      
      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
      
      return search_url;     // Returns True when one of the 3 main cells (First Name, Last Name, Email) is empty and false when ALL of them are not empty.
      
    } else if (last_name_check === true && company_name_check === true) {
      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];
      
      var search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+email+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
      return search_url;
      
    } else if (last_name_check === true) {
      
      company_name = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].replace(/[^a-z\d\s]+/gi, "").split(' ').slice(0,2).join('+');
//      company_name = company_name_dirty.replace(/[^a-zA-Z ]/g, " ");
      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];
      
      search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+company_name+"+"+email+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
      
      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
      
      return search_url;
      
      } else if (company_name_check === true) {
        
      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];      
      last_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].replace(/[^a-zA-Z ]/g, "");
      last_name = last_name_dirty.replace(/ /g,"+");
      
      search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+email+"+"+last_name+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
      
      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
         
         return search_url;
    } else {
      
      company_name = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].replace(/[^a-z\d\s]+/gi, "").split(' ').slice(0,2).join('+');
//      company_name = company_name_dirty.replace(/[^a-z\d\s]+/gi, "");
      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];
      last_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].replace(/[^a-zA-Z ]/g, "");
      last_name = last_name_dirty.replace(/ /g,"+");
      
      search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+company_name+"+"+email+"+"+last_name+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
      
      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
      
      return search_url;
    }
  } else {
    
    Logger.log("There is already an initial URL at ROW "+row+". The create_Search_URL() function did not start at all");
    search_url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("F"+row).getValues();
    return search_url;
  }
  
  return true;
}







///* ----------------CONSTRUCT SEARCH URL------------------ */
//
//// INCLUDE FIRST NAME IN THE SEARCH
//
//
//function create_Search_URL(row) {
//  
//  // Get Company Name Cell - Check if it is empty
//  var company_name_check= SpreadsheetApp.getActiveSheet().getRange("E"+row).isBlank();
//  
//  
//  // Get First Name Cell - Check if it is empty
//  var first_name_check= SpreadsheetApp.getActiveSheet().getRange("C"+row).isBlank();
//  
//  // Get Last Name Cell - Check if it is empty
//  var last_name_check= SpreadsheetApp.getActiveSheet().getRange("D"+row).isBlank();
//  
//  // Get Email Address Cell - Check if it is empty
//  var email_address_check= SpreadsheetApp.getActiveSheet().getRange("B"+row).isBlank();
//  
//  // Get Initial URL Cell - Check if it is empty
//  var search_url_check= SpreadsheetApp.getActiveSheet().getRange("F"+row).isBlank();
//  
//  var company_name_dirty = "";
//  var company_name = "";
//  var first_name_dirty = "";
//  var first_name = "";
//  var last_name_dirty = "";
//  var last_name = "";
//  var search_url = "";
//  var email = "";
//  
//  if (search_url_check === true) {
//    
//    if (email_address_check === true) {
//      
//      var search_url = "No Result";
//      
//      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
//      
//      return search_url;     // Returns True when one of the 3 main cells (First Name, Last Name, Email) is empty and false when ALL of them are not empty.
//      
//    } else if (first_name_check && last_name_check) {
//      var search_url = "No Result";
//      
//      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
//      
//      return search_url;     // Returns True when one of the 3 main cells (First Name, Last Name, Email) is empty and false when ALL of them are not empty.
//      
//    } else if (last_name_check === true && company_name_check === true) {
//      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];
//      
//      first_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].replace(/ /g,"+");
//      first_name = first_name_dirty.replace(/[^a-zA-Z ]/g, "");
//      
//      var search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+first_name+"+"+email+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
//      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
//      return search_url;
//      
//    } else if (last_name_check === true) {
//      
//      company_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].split(' ').slice(0,2).join('+');
//      company_name = company_name_dirty.replace(/[^a-zA-Z ]/g, " ");
//      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];
//      
//      first_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].replace(/ /g,"+");
//      first_name = first_name_dirty.replace(/[^a-zA-Z ]/g, "");
//      
//      search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+first_name+"+"+company_name+"+"+email+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
//      
//      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
//      
//      return search_url;
//      
//    } else if (company_name_check === true) {
//      
//      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];      
//      last_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].replace(/ /g,"+");
//      last_name = last_name_dirty.replace(/[^a-zA-Z ]/g, "");
//      
//      first_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].replace(/ /g,"+");
//      first_name = first_name_dirty.replace(/[^a-zA-Z ]/g, "");
//      
//      search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+email+"+"+first_name+"+"+last_name+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
//      
//      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
//      
//      return search_url;
//    } else {
//      
//      company_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].split(' ').slice(0,2).join('+');
//      company_name = company_name_dirty.replace(/[^a-zA-Z ]/g, " ");
//      email = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "B"+row).values[0][0];
//      last_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].replace(/ /g,"+");
//      last_name = last_name_dirty.replace(/[^a-zA-Z ]/g, "");
//      
//      first_name_dirty = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].replace(/ /g,"+");
//      first_name = first_name_dirty.replace(/[^a-zA-Z ]/g, "");
//      
//      search_url = "https://api.serpwow.com/live/search?api_key=FE155CEC036549CFB1BA722FA69CC303&q="+first_name+"+"+last_name+"+"+company_name+"+linkedin&google_domain=google.nl&location=Netherlands&include_answer_box=false&url=&gl=nl&hl=en";
//      
//      SpreadsheetApp.getActiveSheet().getRange("F"+row).setValue(search_url);
//      
//      return search_url;
//    }
//  } else {
//    
//    Logger.log("There is already an initial URL at ROW "+row+". The create_Search_URL() function did not start at all");
//    search_url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("F"+row).getValues();
//    return search_url;
//  }
//  
//  return true;
//}








// Set Number of row - Create Initial URL
function Process_All_Rows_Create_Search_URL() {
  
  var i = 0;
  
  let lastIndexedRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("A2").getValues(); // Get the value of last row that has been indexed (spreadsheet name: variables, column/cell A2)
  let lastSpreadsheetRow = SpreadsheetApp.openById("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE").getSheetByName("Sheet1").getLastRow(); //Get the last used row of the spreadsheet
  
  for (row = lastIndexedRow; row <= lastSpreadsheetRow; row++) {
    Logger.log(row);
    if (i <=9) { //Change i to change the maximum succesfull rows you want to get per run. 
      
      if (create_Search_URL(row)) {
        
        i++;
        Logger.log(i);
      }       
      
    } else {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("A2").setValue(row);
      
      break;
      
    }
    
  }
  
//  if (row > lastSpreadsheetRow) {
//    Logger.log("We reached the last Row of Create Initial URL. I will stop now!");
//    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("A2").setValue(row);
//  }
  
}






/*----------------- END OF CONSTRUCT SEARCH URL-------------------*/





/*------------------FILTER RESULTS FUNCTION-------------*/


/* Filter Function*/

function filterResults(arrResults,company_1,company_2,firstName,lastName) {
  const requestedTerm = "linkedin.com/in/";
  let finalUrl = "Initial Value";
  
  const filterBasedOnUrl = arrResults.filter(function(item) {
    return item.link.includes(requestedTerm)})
  
  Logger.log("The Initial Filtered Results are: "+filterBasedOnUrl);
  
  //CAPITALIZE TITLE ON THE FILTERED RESULTS(ARRAY)
  const upperCaseTitle = filterBasedOnUrl.map(function(item){ return item.title.toUpperCase()});
  
  Logger.log("The Filtered Uppercase Titles are: "+upperCaseTitle);
  
  //CAPITALIZE SNIPPET ON THE FILTERED RESULTS(ARRAY)
  
  const upperCaseSnippet = filterBasedOnUrl.map(function(item){
    if (item.snippet != null){
    return item.snippet.toUpperCase();
      Logger.log("The NON empty snippet returned: "+item.snippet);
    } else {
      item.snippet = "null";
      return item.snippet;
      Logger.log("The empty snippet returned: "+item.snippet);
    }  
  });
  
  Logger.log("The Filtered Uppercase Snippets are: "+upperCaseSnippet);
  
  const arrKeywords = [];
  
  
  
  for (let i = 0; i<filterBasedOnUrl.length; i++) {
    arrKeywords[i] = filterBasedOnUrl[i].keywordsMatch;
  }
  
  for (let i = 0; i < filterBasedOnUrl.length; i++) {
    filterBasedOnUrl[i].title = upperCaseTitle[i];
    filterBasedOnUrl[i].snippet = upperCaseSnippet[i];
    
  }
  
  
  Logger.log("The Final Filtered Results are: "+filterBasedOnUrl);
  
  
  //Filter Based on First Name & Last Name & Company
  const superMatch = filterBasedOnUrl.filter(function(item) {
    let fName = item.title.includes(firstName);
    let lName = item.title.includes(lastName);
    let cName = item.title.includes(company_1);
    let cName2 = item.title.includes(company_2);
    let cNameSnippet1 = item.snippet.includes(company_1);
    let cNameSnippet2 = item.snippet.includes(company_2);
    
    
    
    if ( fName && lName && cName && cName2) {
      return item;
    } else if(fName && lName) {
      if(cName || cName2) {
        return item;
      } else if (cNameSnippet1 || cNameSnippet2){
        return item;
      }
    }
    else {
      Logger.log("No Match for fName & lName & cName & cName2. There is not a SUPER match");
    }
  })
  
  Logger.log("The SUPER MATCH ARRAY is:");
  Logger.log(superMatch);
  Logger.log("******************************");
  
  
  // Store the final Result in this Object
  const urlNscore = {link: "No Link",
                     score: "No Score"};
  
  
  let superMatchScore;
  
  if(superMatch.length) {
    if(superMatch.length === 3) {
      superMatchScore = 40;
    } else if (superMatch.length === 2) {
      superMatchScore = 85;
    } else if (superMatch.length === 1){
      superMatchScore = 100;
      Logger.log("There is a SuperMatch "+superMatchScore+"%. The Url we need is "+superMatch[0].link)
      urlNscore.link = superMatch[0].link;
      urlNscore.score = superMatchScore;
      return urlNscore;
    } else {
      superMatchScore = 26;
    }
  } else {
    superMatchScore = 0;
  }
  
  Logger.log("Supermatch Score: "+superMatchScore);
  
  
  
  /* END OF SUPERMATCH*/
  
  
  
  
  //Filter Based on First Name OR Last Name & Company
  const strongMatch = filterBasedOnUrl.filter(function(item) {
    let fName = item.title.includes(firstName);
    let lName = item.title.includes(lastName);
    let cName = item.title.includes(company_1);
    let cName2 = item.title.includes(company_2);
    let cNameSnippet1 = item.snippet.includes(company_1);
    let cNameSnippet2 = item.snippet.includes(company_2);
    
    
    
    if ( fName || lName) {
      if (cName || cName2){
        return item;
      } else if (cNameSnippet1 || cNameSnippet2) {
        return item;
      } else { Logger.log("Fname or Lname found on title cName1 or cName2 were not on Title neither on Snippet! There is not a STRONG Match");
             }} else {
               Logger.log("No Match for fName || lName. There is not a STRONG Match!");
             }
  })
  
  let strongMatchScore;
  
  if(strongMatch.length) {
    if(strongMatch.length === 3) {
      strongMatchScore = 30;
    } else if (strongMatch.length === 2) {
      strongMatchScore = 60;
    } else if (strongMatch.length === 1){
      strongMatchScore = 90;
      Logger.log("There is a StrongMatch "+strongMatchScore+"%. The Url we need is "+strongMatch[0].link)
      urlNscore.link = strongMatch[0].link;
      urlNscore.score = strongMatchScore;
      return urlNscore;
    } else {
      strongMatchScore = 15;
    }
  } else {
    strongMatchScore = 0;
  }
  
  Logger.log("Strongmatch Score: "+strongMatchScore);
  
  /*END OF STRONG MATCH*/
  
  
  if ( superMatchScore !== strongMatchScore) {
    if(superMatchScore > strongMatchScore) {
      Logger.log("The winner is the SUPER match with a score of "+superMatchScore+". The highest score of the STRONG match was: "+strongMatchScore+". The final URL is: "+superMatch[0].link);
      urlNscore.link = superMatch[0].link;
      urlNscore.score = superMatchScore;
      return urlNscore;
    }
    if(superMatchScore < strongMatchScore) {
      Logger.log("The winner is the STRONG match with a score of "+strongMatchScore+". The highest score of the SUPER match was: "+superMatchScore+". The final URL is: "+strongMatch[0].link);
      urlNscore.link = strongMatch[0].link;
      urlNscore.score = strongMatchScore;
      return urlNscore;
    }
  } else {
    
    
    //Filter Based on First Name AND Last Name (NO COMPANY NAME INVOLVED)
    const weakMatch = filterBasedOnUrl.filter(function(item) {
      let fName = item.title.includes(firstName);
      let lName = item.title.includes(lastName);
      
      
      if ( fName && lName) {
        return item;
      } else {
        Logger.log("No Match for fName && lName. There is not a WEAK Match!");
      }
    })
    
    
    let weakMatchScore;
    
    if(weakMatch.length) {
      if(weakMatch.length === 3) {
        weakMatchScore = 15;
        Logger.log("There is a WeakMatch "+weakMatchScore+"%. The Url we need is "+weakMatch[0].link)
        urlNscore.link = weakMatch[0].link;
        urlNscore.score = weakMatchScore;
        return urlNscore;
      } else if (weakMatch.length === 2) {
        weakMatchScore = 20;
        Logger.log("There is a WeakMatch "+weakMatchScore+"%. The Url we need is "+weakMatch[0].link)
        urlNscore.link = weakMatch[0].link;
        urlNscore.score = weakMatchScore;
        return urlNscore;
      } else if (weakMatch.length === 1){
        weakMatchScore = 25;
        Logger.log("There is a WeakMatch "+weakMatchScore+"%. The Url we need is "+weakMatch[0].link)
        urlNscore.link = weakMatch[0].link;
        urlNscore.score = weakMatchScore;
        return urlNscore;
      } else {
        weakMatchScore = 11;
        Logger.log("There is a WeakMatch "+weakMatchScore+"%. The Url we need is "+weakMatch[0].link)
        urlNscore.link = weakMatch[0].link;
        urlNscore.score = weakMatchScore;
        return urlNscore;
      }
    } else {
      weakMatchScore = 0;
    }
    
    Logger.log("The Weak Match Score is :"+weakMatchScore);
    
    
    
    //Filter Based on First Name OR Last Name (NO COMPANY NAME INVOLVED)
    const veryWeakMatch = filterBasedOnUrl.filter(function(item) {
      let fName = item.title.includes(firstName);
      let lName = item.title.includes(lastName);
      
      
      if ( fName || lName) {
        return item;
      } else {
        Logger.log("No Match for fName || lName. There is not a VERY WEAK Match!");
      }
    })
    
    Logger.log("The veryWeak Match is :"+veryWeakMatch);
    
    
    let veryWeakMatchScore;
    
    if(veryWeakMatch.length) {
      if(veryWeakMatch.length === 3) {
        veryWeakMatchScore = 1;
        Logger.log("There is a VeryWeakMatch "+veryWeakMatchScore+"%. The Url we need is "+veryWeakMatch[0].link)
        urlNscore.link = veryWeakMatch[0].link;
        urlNscore.score = veryWeakMatchScore;
        return urlNscore;
      } else if (veryWeakMatch.length === 2) {
        veryWeakMatchScore = 5;
        Logger.log("There is a VeryWeakMatch "+veryWeakMatchScore+"%. The Url we need is "+veryWeakMatch[0].link)
        urlNscore.link = veryWeakMatch[0].link;
        urlNscore.score = veryWeakMatchScore;
        return urlNscore;
      } else if (veryWeakMatch.length === 1){
        veryWeakMatchScore = 10;
        Logger.log("There is a VeryWeakMatch "+veryWeakMatchScore+"%. The Url we need is "+veryWeakMatch[0].link)
        urlNscore.link = veryWeakMatch[0].link;
        urlNscore.score = veryWeakMatchScore;
        return urlNscore;
      } else {
        veryWeakMatchScore = 0.5;
        Logger.log("There is a VeryWeakMatch "+veryWeakMatchScore+"%. The Url we need is "+veryWeakMatch[0].link)
        urlNscore.link = veryWeakMatch[0].link;
        urlNscore.score = veryWeakMatchScore;
        return urlNscore;
      }
    } else {
      veryWeakMatchScore = 0;
      Logger.log("There was not a Match at all!!! We are gonna return No Result as the final URL");
      urlNscore.link = "No Result";
      urlNscore.score = "No Score";
      return urlNscore;      
    }
    
    Logger.log("The Very Weak Match Score is :"+veryWeakMatchScore);
    
  }
  
  //  Logger.log(filterBasedOnUrl);
  
}




/*-----------------------END OF FILTER FUNCTION-------------------*/






/*------------------ RUN SEARCH USING THE SEARCH URL AND GET THE RAW RESULTS--------------*/





/* Get LinkedIn Url From Search Results*/

function getLinkedinUrl(row) {
  
  //  const jsonResults = BingSearchSerpWowScraping.resultsGoogle();
  //  const jsonResults = BingSearchSerpWowScraping.resultsMarkusBucher();
  //  const jsonResults = BingSearchSerpWowScraping.resultsLorenzo();
  //  const jsonResults = BingSearchSerpWowScraping.resultsChristian();
  //  const jsonResults = BingSearchSerpWowScraping.resultsBelasik();
  //  const jsonResults = BingSearchSerpWowScraping.resultsAlbada();
  
  let firstName;
  let company_name_1;
  let company_name_2;
  let lastName;
  
  
  let url_string = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "F"+row).values[0][0];
  
  if (url_string !== "No Result") {
    
    
    // Get Company Name 1st string - Check if it is empty
    
    let company = SpreadsheetApp.getActiveSheet().getRange("E"+row).isBlank();
    
    if (!company) {
      company_name_1 = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].toUpperCase().split(' ').slice(0,1);
      company_name_2 = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].toUpperCase().split(' ').slice(1,2);
    } else {      
      company_name_1 = "NoEmpty";
      company_name_2 = "NoEmpty";
    }
    
    
    // Get First Name - Check if it is empty
    
    var first_Name = SpreadsheetApp.getActiveSheet().getRange("C"+row).isBlank();
    
    if (!first_Name) {
      firstName = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].toUpperCase();
    } else {
      firstName = "NoEmpty";
    }
    
    Logger.log(firstName);
    
    
    // Get Last Name - Check if it is empty
    
    var last_Name = SpreadsheetApp.getActiveSheet().getRange("D"+row).isBlank();
    
    if (!last_Name) {
      lastName = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].toUpperCase();
    } else {
      lastName = "NoEmpty";
    }
    
    Logger.log(lastName);
    
    
    //    var company_name_1 = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].toUpperCase().split(' ').slice(0,1);
    //    var company_name_2 = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "E"+row).values[0][0].toUpperCase().split(' ').slice(1,2);
    //    const firstName = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "C"+row).values[0][0].toUpperCase();
    //    const lastName = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "D"+row).values[0][0].toUpperCase();
    
    
    
    let finalUrlNscore;
    
    var response = UrlFetchApp.fetch(url_string);
    var content = response.getContentText();
    var jsonResults = JSON.parse(content);
    
    
    let topFiveResults = [];
    
    let linkedinUrl = "";
    let resultTitle = "";
    let matchedKeywords = [];
    let snippet;
    
    Logger.log("The Length is : "+jsonResults["organic_results"].length);
    
    
    for (let i =0; (i<jsonResults["organic_results"].length) && (i<=4); i++) {
      
      Logger.log(i);
      
      linkedinUrl = jsonResults["organic_results"][i]["link"];
      resultTitle = jsonResults["organic_results"][i]["title"];
      matchedKeywords = jsonResults["organic_results"][i]["snippet_matched"];
      snippet = jsonResults["organic_results"][i]["snippet"];
      topFiveResults.push({
        link: linkedinUrl,
        title: resultTitle,
        keywordsMatch: matchedKeywords,
        snippet: snippet
      });
    }
    
 Logger.log(topFiveResults);
    
    
    
    finalUrlNscore = filterResults(topFiveResults,company_name_1,company_name_2,firstName,lastName);
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("G"+row).setValue(finalUrlNscore.link);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("H"+row).setValue(finalUrlNscore.score);
    
    return true;
  } else {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("G"+row).setValue("No Result");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("H"+row).setValue("No Score");
    return true;
  }
  
}


/* Process all rows and Get linkedIn Url*/

function Process_All_Rows_Get_LinkedIn_Url () { 
  var i = 0;
  
  let lastIndexedRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("B2").getValues(); // Get the value of last row that has been indexed (spreadsheet name: variables, column/cell C2)
  let lastSpreadsheetRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("A2").getValues() - 1; //Get the last row that has been scraped and converted into Public URL
  
  for (row = lastIndexedRow; row <= lastSpreadsheetRow; row++) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("B2").setValue(row);
    if (i <= 3) {  //Change this number to change the max number of created profiles
      
      if (getLinkedinUrl(row)) {
        
        i++;
        Logger.log(i);
      }    
      
    } else {
      
      break;
    }
    
  }
}



/* CHECK IF OBJECT IS EMPTY*/
  
  function objIsEmpty(obj) {
    if (Object.keys(obj).length === 0) {
      Logger.log("Object is empty");
      return true; 
    } else {      
      return false;
    } 
  }
  
  
    /* CHECK IF ARRAY IS EMPTY */
  function arrIsEmpty(arr) {
    if (arr.length === 0) {
      Logger.log("Array is empty");
      return true; 
    } else {     
      return false;
    } 
  }
  

  



/*----------------END GET RUN SEARCH AND GET THE RAW RESULTS------------*/





/*----------------- GET LINKEDIN URL-------------------*/

function Get_LinkedIn_URL(row) {
  
  //Check if the cell H+row is blank
  
  var finalLinkedinUrlCheck = SpreadsheetApp.getActiveSheet().getRange("G"+row).isBlank();
  
  if ( finalLinkedinUrlCheck != true ) {
    
    var cell_value = Sheets.Spreadsheets.Values.get("1KZjcesDjeT2c6-46qQYv_IkaSU3242uAkUWmeYb1cJc", "G"+row).values[0];
    
//    Logger.log(cell_value);
    return cell_value;
    
  } else {
    
    var cell_value = 0;
    
    Logger.log("The cell_value var gave back the number "+cell_value+" for ROW "+row+ ". So there is not a Final URL yet!");
    
    return cell_value;
  }
  
}

/*------------------END GET THE LINKEDIN URL------------------*/







/* ---------------- CREATE HUMANTIC PROFILE -----------------*/



//Create Humantic AI Profile using the LinkedIn URL
function create_Humantic_Profile(row) {
  
  var url_string = Get_LinkedIn_URL(row);
  
  if (url_string != 0) {
    
    //Get Humantic Profile Field Value
    var humantic_profile = Sheets.Spreadsheets.Values.get("1KZjcesDjeT2c6-46qQYv_IkaSU3242uAkUWmeYb1cJc", "I"+row).values[0];
    
    
    if (url_string != "No Result" && humantic_profile != "Yes") {
      
      //Chenage Humantic Profile Field Value To Yes
      SpreadsheetApp.getActiveSheet().getRange("I"+row).setValue("Yes");
      
      //Create Humantic Profile
      var create_humantic_profile = "https://api.humantic.ai/v1/user-profile/create?apikey=chrexec_edf11eb1e9045701bd531c600197ae5d&userid="+url_string;
      
      var Fetch_Profile = UrlFetchApp.fetch(create_humantic_profile);
      
      return true;
      
    } else  {
      //Chenage Humantic Profile Field Value To Yes
      SpreadsheetApp.getActiveSheet().getRange("I"+row).setValue("Yes");
      Logger.log("Conditions are not met for ROW "+row+". Either URL_STRING is No Result or Humantic Profile is already turned to Yes in ROW "+row+".");
      return true;
    }
    
    
  } else {
    
    Logger.log("Final URL To Process at ROW "+row+" is empty! I didn't even started the process of creating a Profile in create_Humantic_Profile() function for ROW "+row+"!");
    
    return true;
    
  }
}




// Process All Rows and CREATE Humantic Profile

function Process_All_Rows_Create_Humantic_Profile() {
  
  var i = 0;
  
  let lastIndexedRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("C2").getValues(); // Get the value of last row that has been indexed (spreadsheet name: variables, column/cell B2)
  let lastSpreadsheetRow = SpreadsheetApp.openById("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE").getSheetByName("Sheet1").getLastRow(); //Get the last used row of the spreadsheet
  
  for (row = lastIndexedRow; row < lastSpreadsheetRow; row++) {
    
    if ( i < 4) {  //Change this number to change the max number of created profiles
      
      if (create_Humantic_Profile(row)) {
        
        i++;
        Logger.log(i);
        
      }
    } else {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("C2").setValue(row);
      break;
    }
    
  }
  
//  if (row > lastSpreadsheetRow) {
//    Logger.log("We reached the last Row of Humantic Create Profile. I will stop now!");
//    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("B2").setValue(row);
//  }
  
}


/*-----------------END HUMANTIC PROFILE CREATION----------------*/







/*-----------------GET HUMANTIC DATA----------------*/


function getData(url,row) {
  
  var url_string = url;
  
  var response = UrlFetchApp.fetch(url_string);
  var content = response.getContentText();
  var json = JSON.parse(content);
  
  const isSuccesful = json.message;
  
  if (isSuccesful == "Success") {
    var resultsFull = json.results;
    var firstName = json.results["first_name"];
    var lastName = json.results["last_name"];
    var educationArr = json.results.education;
    const persona = json.results.persona;
    const workHistory = json["results"]["work_history"];
    let eduSchool = "";
    let organisation = "";
    let jobTitle = "";
    const personalitySum = json["results"]["personality_analysis"]["summary"];
    const personaObj = json["results"]["persona"];
    const personaSalesObj = json["results"]["persona"]["sales"];
    let salesComAdvice = "";
    
    // PRINT FIRST & LAST NAME
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("K"+row).setValue(firstName);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("L"+row).setValue(lastName);
    
    
    // EDUCATION
    if (!arrIsEmpty(educationArr)) {
      eduSchool = educationArr[0].school;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("M"+row).setValue(eduSchool);
    }
    
    // WORK HISTORY RECORDS
    if (!arrIsEmpty(workHistory)) {
      organisation = workHistory[0].organization;
      jobTitle = workHistory[0].title;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("N"+row).setValue(organisation);
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("O"+row).setValue(jobTitle);
    }
    
    // PERSONALITY RECORDS
    
    // DISC RECORDS
    if (!objIsEmpty(personalitySum)) {
      if(!objIsEmpty(personalitySum.disc)) {
        const discDescr = [];
        const discLabel = [];
        for (let i = 0; i <= personalitySum.disc.description.length -1; i++) {
          discDescr[i] = personalitySum.disc.description[i];
          discLabel[i] = personalitySum.disc.label[i];
          Logger.log(discDescr);
          Logger.log(discLabel);
        }
        // PRINT DISC DESCRIPTION AND LABEL
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("P"+row).setValue(discDescr.toString());
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("Q"+row).setValue(discLabel.toString());
      }
      
      // OCEAN RECORDS
      if(!objIsEmpty(personalitySum.ocean)) {
        const oceanDescr = [];
        const oceanLabel = [];
        for (let i = 0; i <= personalitySum.ocean.description.length -1; i++) {
          oceanDescr[i] = personalitySum.ocean.description[i];
          oceanLabel[i] = personalitySum.ocean.label[i];          
          Logger.log(oceanDescr);
          Logger.log(oceanLabel);
        }
        // PRINT OCEAN DESCRIPTION AND LABEL
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("R"+row).setValue(oceanDescr.toString());
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("S"+row).setValue(oceanLabel.toString());
      }
    }
    
    // SALES COMMUNICATION ADVICE
    if (!objIsEmpty(personaObj)) {
      if (!objIsEmpty(personaSalesObj)) {
        let salesComAdviceType = [];
        let salesComAdviceDesc = [];
        let salesComAdviceAdj = [];
        let salesComAdviceSay = [];
        let salesComAdviceAvoid = [];
        let salesComAdviceKeyTraits = "";
        salesComAdvice = json["results"]["persona"]["sales"]["communication_advice"];
        
        //TYPE
        for (let i = 0; i <= salesComAdvice["_type"].length -1; i++) {        
        salesComAdviceType[i] = salesComAdvice["_type"][i];
        }
        
        //DESCRIPTION
        for (let i = 0; i <= salesComAdvice["description"].length -1; i++) {
        salesComAdviceDesc[i] = salesComAdvice["description"][i];
        }
        
        //ADJECTIVES
        for (let i = 0; i <= salesComAdvice["adjectives"].length -1; i++) {
        salesComAdviceAdj[i] = salesComAdvice["adjectives"][i];
        }
        
        //WHAT TO SAY
        for (let i = 0; i <= salesComAdvice["what_to_say"].length -1; i++) {
        salesComAdviceSay[i] = salesComAdvice["what_to_say"][i];
        }
        
        //WHAT TO AVOID
        for (let i = 0; i <= salesComAdvice["what_to_avoid"].length -1; i++) {
        salesComAdviceAvoid[i] = salesComAdvice["what_to_avoid"][i];
        }
        
        //KEY TRAITS
        salesComAdviceKeyTraits = JSON.stringify(salesComAdvice["key_traits"]);
        
        
        //PRINT THE WHOLE SALES OBJECT
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("T"+row).setValue(JSON.stringify(salesComAdvice));
        
        //PRINT TYPE
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("U"+row).setValue(salesComAdviceType.toString());
        
        //PRINT DESCRIPTION
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("V"+row).setValue(salesComAdviceDesc.toString());
        
        //PRINT ADJECTIVES
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("W"+row).setValue(salesComAdviceAdj.toString());
        
        //PRINT WHAT TO SAY
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("X"+row).setValue(salesComAdviceSay.toString());
        
        //PRINT WHAT TO AVOID
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("Y"+row).setValue(salesComAdviceAvoid.toString());
        
        //PRINT KEY TRAITS
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange("Z"+row).setValue(salesComAdviceKeyTraits.toString());
      }
      
      
    }
    
    return true;
  } else {
    Logger.log("There was not a success response from the Humantic API");
    return true;
    
  }
  
}


/*----------------END GET HUMANTIC DATA FUNCTION------------*/








/*--------------HUMANTIC PROFILE FETCHING - RETRIEVE PROFILE INFO--------------*/


// Call Humantic AI API, pass the LinkedIn URL and fetch Profile data
function Humantic_AI(row) { 
  
  
  var url_string = Get_LinkedIn_URL(row);
  
  if (url_string != 0) {
    
    //Get the Processed Value
    var processed = Sheets.Spreadsheets.Values.get("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE", "J"+row).values[0];
    
    if (url_string != "No Result" && processed != "Yes")  {
      
      url_string = "https://api.humantic.ai/v1/user-profile?apikey=chrexec_edf11eb1e9045701bd531c600197ae5d&userid="+url_string;
      
      //Change processed Status to "Yes"
      SpreadsheetApp.getActiveSheet().getRange("J"+row).setValue("Yes");
      
      if(getData(url_string,row)) {
        return true;
      }
      
    } else {
      
      //Change processed Status to "Yes"
      SpreadsheetApp.getActiveSheet().getRange("J"+row).setValue("Yes");
      
      return true;
      
    }
  } else {
    Logger.log("There is not a Final URL yet! Profile fetching didn't start at all at ROW "+row+"!");
    
    return true;
  }
  
}



// Set Number of row and Run Humantic Profile Fetching
function Process_All_Rows_Fetch_Humantic_Profile_Info() {
  var i = 0;
  
  let lastIndexedRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("D2").getValues(); // Get the value of last row that has been indexed (spreadsheet name: variables, column/cell C2)
  let lastSpreadsheetRow = SpreadsheetApp.openById("1F-w_yv3eZrVTG4F_0CVmsfXBzGmmXgwbfwesr1v6nZE").getSheetByName("Sheet1").getLastRow(); //Get the last used row of the spreadsheet
  
  for (row = lastIndexedRow; row <= lastSpreadsheetRow; row++) {
    if (i <= 4) {  //Change this number to change the max number of created profiles
      
      if (Humantic_AI(row)) {
        
        i++;
        Logger.log(i);
      }    
      
    } else {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variables').getRange("D2").setValue(row);
      break;
    }
    
  }
}


/*-----------------END PROFILE FETCHING-------------*/
