//This module is used to conver the excel document into json//
var excelParser  = require('xlsx-to-json-lc');
//The lodash module has built in utilities to sort through data //
var _ = require('lodash');

//The excelParser requires an input name, output name, and a boolean for headers/
excelParser({
  input: 'sample.xlsx',
  output: 'sample.json',
  loweCaseHeaders: true

// this is a callback function to the excelParser. It is filtering through the data and pushing the results into the results object. //
}, function(err, results) {
  if (err) {
    console.log(err);
  } else {

    var array = [];

// the for loop below is looping through the file and adding any object that is <= Sept 6, 2010. Then it is pushing every object that has a value < Sept 06, 2010 into that array and parsing out the words in the words key for that object.//
    for (var i = 0; i < results.length; i++) {
      if (results[i].start_date <= 1283731200) {
        array.push(results[i]);
      }
  };

//Here i am sorting the start_date column by ascending order//
  array = _.sortBy(array, 'start_date', ['asc']);

  var phrase = '';

// Here, the loop is going through the objects stored in the array, which in turn gets the value at each words key and adds it to a string that is stored in the phrase variable. Finally, the phrase is logged.//

for (var i = 0; i < array.length; i++) {
    phrase += array[i].words + ' '
  }
  console.log(phrase);
  }
});
