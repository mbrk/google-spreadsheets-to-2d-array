# google-spreadsheets-to-2d-array
simple node module to transform all worksheets in a google spreadsheet into a 2 dimensional array accessible
via x (=columns) and y (=rows) coordinates

## Installation

    npm install google-spreadsheets-to-2d-array

## Usage


	// configuration includes:
	//  - credentials for your google service account
	//  - number of rows to extract - otherwise all rows are extracted
	//  - number of columns to extract - otherwise all cols are extrcted

	var config = {
		credentials: require('./credentials.json'),
		rows: 20,
		cols: 10
	};

	var sheet = require('google-spreadsheets-to-2d-array')(config);

	var id = '<your-spreadsheet-id>';

	// load a spreadsheet
	sheet.load(id)
    	.then(function(r){
    		console.log(r.sheet(0).get(0,1));
    		console.log(r.sheet(1).get(0,1));
    		console.log(r.sheet(0).data.length);
    	})
    	.catch(function(err){
    		console.error(err);
    	});

	// load, change and save a spreadsheet
	sheet.load(id)
    	.then(function(r){
    		// make changes to worksheet 1
    		r.sheet(0).set(0,10,'Value 1');
    		r.sheet(0).set(2,10,'Value 2');
    		return sheet.save(id, r);

    	})
    	.then(function(r){
    		console.log('success');
    	})
    	.catch(function(err){
    		console.error(err);
    	});



## Thanks

Thanks to https://github.com/theoephraim/node-google-spreadsheet for doing the real work.


## Changelog

v0.1.0 added saving capability, renamed #extract() to #load()
v0.0.1 basic implementation