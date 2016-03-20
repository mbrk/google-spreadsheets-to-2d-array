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

	sheet.extract(id)
		.then(function(r){
			console.log(r.sheet(0).get(0,1));
			console.log(r.sheet(1).get(0,1));
		});



