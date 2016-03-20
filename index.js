var GoogleSpreadsheet = require("google-spreadsheet");
var async = require('async');
var Promise = require('bluebird');
var config;
var logger;
var worksheets;
var maxRows = null;
var maxCols = null;

var Spreadsheet2D = function(id, worksheets){
	var self = this;

	self.id = id;
	self.worksheets = [];

	worksheets.forEach(function(s){
		self.worksheets.push(new Worksheet2D(s));
	});

	self.sheet = function(idx){
		return self.worksheets[idx];
	}
};

var Worksheet2D = function(ws){
	var self = this;

	self.title = ws.title;
	self.id = ws.title;
	self.url = ws.url;
	self.data = ws.extractedCells;

	self.get = function(row, col){
		return self.data[row][col];
	}
};




var workSheetTo2DArray = function (ws, cb) {

	rows = maxRows || ws.rowCount;
	cols = maxCols || ws.colCount;

	logger.info('extracting sheet: ' + ws.title, rows + 'x' + cols, 'from', ws.rowCount + 'x' + ws.colCount);

	// create 2d array with given size populated with null
	ws.extractedCells = Array.apply(null, Array(rows)).map(function () {
		return Array.apply(null, Array(cols)).map(function () {
			return null
		})
	});


	var query = {
		'min-row': 1,
		'max-row': rows,
		'return-empty': true
	};

	ws.getCells(query, function (err, cells) {
		cells.forEach(function(c){
			ws.extractedCells[c.row - 1][c.col - 1] = c.value;
		});
		cb(null, ws);
	});
};


var extractContents = function (id) {

	// with which doc are we working
	var doc = new GoogleSpreadsheet(id);

	return new Promise(function(resolve, reject){
		// function names are just for understanding - could be anonymous
		async.series([
			function authorize(step) {
				doc.useServiceAccountAuth(config.credentials, step);
			},
			function loadSheets(step) {
				doc.getInfo(function (err, info) {
					if (err) {
						logger.error(err);
						step(err, null);
					}
					worksheets = info.worksheets;
					logger.info('Loaded doc: "' + info.title + '" by ' + info.author.email);
					logger.info('found', worksheets.length, 'worksheets: ', worksheets.map(function(e){return e.title}));
					step();
				});
			},
			function to2DArray(step){
				async.map(worksheets, workSheetTo2DArray, function(err, result){
					if(err){
						logger.error(err);
						step(err);
					}
					step(null, result);
				});
			}
		], function(err, result){
			if(err){
				reject(err);
			}



			resolve(new Spreadsheet2D(id, worksheets));
		});
	});

};

module.exports = function(cnf){
	config = cnf;
	logger = cnf.logger || console;
	maxRows = cnf.rows || maxRows;
	maxCols = cnf.cols || maxCols;

	return {
		extract: extractContents
	};
};