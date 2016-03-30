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
	};

	self.sheetById = function(id){
		var result = null;
		self.worksheets.forEach(function(s){
			if(s.id === id){
				result = s;
			}
		});
		return result;
	}
};

var Worksheet2D = function(ws){
	var self = this;

	self.title = ws.title;
	self.id = ws.id;
	self.url = ws.url;
	self.data = ws.extractedCells;

	self.get = function(row, col){
		return self.data[row][col];
	};

	self.set = function(row, col, value){
		self.data[row][col] = value;
	};

};


var array2DToWorksheet = function (ws, cb) {

	if(ws.dataToSave == null){
		cb(null, ws);
		return;
	}

	rows = maxRows || ws.rowCount;
	cols = maxCols || ws.colCount;

	logger.info('setting contents of sheet: ' + ws.title);

	var query = {
		'min-row': 1,
		'max-row': rows,
		'min-cols': 1,
		'max-cols': cols,
		'return-empty': true
	};

	ws.getCells(query, function (err, cells) {
		cells.forEach(function(c){
			c.value = ws.dataToSave[c.row-1][c.col-1];
		});
		ws.dataToSave = null;
		logger.info('saving sheet: ', ws.title );
		ws.bulkUpdateCells(cells, cb);
	});


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


var saveContents = function(id, spreadsheet2d){
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
					logger.info('Loaded doc for saving: "' + info.title + '" by ' + info.author.email);
					logger.info('found', worksheets.length, 'worksheets: ', worksheets.map(function(e){return e.id + ': ' + e.title}));

					// attach dirty data
					worksheets.forEach(function(ws){
						ws.dataToSave = spreadsheet2d.sheetById(ws.id).data;
					});

					//logger.info(spreadsheet2d.worksheets.map(function(s){return s.id + ': ' + s.title}));
					step();
				});
			},
			function to2DArray(step){
				async.map(worksheets, array2DToWorksheet, function(err, result){
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

var loadContents = function (id) {

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
					logger.info('found', worksheets.length, 'worksheets: ', worksheets.map(function(e){return e.id + ': ' + e.title}));
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
		load: loadContents,
		save: saveContents
	};
};