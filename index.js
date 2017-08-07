var _ = require('lodash');
var GulpUtil = require('gulp-util');
var exec = require('sync-exec');
var xlsx = require('xlsx');
var fs = require('fs');
var mkdirp = require('mkdirp');
var getDirName = require('path').dirname;
var filendir = require('filendir');

var Converter = function() {
    this.allowedFiles = ['xlsx', 'xls', 'csv'];
    this.outputdir = './output/';
    this.fileSuffix = '-emails.csv'
    this.directory = null;
    this.totalFiles = 0;
};

Converter.prototype.init = function() {

    this.directory = GulpUtil.env.directory;
    this.outputdir = GulpUtil.env.outputdir;

    if (!this.directory) {
        console.log('No directory found');
        process.exit()
    }

    if (!this.outputdir) {
        this.outputdir = './output/';
        console.log('Output directory set to ./output.');
    }

    var files = this.readDirectory(this.directory);

    this.totalFiles = files.length;

    console.log(files.length + ' files to be converted to csv.');

    _.each(files, this.convertFiletoCsv.bind(this));
};

Converter.prototype.convertFiletoCsv = function(file, index) {

    console.log('Processing file ' + (index + 1) + ' out of ' + this.totalFiles + ' files.');

    var path = this.directory + file;
    var workbook = xlsx.readFile(path);
    var wfile = this.outputdir + file + this.fileSuffix;
    var fdir = wfile.replace(/ /g, '-');

    var em = this.extractEmailsFromString(JSON.stringify(workbook));
    filendir.writeFileSync(wfile.replace(/ /g, '-'), em.join('\n'));
};

Converter.prototype.extractEmailsFromString = function(text) {

    var matches = text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
    if (!matches)
        return [];

    return this.removeDuplicateValues(matches);
};

Converter.prototype.removeDuplicateValues = function(values) {

    return values.filter(function(item, index, inputArray) {
        return inputArray.indexOf(item) == index;
    });
};

Converter.prototype.readDirectory = function(dir, filelist) {

    var files = fs.readdirSync(dir);
    filelist = filelist || [];
    files.forEach(function(file) {
        if (fs.statSync(dir + file).isDirectory()) {
            filelist = this.readDirectory(dir + file + '/', filelist);
        } else {
            if (this.isAllowed(file)) {
                filelist.push(dir.replace(this.directory, '') + file);
            }
        }
    }.bind(this));
    return filelist;
};

Converter.prototype.isAllowed = function(name) {

    if (_.indexOf(this.allowedFiles, String(name).split('.').pop()) == -1) {
        return false;
    }
    return true;
};

var app = new Converter();
app.init();
