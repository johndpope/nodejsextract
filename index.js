var _ = require('lodash');
var GulpUtil = require('gulp-util');
var exec = require('child_process').exec;
var xlsx = require('xlsx');
var fs = require('fs');
var mkdirp = require('mkdirp');
var getDirName = require('path').dirname;
var filendir = require('filendir');
var moment = require('moment');
var stringify = require('json-stable-stringify');
var fileExists = require('file-exists');

var Converter = function() {
    this.startTime = null;
    this.allowedFiles = ['xlsx', 'xls', 'csv'];
    this.outputdir = './output/';
    this.fileSuffix = '-emails.csv'
    this.directory = null;
    this.totalFiles = 0;
};

Converter.prototype.init = function() {

    this.startTime = new Date();
    this.directory = GulpUtil.env.directory;
    this.outputdir = GulpUtil.env.outputdir;

    if (!this.directory) {
        console.log(this.getDateTimeSince(this.startTime) + ' ::: No directory found');
        process.exit()
    }

    if (!this.outputdir) {
        this.outputdir = './output/';
        console.log(this.getDateTimeSince(this.startTime) + ' ::: Output directory set to ./output.');
    }

    var files = this.readDirectory(this.directory);

    this.totalFiles = files.length;

    console.log(this.getDateTimeSince(this.startTime) + ' ::: ' + files.length + ' files to be converted to csv.');

    _.each(files, this.convertFiletoCsv.bind(this));
};

Converter.prototype.convertFiletoCsv = function(file, index) {

    console.log(this.getDateTimeSince(this.startTime) + ' ::: Processing file ' + (index + 1) + ' out of ' + this.totalFiles + ' files.');
    var em = [];
    var path = this.directory + file;
    var wfile = this.outputdir + file + this.fileSuffix;
    var fdir = wfile.replace(/ /g, '-');
    var filename = fdir.replace(/^.*[\\\/]/, '')

    if (fileExists.sync(wfile.replace(/ /g, '-'))) {
        console.log(this.getDateTimeSince(this.startTime) + ' ::: File already exists ' + (index + 1));
        return;
    }

    var workbook = xlsx.readFile(path);
    if (workbook) {
        try {
            em = this.extractEmailsFromString(JSON.stringify(workbook));
        } catch (e) {
            console.log(this.getDateTimeSince(this.startTime) + ' ::: Error ' + (index + 1));
            return;
        }
    }

    var fcsv = em.join('\n');
    var combinef = Date.now() + '-combine.csv';

    //combine xlsx
    if (fileExists.sync(combinef)) {
        fs.appendFileSync(combinef, '\n' + fcsv)
    } else {
        filendir.writeFileSync(combinef, fcsv);
    }

    exec('sort ' + combinef + ' | uniq -u',
        function(error, stdout, stderr) {
            fs.writeFileSync(combinef, stdout);
        });

    //write per xlsx
    filendir.writeFileSync(wfile.replace(/ /g, '-'), fcsv);

    console.log(this.getDateTimeSince(this.startTime) + ' ::: File ' + (index + 1) + ' - Unique Emails Found: ' + em.length + ' ::: ' + filename);
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

Converter.prototype.getDateTimeSince = function(target) { // target should be a Date object
    return moment(target).from(new Date(), true);
}

var app = new Converter();
app.init();
