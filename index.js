var _ = require('lodash');
var GulpUtil = require('gulp-util');
var fs = require('fs');
var exec = require('sync-exec');
var xlsx = require('xlsx');

var Converter = function() {
    this.allowedFiles = ['xlsx', 'xls', 'csv'];
    this.directory = null;
};

Converter.prototype.init = function() {

    this.directory = GulpUtil.env.directory;
    this.outputdir = GulpUtil.env.outputdir;

    if (!this.directory) {
        console.log('No directory found');
        process.exit()
    }

    if (!this.outputdir) {
        this.outputdir = './output';
        console.log('Output directory set to ./output.');
    }

    var files = this.readDirectory(this.directory);

    //_.each(files, this.convertFiletoCsv.bind(this));

    this.convertFiletoCsv(files[300]);

    console.log(files.length + ' files to be converted to csv.');
};

Converter.prototype.convertFiletoCsv = function(file) {
    var path = this.directory + file;
    var workbook = xlsx.readFile(path);
    var em = this.extractEmailsFromString(JSON.stringify(workbook));
    fs.writeFileSync('output.csv', em.join('\n'));
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
