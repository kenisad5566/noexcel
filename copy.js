
let path = require('path')
var shell = require('shelljs');
let src = path.resolve('./src/excel4node')
let dist =  path.resolve('./dist')

shell.cp('-R', src, dist);