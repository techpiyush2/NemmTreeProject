let express = require('express');
let app = express();
let path = require('path');
let multer = require('multer');
let xlstojson = require("xls-to-json");
let xlsxtojson = require("xlsx-to-json");
const staticpath = path.join(__dirname, "/Public")

app.use(express.static(staticpath));

let storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length - 1])
    }
});

let upload = multer({ //multer settings
    storage: storage,
    fileFilter: function (req, file, callback) { //file filter
        if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length - 1]) === -1) {
            return callback(new Error('Wrong extension type'));
        }
        callback(null, true);
    }
}).single('file');

/** API path that will upload the files */
app.post('/upload', function (req, res) {
    var exceltojson;
    upload(req, res, function (err) {
        if (err) {
            console.log({ error_code: 1, err_desc: err });
            return;
        }
        /** Multer gives us file info in req.file object */
        if (!req.file) {
            console.log({ error_code: 1, err_desc: "No file passed" });
            return;
        }
        /** Check the extension of the incoming file and 
         */
        if (req.file.originalname.split('.')[req.file.originalname.split('.').length - 1] === 'xlsx') {
            exceltojson = xlsxtojson;
        } else {
            exceltojson = xlstojson;
        }
        console.log(req.file.path);
        try {
            exceltojson({
                input: req.file.path,
                output: null,
                lowerCaseHeaders: true
            }, function (err, result) {
                if (err) {
                    console.log(res.json({ error_code: 1, err_desc: err, data: null }));
                }
                console.log({ error_code: 0, err_desc: null, data: result });
            });
        } catch (e) {
            console.log({ error_code: 1, err_desc: "Corupted excel file" });
        }
    })

});

app.get('/', function (req, res) {
    res.sendFile(__dirname + "/index.html");
});

app.listen('3000', function () {
    console.log('running on 3000...');
});