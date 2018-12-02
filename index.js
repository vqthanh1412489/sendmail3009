const express = require('express');
const parser = require('body-parser');
const nodemailer = require('nodemailer');
const fs = require('fs');
var multer = require('multer');
let exceltojson;

var xlstojson = require("xls-to-json-lc");
var xlsxtojson = require("xlsx-to-json-lc");

var storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length - 1])
    }
});

var upload = multer({
    storage,
    fileFilter: function (req, file, callback) { //file filter
        if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length - 1]) === -1) {
            return callback(new Error('Wrong extension type'));
        }
        callback(null, true);
    }
}).single('file');

const app = express();

app.use(parser.json());
app.use(parser.urlencoded({ extended: true }));

app.set('view engine', 'ejs');
app.set('views', './views');
app.use(express.static('./public'));

app.listen(process.env.PORT || 3000, console.log('Start!'));

app.get('/', (req, res) => {
    res.render('upload');
});
app.get('/send', (req, res) => {
    res.render('send');
});
app.get('/success', (req, res) => {
    res.render('success');
});

let manageFile;
app.post('/upload', (req, res) => {
    upload(req, res, (err) => {
        if (err) {
            res.render('405');
            return;
        }
        if (!req.file) {
            res.render('404');
            return;
        }
        manageFile = req.file;
        res.redirect('/send');
    });
});


app.post('/send', (req, res) => {
    if (manageFile.originalname.split('.')[manageFile.originalname.split('.').length - 1] === 'xlsx') {
        exceltojson = xlsxtojson;
    } else {
        exceltojson = xlstojson;
    }

    try {
        exceltojson(
            {
                input: manageFile.path,
                output: null, //since we don't need output.json
                lowerCaseHeaders: true
            }, function (err, result) {
                if (err) {
                    return res.render('405');
                }

                const { subject, message } = req.body;

                let transporter = nodemailer.createTransport({
                    service: 'Gmail',
                    auth: {
                        user: 'vqthanh1@gmail.com', // generated ethereal user
                        pass: 'thanh1234' // generated ethereal password
                    }
                });
                let arrEmail = [];
                result.forEach(element => {
                    arrEmail.push(element.email);
                });
                let mailOptions = {
                    from: '"Vqthanh1 👻"', // sender address
                    to: arrEmail.join(','), // list of receivers
                    subject: subject, // Subject line
                    html: `<b>${message}</b>` // html body
                };
                transporter.sendMail(mailOptions, (error, info) => {
                    if (error) {
                        res.render('405');
                    }
                });

                try {
                    fs.unlinkSync(manageFile.path);
                } catch(e) {
                    console.log(e);
                }

                manageFile = null;
                res.redirect('success');
            });
    } catch (e) {
        res.render('brokenFile');
    }
});





