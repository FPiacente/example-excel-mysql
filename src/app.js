const express = require('express')
const app = express()
const bodyparser = require('body-parser')
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
const mysql = require('mysql2')
const multer = require('multer')
const path = require('path')
//use express static folder
app.use(express.static("./public"))
// body-parser middleware use
app.use(bodyparser.json())
app.use(bodyparser.urlencoded({
    extended: true
}))
// Database connection
const db = mysql.createConnection({
    host: "localhost",
    user: "isi",
    password: "BelimoISI",
    database: "isi"
})
db.connect(function (err) {
    if (err) {
        return console.error('error: ' + err.message);
    }
    console.log('Connected to the MySQL server.');
})
// Multer Upload Storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, '../src' + '/uploads/')
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname)
    }
});
const upload = multer({storage: storage});
//! Routes start
//route for Home page
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});
// -> Express Upload RestAPIs
app.post('/uploadfile', upload.single("uploadfile"), (req, res) =>{
    importExcelData2MySQL('../src' + '/uploads/' + req.file.filename);
    console.log(res);
});
// -> Import Excel Data to MySQL database
function importExcelData2MySQL(filePath){
// File path.
    readXlsxFile(filePath).then((rows) => {
// `rows` is an array of rows
// each row being an array of cells.
        console.log(rows);
        /**
         [ [ 'Id', 'Name', 'Address', 'Age' ],
         [ 1, 'john Smith', 'London', 25 ],
         [ 2, 'Ahman Johnson', 'New York', 26 ]
         */
// Remove Header ROW
        rows.shift();
// Open the MySQL connection
        db.connect((error) => {
            if (error) {
                console.error(error);
            } else {
                let query = 'INSERT INTO isilo2 (erstellungsdatum, isiid, idocnr, refidocnr, nachrichtentyp, status, fehlercode, fehlertext, aktposition, tanr, tapos, gebindenr, quellgebinde, quelltyp, zielgebinde, zieltyp, menge, restmenge, mengeneinheit, invdiff, artikel, charge, nio, ziel, prio, bearbeitungsdatum, datum, wickelprogramm, warenstatus, werk, lagerort, sonderbestkennz, sonderbestnr) VALUES ?';
                db.query(query, [rows], (error, response) => {
                    console.log(error || response);
                    /**
                     OkPacket {
fieldCount: 0,
affectedRows: 5,
insertId: 0,
serverStatus: 2,
warningCount: 0,
message: '&Records: 5  Duplicates: 0  Warnings: 0',
protocol41: true,
changedRows: 0 }
                     */
                });
            }
        });
    })
}
// Create a Server
let server = app.listen(8080, function () {
    let host = server.address().address
    let port = server.address().port
    console.log("App listening at http://%s:%s", host, port)
})