const path = require("path");
const fs = require("fs");
const express = require("express");
const bodyParser = require('body-parser');
const fileupload = require("express-fileupload");
const xlsx = require("xlsx");
const app = express();
const port = 80;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(fileupload());
app.use(express.static(path.join(__dirname, '/..', 'views')));
app.set('view engine', 'ejs');
app.engine('html', require('ejs').renderFile);

const outPath = path.join(path.dirname(__dirname), "out.xlsx");

let data = getData();

let line = "";



app.get('/', (req, res) => {
    res.render('index', {
        "data": data,
        "line": line
    });
    line = "";
});

app.post('/upload', (req, res) => {
    if (typeof req.files !== 'undefined') {
        data = xlsx.read(req.files.file.data, { type: 'buffer' });
        console.log("data get");
        savexlsx();
    }
    res.redirect("/");
});

app.get('/reset', (req, res) => {
    try {
        removexlsx();
    } catch {
        null;
    }
    res.redirect("/");
});

app.get('/lend', (req, res) => {
    // TODO
    line = "대출 되었습니다.";
    console.log(line);
    savexlsx();
    res.redirect("/");
});

app.get('/return', (req, res) => {
    // TODO
    line = "반납 되었습니다.";
    console.log(line);
    savexlsx();
    res.redirect("/");
});

app.get('/renew', (req, res) => {
    // TODO
    line = "갱신 되었습니다.";
    console.log(line);
    savexlsx();
    res.redirect("/");
});

app.get('/save', (req, res) => {
    console.log(outPath);
    savexlsx();
    res.sendFile(outPath);
});

app.listen(port || 80);

console.log(`listening at ${port}!`);
console.log(`http://localhost`);

function getData() {
    if (fs.existsSync(outPath)) {
        return xlsx.readFile(outPath);
    }
    return null;
}

function savexlsx() {
    xlsx.writeFile(data, outPath);
}

function removexlsx() {
    fs.unlinkSync(outPath);
}