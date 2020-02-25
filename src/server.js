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

let data;

try {
    data = JSON.parse(fs.readFileSync("../data.json")) || {};
} catch {
    data = {};
}

let line = "";


function savexlsx() {
    fs.writeFileSync("../data.json", JSON.stringify(data));
}

function removexlsx() {
    data = {};
    savexlsx();
}


app.get('/', (req, res) => {
    res.render('index', {
        "data": data,
        "line": line
    });
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
    removexlsx();
    res.redirect("/");
})

app.get('/lend', (req, res) => {
    res.redirect("/");
})

app.get('/return', (req, res) => {
    res.redirect("/");
})

app.get('/renew', (req, res) => {
    res.redirect("/");
})

app.listen(port || 80);

console.log(`listening at ${port}!`);
console.log(`http://localhost`);