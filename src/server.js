const path = require("path");
const fs = require("fs");
const open = require('open');
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
    // console.log(data);
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
    } catch (e) {
        console.error(e);
    }
    res.redirect("/");
});

function findByRow(person, book) {
    const sheets = data.Sheets;
    const result = [];
    for (const sheetname in sheets) {
        const sheet = sheets[sheetname];
        console.log(sheet);
        if (sheet['!ref'] == undefined) {
            return null;
        }
        const range = xlsx.utils.decode_range(sheet['!ref']);
        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            const row = [];
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                const nextCell = sheet[
                    xlsx.utils.encode_cell({
                        r: rowNum,
                        c: colNum
                    })
                ];
                if (typeof nextCell === 'undefined') {
                    row.push(void 0);
                } else row.push(nextCell.v);
            }
            console.log(row);
            if (person && book) {
                if (row.includes(person) && row.includes(book)) {
                    return result.push(row);
                }
            } else if (person && !book) {
                if (row.includes(person)) {
                    return result.push(row);
                }
            } else if (!person && book) {
                if (row.includes(book)) {
                    return result.push(row);
                }
            } else {
                result.push(row);
            }

        }
    }
    if (result.length === 0) {
        return null;
    }
    return result;
};

const appendByRow = (person, book) => {
    const sheets = data.Sheets;
    for (const sheetname in sheets) {
        const sheet = sheets[sheetname];
        const date = new Date();
        date.setDate(date.getDate() + 7);
        console.log(sheet);
        xlsx.utils.sheet_add_aoa(sheet, [
            [person, book, date.toUTCString()]
        ], {
            origin: -1
        });
    };
}

const renewByRow = (person, book) => {
    const sheets = data.Sheets;
    for (const sheetname in sheets) {
        const sheet = sheets[sheetname];
        console.log(sheet);
        const range = xlsx.utils.decode_range(sheet['!ref']);
        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            const row = [];
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                const nextCell = sheet[
                    xlsx.utils.encode_cell({
                        r: rowNum,
                        c: colNum
                    })
                ];
                if (typeof nextCell === 'undefined') {
                    row.push(void 0);
                } else row.push(nextCell.v);
            }
            if (row.includes(person) && row.includes(book)) {
                const cell = sheet[
                    xlsx.utils.encode_cell({
                        r: rowNum,
                        c: range.e.c
                    })
                ];
                const date = new Date(cell.v);
                date.setDate(date.getDate() + 7);
                cell.v = date.toUTCString();
            }
        }
    }
};

function deleteByRow(person, book) {
    const sheets = data.Sheets;
    for (const sheetname in sheets) {
        const sheet = sheets[sheetname];
        console.log(sheet);
        const range = xlsx.utils.decode_range(sheet['!ref']);
        let overwrite = false;
        let rowNum;
        for (rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            const row = [];
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                const nextCell = sheet[
                    xlsx.utils.encode_cell({
                        r: rowNum,
                        c: colNum
                    })
                ];
                if (typeof nextCell === 'undefined') {
                    row.push(void 0);
                } else row.push(nextCell.v);
            }
            console.log(row);
            if (row.includes(person) && row.includes(book)) {
                overwrite = true;
                break;
            }
        }
        if (overwrite) {
            for (; rowNum <= range.e.r; rowNum++) {
                for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                    sheet[
                        xlsx.utils.encode_cell({
                            r: rowNum,
                            c: colNum
                        })
                    ] = sheet[
                        xlsx.utils.encode_cell({
                            r: rowNum + 1,
                            c: colNum
                        })
                    ];

                }
            }

            range.e.r--;
            sheet['!ref'] = xlsx.utils.encode_range(range.s, range.e);
            return true;
        }
        return false;
    };
}

app.get('/lend', (req, res) => {
    console.log(req.query);
    if (findByRow(req.query.person, req.query.book) != null) {
        line = "중복되는 대출이 있습니다.";
        console.log(line);
        savexlsx();
        res.redirect("/");
        return;
    }
    appendByRow(req.query.person, req.query.book);
    line = "대출 되었습니다.";
    console.log(line);
    savexlsx();
    res.redirect("/");
});

app.get('/return', (req, res) => {
    if (findByRow(req.query.person, req.query.book) == null) {
        line = "대출이 없습니다.";
        console.log(line);
        savexlsx();
        res.redirect("/");
        return;
    }
    deleteByRow(req.query.person, req.query.book);
    line = "반납 되었습니다.";
    console.log(line);
    savexlsx();
    res.redirect("/");
});

app.get('/renew', (req, res) => {

    if (findByRow(req.query.person, req.query.book) == null) {
        line = "대출이 없습니다.";
        console.log(line);
        savexlsx();
        res.redirect("/");
        return;
    }

    renewByRow(req.query.person, req.query.book);
    line = "갱신 되었습니다.";
    console.log(line);
    savexlsx();
    res.redirect("/");
});

app.get('/save', (req, res) => {
    console.log(outPath);
    savexlsx();
    res.download(outPath, req.query["filename"]);
});

app.listen(port || 80);

console.log(`listening at ${port}!`);
console.log(`http://localhost`);

(async() => {
    await open(`http://localhost`);
})();


function getData() {
    if (fs.existsSync(outPath)) {
        return xlsx.readFile(outPath);
    }
    return null;
}

function savexlsx() {
    xlsx.writeFile(data, outPath, { bookType: "xlsx" });
}

function removexlsx() {
    fs.unlinkSync(outPath);
    data = null;
}