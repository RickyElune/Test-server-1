var createError = require("http-errors");
var express = require("express");
var path = require("path");
var cookieParser = require("cookie-parser");
var logger = require("morgan");
var XLSX = require("xlsx");
var serveIndex = require("serve-index");
const multer = require("multer");
const cors = require("cors");

var esgReportingData = require("./public/json/esgReporting.json");

var indexRouter = require("./routes/index");
var usersRouter = require("./routes/users");

var app = express();

// view engine setup
app.set("views", path.join(__dirname, "views"));
app.set("view engine", "jade");

app.use(cors());
app.use(logger("dev"));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, "public")));
app.use("/sheets", serveIndex(__dirname + "/public/sheets"));

app.use("/", indexRouter);
app.use("/users", usersRouter);

const fileStorageEngine = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, "./public/sheets"); //important this is a direct path fron our current file to storage location
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    },
});

const upload = multer({ storage: fileStorageEngine });

// Single File Route Handler
app.post("/single", upload.single("sheet"), async (req, res) => {
    const excelFile = req.file;
    console.log(excelFile);
    const filePath = excelFile.path;
    const workbook = XLSX.readFile(filePath);
    // Convert the XLSX to JSON
    let worksheets = {};
    for (const sheetName of workbook.SheetNames) {
        worksheets[sheetName] = XLSX.utils.sheet_to_json(
            workbook.Sheets[sheetName]
        );
    }

    let sheetData = worksheets.Sheet1;
    for (const e of sheetData) {
        delete e.__EMPTY;
    }

    // Show the data as JSON
    console.log("json:\n", JSON.stringify(sheetData), "\n\n");

    esgReportingData.esgreport.h23303a21.a21.electricity = sheetData[0];
    esgReportingData.esgreport.h23303a22.a22.water = sheetData[1];
    esgReportingData.esgreport.h23301a12.a12.ghg = sheetData[2];

    res.send(esgReportingData);
});

// Multiple Files Route Handler
app.post("/multiple", upload.array("sheets", 3), (req, res) => {
    console.log(req.files);
    res.send("Multiple Files Upload Success");
});

// catch 404 and forward to error handler
app.use(function (req, res, next) {
    next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
    // set locals, only providing error in development
    res.locals.message = err.message;
    res.locals.error = req.app.get("env") === "development" ? err : {};

    // render the error page
    res.status(err.status || 500);
    res.render("error");
});

app.listen(5000);

module.exports = app;
