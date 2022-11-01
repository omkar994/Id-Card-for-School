const express = require('express');
const app = express();
const path = require('path');
const engine = require('ejs-mate');
const methodOverride = require('method-override');
const mongoose = require('mongoose');
const Student = require('./models/student');
const { Parser } = require('json2csv');
const exceljs = require('exceljs');
const moment = require('moment');

mongoose.connect('mongodb://localhost:27017/Student-DB', {
    useNewUrlParser: true,
    useUnifiedTopology: true,

})
    .then(() => {
        console.log("Database Connected!!!");
    })
    .catch((err) => {
        console.log("ERROR! Connecting to Database.");
        console.log(err);
    });

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));
app.engine('ejs', engine);
app.use(methodOverride('_method'));
app.locals.moment = require('moment');


app.get('/idform', (req, res) => {
    res.render('form');
});

app.post('/student/search', async (req, res) => {
    const stduClass = req.body.class;
    const { division, rollNo } = req.body;

    await Student.find({ class: stduClass, division: division, rollNo: rollNo }, function (err, data) {
        if (err) {
            console.log(err);
        }
        else {
            output = data;
            if (output.length == 0) {
                res.send('Not Found');
            }
            else
                res.render('search', { output });
        }
    }).clone().catch(function (err) { console.log(err) });
});

app.get('/getcsv', async (req, res) => {

    await Student.find({}, function (err, data) {
        if (err) {
            console.log(err);
        }
        else {
            output = data;
            if (output.length == 0) {
                res.send('Not Found');
            }
            else {
                const json2csvParser = new Parser();
                const csv = json2csvParser.parse(output);
                res.attachment('student info.csv');
                res.status(200).send(csv);
            }

        }
    }).clone().catch(function (err) { console.log(err) });

})

app.get('/student/:id/edit', async (req, res) => {
    const student = await Student.findById({ _id: req.params.id });
    //console.log(student);
    res.render('edit', { student });

});

app.post('/student', async (req, res) => {
    const stduClass = req.body.class;
    const { division, rollNo } = req.body;
    const student = await Student.find({ rollNo: rollNo, class: stduClass, division: division });
    if (student.length !== 0) {
        res.send('Already Registered');
    }
    else {
        const student = new Student(req.body);
        await student.save();
        res.render('summary', { student });
    }
});

app.get('/students', async (req, res) => {
    const students = await Student.find({});
    res.render('students', { students });
});

app.put('/student/:id', async (req, res) => {
    const { id } = req.params;
    await Student.findByIdAndUpdate(id, { ...req.body });
    res.redirect('/students');
});

app.get('/student/export', async (req, res) => {
    try {
        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet('Students');

        worksheet.columns = [
            { headers: 'Name', key: 'stduName' },
            { headers: 'Class', key: 'class' },
            { headers: 'Division', key: 'division' },
            { headers: 'Photo Ref', key: 'photoRefNo' },
            { headers: 'Birth Date', key: 'bdate' },
            { headers: 'Blood Group', key: 'bloodGrp' },
            { headers: 'Contact', key: 'contactNo' },
            { headers: 'Address', key: 'address' }
        ];

        const students = await Student.find({});

        students.forEach((student) => {
            worksheet.addRow(student);
        });

        worksheet.getRow(1).eachCell((cell) => {
            cell.font = { bold: false };
        });

        res.setHeader(

            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheatml.sheet'
        );

        res.setHeader(
            'Content-Disposition',
            `attachment; filename=students.xlsx`
        );

        return workbook.xlsx.write(res).then(() => {
            res.status(200);
        });
    }
    catch (err) {
        console.log(err.message);
    }
});

app.listen(3000, () => {
    console.log('Serving on PORT 3000');
});
