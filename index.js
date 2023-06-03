const express = require('express');
const xlsx = require('xlsx');
const ejs = require('ejs');
const path = require('path');
const axios = require('axios');

const app = express();
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'views/public')));
app.use(express.static('D:/Search_by_Excel_file'));

app.get('/search', (req, res) => {
  const searchTerm = req.query.term;
  const workbook = xlsx.readFile('D:/Search_by_Excel_file/test.xlsx');
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const columnNames = Object.keys(worksheet);
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  const matchingRows = data.filter((row, index) => {
    if (index === 0) {
      return false;
    }
    return row.some(cell =>
      cell.toString().toLowerCase().includes(searchTerm.toLowerCase())
    );
  });
  const noResults = matchingRows.length === 0;
  res.render('results', { columns: data[0], rows: matchingRows, noResults });
});

async function fetchShabbatTimes() {
  try {
    const response = await axios.get('https://www.hebcal.com/shabbat?cfg=json&geonameid=3448439&M=on');
    const data = response.data;
    // כאן תוכל לעבד את התוצאות שנמצאות במשתנה data
  
    return data.items;
  } catch (error) {
    console.error(error);
    return null;
  }
}





app.get('/', async (req, res) => {
  try {
    const workbook = xlsx.readFile('D:/Search_by_Excel_file/test.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const columnNames = Object.keys(worksheet);

    const formFields = columnNames.map(column => {
      return `<input type="text" name="${column}" placeholder="${column}" required>`;
    });

    const formHTML = `
      <div id="id01" class="modal">
        <form id="addBookForm" action="/add-book" method="POST" onsubmit="writeToExcel(event)">
          <h1>נא להכניס ספר חדש</h1>
          ${formFields.join('\n')}
          <button type="submit" class="btn btn-primary">הוסף ספר</button>
        </form>
      </div>
    `;

    const shabbatData = await fetchShabbatTimes();

    res.render('search-form', { shabbatData, formHTML });
  } catch (error) {
    console.error(error);
    res.render('search-form', { shabbatData: null });
  }
});

app.post('/search', (req, res) => {
  const searchTerm = req.body.term;
  const workbook = xlsx.readFile('D:/Search_by_Excel_file/test.xlsx');
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  const matchingRows = data.filter((row, index) => {
    if (index === 0) {
      return false;
    }
    return row.some(cell =>
      cell.toString().toLowerCase().includes(searchTerm.toLowerCase())
    );
  });
  const noResults = matchingRows.length === 0;
  res.render('results', { columns: data[0], rows: matchingRows, noResults });
});

// Function to write form data to Excel
app.post('/write-to-excel', (req, res) => {
  const data = req.body; // נתוני הטופס שנשלחו
  console.log(data);
  const workbook = xlsx.readFile('D:/Search_by_Excel_file/test.xlsx');
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const newRow = xlsx.utils.sheet_add_json(worksheet, [data], { skipHeader: true, origin: -1 });
  xlsx.writeFile(workbook, 'D:/Search_by_Excel_file/test.xlsx');
  res.sendStatus(200); // תשובת ההצלחה של הבקשה
});


function writeToExcel(event) {
  event.preventDefault(); // Prevent form submission

  // Get form data
  const form = document.getElementById('addBookForm');
  const formData = new FormData(form);

  // Convert form data to JSON object
  const jsonData = {};
  formData.forEach((value, key) => {
    jsonData[key] = value;
  });

  // Send the JSON data to the server for writing to Excel
  fetch('/write-to-excel', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(jsonData)
  })
  .then(response => {
    // Handle response from the server
    if (response.ok) {
      alert('הנתונים נכתבו בהצלחה לקובץ אקסל');
    } else {
      alert('אירעה שגיאה בכתיבת הנתונים לקובץ אקסל');
    }
  })
  .catch(error => {
    console.error('אירעה שגיאה:', error);
    alert('אירעה שגיאה בכתיבת הנתונים לקובץ אקסל');
  });
}

app.post('/add-book', (req, res) => {
  const data = req.body; // נתוני הטופס שנשלחו
  console.log(data);
  const workbook = xlsx.readFile('D:/Search_by_Excel_file/test.xlsx');
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const newRow = xlsx.utils.sheet_add_json(worksheet, [data], { skipHeader: true, origin: -1 });
  xlsx.writeFile(workbook, 'D:/Search_by_Excel_file/test.xlsx');
  res.redirect('/'); // Redirect the user back to the home page
});

app.post('/delete-row', (req, res) => {
  const rowIndex = req.body.rowIndex;
  console.log("rowIndex:", rowIndex);
  let workbook = xlsx.readFile('D:/Search_by_Excel_file/test.xlsx');
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // מחיקת השורה מהמערך
  data.splice(rowIndex, 1);

  // כתיבת המערך המעודכן לגיליון האקסל
  const newWorksheet = xlsx.utils.aoa_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

  // שמירת הגיליון המעודכן לקובץ
  xlsx.writeFile(workbook, 'D:/Search_by_Excel_file/test.xlsx');

  res.redirect('/');
});







const port = 3000;
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
