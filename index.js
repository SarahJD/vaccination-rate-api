const express = require('express'),
morgan = require('morgan'), // logging middleware
fs = require('fs'),
path = require('path'),
fetch = require('node-fetch'),
excelToJson = require('convert-excel-to-json');

const app = express();

// Log to the file access.log
const accessLogStream = fs.createWriteStream(path.join(__dirname, 'access.log'), { flags: 'a' })
app.use(morgan('common', { stream: accessLogStream }));

// // for error handling
// app.use((err, req, res, next) => {
//   console.error(err.stack);
//   res.status(500.send('something broke!'));
// });


// Fetch xlsx-file from rki.de
  const url = 'https://www.rki.de/DE/Content/InfAZ/N/Neuartiges_Coronavirus/Daten/Impfquotenmonitoring.xlsx?__blob=publicationFile';
  const rkiFile = `./rki-data.xlsx`;
  const re = /Gesamt/;

  async function download() {
    const response = await fetch(url);
    const buffer = await response.buffer();
    fs.writeFile(rkiFile, buffer, () => 
      console.log('finished downloading!'));
  }

  download();

// Convert xlsx-file to JSON
  const result = excelToJson({
    sourceFile: rkiFile,
    header:{
        rows: 1
    }
  });

  let key;
  for(key in result){
    const found = key.match(re);
    console.log(found);
  }
  

  

// // Endpoint 
// app.get('', (req, res) => {
//   Movies.find()
//     .then((movies) => {
//       res.status(201).json(movies);
//     })
//     .catch((err) => {
//       console.error(err);
//       res.status(500).send('Error: ' + err);
//     });
// });


const PORT = 4000;

app.listen(PORT, () => {
  console.log(`Server is listening at port: ${PORT}`);
});