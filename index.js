const { response } = require('express');
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

// Fetch xlsx-file from rki.de
const url = 'https://www.rki.de/DE/Content/InfAZ/N/Neuartiges_Coronavirus/Daten/Impfquotenmonitoring.xlsx?__blob=publicationFile';
const rkiFile = `./rki-data.xlsx`;
let result;
const re = /Gesamt/;
let Germany;
let dateStatus;

console.log('started');

function download() {
    fetch(url)
    .then(res => {
      const dest = fs.createWriteStream(rkiFile);
      res.body.pipe(dest);
    });
  convertToJson();
}

// Convert xlsx-file to JSON
function convertToJson() {
  console.log('finished downloading!');
  result = excelToJson({
    sourceFile: rkiFile,
    header:{
        rows: 1
    }
  })
  createJsonObject();
}

// Find the right sheet which name starts with "Gesamt"
function createJsonObject () {
  let filteredData;
  for(let key in result){
    if(key.match(re)){
      filteredData = result[key];
    } else if (key === 'Erläuterung'){
      dateStatus = result[key][1]['A'];
    }
  }

Germany = {
  dateStatus: dateStatus,
  fedStates: {} 
}

  for(let row in filteredData){
    switch (filteredData[row]['B']){
      case 'Gesamt': Germany.Total = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break;
      case 'Baden-Württemberg': Germany.fedStates.BadenWürttemberg = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
        }
        break;
      case 'Bayern': Germany.fedStates.Bavaria = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break;
      case 'Berlin': Germany.fedStates.Berlin = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break;
      case 'Brandenburg': Germany.fedStates.Brandenburg = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break;
    case 'Bremen': Germany.fedStates.Bremen = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break;
     case 'Hamburg': Germany.fedStates.Hamburg = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break; 
    case 'Hessen': Germany.fedStates.Hesse = {
        'Total number of vaccine doses administered to date': filteredData[row]['C'],
        'First vaccination Total': filteredData[row]['D'],
        'First vaccination BioNTech': filteredData[row]['E'],
        'First vaccination Moderna': filteredData[row]['F'],
        'First vaccination AstraZeneca': filteredData[row]['G'],
        'First vaccination Difference to previous day': filteredData[row]['H'],
        'First vaccination Vaccination rate': filteredData[row]['I'],
        'Second vaccination Total': filteredData[row]['J'],
        'Second vaccination BioNTech': filteredData[row]['K'],
        'Second vaccination Moderna': filteredData[row]['L'],
        'Second vaccination Difference to previous day': filteredData[row]['M'],
        'Second vaccination Vaccination rate': filteredData[row]['N']
      }
        break;
    case 'Mecklenburg-Vorpommern': Germany.fedStates.MecklenburgVorpommern = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break;
    case 'Niedersachsen': Germany.fedStates.LowerSaxony = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break;
    case 'Nordrhein-Westfalen': Germany.fedStates.NorthRhineWestphalia = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break;
    case 'Rheinland-Pfalz': Germany.fedStates.RhinelandPalatinate = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break;  
    case 'Saarland': Germany.fedStates.Saarland = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break; 
    case 'Sachsen': Germany.fedStates.Saxony = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break;            
    case 'Sachsen-Anhalt': Germany.fedStates.SaxonyAnhalt = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break; 
    case 'Schleswig-Holstein': Germany.fedStates.SchleswigHolstein = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break; 
    case 'Thüringen': Germany.fedStates.Thuringia = {
      'Total number of vaccine doses administered to date': filteredData[row]['C'],
      'First vaccination Total': filteredData[row]['D'],
      'First vaccination BioNTech': filteredData[row]['E'],
      'First vaccination Moderna': filteredData[row]['F'],
      'First vaccination AstraZeneca': filteredData[row]['G'],
      'First vaccination Difference to previous day': filteredData[row]['H'],
      'First vaccination Vaccination rate': filteredData[row]['I'],
      'Second vaccination Total': filteredData[row]['J'],
      'Second vaccination BioNTech': filteredData[row]['K'],
      'Second vaccination Moderna': filteredData[row]['L'],
      'Second vaccination Difference to previous day': filteredData[row]['M'],
      'Second vaccination Vaccination rate': filteredData[row]['N']
    }
      break; 
  }
}
}

 app.get('/vaccinations', (req, res) => {
   download();
   res.status(200).json(Germany);
});

const PORT = process.env.PORT || 4000;

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is listening at port: ${PORT}`);
});
