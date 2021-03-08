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

// Find the right sheet which name starts with "Gesamt"
let filteredData;
for(let key in result){
  if(key.match(re)){
    filteredData = result[key];
  }
}

let Germany

  for(let row in filteredData){
    switch (filteredData[row]['B']){
      case 'Baden-Württemberg': 
      Germany = { "Baden-Württemberg": {
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
      }
        break;
      case 'Bayern': Germany.Bavaria = {
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
      case 'Berlin': Germany.Berlin = {
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
      case 'Brandenburg': Germany.Brandenburg = {
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
    case 'Bremen': Germany.Bremen = {
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
     case 'Hamburg': Germany.Hamburg = {
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
    case 'Hessen': Germany.Hesse = {
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
    case 'Mecklenburg-Vorpommern': Germany.MecklenburgVorpommern = {
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
    case 'Niedersachsen': Germany.LowerSaxony = {
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
    case 'Nordrhein-Westfalen': Germany.NorthRhineWestphalia = {
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
    case 'Rheinland-Pfalz': Germany.RhinelandPalatinate = {
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
    case 'Saarland': Germany.Saarland = {
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
    case 'Sachsen': Germany.Saxony = {
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
    case 'Sachsen-Anhalt': Germany.SaxonyAnhalt = {
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
    case 'Schleswig-Holstein': Germany.SchleswigHolstein = {
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
    case 'Thüringen': Germany.Thuringia = {
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
  }
}

app.get('/vaccinations', (req, res) => {
      res.status(200).json(Germany);
    })

const PORT = process.env.PORT || 4000;

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is listening at port: ${PORT}`);
});
