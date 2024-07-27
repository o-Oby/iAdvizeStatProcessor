import express from 'express';
import bodyParser from 'body-parser';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import writeXlsxFile from 'write-excel-file/node';

// Convert `import.meta.url` to a path
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = 3001;

// Set up CORS options
const corsOptions = {
  origin: 'http://127.0.0.1',
  optionsSuccessStatus: 200
};

app.use(cors(corsOptions)); // Enable CORS
app.use(bodyParser.json()); // Parse JSON bodies
app.use(express.static(path.join(__dirname, 'public'))); // Serve static files

// Function to get current month and previous month in French
function getCurrentAndPreviousMonth() {
  const months = [
    "janvier", "février", "mars", "avril", "mai", "juin", "juillet",
    "août", "septembre", "octobre", "novembre", "décembre"
  ];
  const now = new Date();
  const currentMonth = months[now.getMonth()] + " " + now.getFullYear();
  now.setMonth(now.getMonth() - 1);
  const previousMonth = months[now.getMonth()] + " " + now.getFullYear();
  return { currentMonth, previousMonth };
}

// Function to process the input text and extract the required data
function processInputText(inputText) {
  const inputLines = inputText.split('\n'); // Split input text into lines
  const { currentMonth, previousMonth } = getCurrentAndPreviousMonth();

  let data = [];
  let currentData = {};

  inputLines.forEach((line, index) => {
    line = line.trim();
    if (line.match(/janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre|Mois en cours|Mois dernier/)) {
      // Save previous data if exists
      if (Object.keys(currentData).length > 0) {
        data.push(currentData);
      }
      // Replace "Mois en cours" with current month and "Mois dernier" with previous month
      if (line === "Mois en cours") {
        currentData = { mois: currentMonth };
      } else if (line === "Mois dernier") {
        currentData = { mois: previousMonth };
      } else {
        currentData = { mois: line };
      }
    } else if (line.startsWith("Conversations") && !currentData.conversations) {
      currentData.conversations = parseInt(inputLines[index + 1].trim(), 10);
    } else if (line.startsWith("Durée de traitement") && !currentData.duree) {
      currentData.duree = inputLines[index + 1].trim();
    }
  });

  // Push the last data entry if it exists
  if (Object.keys(currentData).length > 0) {
    data.push(currentData);
  }

  return data;
}

// Function to convert duration string to minutes and seconds
function durationToParts(duration) {
  const parts = duration.match(/(\d+)(?=min)|(\d+)(?=sec)/g);
  const minutes = parseInt(parts[0], 10) || 0;
  const seconds = parseInt(parts[1], 10) || 0;
  return { minutes, seconds };
}

// Function to generate the formula for total hours in French format
function generateTotalHoursFormula(conversations, duration) {
  const { minutes, seconds } = durationToParts(duration);
  return `=${conversations}*(${minutes}+${seconds}/60)/60`;
}

// Function to create an Excel file from the extracted data
async function createExcel(data, outputPath) {
  const rows = [
    [
      { value: 'Données iAdvize - Tableau de bord', fontWeight: 'bold', align: 'center', span: 4, fontSize: 14 }
    ],
    [
      { value: 'Conversations', fontWeight: 'bold', align: 'center', backgroundColor: '#D3D3D3', borderStyle: 'medium' },
      { value: 'Mois', fontWeight: 'bold', align: 'center', backgroundColor: '#D3D3D3', borderStyle: 'medium' },
      { value: 'Durée moyenne de traitement (DMT)', fontWeight: 'bold', align: 'center', backgroundColor: '#D3D3D3', borderStyle: 'medium' },
      { value: "Nombre d'heures totales", fontWeight: 'bold', align: 'center', backgroundColor: '#D3D3D3', borderStyle: 'medium' }
    ]
  ];

  let yearlyData = {};

  data.forEach((entry, index) => {
    const totalHoursFormula = generateTotalHoursFormula(entry.conversations, entry.duree);
    const year = entry.mois.split(' ')[1];
    if (!yearlyData[year]) {
      yearlyData[year] = 0;
    }
    yearlyData[year] += entry.conversations * (parseInt(entry.duree.split('min')[0], 10) + parseInt(entry.duree.split('min')[1].split('sec')[0], 10) / 60) / 60;

    rows.push([
      { type: Number, value: entry.conversations, align: 'center', borderStyle: 'thin' },
      { type: String, value: entry.mois, align: 'center', borderStyle: 'thin' },
      { type: String, value: entry.duree, align: 'center', borderStyle: 'thin' },
      { type: 'Formula', value: totalHoursFormula, align: 'center', format: '0.00', borderStyle: 'thin' }
    ]);
  });

  rows.push([]);
  rows.push([
    { value: 'Total Annuel', fontWeight: 'bold', align: 'center', backgroundColor: '#D3D3D3', borderStyle: 'medium' },
    { borderStyle: 'medium' }, // Empty cell with border
    { borderStyle: 'medium' }, // Empty cell with border
    { borderStyle: 'medium' }  // Empty cell with border
  ]);

  for (let year in yearlyData) {
    rows.push([
      { type: String, value: `Total pour ${year}`, align: 'center', fontWeight: 'bold', borderStyle: 'thin' },
      { borderStyle: 'thin' }, // Empty cell with border
      { borderStyle: 'thin' }, // Empty cell with border
      { type: Number, value: parseFloat(yearlyData[year].toFixed(2)), format: '0.00', align: 'center', fontWeight: 'bold', borderStyle: 'thin' }
    ]);
  }

  await writeXlsxFile(rows, {
    filePath: outputPath,
    columns: [
      { width: 30 },
      { width: 30 },
      { width: 40 },
      { width: 30 }
    ]
  });

  console.log(`Le fichier Excel a été créé à l'emplacement ${outputPath}`);
}

// Endpoint to process input text and generate Excel
app.post('/process-iadvize', async (req, res) => {
  const { inputText } = req.body;
  
  if (!inputText) {
    return res.status(400).send('No input text provided');
  }

  const outputPath = path.join(__dirname, 'output.xlsx');

  // Process the input text
  const data = processInputText(inputText);

  // Create the Excel file
  await createExcel(data, outputPath);

  // Send the generated Excel file to the client
  res.download(outputPath, 'output.xlsx', (err) => {
    if (err) {
      console.error(err);
      res.status(500).send('Error generating Excel file');
    }
  });
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on http://0.0.0.0:${port}`);
});
