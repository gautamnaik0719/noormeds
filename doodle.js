import express from "express";
import { google } from "googleapis";
import cors from "cors";

const app = express();
const port = 3000;

// Middleware setup
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors({ origin: "*" }));
app.use(express.static(".")); // Serve static files (e.g., logo) from 'public' folder

// Google Sheets setup
const creds = JSON.parse(process.env.GOOGLE_SHEETS_CREDENTIALS);
const spreadsheetId = process.env.GOOGLE_SHEET_ID;

let auth;
try {
  auth = new google.auth.JWT({
    email: creds.client_email,
    key: creds.private_key.replace(/\\n/g, "\n"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  await auth.authorize();
  console.log("Google Sheets authentication successful");
} catch (error) {
  console.error("Google Sheets authentication failed:", error);
  throw new Error("Authentication setup failed");
}

const sheets = google.sheets({ version: "v4", auth });

// Utility: Get sheet ID by name (robust to invisible/trailing spaces and case)
async function getSheetIdByName(sheetName) {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties",
    });
    console.log("Sheets returned by API:");
    response.data.sheets.forEach((s) => {
      console.log(
        `Title: "${s.properties.title}", ID: ${s.properties.sheetId}`,
      );
    });

    // Log all available sheet names for debugging
    response.data.sheets.forEach((s) => {
      console.log("Found sheet:", `"${s.properties.title}"`);
    });

    // Compare using trim() and toLowerCase() to avoid invisible/trailing space issues
    const sheet = response.data.sheets.find(
      (s) =>
        s.properties.title.trim().toLowerCase() ===
        sheetName.trim().toLowerCase(),
    );
    console.log(sheet);

    if (!sheet) {
      console.error(
        `Sheet with name "${sheetName}" not found. Double-check for extra spaces or invisible characters.`,
      );
    }

    console.log(sheet.properties.sheetId);
    return sheet.properties.sheetId;
  } catch (error) {
    console.error("Error fetching sheet ID:", error);
    return null;
  }
}

// Utility: Fetch sheet data (defaults to columns A:D)
async function getSheetData(sheetName, range = "A:D") {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!${range}`,
    });
    return response.data.values || [];
  } catch (error) {
    console.error(`Error fetching data from ${sheetName}:`, error);
    return [];
  }
}

// Utility: Filter data by medication name (case-insensitive)
async function getFilteredData(searchName) {
  const sheetNames = ["File Meds", "Closet Meds"];
  const allData = [];
  for (const sheetName of sheetNames) {
    const data = await getSheetData(sheetName);
    if (data.length > 0) {
      data.slice(1).forEach((row, index) => {
        if (row.length >= 4) {
          allData.push({
            sheetName,
            rowIndex: index + 2, // +2: skip header and zero-based index
            name: row[0],
            dose: row[1],
            location: row[2],
            quantity: row[3],
          });
        }
      });
    }
  }
  if (!searchName) return [];
  return allData.filter((item) =>
    (item.name || "").toLowerCase().includes((searchName || "").toLowerCase()),
  );
}

// Renders the inventory page with optional search results
function renderInventoryPage({ resultsSection, name, locationOptions }) {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Check out our Medication Inventory!</title>
  <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #F37021;
      --light: #FFF9F5;
      --dark: #333;
      --border: #E8E8E8;
    }
    body {
      font-family: 'Open Sans', sans-serif;
      margin: 0;
      padding: 0;
      background: white;
      color: var(--dark);
    }
    header {
      background: white;
      padding: 1rem 2rem;
      border-bottom: 1px solid var(--border);
      max-width: 1000px;
      margin: 0 auto;
      display: flex;
      align-items: center;
      gap: 1rem;
    }
    .logo {
      height: 60px;
    }
    h1 {
      color: var(--dark);
      font-weight: 600;
      margin: 0;
      font-size: 1.8rem;
    }
    .container {
      max-width: 1000px;
      margin: 0 auto;
      padding: 1rem;
    }
    form {
      background: white;
      padding: 1.5rem;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      margin-bottom: 2rem;
      border: 1px solid var(--border);
    }
    label {
      display: block;
      margin-bottom: 0.5rem;
      font-weight: 600;
    }
    input[type="text"], input[type="number"], select {
      width: 100%;
      padding: 0.5rem;
      margin-bottom: 1rem;
      border: 1px solid var(--border);
      border-radius: 4px;
      font-family: 'Open Sans', sans-serif;
    }
    button {
      background: var(--primary);
      color: white;
      border: none;
      padding: 0.5rem 1.25rem;
      border-radius: 4px;
      cursor: pointer;
      font-family: 'Open Sans', sans-serif;
      font-weight: 600;
      font-size: 1rem;
      transition: background 0.2s;
    }
    button:hover {
      background: #E05A1A;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 2rem;
      background: white;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      border: 1px solid var(--border);
    }
    th, td {
      padding: 0.75rem 1rem;
      text-align: left;
      border-bottom: 1px solid var(--border);
    }
    th {
      background: var(--primary);
      color: white;
      font-weight: 600;
    }
    .qty-btn {
      width: 3.5em;
      text-align: center;
      padding: 0.3rem;
      border: 1px solid var(--border);
      border-radius: 4px;
      font-family: 'Open Sans', sans-serif;
    }
    .qty-controls {
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    .no-results {
      background: white;
      padding: 1rem;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      text-align: center;
      margin-bottom: 2rem;
      border: 1px solid var(--border);
    }
    .divider {
      border-top: 4px solid var(--primary);
      margin: 30rem 0 2rem 0;
    }
    .add-section-title {
      text-align: center;
      color: var(--primary);
      font-size: 1.3rem;
      font-weight: 600;
      margin-bottom: 1rem;
      margin-top: 0.5rem;
    }
    @media (max-width: 768px) {
      header {
        text-align: center;
        flex-direction: column;
      }
      .logo {
        margin-bottom: 1rem;
      }
    }
    footer {
      text-align: center;
      padding: 1.5rem 0;
      background: #f9f9f9;
      color: #666;
      margin-top: 2rem;
      font-size: 0.9rem;
      border-top: 1px solid var(--border);
    }
  </style>
</head>
<body>
  <header>
    <img src="/noor-logo.jpg" alt="SLO Noor Foundation Logo" class="logo">
    <h1>Check out our Medication Inventory!</h1>
  </header>
  <div class="container">
    <form action="/search" method="GET">
      <label>Search Medication Name</label>
      <input type="text" name="name" required value="${name}">
      <button type="submit">Search</button>
    </form>
    ${resultsSection}
    <div class="divider"></div>
    <div class="add-section-title">Add Medication to Our Inventory!</div>
    <form action="/add-medication" method="POST">
      <label>Medication Name</label>
      <input type="text" name="name" required>
      <label>Dose</label>
      <input type="text" name="dose" required>
      <label>Location</label>
      <select name="location" required>
        <option value="">-- Select Location --</option>
        ${locationOptions}
      </select>
      <label>Quantity</label>
      <input type="number" name="quantity" required min="1">
      <button type="submit">Submit</button>
    </form>
  </div>
  <footer>
    <p>© 2025 SLO Noor Foundation. All rights reserved.</p>
  </footer>
</body>
</html>
  `;
}

// Homepage: only show Add Medication form and search bar, not all meds
app.get("/", async (req, res) => {
  const sheetNames = ["File Meds", "Closet Meds"];
  const allLocations = new Set();
  for (const sheetName of sheetNames) {
    const data = await getSheetData(sheetName);
    if (data.length > 0) {
      data.slice(1).forEach((row) => {
        if (row[2]) allLocations.add(row[2]);
      });
    }
  }
  let locationOptions = Array.from(allLocations)
    .map((loc) => `<option value="${loc}">${loc}</option>`)
    .join("");
  if (!locationOptions) {
    locationOptions = `<option value="">No locations available</option>`;
  }

  // No resultsSection by default
  res.send(
    renderInventoryPage({ resultsSection: "", name: "", locationOptions }),
  );
});

// Search route: show only relevant medications
app.get("/search", async (req, res) => {
  const { name } = req.query;
  const sheetNames = ["File Meds", "Closet Meds"];
  const allLocations = new Set();
  for (const sheetName of sheetNames) {
    const data = await getSheetData(sheetName);
    if (data.length > 0) {
      data.slice(1).forEach((row) => {
        if (row[2]) allLocations.add(row[2]);
      });
    }
  }
  let locationOptions = Array.from(allLocations)
    .map((loc) => `<option value="${loc}">${loc}</option>`)
    .join("");
  if (!locationOptions) {
    locationOptions = `<option value="">No locations available</option>`;
  }

  const data = await getFilteredData(name);
  let resultsSection = "";
  if (data.length > 0) {
    resultsSection = `
      <form id="updateForm" action="/update" method="POST">
        <table>
          <tr><th>Name</th><th>Dose</th><th>Location</th><th>Quantity</th><th>Amount Used</th></tr>
          ${data
            .map(
              (item, i) => `
            <tr>
              <td>${item.name}</td>
              <td>${item.dose}</td>
              <td>${item.location}</td>
              <td>${item.quantity}</td>
              <td>
                <div class="qty-controls">
                  <button type="button" onclick="decQty(${i})">-</button>
                  <input type="number" min="0" max="${item.quantity}" value="0" id="qty${i}" class="qty-btn" onchange="updateQty(${i})">
                  <button type="button" onclick="incQty(${i})">+</button>
                  <input type="hidden" name="items[${i}][sheetName]" value="${item.sheetName}">
                  <input type="hidden" name="items[${i}][rowIndex]" value="${item.rowIndex}">
                  <input type="hidden" name="items[${i}][name]" value="${item.name}">
                  <input type="hidden" name="items[${i}][quantity]" value="${item.quantity}">
                  <input type="hidden" name="items[${i}][qty]" id="qtyHidden${i}" value="0">
                </div>
              </td>
            </tr>
          `,
            )
            .join("")}
        </table>
        <button type="submit">Submit</button>
      </form>
      <script>
        function updateQty(index) {
          let el = document.getElementById('qty' + index);
          let hidden = document.getElementById('qtyHidden' + index);
          hidden.value = el.value;
        }
        function decQty(index) {
          let el = document.getElementById('qty' + index);
          if (el.value > 0) el.value--;
          updateQty(index);
        }
        function incQty(index) {
          let el = document.getElementById('qty' + index);
          if (el.value < parseInt(el.max)) el.value++;
          updateQty(index);
        }
      </script>
    `;
  } else {
    resultsSection = `<div class="no-results"><p>No results found for "${name}".</p></div>`;
  }

  res.send(renderInventoryPage({ resultsSection, name, locationOptions }));
});

// Update medication quantities and delete row if quantity reaches zero
app.post("/update", async (req, res) => {
  const { items } = req.body;
  if (!items || !Array.isArray(items)) {
    console.log("No items or invalid items array in request");
    return res.status(400).send("No items to update");
  }
  for (const item of items) {
    if (!item.sheetName || !item.rowIndex || item.qty === undefined) {
      console.log("Missing required item fields:", item);
      continue;
    }
    const qtyToTake = parseInt(item.qty) || 0;
    const currentQty = parseInt(item.quantity) || 0;
    const newQty = Math.max(0, currentQty - qtyToTake);

    console.log(
      `Updating item: ${item.name}, current qty: ${currentQty}, qty to take: ${qtyToTake}, new qty: ${newQty}`,
    );

    if (newQty <= 0) {
      try {
        const sheetId = await getSheetIdByName(item.sheetName);
        console.log(sheetId);
        if (sheetId === undefined || sheetId === null) {
          console.log(`Sheet ID not found for: ${item.sheetName}`);
          continue;
        }
        console.log(
          `Deleting row ${item.rowIndex} from ${item.sheetName} (sheetId: ${sheetId})`,
        );
        // Adjust for zero-based index and header row
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: {
            requests: [
              {
                deleteDimension: {
                  range: {
                    sheetId,
                    dimension: "ROWS",
                    startIndex: item.rowIndex - 1, // header is row 1
                    endIndex: item.rowIndex,
                  },
                },
              },
            ],
          },
        });
      } catch (error) {
        console.error("Error deleting row:", error);
      }
    } else {
      try {
        console.log(
          `Updating quantity to ${newQty} for ${item.name} in ${item.sheetName}, row ${item.rowIndex}`,
        );
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${item.sheetName}!D${item.rowIndex}`,
          valueInputOption: "RAW",
          resource: { values: [[newQty]] },
        });
      } catch (error) {
        console.error("Error updating row:", error);
      }
    }
  }
  res.redirect("/");
});

// Add new medication or update existing
app.post("/add-medication", async (req, res) => {
  const { name, dose, location, quantity } = req.body;
  const sheetNames = ["File Meds", "Closet Meds"];
  let found = false;
  for (const sheetName of sheetNames) {
    const data = await getSheetData(sheetName);
    if (data.length > 0) {
      const rowIndex = data
        .slice(1)
        .findIndex(
          (row) =>
            (row[0] || "").toLowerCase() === (name || "").toLowerCase() &&
            (row[1] || "").toLowerCase() === (dose || "").toLowerCase() &&
            (row[2] || "").toLowerCase() === (location || "").toLowerCase(),
        );

      if (rowIndex !== -1) {
        const actualRowIndex = rowIndex + 2;
        const currentQty = parseInt(data[rowIndex + 1][3]) || 0;
        const newQty = currentQty + parseInt(quantity);
        console.log(
          `Adding to existing: ${name} ${dose} at ${location} in ${sheetName}. Old qty: ${currentQty}, add: ${quantity}, new qty: ${newQty}`,
        );
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!D${actualRowIndex}`,
          valueInputOption: "RAW",
          resource: { values: [[newQty]] },
        });
        found = true;
        break;
      }
    }
  }
  if (!found) {
    console.log(
      `Adding NEW: ${name} ${dose} at ${location} qty ${quantity} to File Meds`,
    );
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "File Meds!A:D",
      valueInputOption: "RAW",
      resource: { values: [[name, dose, location, quantity]] },
    });
  }
  res.redirect("/");
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
