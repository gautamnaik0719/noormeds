import express from "express";
import { google } from "googleapis";
import cors from "cors";

const app = express();
const port = 3000;

function normalizeDose(dose) {
  return (dose || "").replace(/\s+/g, "").toLowerCase();
}

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
    const sheet = response.data.sheets.find(
      (s) =>
        s.properties.title.trim().toLowerCase() ===
        sheetName.trim().toLowerCase(),
    );
    if (!sheet) {
      console.error(`Sheet "${sheetName}" not found.`);
      return null;
    }
    return sheet.properties.sheetId;
  } catch (error) {
    console.error("Error fetching sheet ID:", error);
    return null;
  }
}

// Log Activity function (unchanged)
async function logActivity({ action, name, dose, location, quantity }) {
  const date = new Date();
  const formatted = date.toLocaleString("en-US", {
    timeZone: "America/Los_Angeles",
  });
  const sheetId = await getSheetIdByName("Activity Records");
  if (sheetId == null)
    throw new Error('Could not find "Activity Records" sheet');

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [
        {
          insertDimension: {
            range: { sheetId, dimension: "ROWS", startIndex: 1, endIndex: 2 },
            inheritFromBefore: false,
          },
        },
      ],
    },
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: "Activity Records!A2:F2",
    valueInputOption: "RAW",
    resource: { values: [[formatted, action, name, dose, location, quantity]] },
  });
}

// Utility: Append to Past Medication tab
async function addToPastMedication({ name, dose, location }) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Past Medication!A:C",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      resource: {
        values: [[name, dose, location]],
      },
    });
  } catch (error) {
    console.error("Error appending to Past Medication:", error);
  }
}

// Utility: Remove medication from Past Medication by name, dose, location
async function removeFromPastMedication({ name, dose, location }) {
  try {
    // Fetch all Past Medication rows
    const data = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Past Medication!A:C",
    });
    const rows = data.data.values || [];
    if (rows.length === 0) return;

    // Find matching rows to remove (return indices)
    const rowsToRemove = [];
    rows.forEach((row, i) => {
      if (
        (row[0] || "").toLowerCase() === name.toLowerCase() &&
        (row[1] || "").toLowerCase() === dose.toLowerCase() &&
        (row[2] || "").toLowerCase() === location.toLowerCase()
      ) {
        // i is zero indexed, header row is row 1, so actual sheet row = i+1
        rowsToRemove.push(i + 1);
      }
    });

    if (rowsToRemove.length === 0) return;

    // Sheets batchUpdate DELETE requests need ranges descending order to avoid shift issues
    rowsToRemove.sort((a, b) => b - a);

    const sheetId = await getSheetIdByName("Past Medication");
    if (sheetId === null) {
      console.error('Could not find "Past Medication" sheet.');
      return;
    }

    for (const rowIndex of rowsToRemove) {
      // Delete each matching row (rowIndex - 1 for zero-based index in batchUpdate)
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId,
                  dimension: "ROWS",
                  startIndex: rowIndex - 1, // zero based
                  endIndex: rowIndex,
                },
              },
            },
          ],
        },
      });
    }
  } catch (error) {
    console.error("Error removing from Past Medication:", error);
  }
}

// Utility: Fetch sheet data (default columns A:D)
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

// Sort a given sheet by Location (column C) A→Z
async function sortSheetByLocation(sheetName) {
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties",
    });
    const sheet = resp.data.sheets.find(
      (s) =>
        s.properties.title.trim().toLowerCase() ===
        sheetName.trim().toLowerCase(),
    );
    if (!sheet) {
      console.error(`No sheet ID found for '${sheetName}'`);
      return;
    }

    const sheetId = sheet.properties.sheetId;
    const totalRows = sheet.properties.gridProperties.rowCount;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: {
        requests: [
          {
            sortRange: {
              range: {
                sheetId,
                startRowIndex: 1, // skip header row
                endRowIndex: totalRows, // go to bottom of sheet
                startColumnIndex: 0,
                endColumnIndex: 4, // columns A–D
              },
              sortSpecs: [
                { dimensionIndex: 2, sortOrder: "ASCENDING" }, // Location
                { dimensionIndex: 0, sortOrder: "ASCENDING" }, // Name
              ],
            },
          },
        ],
      },
    });
    console.log(`Sorted '${sheetName}' by Location`);
  } catch (err) {
    console.error(`Error sorting ${sheetName}:`, err);
  }
}

// NEW: Fetch ordered location list from Location Catalog sheet (column A)
async function getLocationCatalogOrder() {
  try {
    const data = await getSheetData("Location Catalog", "A:A");
    // Skip header, filter blanks
    return data
      .slice(1)
      .map((r) => r[0])
      .filter(Boolean);
  } catch (err) {
    console.error("Error fetching Location Catalog:", err);
    return [];
  }
}

// Utility: Combine meds from File Meds and Closet Meds for search/filter
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
            rowIndex: index + 2, // +2 to skip header and zero-based
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
    (item.name || "").toLowerCase().includes(searchName.toLowerCase()),
  );
}

function isAngieMode(name) {
  return name.startsWith("sparkles++");
}
function stripPlusPlus(name) {
  return name.startsWith("sparkles++")
    ? name.substring("sparkles++".length)
    : name;
}

// Fetch meds from Angie Stash
async function getAngiesStashFiltered(searchName) {
  const data = await getSheetData("Angie Stash", "A:C"); // Name, Dose, Quantity
  if (data.length < 2) return [];
  return data
    .slice(1)
    .filter((row) =>
      (row[0] || "").toLowerCase().includes(searchName.toLowerCase()),
    )
    .map((row, idx) => ({
      sheetName: "Angie Stash",
      rowIndex: idx + 2,
      name: row[0] || "",
      dose: row[1] || "",
      location: "Angie Stash",
      quantity: row[2] || "0",
    }));
}

// Utility: Fetch meds from Past Medication matching search (for Add Med search)
async function getPastMedicationFiltered(searchName, angieMode = false) {
  const data = await getSheetData("Past Medication", "A:C"); // Name, Dose, Location
  if (data.length < 2) return [];
  return data
    .slice(1)
    .filter((row) => {
      const matchesName = (row[0] || "")
        .toLowerCase()
        .includes(searchName.toLowerCase());
      if (!matchesName) return false;
      const location = (row[2] || "").trim().toLowerCase();
      if (angieMode) {
        // only Angie Stash rows
        return location === "angie stash";
      } else {
        // exclude Angie Stash rows
        return location !== "angie stash";
      }
    })
    .map((row, idx) => ({
      rowIndex: idx + 2,
      name: row[0] || "",
      dose: row[1] || "",
      location: row[2] || "",
      quantity: "0",
    }));
}

// Render page - full HTML (includes Add Medication search showing past meds with location dropdown)
function renderInventoryPage({
  resultsSection,
  name,
  locationOptions,
  quickAddResultsSection = "",
}) {
  // Helper: Render options for location dropdown with selected one marked
  function renderLocationOptions(selectedLocation) {
    return locationOptions
      .split("</option>")
      .filter(Boolean)
      .map((optionHtml) => {
        // Parse value attribute in option
        const valueMatch = optionHtml.match(/value="([^"]*)"/);
        if (!valueMatch) return optionHtml + "</option>";
        const val = valueMatch[1];
        if (val === selectedLocation) {
          // Add selected attribute if matches
          if (optionHtml.includes(" selected")) return optionHtml + "</option>";
          return optionHtml.replace(/>/, ' selected="selected">') + "</option>";
        }
        return optionHtml + "</option>";
      })
      .join("");
  }

  return `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Check out our Medication Inventory!</title>
<link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600&display=swap" rel="stylesheet" />
<style>
  :root {
    --primary: #F37021; /* Orange */
    --add-primary: #28a745; /* Green */
    --light: #FFF9F5;
    --dark: #333;
    --border: #E8E8E8;
  }
  body {
    font-family: 'Open Sans', sans-serif;
    margin: 0; padding: 0;
    background: white;
    color: var(--dark);
  }
  header {
    background: white;
    padding: 1rem 2rem;
    border-bottom: 1px solid var(--border);
    max-width: 1000px;
    margin: auto;
    display: flex;
    align-items: center;
    gap: 1rem;
  }
  .logo { height: 60px; }
  h1 { margin: 0; color: var(--dark); font-weight: 600; font-size: 1.8rem; }
  .container { max-width: 1000px; margin: auto; padding: 1rem; }
  form {
    background: white;
    padding: 1.5rem;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    margin-bottom: 2rem;
    border: 1px solid var(--border);
  }
  label { display: block; margin-bottom: 0.5rem; font-weight: 600; }
  input[type="text"], input[type="number"], select {
    width: 100%; padding: 0.5rem; margin-bottom: 1rem;
    border: 1px solid var(--border); border-radius: 4px;
    font-family: 'Open Sans', sans-serif;
  }
  button {
    background: var(--primary);
    color: white;
    border: none; padding: 0.5rem 1.25rem;
    border-radius: 4px; cursor: pointer;
    font-family: 'Open Sans', sans-serif; font-weight: 600;
    font-size: 1rem; transition: background 0.2s;
  }
  button:hover { background: #E05A1A; }

  /* Orange top table */
  .top-table th {
    background: var(--primary);
    color: white;
    font-weight: 600;
  }

  /* Green Add CURRENT Medication table */
  .add-table th {
    background: var(--add-primary);
    color: white;
    font-weight: 600;
  }
  .add-table {
    border-color: var(--add-primary);
  }
  /* Green +/- buttons inside Add CURRENT Medication table */
  .add-table .qty-controls button {
    background: var(--add-primary);
    color: white;
    border: none;
  }
  .add-table .qty-controls button:hover {
    background: #218838; /* darker green on hover */
  }

  /* Green "Add Quantity" submit button under Add CURRENT Medication table */
  .add-table + button,
  .add-current-med button[type="submit"] {
    background: var(--add-primary);
    color: white;
  }
  .add-table + button:hover,
  .add-current-med button[type="submit"]:hover {
    background: #218838;
  }

  /* New button style for Add CURRENT Medication section */
  form.add-current-med .add-search-btn {
    background: var(--add-primary);
  }
  form.add-current-med .add-search-btn:hover {
    background: #218838;
  }
  form.add-current-med label {
    color: var(--add-primary);
    font-weight: 700;
    font-size: 1.1rem;
    margin-bottom: 0.7rem;
  }

  table {
    width: 100%; border-collapse: collapse; margin-bottom: 2rem;
    background: white; border-radius: 8px; overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05); border: 1px solid var(--border);
  }
  th, td {
    padding: 0.75rem 1rem; text-align: left;
    border-bottom: 1px solid var(--border); vertical-align: middle;
  }
  .qty-btn {
    width: 3.5em; text-align: center; padding: 0.3rem;
    border: 1px solid var(--border); border-radius: 4px;
  }
  .qty-controls { display: flex; gap: 0.5rem; align-items: center; }
  .no-results {
    background: white; padding: 1rem; border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05); text-align: center;
    margin-bottom: 2rem; border: 1px solid var(--border);
  }
  .divider { border-top: 4px solid var(--primary); margin: 2rem 0; }
  .add-section-title {
    text-align: center; font-size: 1.3rem; font-weight: 600;
    margin-bottom: 1rem; margin-top: 0.5rem;
  }
  .subsection-title {
    font-weight: 700; font-size: 1.2rem; margin-bottom: 1rem;
    color: var(--add-primary); border-bottom: 2px solid var(--add-primary);
    padding-bottom: 0.25rem;
  }
  @media (max-width: 768px) {
    header { text-align: center; flex-direction: column; }
    .logo { margin-bottom: 1rem; }
    table, form { font-size: 0.9rem; }
  }
  footer {
    text-align: center; padding: 1.5rem 0; background: #f9f9f9;
    color: #666; margin-top: 2rem; font-size: 0.9rem;
    border-top: 1px solid var(--border);
  }
</style>
</head>
<body>
  <header>
    <img src="/noor-logo.jpg" alt="SLO Noor Foundation Logo" class="logo" />
    <h1>Check out our Medication Inventory!</h1>
  </header>

  <div class="container">
    <!-- Top search -->
    <form action="/search" method="GET">
      <label>Search Medication Name</label>
      <input type="text" name="name" required value="${name}" />
      <button type="submit">Search</button>
    </form>

    <!-- Example top table wrapper -->
    <div class="top-table-wrapper">
      ${resultsSection.replace("<table", '<table class="top-table"')}
    </div>

    <div style="margin-top: 23rem;">
      <div class="divider"></div>
      <div class="add-section-title">Add Medication to Our Inventory!</div>

      <div class="subsection-title">Add CURRENT Medication</div>
      <form action="/quick-add" method="GET" class="add-current-med">
        <label>Search Medication Name</label>
        <input type="text" name="name" id="quickAddNameInput"
               list="quickAddNamesList" required autocomplete="off" />
        <datalist id="quickAddNamesList"></datalist>
        <button type="submit" class="add-search-btn">Search</button>
      </form>

      <div id="quickAddResults">
        ${quickAddResultsSection.replace("<table", '<table class="add-table"')}
      </div>

      <div class="subsection-title" style="color: var(--primary); border-bottom-color: var(--primary); margin-top: 2rem;">
        Add NEW Medication
      </div>
      <form action="/add-medication" method="POST" class="add-new-med">
        <label>Medication Name</label>
        <input type="text" name="name" id="medNameInput"
               list="medNamesList" required autocomplete="off" />
        <datalist id="medNamesList"></datalist>

        <label>Dose</label>
        <input type="text" name="dose" id="doseInput" list="doseList" required />
        <datalist id="doseList"></datalist>

        <label>Location</label>
        <select name="location" required>
          <option value="">-- Select Location --</option>
          ${locationOptions}
        </select>

        <label>Quantity</label>
        <input type="number" name="quantity" required min="1" />
        <button type="submit">Submit</button>
      </form>
    </div>
  </div>

  <footer>
    <p>© 2025 SLO Noor Foundation. All rights reserved.</p>
  </footer>

<script>
document.addEventListener("DOMContentLoaded", function () {
  fetch("/all-med-names")
    .then(res => res.json())
    .then(names => {
      let datalist = document.getElementById("medNamesList");
      datalist.innerHTML = "";
      names.forEach(name => {
        const opt = document.createElement("option");
        opt.value = name;
        datalist.appendChild(opt);
      });
      let quickAddList = document.getElementById("quickAddNamesList");
      if (quickAddList) {
        quickAddList.innerHTML = "";
        names.forEach(name => {
          const opt = document.createElement("option");
          opt.value = name;
          quickAddList.appendChild(opt);
        });
      }
    });

  const medNameInput = document.getElementById("medNameInput");
  const doseList = document.getElementById("doseList");
  function fetchDoseSuggestions() {
    const medName = medNameInput.value.trim();
    doseList.innerHTML = "";
    if (medName !== "") {
      fetch("/doses-for-name?name=" + encodeURIComponent(medName))
        .then(res => res.json())
        .then(doses => {
          doseList.innerHTML = "";
          doses.forEach(dose => {
            const opt = document.createElement("option");
            opt.value = dose;
            doseList.appendChild(opt);
          });
        });
    }
  }
  medNameInput.addEventListener("input", fetchDoseSuggestions);
  medNameInput.addEventListener("change", fetchDoseSuggestions);

  const quickAddResults = document.getElementById("quickAddResults");
  if (quickAddResults && quickAddResults.innerHTML.trim() !== "") {
    quickAddResults.scrollIntoView({ behavior: "smooth" });
  }
});
</script>
</body>
</html>
  `;
}

// Helper function to render location options with selected attribute - used INSIDE renderInventoryPage
function renderLocationOptions(selectedLocation) {
  // Must match options from ${locationOptions} in server (called during HTML construction)
  const allLocations = selectedLocation ? [selectedLocation] : [];
  // This function should really be consistent with locationOptions at server side
  // Here we return a minimal example; you can customize to match exactly your locationOptions
  const exampleLocations = [
    "Cabinet 1",
    "Shelf A",
    "Fridge",
    "Locker",
    "File Meds",
    "Closet Meds",
  ];
  // Unique, add selectedLocation if missing
  if (selectedLocation && !exampleLocations.includes(selectedLocation)) {
    exampleLocations.push(selectedLocation);
  }
  return exampleLocations
    .map(
      (loc) =>
        `<option value="${loc}"${loc === selectedLocation ? ' selected="selected"' : ""}>${loc}</option>`,
    )
    .join("");
}

// Homepage route
app.get("/", async (req, res) => {
  const orderedLocations = await getLocationCatalogOrder();
  let locationOptions = orderedLocations
    .map((loc) => `<option value="${loc}">${loc}</option>`)
    .join("");
  if (!locationOptions)
    locationOptions = `<option value="">No locations available</option>`;
  res.send(
    renderInventoryPage({ resultsSection: "", name: "", locationOptions }),
  );
});

// Search route — excludes Past Medication
app.get("/search", async (req, res) => {
  let { name } = req.query;
  if (!name) return res.redirect("/");

  const angieMode = isAngieMode(name);
  const cleanName = stripPlusPlus(name);

  const orderedLocations = await getLocationCatalogOrder();
  let locationOptions = orderedLocations
    .map((loc) => `<option value="${loc}">${loc}</option>`)
    .join("");
  if (!locationOptions)
    locationOptions = `<option value="">No locations available</option>`;

  let data = angieMode
    ? await getAngiesStashFiltered(cleanName)
    : await getFilteredData(cleanName);

  // pass cleanName to renderInventoryPage so "sparkles++" doesn’t show in the input

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
                  <input type="number" min="0" max="${item.quantity}" value="0" id="qty${i}" class="qty-btn" onchange="updateQty(${i})" />
                  <button type="button" onclick="incQty(${i})">+</button>

                  <input type="hidden" name="items[${i}][sheetName]" value="${item.sheetName}" />
                  <input type="hidden" name="items[${i}][rowIndex]" value="${item.rowIndex}" />
                  <input type="hidden" name="items[${i}][name]" value="${item.name}" />
                  <input type="hidden" name="items[${i}][quantity]" value="${item.quantity}" />
                  <input type="hidden" name="items[${i}][qty]" id="qtyHidden${i}" value="0" />
                  <input type="hidden" name="items[${i}][dose]" value="${item.dose}" />
                  <input type="hidden" name="items[${i}][location]" value="${item.location}" />
                </div>
              </td>
            </tr>`,
            )
            .join("")}
        </table>
        <button type="submit">Submit</button>
      </form>
      <script>
        function updateQty(index) {
          let el = document.getElementById("qty" + index);
          let hidden = document.getElementById("qtyHidden" + index);
          hidden.value = el.value;
        }
        function decQty(index) {
          let el = document.getElementById("qty" + index);
          if (el.value > 0) el.value--;
          updateQty(index);
        }
        function incQty(index) {
          let el = document.getElementById("qty" + index);
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

// Quick Add (Add Medication search) includes Past Medication meds with location dropdown
app.get("/quick-add", async (req, res) => {
  const { name } = req.query;
  if (!name) return res.redirect("/");

  // Detect Angie mode and strip the sparkles++ key from the search string
  const angieMode = isAngieMode(name);
  const cleanName = stripPlusPlus(name);

  // Build location dropdown from Location Catalog
  const orderedLocations = await getLocationCatalogOrder();
  let locationOptions = orderedLocations
    .map((loc) => `<option value="${loc}">${loc}</option>`)
    .join("");
  if (!locationOptions) {
    locationOptions = `<option value="">No locations available</option>`;
  }

  let currentData = [];
  let pastData = [];

  if (angieMode) {
    // Current meds = ONLY Angie’s Stash
    currentData = await getAngiesStashFiltered(cleanName);
    // Past meds = only those in Past Medication with location === "Angie Stash"
    pastData = await getPastMedicationFiltered(cleanName, true);
  } else {
    // Current meds = File Meds + Closet Meds
    currentData = await getFilteredData(cleanName);
    // Past meds = only those in Past Medication with location !== "Angie Stash"
    pastData = await getPastMedicationFiltered(cleanName, false);
  }

  // Build results table with both current + past meds
  let quickAddResultsSection = "";
  if (currentData.length > 0 || pastData.length > 0) {
    quickAddResultsSection = `
      <form id="quickAddForm" action="/quick-add-update" method="POST">
      <table>
        <tr>
          <th>Name</th><th>Dose</th><th>Location</th><th>Quantity</th><th>Amount to Add</th>
        </tr>
        ${currentData
          .map(
            (item, i) => `
            <tr>
              <td>${item.name}</td>
              <td>${item.dose}</td>
              <td>${item.location}</td>
              <td>${item.quantity}</td>
              <td>
                <div class="qty-controls">
                  <button type="button" onclick="quickDecQty(${i})">-</button>
                  <input type="number" min="0" value="0" id="addQty${i}" class="qty-btn" onchange="quickUpdateQty(${i})" />
                  <button type="button" onclick="quickIncQty(${i})">+</button>
                  <input type="hidden" name="items[${i}][sheetName]" value="${item.sheetName}" />
                  <input type="hidden" name="items[${i}][rowIndex]" value="${item.rowIndex}" />
                  <input type="hidden" name="items[${i}][name]" value="${item.name}" />
                  <input type="hidden" name="items[${i}][dose]" value="${item.dose}" />
                  <input type="hidden" name="items[${i}][location]" value="${item.location}" />
                  <input type="hidden" name="items[${i}][quantity]" value="${item.quantity}" />
                  <input type="hidden" name="items[${i}][addQty]" id="addQtyHidden${i}" value="0" />
                  <input type="hidden" name="items[${i}][originalLocation]" value="${item.location}" />
                </div>
              </td>
            </tr>
          `,
          )
          .join("")}
        ${pastData
          .map((item, idx) => {
            const index = currentData.length + idx; // continue index count
            const locOptions = orderedLocations
              .map(
                (loc) =>
                  `<option value="${loc}"${
                    loc === item.location ? " selected" : ""
                  }>${loc}</option>`,
              )
              .join("");
            return `
            <tr>
              <td>${item.name}</td>
              <td>${item.dose}</td>
              <td>
                <select name="items[${index}][location]" >
                  ${locOptions}
                </select>
              </td>
              <td>0</td>
              <td>
                <div class="qty-controls">
                  <button type="button" onclick="quickDecQty(${index})">-</button>
                  <input type="number" min="0" value="0" id="addQty${index}" class="qty-btn" onchange="quickUpdateQty(${index})" />
                  <button type="button" onclick="quickIncQty(${index})">+</button>
                  <input type="hidden" name="items[${index}][sheetName]" value="Past Medication" />
                  <input type="hidden" name="items[${index}][rowIndex]" value="${item.rowIndex}" />
                  <input type="hidden" name="items[${index}][name]" value="${item.name}" />
                  <input type="hidden" name="items[${index}][dose]" value="${item.dose}" />
                  <input type="hidden" name="items[${index}][quantity]" value="0" />
                  <input type="hidden" name="items[${index}][addQty]" id="addQtyHidden${index}" value="0" />
                  <input type="hidden" name="items[${index}][originalLocation]" value="${item.location}" />
                </div>
              </td>
            </tr>`;
          })
          .join("")}
      </table>
      <button type="submit">Add Quantity</button>
      </form>
      <script>
        function quickUpdateQty(index) {
          let el = document.getElementById("addQty" + index);
          let hidden = document.getElementById("addQtyHidden" + index);
          hidden.value = el.value;
        }
        function quickDecQty(index) {
          let el = document.getElementById("addQty" + index);
          if (el.value > 0) el.value--;
          quickUpdateQty(index);
        }
        function quickIncQty(index) {
          let el = document.getElementById("addQty" + index);
          el.value++;
          quickUpdateQty(index);
        }
      </script>
    `;
  } else {
    quickAddResultsSection = `<div class="no-results"><p>No results found for "${name}".</p></div>`;
  }

  res.send(
    renderInventoryPage({
      resultsSection: "",
      name, // keep original in the search bar
      locationOptions,
      quickAddResultsSection,
    }),
  );
});

// Quick Add POST handler (updated to handle Past Medication removal if qty added)
app.post("/quick-add-update", async (req, res) => {
  const { items } = req.body;
  console.log(JSON.stringify(req.body, null, 2));
  if (!items || !Array.isArray(items)) return res.redirect("/");

  for (const idx in items) {
    const item = items[idx];
    if (!item.name || !item.addQty) continue;

    const addQty = parseInt(item.addQty) || 0;
    if (addQty <= 0) continue;

    const angieMode = isAngieMode(item.name);
    const cleanName = stripPlusPlus(item.name);

    // 🚀 Direct Angie mode (typed sparkles++)
    if (angieMode) {
      const angieData = await getAngiesStashFiltered(cleanName);
      const existingMed = angieData.find(
        (med) =>
          med.name.toLowerCase() === cleanName.toLowerCase() &&
          med.dose.toLowerCase() === (item.dose || "").toLowerCase(),
      );

      if (existingMed) {
        const currentQty = parseInt(existingMed.quantity) || 0;
        const newQty = currentQty + addQty;
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `Angie Stash!C${existingMed.rowIndex}`, // ✅ Quantity column C
          valueInputOption: "RAW",
          resource: { values: [[newQty]] },
        });
      } else {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: "Angie Stash!A:C",
          valueInputOption: "RAW",
          resource: { values: [[cleanName, item.dose, addQty]] },
        });
      }

      await logActivity({
        action: "ADD",
        name: cleanName,
        dose: item.dose,
        location: "Angie Stash",
        quantity: addQty,
      });

      continue;
    }

    // 🗂 Past Medication item
    if (item.sheetName === "Past Medication") {
      // Detect if original location is Angie Stash
      console.log(item.originalLocation);
      const isAngiePast =
        (item.originalLocation || "").trim().toLowerCase() === "angie stash";

      // Remove the row from Past Medication immediately
      await removeFromPastMedication({
        name: item.name,
        dose: item.dose,
        location: item.originalLocation,
      });

      console.log(isAngiePast);
      if (isAngiePast) {
        console.log("Reached");
        // Add directly to Angie Stash, ignoring location input
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: "Angie Stash!A:C",
          valueInputOption: "RAW",
          resource: {
            values: [[item.name, item.dose, addQty]],
          },
        });
        await logActivity({
          action: "ADD",
          name: item.name,
          dose: item.dose,
          location: "Angie Stash",
          quantity: addQty,
        });
      } else {
        const location = item.location || "";
        // Normal behavior: add to File Meds or Closet Meds depending on location input
        const targetSheet = location.toLowerCase().includes("closet")
          ? "Closet Meds"
          : "File Meds";

        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${targetSheet}!A:D`,
          valueInputOption: "RAW",
          resource: {
            values: [[item.name, item.dose, location, addQty]],
          },
        });
        await sortSheetByLocation(targetSheet);
        await logActivity({
          action: "ADD",
          name: item.name,
          dose: item.dose,
          location,
          quantity: addQty,
        });
      }
      continue;
    }
    // 📊 Normal current stock update
    const currentQty = parseInt(item.quantity) || 0;
    const newQty = currentQty + addQty;

    if (item.sheetName === "Angie Stash") {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `Angie Stash!C${item.rowIndex}`, // ✅ Quantity column C
        valueInputOption: "RAW",
        resource: { values: [[newQty]] },
      });
    } else {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${item.sheetName}!D${item.rowIndex}`, // Normal: Quantity column D
        valueInputOption: "RAW",
        resource: { values: [[newQty]] },
      });
    }

    await logActivity({
      action: "ADD",
      name: item.name,
      dose: item.dose || "",
      location: item.location || "",
      quantity: addQty,
    });
  }

  res.redirect("/");
});

// Remove/Use Medications - unchanged
app.post("/update", async (req, res) => {
  const { items } = req.body;
  if (!items || !Array.isArray(items))
    return res.status(400).send("No items to update");
  for (const item of items) {
    if (!item.sheetName || !item.rowIndex || item.qty === undefined) continue;
    const qtyToTake = parseInt(item.qty) || 0;
    if (!qtyToTake) continue;
    const currentQty = parseInt(item.quantity) || 0;
    const newQty = Math.max(0, currentQty - qtyToTake);

    await logActivity({
      action: "REMOVE",
      name: item.name,
      dose: item.dose || "",
      location: item.location || "",
      quantity: qtyToTake,
    });

    if (newQty <= 0) {
      try {
        await addToPastMedication({
          name: item.name,
          dose: item.dose || "",
          location: item.location || "",
        });
        const sheetId = await getSheetIdByName(item.sheetName);
        if (sheetId === undefined || sheetId === null) continue;
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: {
            requests: [
              {
                deleteDimension: {
                  range: {
                    sheetId,
                    dimension: "ROWS",
                    startIndex: item.rowIndex - 1,
                    endIndex: item.rowIndex,
                  },
                },
              },
            ],
          },
        });
      } catch (error) {
        console.error(
          "Error deleting row and moving to Past Medication:",
          error,
        );
      }
    } else {
      try {
        if (item.sheetName === "Angie Stash") {
          // Quantity in Angie Stash is column C
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `Angie Stash!C${item.rowIndex}`,
            valueInputOption: "RAW",
            resource: { values: [[newQty]] },
          });
        } else {
          // All other sheets: quantity is column D
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `${item.sheetName}!D${item.rowIndex}`,
            valueInputOption: "RAW",
            resource: { values: [[newQty]] },
          });
        }
      } catch (error) {
        console.error("Error updating row:", error);
      }
    }
  }
  res.redirect("/");
});

// Add new medication or update existing (unchanged)
app.post("/add-medication", async (req, res) => {
  let { name, dose, location, quantity } = req.body;
  const angieMode = isAngieMode(name);
  name = stripPlusPlus(name);

  if (angieMode) {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Angie Stash!A:C",
      valueInputOption: "RAW",
      resource: { values: [[name, dose, quantity]] },
    });
    await logActivity({
      action: "ADD",
      name,
      dose,
      location: "Angie Stash",
      quantity,
    });
    return res.redirect("/");
  }

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
            normalizeDose(row[1]) === normalizeDose(dose) &&
            (row[2] || "").toLowerCase() === (location || "").toLowerCase(),
        );
      if (rowIndex !== -1) {
        const actualRowIndex = rowIndex + 2;
        const currentQty = parseInt(data[rowIndex + 1][3]) || 0;
        const newQty = currentQty + parseInt(quantity);
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!D${actualRowIndex}`,
          valueInputOption: "RAW",
          resource: { values: [[newQty]] },
        });
        await logActivity({ action: "ADD", name, dose, location, quantity });
        found = true;
        break;
      }
    }
  }
  if (!found) {
    // If adding new med, and it exists in Past Medication, remove it first (Location might differ)
    await removeFromPastMedication({ name, dose, location });

    // Decide target sheet based on location text
    const targetSheet = location.toLowerCase().includes("closet")
      ? "Closet Meds"
      : "File Meds";

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${targetSheet}!A:D`,
      valueInputOption: "RAW",
      resource: { values: [[name, dose, location, quantity]] },
    });
    await sortSheetByLocation(targetSheet);
    await logActivity({ action: "ADD", name, dose, location, quantity });
  }
  res.redirect("/");
});

// Suggestion for Medication Names
app.get("/all-med-names", async (req, res) => {
  const sheetNames = ["File Meds", "Closet Meds"];
  const namesSet = new Set();
  for (const sheetName of sheetNames) {
    const data = await getSheetData(sheetName);
    if (data.length > 0) {
      data.slice(1).forEach((row) => {
        if (row[0]) namesSet.add(row[0]);
      });
    }
  }
  res.json(Array.from(namesSet));
});

// Suggestion for Dose based on Med Name
app.get("/doses-for-name", async (req, res) => {
  const { name } = req.query;
  if (!name) return res.json([]);
  const sheetNames = ["File Meds", "Closet Meds"];
  const dosesSet = new Set();
  for (const sheetName of sheetNames) {
    const data = await getSheetData(sheetName);
    if (data.length > 1) {
      data.slice(1).forEach((row) => {
        if (
          (row[0] || "").trim().toLowerCase() === name.trim().toLowerCase() &&
          row[1]
        ) {
          dosesSet.add(row[1]);
        }
      });
    }
  }
  res.json(Array.from(dosesSet));
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
