const excelInput = document.getElementById("excel-input");
const label = document.getElementById("excel-input-label");
const snowContainer = document.getElementById("snow-container");
defaultLabel();
summonSnow();

excelInput.addEventListener("change", loadExcel);

/* Excel loading */

function summonSnow() {
  snowContainer.innerHTML = Array.from(
    { length: 11 },
    () => `
  <div class="snowflake">
    <div class="inner">❅</div>
  </div>
`,
  ).join("");
}

function defaultLabel() {
  label.innerText = "Select a file";
}

async function loadExcel(event) {
  const file = event.target.files[0];
  if (!file) {
    defaultLabel();
    return;
  }
  label.innerText = file.name;

  const reader = new FileReader();

  reader.onload = async (event) => {
    try {
      const excelData = new Uint8Array(event.target.result);

      const workbook = XLSX.read(excelData, {
        type: "array",
      });

      const sheetName = workbook.SheetNames.find(
        (name) => name.toLowerCase() === "data",
      );

      if (!sheetName) {
        throw new Error('Could not find the "Data" worksheet.');
      }

      const sheet = workbook.Sheets[sheetName];
      const zip = await JSZip.loadAsync(excelData);

      const tableNames = ["Details", "Personal", "Data"];

      const tables = await Promise.all(
        tableNames.map((name) => getExcelTable(zip, name)),
      );

      if (tables.some((table) => !table)) {
        throw new Error("One or more Excel tables could not be found.");
      }

      const [detailsRows, personalRows, dataRows] = tables.map((table) =>
        XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          range: table.range,
          defval: "",
          raw: false,
        }),
      );

      renderQuoteItems(convertQuoteItems(extractTable(dataRows)));
      renderPersonalItems(extractTable(personalRows));
      renderDetails(extractTable(detailsRows));
    } catch (error) {
      console.error("Failed to load Excel workbook:", error);
      alert(error.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

/* Excel table lookup */

async function getExcelTable(zip, tableName) {
  const tableFiles = Object.keys(zip.files).filter((path) =>
    /^xl\/tables\/table\d+\.xml$/i.test(path),
  );

  for (const path of tableFiles) {
    const xml = await zip.file(path).async("text");

    const document = new DOMParser().parseFromString(xml, "application/xml");

    const table = document.documentElement;
    const name = table.getAttribute("displayName");

    if (name === tableName) {
      return {
        name,
        range: table.getAttribute("ref"),
      };
    }
  }

  return null;
}

/* Data extraction */

function extractTable(rows) {
  const headers = rows[0];

  return Object.fromEntries(
    headers.map((header, index) => [
      header,
      rows.slice(1).map((row) => row[index] ?? ""),
    ]),
  );
}

function getColumnValues(table, key) {
  return (table[key] ?? [])
    .map((value) => String(value).trim())
    .filter(Boolean);
}

function getFirstValue(table, key) {
  return getColumnValues(table, key)[0] ?? "";
}

function getHtmlString(table, key) {
  return getColumnValues(table, key)
    .map((item) => `${item}`)
    .join("<br />");
}

/* Quote items */

function convertQuoteItems(table) {
  return table.Description.map((description, index) => ({
    description: String(description).trim(),
    quantity: Number(table.Quantity[index]),
    unitPrice: Number(String(table.Unit_Price[index]).replace(/[R,\s]/g, "")),
  })).filter((item) => item.description);
}

function formatMoney(value) {
  return `R ${value.toFixed(2)}`;
}

function renderQuoteItems(items) {
  const container = document.getElementById("quote-items");
  const fragment = document.createDocumentFragment();

  let total = 0;

  items.forEach((item) => {
    const itemTotal = item.quantity * item.unitPrice;
    total += itemTotal;

    const row = document.createElement("tr");

    const values = [
      item.description,
      item.quantity,
      formatMoney(item.unitPrice),
      formatMoney(itemTotal),
    ];

    values.forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = value;
      row.appendChild(cell);
    });

    fragment.appendChild(row);
  });

  container.replaceChildren(fragment);

  document.getElementById("quote-total").innerHTML = formatMoney(total);
}

/* Personal information */

function renderPersonalItems(personal) {
  document.getElementById("business-name").innerHTML = getHtmlString(
    personal,
    "My_Name",
  );

  document.getElementById("business-address").innerHTML = getHtmlString(
    personal,
    "My_Address",
  );

  document.getElementById("business-phone").innerHTML = getHtmlString(
    personal,
    "My_Phone",
  );

  document.getElementById("business-email").innerHTML = getHtmlString(
    personal,
    "My_Email",
  );

  document.getElementById("customer-name").innerHTML = getHtmlString(
    personal,
    "To_Who",
  );

  document.getElementById("customer-address").innerHTML = getHtmlString(
    personal,
    "Their_Address",
  );

  document.getElementById("quote-date").innerHTML =
    new Date().toLocaleDateString("en-ZA", {
      day: "numeric",
      month: "long",
      year: "numeric",
    });

  renderBankingDetails(personal);
}

function renderBankingDetails(personal) {
  const container = document.getElementById("banking-details");
  const fragment = document.createDocumentFragment();

  getColumnValues(personal, "Banking_Details").forEach((item) => {
    const paragraph = document.createElement("p");
    paragraph.textContent = item;
    fragment.appendChild(paragraph);
  });

  container.replaceChildren(fragment);
}

/* Quote details */

function renderDetails(details) {
  document.getElementById("details-element").textContent =
    `Details : ${getFirstValue(details, "Details")}`;
}
