let workbook;
let rawData = [];
let confirmedMapping = {}; 
let finalExportData = []; 

const defaultPricing = {
    pkr: [{th: 4000, add: 500}, {th: 2000, add: 300}, {th: 1000, add: 300}],
    usdDivisor: 230,
    usd: [{th: 5000, add: 5}, {th: 1500, add: 3}, {th: 1000, add: 2}],
    prefix: "$"
};

const columnAliases = {
    "Websites": ["websites", "website", "domain", "url", "site", "link", "blog"],
    "Price (PKR)": ["gp pkr", "gp price", "price pkr", "guest post price", "cost pkr", "pkr", "price", "cost", "rates"],
    "Status": ["status", "type", "package", "category"],
    "DA": ["da", "domain authority", "moz da", "pa/da"],
    "DR": ["dr", "domain rating", "ahrefs dr", "ur/dr"],
    "Traffic": ["traffic", "organic traffic", "monthly traffic", "ahrefs traffic"],
    "TAT": ["tat", "turnaround time", "time", "delivery", "duration"],
    "Links": ["links", "link type", "do follow", "dofollow", "follow"]
};

const outputColumns = ["Websites", "Status", "DA", "DR", "Traffic", "Price", "Price Dollar", "TAT", "Links"];

function getSafeId(colName) {
    return "map_" + colName.replace(/[^a-zA-Z0-9]/g, '');
}

document.getElementById('toggleSettingsBtn').addEventListener('click', function() {
    const el = document.getElementById('settingsArea');
    el.style.display = el.style.display === 'none' ? 'block' : 'none';
});

document.getElementById('restoreDefaultsBtn').addEventListener('click', function() {
    if(!confirm("Reset all pricing rules to default?")) return;
    document.getElementById('pkrThresh1').value = defaultPricing.pkr[0].th;
    document.getElementById('pkrAdd1').value = defaultPricing.pkr[0].add;
    document.getElementById('pkrThresh2').value = defaultPricing.pkr[1].th;
    document.getElementById('pkrAdd2').value = defaultPricing.pkr[1].add;
    document.getElementById('pkrThresh3').value = defaultPricing.pkr[2].th;
    document.getElementById('pkrAdd3').value = defaultPricing.pkr[2].add;
    
    document.getElementById('usdDivisor').value = defaultPricing.usdDivisor;
    document.getElementById('usdThresh1').value = defaultPricing.usd[0].th;
    document.getElementById('usdAdd1').value = defaultPricing.usd[0].add;
    document.getElementById('usdThresh2').value = defaultPricing.usd[1].th;
    document.getElementById('usdAdd2').value = defaultPricing.usd[1].add;
    document.getElementById('usdThresh3').value = defaultPricing.usd[2].th;
    document.getElementById('usdAdd3').value = defaultPricing.usd[2].add;
    document.getElementById('currencyPrefix').value = defaultPricing.prefix;
});

const tabExcelBtn = document.getElementById('tabExcelBtn');
const tabGoogleBtn = document.getElementById('tabGoogleBtn');
const uploadExcelSection = document.getElementById('uploadExcelSection');
const googleSheetSection = document.getElementById('googleSheetSection');
const sheetSelectionArea = document.getElementById('sheetSelectionArea');

tabExcelBtn.addEventListener('click', () => {
    tabExcelBtn.classList.add('active'); tabGoogleBtn.classList.remove('active');
    uploadExcelSection.style.display = 'block'; googleSheetSection.style.display = 'none';
});
tabGoogleBtn.addEventListener('click', () => {
    tabGoogleBtn.classList.add('active'); tabExcelBtn.classList.remove('active');
    googleSheetSection.style.display = 'block'; uploadExcelSection.style.display = 'none';
});

document.getElementById('fileUpload').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        workbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
        populateSheetDropdown();
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('fetchGoogleBtn').addEventListener('click', async function() {
    const urlInput = document.getElementById('googleUrlInput').value.trim();
    const errorMsg = document.getElementById('googleErrorMsg');
    errorMsg.style.display = 'none';

    if (!urlInput.includes("docs.google.com/spreadsheets")) {
        errorMsg.textContent = "Please enter a valid Google Sheets URL.";
        errorMsg.style.display = 'block'; return;
    }

    try {
        const sheetId = urlInput.match(/\/d\/(.*?)(\/|$)/)[1];
        const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`;
        
        const response = await fetch(exportUrl);
        if (!response.ok) throw new Error("Could not access sheet. Ensure it is public.");
        
        const csvText = await response.text();
        workbook = XLSX.read(csvText, {type: 'string'});
        populateSheetDropdown();
    } catch (error) {
        errorMsg.textContent = "Error: " + error.message;
        errorMsg.style.display = 'block';
    }
});

function detectAndSetHeaderRow() {
    const worksheet = workbook.Sheets[document.getElementById('sheetSelect').value];
    const rawArrays = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
    
    let detectedRow = 1;
    
    // Scan first 25 rows
    for(let i = 0; i < Math.min(rawArrays.length, 25); i++) {
        if(!rawArrays[i]) continue;
        
        let matchCount = 0;
        let cells = rawArrays[i].map(c => String(c).toLowerCase().trim());
        
        // Exact column matching, not substring
        for(let cell of cells) {
            if (cell === "website" || cell === "websites" || cell === "domain" || cell === "url" || 
                cell === "da" || cell === "dr" || cell === "traffic" || cell === "status" || 
                cell === "tat" || cell.includes("pkr") || cell === "price" || cell === "cost") {
                matchCount++;
            }
        }
        
        // Agar 2 ya 2 se zyada headers exactly match ho jayen tabhi ye row header banegi
        if (matchCount >= 2) {
            detectedRow = i + 1; 
            break;
        }
    }
    
    const headerInput = document.getElementById('headerRow');
    if(headerInput) headerInput.value = detectedRow;
}

function populateSheetDropdown() {
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.innerHTML = '';
    workbook.SheetNames.forEach(name => {
        const option = document.createElement('option');
        option.text = name; option.value = name;
        sheetSelect.appendChild(option);
    });
    sheetSelectionArea.style.display = 'block';
    document.getElementById('mappingArea').style.display = 'none';
    document.getElementById('processingArea').style.display = 'none';

    detectAndSetHeaderRow();
}

document.getElementById('sheetSelect').addEventListener('change', detectAndSetHeaderRow);

document.getElementById('loadSheetBtn').addEventListener('click', function() {
    const worksheet = workbook.Sheets[document.getElementById('sheetSelect').value];
    
    let headerInput = document.getElementById('headerRow');
    let headerIndex = 0;
    
    if(headerInput && headerInput.value) {
        headerIndex = parseInt(headerInput.value) - 1;
        if(headerIndex < 0) headerIndex = 0;
    }
    
    rawData = XLSX.utils.sheet_to_json(worksheet, {defval: "", range: headerIndex}); 
    
    if (rawData.length > 0) {
        setupMapping(rawData);
    } else {
        alert("No data found! Check if the sheet is empty or 'Headers are on Row' number is wrong.");
    }
});

function setupMapping(data) {
    const mappingBody = document.getElementById('mappingBody');
    mappingBody.innerHTML = '';
    const vendorHeaders = Object.keys(data[0]);
    const sampleRow = data[0];

    Object.keys(columnAliases).forEach(linkifyCol => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td><strong>${linkifyCol}</strong></td>`;

        const tdSelect = document.createElement('td');
        const select = document.createElement('select');
        select.id = getSafeId(linkifyCol);
        select.innerHTML = `<option value="">-- Select Column --</option>`;
        
        let bestGuess = "";
        vendorHeaders.forEach(vh => {
            if (columnAliases[linkifyCol].includes(vh.toLowerCase().trim())) bestGuess = vh;
            select.innerHTML += `<option value="${vh}">${vh}</option>`;
        });

        if (bestGuess) select.value = bestGuess;
        tdSelect.appendChild(select);
        tr.appendChild(tdSelect);

        const tdSample = document.createElement('td');
        tdSample.className = "sample-data";
        tdSample.textContent = bestGuess ? sampleRow[bestGuess] : "None";
        tr.appendChild(tdSample);

        select.addEventListener('change', e => tdSample.textContent = e.target.value ? sampleRow[e.target.value] : "None");
        mappingBody.appendChild(tr);
    });
    document.getElementById('mappingArea').style.display = 'block';
}

function cleanDomain(url) {
    if (!url || typeof url !== 'string') return "";
    let clean = url.trim().toLowerCase().replace(/\badvance\b/g, '').trim();
    clean = clean.replace(/^https?:\/\//, '').replace(/^www\./, '');
    return clean.split('/')[0];
}

function calcLinkifyPKR(orig) {
    const p1 = Number(document.getElementById('pkrThresh1').value);
    const a1 = Number(document.getElementById('pkrAdd1').value);
    const p2 = Number(document.getElementById('pkrThresh2').value);
    const a2 = Number(document.getElementById('pkrAdd2').value);
    const p3 = Number(document.getElementById('pkrThresh3').value);
    const a3 = Number(document.getElementById('pkrAdd3').value);

    if (orig >= p1) return orig + a1;
    if (orig >= p2) return orig + a2;
    if (orig >= p3) return orig + a3;
    return orig;
}

function calcLinkifyUSD(adjustedPkr) {
    const div = Number(document.getElementById('usdDivisor').value);
    const p1 = Number(document.getElementById('usdThresh1').value);
    const a1 = Number(document.getElementById('usdAdd1').value);
    const p2 = Number(document.getElementById('usdThresh2').value);
    const a2 = Number(document.getElementById('usdAdd2').value);
    const p3 = Number(document.getElementById('usdThresh3').value);
    const a3 = Number(document.getElementById('usdAdd3').value);
    const prefix = document.getElementById('currencyPrefix').value;

    let baseUsd = Math.ceil(adjustedPkr / div); 
    let addUsd = 0;
    
    if (adjustedPkr > p1) addUsd = a1;
    else if (adjustedPkr > p2) addUsd = a2;
    else if (adjustedPkr >= p3) addUsd = a3;

    return prefix + (baseUsd + addUsd);
}

document.getElementById('confirmMappingBtn').addEventListener('click', function() {
    confirmedMapping = {};
    Object.keys(columnAliases).forEach(col => confirmedMapping[col] = document.getElementById(getSafeId(col)).value);

    if (!confirmedMapping["Price (PKR)"] || !confirmedMapping["Websites"]) {
        alert("Action Required: Both 'Websites' and 'Price (PKR)' must be mapped.");
        return;
    }

    finalExportData = [];
    let seenDomains = new Set();
    let stats = { total: rawData.length, valid: 0, duplicates: 0, invalid: 0 };

    rawData.forEach(row => {
        let domain = cleanDomain(row[confirmedMapping["Websites"]]);
        let rawPrice = row[confirmedMapping["Price (PKR)"]];
        
        let strPrice = String(rawPrice).replace(/,/g, '');
        let priceMatch = strPrice.match(/\d+(\.\d+)?/);
        let price = priceMatch ? parseFloat(priceMatch[0]) : NaN;

        if (!domain || isNaN(price) || price <= 0) { stats.invalid++; return; }
        if (seenDomains.has(domain)) { stats.duplicates++; return; }

        seenDomains.add(domain);
        stats.valid++;

        let finalRow = {};
        outputColumns.forEach(col => finalRow[col] = ""); 

        finalRow["Websites"] = domain;
        if(confirmedMapping["Status"]) finalRow["Status"] = row[confirmedMapping["Status"]];
        if(confirmedMapping["DA"]) finalRow["DA"] = row[confirmedMapping["DA"]];
        if(confirmedMapping["DR"]) finalRow["DR"] = row[confirmedMapping["DR"]];
        if(confirmedMapping["Traffic"]) finalRow["Traffic"] = row[confirmedMapping["Traffic"]];
        
        let linkifyPKR = calcLinkifyPKR(price); 
        finalRow["Price"] = linkifyPKR;
        finalRow["Price Dollar"] = calcLinkifyUSD(linkifyPKR); 

        if(confirmedMapping["TAT"]) finalRow["TAT"] = row[confirmedMapping["TAT"]];
        if(confirmedMapping["Links"]) finalRow["Links"] = row[confirmedMapping["Links"]];

        finalExportData.push(finalRow);
    });

    document.getElementById('statTotal').textContent = stats.total;
    document.getElementById('statValid').textContent = stats.valid;
    document.getElementById('statDuplicates').textContent = stats.duplicates;
    document.getElementById('statInvalid').textContent = stats.invalid;

    showFinalPreview();
});

function showFinalPreview() {
    const head = document.getElementById('cleanTableHead');
    const body = document.getElementById('cleanTableBody');
    head.innerHTML = ''; body.innerHTML = '';
    
    if(finalExportData.length === 0) {
        alert("No valid data found after cleaning!"); return;
    }

    const trHead = document.createElement('tr');
    outputColumns.forEach(h => {
        const th = document.createElement('th'); th.textContent = h; trHead.appendChild(th);
    });
    head.appendChild(trHead);

    finalExportData.slice(0, 5).forEach(row => {
        const tr = document.createElement('tr');
        outputColumns.forEach(h => {
            const td = document.createElement('td'); td.textContent = row[h]; tr.appendChild(td);
        });
        body.appendChild(tr);
    });

    document.getElementById('processingArea').style.display = 'block';
    document.getElementById('processingArea').scrollIntoView({ behavior: 'smooth' });
}

document.getElementById('exportExcelBtn').addEventListener('click', function() {
    if(finalExportData.length === 0) return;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(finalExportData, {header: outputColumns});
    XLSX.utils.book_append_sheet(wb, ws, "Linkify Format");
    XLSX.writeFile(wb, "Linkify_Converted_Vendor.xlsx");
});