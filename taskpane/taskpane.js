/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/*Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});*/

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

document.getElementById("loginForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const email = e.target.email.value;
    const password = e.target.password.value;

    // Show spinner
    const loginText = document.getElementById("loginText");
    const loginSpinner = document.getElementById("loginSpinner");
    loginText.style.display = "none";
    loginSpinner.style.display = "block";

    try {
        const res = await fetch("http://localhost:8000/addin/login", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ email, password })
        });

        if (!res.ok) throw new Error("Login failed");

        const data = await res.json();
        const { user, membership } = data;

        document.getElementById("loginScreen").style.display = "none";
        document.getElementById("blockedContainer").style.display = "none";

        if (membership.is_member === "yes") {
            // Save session in localStorage
            localStorage.setItem("user", JSON.stringify(user));
            localStorage.setItem("membership", JSON.stringify(membership));

            document.getElementById("memberContent").style.display = "block";
            document.getElementById("addinUI").style.display = "block";
            document.getElementById("logoutBtn").style.display = "block";

            const profileIcon = document.getElementById("profileIcon");
            const username = document.getElementById("username");

            if (user && user.FullName) {
                username.textContent = user.FullName;
                const initials = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
                profileIcon.textContent = initials;
            }
        } else {
            document.getElementById("blockedContainer").style.display = "flex";
        }

    } catch (err) {
        console.error(err);
        document.getElementById("loginScreen").style.display = "none";
        document.getElementById("memberContent").style.display = "none";
        document.getElementById("blockedContainer").style.display = "flex";
    } finally {
        // Hide spinner and show text again
        loginText.style.display = "block";
        loginSpinner.style.display = "none";
    }
});

window.addEventListener("DOMContentLoaded", () => {
    const savedUser = localStorage.getItem("user");
    const savedMembership = localStorage.getItem("membership");

    if (savedUser && savedMembership) {
        const user = JSON.parse(savedUser);
        const membership = JSON.parse(savedMembership);

        if (membership.is_member === "yes") {
            // Show add-in UI directly
            document.getElementById("memberContent").style.display = "block";
            document.getElementById("addinUI").style.display = "block";
            document.getElementById("logoutBtn").style.display = "block";

            const profileIcon = document.getElementById("profileIcon");
            const username = document.getElementById("username");

            if (user && user.FullName) {
                username.textContent = user.FullName;
                const initials = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
                profileIcon.textContent = initials;
            }

            return; // Exit: don't show login screen
        }
    }

    // No session → show login
    document.getElementById("loginScreen").style.display = "flex";
});




document.getElementById("logoutBtn").addEventListener("click", () => {
    localStorage.removeItem("user");
    localStorage.removeItem("membership");

    document.getElementById("memberContent").style.display = "none";
    document.getElementById("addinUI").style.display = "none";
    document.getElementById("logoutBtn").style.display = "none";
    document.getElementById("loginScreen").style.display = "flex";
});



Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) return;

    const refreshBtn = document.getElementById("refreshBtn");
    const helpBtn = document.getElementById("helpBtn");

    function getExcelColumnLetter(colNum) {
        let letter = "";
        while (colNum > 0) {
            let mod = (colNum - 1) % 26;
            letter = String.fromCharCode(65 + mod) + letter;
            colNum = Math.floor((colNum - mod) / 26);
        }
        return letter;
    }

    // Attach Help button
    helpBtn.onclick = () => {
        const toggle = document.getElementById("dropdownToggle");
        const fincode = toggle.dataset.value;
        if (!fincode) return alert("Select a company first.");

        const url = `https://www.goindiastocks.com/companyinfo/${encodeURIComponent(fincode)}`;
        window.open(url, "_blank");
    };

    // Attach Refresh button
    refreshBtn.onclick = async () => {
        const toggle = document.getElementById("dropdownToggle");
        const fincode = toggle.dataset.value;
        const name = toggle.textContent;

        if (!fincode) return alert("Select a company first.");

        const sectorType = document.getElementById("dropdownToggle").dataset.sector;

        // ✅ Read toggle state
        const modeToggle = document.getElementById("modeToggle");
        const isIndAS = modeToggle.checked; // true = IndAS, false = Detailed

        // ✅ Decide suffix
        const sheetSuffix = isIndAS ? "" : "IND"; // "" for In, "IND" for De

        console.log(sectorType)

        try {
            console.log(`🔄 Refreshing data for: ${fincode} - ${name} | Mode: ${isIndAS ? "In" : "De"}`);

            // === Fetch all APIs in parallel ===
            const [
                cashCResp, cashSResp,
                qplCResp, qplSResp,
                bsCResp, bsSResp,
                plCResp, plSResp,
                keyfResp] = await Promise.all([

                    // Cash Flows
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "C", sector: sectorType, sheet: `CashFlow${sheetSuffix}` })
                    }),
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "S", sector: sectorType, sheet: `CashFlow${sheetSuffix}` })
                    }),

                    // Quarterly Profit & Loss
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "C", sector: sectorType, sheet: `QProfitLoss${sheetSuffix}` })
                    }),
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "S", sector: sectorType, sheet: `QProfitLoss${sheetSuffix}` })
                    }),

                    // Balance Sheet
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "C", sector: sectorType, sheet: `BalanceSheet${sheetSuffix}` })
                    }),
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "S", sector: sectorType, sheet: `BalanceSheet${sheetSuffix}` })
                    }),

                    // Profit & Loss
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "C", sector: sectorType, sheet: `ProfitLoss${sheetSuffix}` })
                    }),
                    fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "S", sector: sectorType, sheet: `ProfitLoss${sheetSuffix}` })
                    }),

                    // Key Financials (single) → does NOT depend on IND suffix
                    fetch("https://transcriptanalyser.com/goindiastock/actuals_forwards", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "C", sector_type: sectorType })
                    })
                ]);

            const [
                cashCData, cashSData,
                qplCData, qplSData,
                bsCData, bsSData,
                plCData, plSData,
                keyfData] = await Promise.all([
                cashCResp.json(), cashSResp.json(),
                qplCResp.json(), qplSResp.json(),
                bsCResp.json(), bsSResp.json(),
                plCResp.json(), plSResp.json(),
                keyfResp.json()
            ]);

            await Excel.run(async (context) => {
                const workbook = context.workbook;
                const sheetNames = ["Key Financials", "Quarterly Data", "Annual Data"];
                workbook.worksheets.load("items/name");
                await context.sync();

                const existingNames = workbook.worksheets.items.map(s => s.name);
                const sheetsMap = {};

                // Create or reset sheets
                for (const sheetName of sheetNames) {
                    let sheet;
                    if (existingNames.includes(sheetName)) {
                        sheet = workbook.worksheets.getItem(sheetName);
                        sheet.getUsedRange()?.clear();
                    } else {
                        sheet = workbook.worksheets.add(sheetName);
                    }
                    sheetsMap[sheetName] = sheet;

                    // Company name in A1
                    const nameCell = sheet.getRange("A1");
                    nameCell.values = [[name]];
                    nameCell.format.font.bold = true;
                    nameCell.format.font.size = 14;
                    nameCell.format.fill.color = "#bed1f8";

                    // Make column A wider
                    sheet.getRange("A:A").format.columnWidth = 180; 
                }

                await context.sync(); // commit sheet setup

                // --- Helper ---
                const getExcelColumnLetter = (colNum) => {
                    let temp = "";
                    let letter = "";
                    while (colNum > 0) {
                        temp = (colNum - 1) % 26;
                        letter = String.fromCharCode(temp + 65) + letter;
                        colNum = (colNum - temp - 1) / 26;
                    }
                    return letter;
                };

                const formatTable = (sheet, startRow, title, headers, values) => {
                    const lastCol = getExcelColumnLetter(headers.length);

                    // Title row
                    const titleCell = sheet.getRange(`A${startRow}`);
                    titleCell.values = [[title]];
                    titleCell.format.font.bold = true;
                    startRow++;

                    // Header row
                    const headerRange = sheet.getRange(`A${startRow}:${lastCol}${startRow}`);
                    headerRange.values = [headers];
                    headerRange.format.fill.color = "#e0e0e0";
                    headerRange.format.font.bold = true;
                    startRow++;

                    // Data rows
                    if (values.length > 0) {
                        const dataRange = sheet.getRange(`A${startRow}:${lastCol}${startRow + values.length - 1}`);
                        dataRange.values = values;
                        startRow += values.length;
                    } else {
                        sheet.getRange(`A${startRow}`).values = [["No data available"]];
                        startRow++;
                    }

                    return startRow + 1; // leave a blank row
                };

                // --- Key Financials ---
                const keySheet = sheetsMap["Key Financials"];
                if (keyfData?.value?.length > 0) {
                    const keyArray = keyfData.value;
                    const staticFields = ["Parameter"];
                    const dynamicHeaders = new Set();

                    keyArray.forEach(row => {
                        Object.keys(row).forEach(key => {
                            if (!staticFields.includes(key) && /^FY\d{4}(E)?$/.test(key)) dynamicHeaders.add(key);
                        });
                    });

                    const sortedDynamicHeaders = Array.from(dynamicHeaders).sort((a, b) =>
                        parseInt(a.replace("FY", "").replace("E", "")) - parseInt(b.replace("FY", "").replace("E", ""))
                    );
                    const keyHeaders = ["Parameter", ...sortedDynamicHeaders];
                    const keyValues = keyArray.map(row => keyHeaders.map(h => row[h] ?? (h === "Parameter" ? row.Parameter ?? "" : "")));

                    // Format the title row
                    const startRow = 3; // same as you pass to formatTable
                    keySheet.getRange(`A${startRow}`).values = [["Key Financials"]];
                    keySheet.getRange(`A${startRow}`).format.font.bold = true;
                    keySheet.getRange(`A${startRow}`).format.font.size = 14;
                    keySheet.getRange(`A${startRow}`).format.fill.color = "#d9ead3";
                    keySheet.getRange(`A${startRow}`).format.horizontalAlignment = "CenterAcrossSelection";

                    // Then format the actual table
                    formatTable(keySheet, startRow, "Key Financials", keyHeaders, keyValues);
                }

                const flattenHierarchicalData = (data, level = 0) => {
                    if (!data) return [];
                    const rows = [];
                    for (const row of data) {
                        rows.push({ ...row, _level: level });
                        if (row.child?.length) {
                            rows.push(...flattenHierarchicalData(row.child, level + 1));
                        }
                    }
                    return rows;
                };


                // --- Helper: Write Standalone + Consolidated side by side ---
                const writeSideBySide = (sheet, startRow, standaloneData, consolidatedData, sectionName) => {
                    const staticFields = ["Parameter"];

                    const getHeaders = (data) => {
                        if (!data?.length) return [];
                        const headers = new Set();
                        data.forEach(row => {
                            Object.keys(row).forEach(k => {
                                if (!staticFields.includes(k) && /^[A-Z][a-z]{2}\d{4}$/.test(k)) headers.add(k);
                            });
                        });
                        return Array.from(headers).sort((a, b) => {
                            const months = { Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11 };
                            const parseDate = s => { const [_, mon, year] = s.match(/^([A-Za-z]+)(\d{4})$/) || []; return new Date(parseInt(year), months[mon]??0); };
                            return parseDate(a) - parseDate(b);
                        });
                    };

                    const sHeaders = getHeaders(standaloneData);
                    const cHeaders = getHeaders(consolidatedData);

                    // Section title
                    sheet.getRange(`A${startRow}`).values = [[sectionName]];
                    sheet.getRange(`A${startRow}`).format.font.bold = true;
                    sheet.getRange(`A${startRow}`).format.font.size = 14;
                    sheet.getRange(`A${startRow}`).format.fill.color = "#d9ead3";
                    sheet.getRange(`A${startRow}`).format.horizontalAlignment = "CenterAcrossSelection";
                    startRow++;

                    const processData = (data, title) => {
                        if (!data?.length) return startRow;

                        const flatData = flattenHierarchicalData(data);
                        const headers = ["Parameter", ...getHeaders(data)];
                        const values = flatData.map(row => headers.map(h => h === "Parameter" 
                            ? " ".repeat(row._level * 4) + (row.Parameter ?? "")
                            : row[h] ?? ""));
                        
                        const dataStartRow = startRow;
                        startRow = formatTable(sheet, dataStartRow, title, headers, values);

                        // Apply smaller font for children (level > 0)
                        flatData.forEach((row, idx) => {
                            if (row._level > 0) {
                                const r = sheet.getRange(`A${dataStartRow + 1 + idx}:${String.fromCharCode(65 + headers.length - 1)}${dataStartRow + 1 + idx}`);
                                r.format.font.size = 11;
                            }
                        });

                        // Corrected logic for outline grouping
                        let i = 0;
                        while (i < flatData.length) {
                            const currentRow = flatData[i];
                            const nextRow = flatData[i + 1];

                            // If the next row is a child (higher level), we are at the start of a potential group
                            if (nextRow && nextRow._level > currentRow._level) {
                                const groupStartRow = dataStartRow + 1 + (i + 1); // Group starts at the child's row
                                let j = i + 1;
                                
                                // Find the end of the contiguous block of children
                                while (j < flatData.length && flatData[j]._level > currentRow._level) {
                                    j++;
                                }

                                const groupEndRow = dataStartRow + 1 + (j - 1);
                                
                                // Ensure a valid range exists before attempting to group
                                if (groupStartRow <= groupEndRow) {
                                    sheet.getRange(`A${groupStartRow}:A${groupEndRow}`).getEntireRow().group();
                                }

                                // Move the main loop's index to the end of the grouped block
                                i = j;
                            } else {
                                // If not a parent, just move to the next row
                                i++;
                            }
                        }

                        return startRow;
                    };

                    startRow = processData(consolidatedData, "Consolidated");
                    startRow = processData(standaloneData, "Standalone");

                    return startRow;
                };

                // --- Quarterly Data ---
                writeSideBySide(sheetsMap["Quarterly Data"], 3, qplSData.value, qplCData.value, "Quarterly P&L");

                // --- Annual Data ---
                let row = 3;
                row = writeSideBySide(sheetsMap["Annual Data"], row, bsSData.value, bsCData.value, "Balance Sheet");
                row++;
                row = writeSideBySide(sheetsMap["Annual Data"], row, cashSData.value, cashCData.value, "Cash Flows");
                row++;
                row = writeSideBySide(sheetsMap["Annual Data"], row, plSData.value, plCData.value, "Detailed P&L");

                await context.sync();
            });

            console.log("✅ Data successfully written to Excel");

        } catch (err) {
            console.error("❌ Error in refreshBtn:", err);
            alert("Failed to fetch company data. Check console for details.");
        }
    };
});
