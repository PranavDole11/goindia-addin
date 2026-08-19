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

async function checkAccessStatus(userId) {
    try {
        const res = await fetch(`https://transcriptanalyser.com/addin/access_status?user_id=${userId}`);
        const data = await res.json();

        const container = document.getElementById("enterpriseRequestContainer");
        const messageParagraph = container.querySelector("p");

        if (data.status === "Approved") {
            messageParagraph.textContent = "Your access has been approved! You can now use the add-in.";
            document.getElementById("requestAccessBtn").style.display = "none";
        } else if (data.status === "Pending") {
            messageParagraph.textContent = "Your request is pending approval. Please wait for the admin to approve it.";
            document.getElementById("requestAccessBtn").style.display = "none";
        } else if (data.status === "Rejected") {
            messageParagraph.textContent = "Your request was rejected. Please contact support for details.";
            document.getElementById("requestAccessBtn").style.display = "none";
        } else if (data.status === "NoRequest") {
            messageParagraph.textContent = "This feature is available only for Enterprise Plan members.\nIf you need access, request it below.";
            document.getElementById("requestAccessBtn").style.display = "block";
        }

        container.style.display = "flex";

    } catch (err) {
        console.error(err);
    }
}

/*document.getElementById("loginForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const email = e.target.email.value;
    const password = e.target.password.value;

    // Show spinner
    const loginText = document.getElementById("loginText");
    const loginSpinner = document.getElementById("loginSpinner");
    loginText.style.display = "none";
    loginSpinner.style.display = "block";

    try {
        const resUser = await fetch("https://www.transcriptanalyser.com/payment/getUser", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ email, password })
        });

        if (!resUser.ok) throw new Error("Login failed");

        const user = await resUser.json();
        if (!user || !user.UserId) throw new Error("Invalid user data");

        const resMembership = await fetch("https://transcriptanalyser.com/payment/check_membership", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ user_id: user.UserId })
        });

        if (!resMembership.ok) throw new Error("Membership check failed");

        const membership = await resMembership.json();

        // Hide all states first
        document.getElementById("loginScreen").style.display = "none";
        document.getElementById("memberContent").style.display = "none";
        document.getElementById("addinUI").style.display = "none";
        document.getElementById("blockedContainer").style.display = "none";
        document.getElementById("enterpriseRequestContainer").style.display = "none";

        // CASE 1: Not a member
        if (membership.is_member === "no") {
            document.getElementById("blockedContainer").style.display = "flex";
            return;
        }

        // CASE 2: Member but not enterprise
        if (!(membership.is_giin_pro === true || membership.is_admin === true || membership.is_enterprise_member === true)) {
            // First, check if user has already been approved in access requests
            try {
                const resStatus = await fetch(`https://transcriptanalyser.com/addin/access_status?user_id=${user.UserId}`);
                const data = await resStatus.json();

                if (data.status === "Approved") {
                    // Treat this user as an enterprise member
                    localStorage.setItem("user", JSON.stringify(user));
                    localStorage.setItem("membership", JSON.stringify(membership));

                    document.getElementById("memberContent").style.display = "block";
                    showAddinUI();
                    updateLogoutBtnVisibility();

                    const profileIcon = document.getElementById("profileIcon");
                    const username = document.getElementById("username");
                    if (user.FullName) {
                        username.textContent = user.FullName;
                        profileIcon.textContent = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
                    }
                    return; // Don't proceed to request access UI
                }
            } catch (err) {
                console.error("Failed to fetch access status:", err);
            }

            // If not approved, show the request access UI
            const container = document.getElementById("enterpriseRequestContainer");
            container.style.display = "flex";  // Show container

            const statusMessage = document.getElementById("statusMessage");  // New message area
            const requestBtn = document.getElementById("requestAccessBtn");
            const checkStatusLink = document.getElementById("checkStatusLink");

            requestBtn.onclick = async () => {
                try {
                    const res = await fetch("https://transcriptanalyser.com/addin/request_access", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({
                            user_id: user.UserId,
                            email: user.Email,
                            full_name: user.FullName
                        })
                    });

                    const data = await res.json();

                    if (!res.ok) throw new Error(data.detail || "Request failed");

                    showWarning(data.message);
                } catch (err) {
                    console.error(err);
                    showWarning("Failed to submit request. Please try again later.");
                }
            };

            checkStatusLink.onclick = async (e) => {
                e.preventDefault();  // Prevent default link behavior

                try {
                    const resStatus = await fetch(`https://transcriptanalyser.com/addin/access_status?user_id=${user.UserId}`);
                    const data = await resStatus.json();

                    if (data.status === "Approved") {
                        statusMessage.textContent = "Approved";
                        statusMessage.style.color = "green";
                    } else if (data.status === "Pending") {
                        statusMessage.textContent = "Pending";
                        statusMessage.style.color = "orange";
                    } else if (data.status === "Rejected") {
                        statusMessage.textContent = "Rejected";
                        statusMessage.style.color = "red";
                    } else if (data.status === "NoRequest") {
                        statusMessage.textContent = "No request found";
                        statusMessage.style.color = "#173760";
                    }
                } catch (err) {
                    console.error(err);
                    showWarning("Failed to fetch status. Please try again later.");
                }
            };
        }

        if (membership.is_giin_pro === true || membership.is_admin === true || membership.is_enterprise_member === true) {
            // Save user details to localStorage
            localStorage.setItem("user", JSON.stringify(user));
            localStorage.setItem("membership", JSON.stringify(membership));

            // Show member-related content
            document.getElementById("memberContent").style.display = "block";
            showAddinUI();
            updateLogoutBtnVisibility();

            // Update profile display
            const profileIcon = document.getElementById("profileIcon");
            const username = document.getElementById("username");
            if (user.FullName) {
                username.textContent = user.FullName;
                profileIcon.textContent = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
            }
        }

        // Update profile
        

    } catch (err) {
        console.error(err);
        document.getElementById("blockedContainer").style.display = "flex";
    } finally {
        loginText.style.display = "block";
        loginSpinner.style.display = "none";
    }
});*/

document.getElementById("loginForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const email = e.target.email.value;
    const password = e.target.password.value;

    // Show spinner
    const loginText = document.getElementById("loginText");
    const loginSpinner = document.getElementById("loginSpinner");
    loginText.style.display = "none";
    loginSpinner.style.display = "block";

    const loginError = document.getElementById("loginError");
    if (loginError) loginError.style.display = "none"; // clear any error from a previous attempt

    try {
        // === Hardcoded exception for test user ===
        if (email === "ut@gmail.com" && password === "123") {
            const user = { UserId: "test123", FullName: "Test User", email };
            const membership = { is_member: "yes", is_enterprise_member: false };

            localStorage.setItem("user", JSON.stringify(user));
            localStorage.setItem("membership", JSON.stringify(membership));

            document.getElementById("loginScreen").style.display = "none";
            document.getElementById("blockedContainer").style.display = "none";

            document.getElementById("memberContent").style.display = "block";
            showAddinUI();
            updateLogoutBtnVisibility();

            const profileIcon = document.getElementById("profileIcon");
            const usernameEl = document.getElementById("username");

            if (user.FullName) {
                usernameEl.textContent = user.FullName;
                const initials = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
                profileIcon.textContent = initials;
            }
            refreshWalletBalance();

            // Skip the rest of the login logic
            return;
        }

        // === Original login logic ===
        // Fetch user data
        const resUser = await fetch("https://www.transcriptanalyser.com/payment/getUser", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ email, password })
        });

        if (!resUser.ok) throw new Error("Login failed");

        const user = await resUser.json();
        if (!user || !user.UserId) throw new Error("Invalid user data");
        logUsageEvent(user.UserId, "login");

        // Check membership
        const resMembership = await fetch("https://transcriptanalyser.com/payment/check_membership", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ user_id: user.UserId })
        });

        if (!resMembership.ok) throw new Error("Membership check failed");

        const membership = await resMembership.json();

        // Hide irrelevant containers
        document.getElementById("loginScreen").style.display = "none";
        document.getElementById("blockedContainer").style.display = "none";

        if (membership.is_giin_pro === true || membership.is_admin === true || membership.is_enterprise_member === true) {
            // Store session data
            localStorage.setItem("user", JSON.stringify(user));
            localStorage.setItem("membership", JSON.stringify(membership));

            // Show relevant UI
            document.getElementById("memberContent").style.display = "block";
            showAddinUI();
            updateLogoutBtnVisibility();

            const profileIcon = document.getElementById("profileIcon");
            const usernameEl = document.getElementById("username");

            if (user.FullName) {
                usernameEl.textContent = user.FullName;
                const initials = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
                profileIcon.textContent = initials;
            }
            refreshWalletBalance();
        } else {
            // Not a member: show blocked container
            document.getElementById("blockedContainer").style.display = "flex";
        }
    } catch (err) {
        console.error(err);

        // "Login failed"/"Invalid user data" mean getUser rejected the credentials themselves —
        // show that plainly and let the user correct it, rather than routing every failure
        // (including a plain wrong password) into the "upgrade to GIIN Club" screen below, which
        // is only actually true for a successful login with an insufficient membership tier (see
        // the "else" branch above). Anything else here (e.g. the membership check itself failing)
        // isn't a credentials problem either — same treatment: report it and let them retry,
        // don't imply they need to upgrade when we don't actually know their membership status.
        if (loginError) {
            loginError.textContent = (err.message === "Login failed" || err.message === "Invalid user data")
                ? "Incorrect email or password."
                : "Something went wrong logging in. Please try again.";
            loginError.style.display = "block";
        }
    } finally {
        // Always restore spinner state
        loginText.style.display = "block";
        loginSpinner.style.display = "none";
    }
});



/*window.addEventListener("DOMContentLoaded", async () => {
    const savedUser = localStorage.getItem("user");
    const savedMembership = localStorage.getItem("membership");

    if (savedUser && savedMembership) {
        const user = JSON.parse(savedUser);
        const membership = JSON.parse(savedMembership);

        try {
            // Check core membership
            const resMembership = await fetch("https://transcriptanalyser.com/payment/check_membership", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ user_id: user.UserId })
            });

            if (!resMembership.ok) throw new Error("Failed to verify membership");

            const updatedMembership = await resMembership.json();

            let isApproved = membership.is_giin_pro === true || membership.is_admin === true || membership.is_enterprise_member === true;

            // If not enterprise member, check AddinAccessRequests
            if (!isApproved) {
                try {
                    const resStatus = await fetch(`https://transcriptanalyser.com/addin/access_status?user_id=${user.UserId}`);
                    const statusData = await resStatus.json();
                    if (statusData.status === "Approved") {
                        isApproved = true;
                    }
                } catch (err) {
                    console.error(err);
                }
            }

            // If user is allowed, store session and show UI
            if (isApproved) {
                localStorage.setItem("user", JSON.stringify(user));
                localStorage.setItem("membership", JSON.stringify(membership));

                document.getElementById("memberContent").style.display = "block";
                showAddinUI();
                updateLogoutBtnVisibility();

                const profileIcon = document.getElementById("profileIcon");
                const username = document.getElementById("username");
                if (user.FullName) {
                    username.textContent = user.FullName;
                    profileIcon.textContent = user.FullName.split(" ").map(n => n[0]).join("").toUpperCase();
                }

                return; // Exit early
            } else {
                // Access revoked or not allowed
                localStorage.removeItem("user");
                localStorage.removeItem("membership");
            }

            // If not approved → show request access UI
            document.getElementById("enterpriseRequestContainer").style.display = "flex";

        } catch (err) {
            console.error("Error verifying session:", err);
            localStorage.removeItem("user");
            localStorage.removeItem("membership");
        }
    }

    // If no valid session, show login screen
    document.getElementById("loginScreen").style.display = "flex";
});*/

const MAX_MODEL_COST_INR = 30; // matches the AI Financial Model card's own displayed cost ceiling
const TEST_BYPASS_USER_ID = "test123"; // ut@gmail.com certification account — see refreshWalletBalance

// Detects an empty/all-zero financial-statement response — a Consolidated-less company doesn't
// necessarily return an EMPTY array, it can return real rows that are all-zero/blank placeholders,
// so a plain "was the array empty" check isn't enough. Used both by fetchKeyFinancials below (to
// decide whether to fall back to Standalone) and by handleRefresh's own probe/sheet-writer for the
// same purpose against the later per-statement fetches.
const isConsolidatedEmpty = (data) => {
    if (!data?.value?.length) return true;
    const ignoredFields = ["Parameter", "child", "Codec", "Info", "parent_id", "hy", "ttm"];
    const checkRow = (row) => {
        const keys = Object.keys(row).filter(k => !ignoredFields.includes(k));
        const parentEmpty = keys.every(k => {
            const val = row[k];
            if (val == null) return true;
            if (typeof val === "number") return val === 0;
            if (typeof val === "string") {
                const v = val.trim().toLowerCase();
                return v === "0" || v === "na" || v === "division by zero" || v === "";
            }
            return true; // ignore other types
        });
        const childrenEmpty = !row.child || isConsolidatedEmpty({ value: row.child });
        return parentEmpty && childrenEmpty;
    };
    return data.value.every(checkRow);
};

// Fetches Key Financials (actuals_forwards) for one company. Prefers Consolidated, but falls back
// to Standalone (isdefault:"yes") when Consolidated isn't available for that company — several
// companies (e.g. banks/NBFCs) report Standalone only. The previous code sent mode:"C"
// unconditionally and never sent "isdefault" at all — for a Consolidated-less company that returned
// nothing instead of falling back, since the API needs isdefault explicitly ("no" for an explicit
// Consolidated request, "yes" when Standalone is being used as the default/fallback presentation).
// NOTE: checking data.value.length alone isn't enough to detect "no Consolidated data" — some
// Consolidated-less companies return real rows that are just all-zero/blank, not an empty array —
// hence isConsolidatedEmpty rather than a plain length check.
// Some companies come back as a clean 200 with an empty/all-zero body when Consolidated isn't
// available (handled by isConsolidatedEmpty above); others — observed on at least one bank — come
// back as an outright 400 Bad Request instead, with no guarantee the error body is even valid JSON.
// tryFetch treats BOTH as "this attempt didn't give us Consolidated data" rather than letting either
// one throw, since an uncaught exception here would abort the entire download instead of falling
// back to Standalone.
async function fetchKeyFinancials(fincode, sectorType) {
    const tryFetch = async (mode, isdefault) => {
        try {
            const res = await fetch("https://transcriptanalyser.com/goindiastock/actuals_forwards", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ fincode, mode, sector_type: sectorType, isdefault }),
            });
            if (!res.ok) return null;
            return await res.json();
        } catch (e) {
            return null; // network failure or a non-JSON error body
        }
    };
    let data = await tryFetch("C", "no");
    if (!data || isConsolidatedEmpty(data)) {
        data = await tryFetch("S", "yes");
    }
    return data;
}

// Fetches the latest available EOD (end-of-day) trading price for a company — used to give the AI
// Financial Model's "current_price" a REAL figure instead of asking the LLM to recall or guess one
// from its own training data (which has no live market data, so a price it reports on its own is
// unavoidably stale, and for a less-covered stock can simply be wrong). current_price is a SINGLE
// scalar fact — today's price, a snapshot — not a per-year series: it has no "value for FY2023" any
// more than "today's date" does, so this only ever needs the single latest point, not history. A
// 5-CALENDAR-day window is enough to reach the latest trading day without pulling years of daily
// data for a number that's only ever used once. Returns { price, date } or null if unavailable
// (invalid fincode, network failure, empty result).
async function fetchLatestPrice(fincode) {
    try {
        const res = await fetch(
            `https://company.accordwebservices.com/Company/GetCompanyGraphEOD_FullD?Type=H&FINCODE=${fincode}&STK=NSE&DateOption=D&DateCount=5&StartDate=&EndDate=&token=Tdo94y6JtwQgFAL8mCvoAwAk3Ueq24ZR`
        );
        if (!res.ok) return null;
        const data = await res.json();
        const rows = data && data.Table;
        if (!Array.isArray(rows) || !rows.length) return null;
        const last = rows[rows.length - 1]; // ascending by date — last entry is the latest trading day
        const price = parseFloat(last.price);
        if (!isFinite(price) || price <= 0) return null;
        const date = String(last.Date || "").split(" ")[0]; // "7/27/2026 12:00:00 AM" -> "7/27/2026"
        return { price, date };
    } catch (e) {
        return null; // network failure or unexpected response shape
    }
}

// Fetches up to ~10 years of DAILY EOD closing prices (same feed/endpoint as fetchLatestPrice,
// just a much wider window) — used as a FALLBACK when sc_year_data has a Market Cap for a
// historical year but no Price for that same year (a real, observed gap — sc_year_data's own
// "price" column isn't always populated even for years the stock genuinely traded). Returns an
// ASCENDING-by-date array of {date, price}, or [] on a failed fetch/empty response. Deliberately
// a single bulk fetch (not one call per missing year) — reused across every year that needs it.
async function fetchAnnualClosingPrices(fincode) {
    try {
        const res = await fetch(
            `https://company.accordwebservices.com/Company/GetCompanyGraphEOD_FullD?Type=H&FINCODE=${fincode}&STK=NSE&DateOption=D&DateCount=3650&StartDate=&EndDate=&token=Tdo94y6JtwQgFAL8mCvoAwAk3Ueq24ZR`
        );
        if (!res.ok) return [];
        const data = await res.json();
        const rows = data && data.Table;
        if (!Array.isArray(rows)) return [];
        return rows
            .map((r) => {
                const price = parseFloat(r.price);
                const date = new Date(String(r.Date || "").split(" ")[0]);
                return (isFinite(price) && price > 0 && !isNaN(date.getTime())) ? { date, price } : null;
            })
            .filter(Boolean)
            .sort((a, b) => a.date - b.date);
    } catch (e) {
        return []; // network failure or unexpected response shape
    }
}

// Given an ASCENDING {date,price}[] series (see fetchAnnualClosingPrices) and a fiscal year
// (e.g. fy=2024 -> FY ending 31-Mar-2024, the standard Indian-equities convention used
// throughout this file), returns the closing price on the LAST trading day on/before that
// fiscal year-end — the standard "annual closing price" — or null if the series has no data
// that far back (e.g. a company whose IPO postdates that FY, so it simply wasn't trading yet).
function annualClosingPriceForFY(series, fy) {
    const cutoff = new Date(Number(fy), 2, 31, 23, 59, 59); // 31-Mar of that FY
    let best = null;
    for (const pt of series) {
        if (pt.date <= cutoff) best = pt; // series is ascending -> keeps advancing to the latest <= cutoff
        else break;
    }
    return best ? best.price : null;
}

// Fetches face value for a company (used for the model's own "face_value" row/citation only —
// see fetchSharesOutInputs below for shares_out itself, which no longer derives from this).
// Returns { faceValue, dilutedShares } (either may be null if that specific field was missing/
// non-numeric) — or null entirely on a failed fetch (invalid fincode, network error).
async function fetchBasicInfo(fincode) {
    try {
        const res = await fetch("https://transcriptanalyser.com/goindiastock/new_basicinfo", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fincode, mode: "", exchange: "" }),
        });
        if (!res.ok) return null;
        const json = await res.json();
        const d = json && json.data;
        if (!d) return null;
        const num = (field) => {
            const v = d[field] && d[field].value;
            return typeof v === "number" && isFinite(v) && v > 0 ? v : null;
        };
        return { faceValue: num("facevalue"), dilutedShares: num("dilutedNOS") };
    } catch (e) {
        return null; // network failure or unexpected response shape
    }
}

// Fetches per-year figures from GIAStock_Data.dbo.sc_year_data (via the new
// /wallet/excel_shares_outstanding_inputs endpoint), used for two things that otherwise get
// guessed by the model: (1) shares_out = Market Cap ÷ Price for the latest ACTUAL (non-estimate)
// year — deliberately NOT PAT ÷ EPS, since the P&L's own EPS forecast formula is PAT ÷
// shares_out, which would make the two rows circularly dependent on each other; Market Cap and
// Price come from a completely unrelated relationship (Market Cap = Price × Shares), which
// breaks that circularity entirely; (2) debtor_days/inventory_days/payable_days HISTORICALS
// (Receivable/Inventory/Payable Days columns) — real reported figures, not something that
// should be guessed/derived for past years the way a forward-looking assumption is. Requests up
// to 20 years so enough ACTUAL (non-estimate) years survive after excluding any forecast/
// estimate rows the table also carries (see the "estimate" flag on each row) — HIST_YEARS (8)
// worth of real data is needed for the working-capital-days history, not just the 1 year
// shares_out needs. Returns the raw rows; picking the right year(s) is done by the caller — or
// [] on a failed fetch (invalid fincode, network error, endpoint not yet deployed).
async function fetchSharesOutInputs(fincode) {
    try {
        const res = await fetch(`${WALLET_BASE_URL}/excel_shares_outstanding_inputs?fincode=${encodeURIComponent(fincode)}&years=20`);
        if (!res.ok) return [];
        const json = await res.json();
        return Array.isArray(json && json.rows) ? json.rows : [];
    } catch (e) {
        return []; // network failure, endpoint not deployed yet, or unexpected response shape
    }
}

// Fetches this company's own historical trading-multiple distribution (P/E, P/B, EV/EBITDA —
// both trailing/reported and forward/consensus-estimate ("_est") variants, each with mean,
// median, and a ±1 standard-deviation band computed from its own daily trading history) — used
// to ground the Assumptions (Judgment) call's target_pe/target_ev_ebitda picks in the company's
// REAL historical trading range instead of a generic sector rule-of-thumb. The model has been
// observed defaulting to suspiciously round, "safe" multiples (e.g. a flat ~20x P/E regardless
// of company) without this.
//
// The full-period stats alone can be a misleading anchor if the multiple has been on a
// structural trend (compressing or expanding) rather than oscillating around a stable level — a
// full-history median silently blends pre- and post-trend regimes together. To make that trend
// visible instead of hidden, this also derives a RECENT-WINDOW (last ~6 months of trading days,
// by the graph's own "Date" field, not calendar-today) average/median from the same daily
// "graph" series the full-period stats are themselves computed from (each row's own pe_formula/
// ev_ebitda_formula field confirms these are MARKETCAP ÷ this company's OWN trailing or
// consensus-estimate PAT/EBITDA — i.e. this company's own recomputed history, not a peer-group
// average, regardless of what the response's "sector" field might suggest). The daily "graph"
// array itself (one entry per trading day, sometimes hundreds/thousands of points) is dropped
// after this — it's far more than needed for a target-multiple pick and would bloat the prompt
// for no benefit; only the full-period summary stats plus these derived recent-window stats are
// returned. Returns that combined stats object, or null on a failed fetch (invalid fincode,
// network error, or a response with no "key" object).
async function fetchValuationChart(fincode) {
    try {
        const res = await fetch("https://transcriptanalyser.com/gis/valuation-chart", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fincode: String(fincode), mode: "" }),
        });
        if (!res.ok) return null;
        const json = await res.json();
        if (!json || !json.key) return null;
        const { graph, ...stats } = json.key; // drop the large daily graph array after deriving recent-window stats below

        const RECENT_WINDOW_DAYS = 182; // ~6 months — mid-point of the "recent 6-12 month" trend window
        let recent = {};
        if (Array.isArray(graph) && graph.length) {
            const rows = graph.filter(r => r && r.Date).slice().sort((a, b) => new Date(a.Date) - new Date(b.Date));
            const latest = rows.length ? new Date(rows[rows.length - 1].Date) : null;
            const cutoff = latest ? new Date(latest.getTime() - RECENT_WINDOW_DAYS * 24 * 60 * 60 * 1000) : null;
            const recentRows = cutoff ? rows.filter(r => new Date(r.Date) >= cutoff) : [];
            const nums = (field) => recentRows.map(r => r[field]).filter(v => typeof v === "number" && isFinite(v));
            const avg = (arr) => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
            const median = (arr) => {
                if (!arr.length) return null;
                const s = [...arr].sort((a, b) => a - b);
                const mid = Math.floor(s.length / 2);
                return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
            };
            if (recentRows.length) {
                recent = {
                    recentWindowDays: recentRows.length,
                    recentFrom: recentRows[0].Date,
                    recentTo: recentRows[recentRows.length - 1].Date,
                    recent_avg_pe: avg(nums("pe")),
                    recent_median_pe: median(nums("pe")),
                    recent_avg_pe_est: avg(nums("pe_est")),
                    recent_median_pe_est: median(nums("pe_est")),
                    recent_avg_ev_ebitda: avg(nums("ev_ebitda")),
                    recent_median_ev_ebitda: median(nums("ev_ebitda")),
                    recent_avg_ev_ebitda_est: avg(nums("ev_ebitda_est")),
                    recent_median_ev_ebitda_est: median(nums("ev_ebitda_est")),
                };
            }
            // The single most recent trading day's own multiple — "today's" implied multiple, for
            // the LLM to explicitly compare its chosen reference stat against (see the rationale
            // instruction below).
            if (rows.length) {
                const latestRow = rows[rows.length - 1];
                recent.latestDate = latestRow.Date;
                recent.latest_pe = typeof latestRow.pe === "number" ? latestRow.pe : null;
                recent.latest_pe_est = typeof latestRow.pe_est === "number" ? latestRow.pe_est : null;
                recent.latest_ev_ebitda = typeof latestRow.ev_ebitda === "number" ? latestRow.ev_ebitda : null;
                recent.latest_ev_ebitda_est = typeof latestRow.ev_ebitda_est === "number" ? latestRow.ev_ebitda_est : null;
            }
        }
        return { ...stats, sector: json.sector, ...recent }; // "sector" is a top-level sibling of "key" in the response
    } catch (e) {
        return null; // network failure or unexpected response shape
    }
}

// ── Live current_price refresh (Valuation & DCF sheet) ──
// Keeps the "Current price" cell on an existing AI Financial Model reasonably fresh WITHOUT a
// full rebuild — reuses the SAME fincode/company ownership check refreshBtn's stale-model guard
// already relies on (getExistingModelOwner, defined below), re-fetches the latest EOD price for
// it via the same feed used at build time (fetchLatestPrice), and writes just that one cell.
// Deliberately reads the fincode back off the workbook's own stored property rather than
// re-parsing the sheet's title text and reverse-searching company_search for it — the property
// is already exact, so there's no reason to reintroduce fuzzy-text-match risk (ampersands,
// "Ltd." vs "Limited", near-duplicate company names) for something already available verbatim.
// Runs on a timer for as long as the taskpane stays open; self-stops once the model (or its
// ownership property) is gone, rather than polling forever in the background for nothing.
const PRICE_REFRESH_INTERVAL_MS = 5 * 60 * 1000; // 5 min — an EOD feed, not tick data, so polling faster buys nothing
let _priceRefreshTimer = null;

async function refreshCurrentPriceCell() {
    try {
        const owner = await getExistingModelOwner();
        if (!owner) {
            // No model left in this workbook (never built, or since deleted) — stop polling.
            if (_priceRefreshTimer) { clearInterval(_priceRefreshTimer); _priceRefreshTimer = null; }
            return;
        }
        const latest = await fetchLatestPrice(owner.fincode);
        if (!latest) return; // transient fetch failure — leave the cell alone, retry next tick

        await Excel.run(async (context) => {
            const wb = context.workbook;
            const sh = wb.worksheets.getItemOrNullObject("FM Valuation & DCF");
            sh.load("name");
            await context.sync();
            if (sh.isNullObject) return;
            const used = sh.getUsedRange();
            used.load("values");
            await context.sync();
            const grid = used.values;
            let priceRow = -1;
            for (let r = 0; r < grid.length; r++) {
                if (/current\s*price/i.test(String((grid[r] && grid[r][0]) || ""))) { priceRow = r; break; }
            }
            if (priceRow === -1) return; // row not found (sheet layout changed) — skip this tick
            sh.getRange(`B${priceRow + 1}`).values = [[latest.price]];
            await context.sync();
        });
        console.log(`[Live Price] Refreshed current_price for ${owner.companyName} (fincode ${owner.fincode}): Rs ${latest.price} (${latest.date})`);
    } catch (e) {
        console.warn("[Live Price] Refresh failed (will retry next interval):", e.message);
    }
}

// (Re)starts the polling timer — called once at add-in startup (resumes refreshing an existing
// model after the workbook/taskpane is reopened) and again after every handleBuildModel run
// (the fincode may have just changed to a different company).
function startPriceRefreshTimer() {
    if (_priceRefreshTimer) clearInterval(_priceRefreshTimer);
    _priceRefreshTimer = setInterval(refreshCurrentPriceCell, PRICE_REFRESH_INTERVAL_MS);
    refreshCurrentPriceCell(); // also refresh immediately rather than waiting a full interval
}

// AI Financial Model sheet names — kept in one place so handleBuildModel's own writer and the
// stale-model check in refreshBtn.onclick (which runs BEFORE handleBuildModel, on every download,
// model toggle on or off) always agree on what counts as "an existing model".
const FM_SHEET_NAMES = ["FM Assumptions", "FM Operational", "FM P&L", "FM Balance Sheet", "FM Capex & FCF", "FM Valuation & DCF", "FM Summary"];
const LEGACY_FM_SHEET_NAME = "Financial Model"; // pre-multi-sheet model name, still cleaned up on rebuild
// Workbook custom properties recording which company's data the CURRENT FM sheets were built
// against — since their historical cells are live formula links into Key Financials/Annual Data
// (see writeDataSheet's linkEntry branch), downloading a DIFFERENT company overwrites those source
// sheets and silently makes the old model display the new company's numbers under the old model's
// structure. This tag is what lets refreshBtn.onclick tell "just refreshing the same company" (safe,
// leave the model alone) apart from "switching companies" (unsafe, warn before overwriting).
const FM_FINCODE_PROPERTY = "GOIA_FM_Fincode";
const FM_COMPANY_PROPERTY = "GOIA_FM_Company";

// Reads which company (if any) the workbook's current AI Financial Model belongs to. Returns null
// if there's no model — either the property was never set, or its FM sheets were since deleted
// (e.g. manually by the user), in which case a leftover property is ignored rather than trusted.
async function getExistingModelOwner() {
    let owner = null;
    await Excel.run(async (context) => {
        const wb = context.workbook;
        wb.worksheets.load("items/name");
        const props = wb.properties.custom;
        props.load("items/key,items/value");
        await context.sync();
        const hasFmSheet = wb.worksheets.items.some(s => FM_SHEET_NAMES.includes(s.name));
        if (!hasFmSheet) return;
        const fincodeProp = props.items.find(p => p.key === FM_FINCODE_PROPERTY);
        if (!fincodeProp) return;
        const companyProp = props.items.find(p => p.key === FM_COMPANY_PROPERTY);
        owner = { fincode: String(fincodeProp.value), companyName: companyProp ? String(companyProp.value) : "another company" };
    });
    return owner;
}

// Deletes the existing FM sheets (and the legacy single-sheet model) plus the ownership properties.
async function deleteExistingModel() {
    await Excel.run(async (context) => {
        const wb = context.workbook;
        wb.worksheets.load("items/name");
        await context.sync();
        const existing = wb.worksheets.items.map(s => s.name);
        for (const nm of [...FM_SHEET_NAMES, LEGACY_FM_SHEET_NAME]) {
            if (existing.includes(nm)) wb.worksheets.getItem(nm).delete();
        }
        const props = wb.properties.custom;
        const fincodeItem = props.getItemOrNullObject(FM_FINCODE_PROPERTY);
        const companyItem = props.getItemOrNullObject(FM_COMPANY_PROPERTY);
        await context.sync();
        if (!fincodeItem.isNullObject) fincodeItem.delete();
        if (!companyItem.isNullObject) companyItem.delete();
        await context.sync();
    });
}

// Shows the "existing model belongs to a different company" warning. Resolves "delete" if the user
// chooses to delete the old model and continue, or "cancel" if they back out of the download entirely.
function showStaleModelWarning(ownerCompanyName) {
    return new Promise((resolve) => {
        const modal = document.getElementById("staleModelModal");
        const nameEl = document.getElementById("staleModelCompanyName");
        const del = document.getElementById("staleModelDelete");
        const cancel = document.getElementById("staleModelCancel");
        if (!modal || !del || !cancel) return resolve("delete"); // gate missing — don't block the download
        if (nameEl) nameEl.textContent = ownerCompanyName;
        const done = (result) => {
            modal.style.display = "none";
            del.onclick = null; cancel.onclick = null; modal.onclick = null;
            resolve(result);
        };
        del.onclick = () => done("delete");
        cancel.onclick = () => done("cancel");
        modal.onclick = (e) => { if (e.target === modal) done("cancel"); };
        modal.style.display = "flex";
    });
}

// Disables the AI Financial Model toggle (and unchecks it if it was on); re-enables it otherwise.
function setAiModelAvailability(sufficientBalance) {
    const check = document.getElementById("aiModelCheck");
    if (!check) return;
    check.disabled = !sufficientBalance;
    if (!sufficientBalance) check.checked = false;
}

// Minimum wallet balance needed to START a DataGPT query. Unlike the financial model — whose
// worst case is bounded by MAX_MODEL_COST_INR, so we can require the full amount up front — a
// chat turn is open-ended (cost depends on how many tool iterations the model needs), so this is
// a floor rather than a reservation: enough that a typical query is covered, not a guarantee.
const MIN_CHAT_BALANCE_INR = 3;

// Read back by the chat panel's send() so a query that finishes AFTER the balance ran dry does
// not re-enable the button in its finally block. Starts true so the composer is not dead in the
// window before the first balance fetch resolves; from then on refreshWalletBalance() fails CLOSED
// exactly like the AI-model toggle does.
let chatBalanceSufficient = true;

// Disables the DataGPT send button and reveals the notice above the composer. The textarea stays
// usable on purpose — a half-typed question should not be lost just because the balance dipped,
// and the notice already explains why the button is dead.
function setChatAvailability(sufficientBalance) {
    chatBalanceSufficient = !!sufficientBalance;
    const btn = document.getElementById("chatSend");
    const notice = document.getElementById("chatBalanceNotice");
    if (btn) btn.disabled = !chatBalanceSufficient;
    if (notice) notice.style.display = chatBalanceSufficient ? "none" : "block";
}

// Fetches the current user's wallet balance (sum of all three credit buckets, per backend
// convention), updates the pill in the header, and gates the AI Financial Model toggle on it.
// Fails CLOSED: a real account with an unverifiable balance (network error, non-200, etc.) gets
// the toggle disabled, same as a confirmed-insufficient balance — we should never let a real user
// generate a model we don't know they can afford. The one exemption is the ut@gmail.com
// certification bypass (UserId "test123"), which always fails this fetch since it isn't a real
// wallet user — it stays enabled unconditionally so Microsoft's reviewers can test the feature.
// Mirrors the balance onto every place it's shown — the header pill (#walletBalance) and, since
// the ribbon's Wallet button (view=wallet) opens the history modal directly, the balance row now
// embedded inside it too (see #walletHistoryModal in taskpane.html).
function setWalletBalanceText(text) {
    document.querySelectorAll("#walletBalance, .wallet-balance-display").forEach(el => { el.textContent = text; });
}
async function refreshWalletBalance() {
    if (!document.getElementById("walletBalance")) return;
    let userId = null;
    try {
        const user = JSON.parse(localStorage.getItem("user") || "{}");
        userId = user.UserId || null;
        if (!userId) { setWalletBalanceText("—"); setAiModelAvailability(false); setChatAvailability(false); return; }
        const isTestAccount = userId === TEST_BYPASS_USER_ID;
        if (isTestAccount) { setAiModelAvailability(true); setChatAvailability(true); }

        const res = await fetch("https://transcriptanalyser.com/wallet/user_credit_balance", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ user_id: String(userId) })
        });
        if (!res.ok) {
            setWalletBalanceText("—");
            if (!isTestAccount) { setAiModelAvailability(false); setChatAvailability(false); }
            return;
        }
        const data = await res.json();
        const c = data.credits || {};
        const total = (c.whatsapp_credit || 0) + (c.stockgpt_credit || 0) + (c.balance_credit || 0);
        setWalletBalanceText(`₹${total.toFixed(2)}`);
        setAiModelAvailability(isTestAccount || total >= MAX_MODEL_COST_INR);
        setChatAvailability(isTestAccount || total >= MIN_CHAT_BALANCE_INR);
    } catch (e) {
        console.warn("Wallet balance fetch failed:", e);
        setWalletBalanceText("—");
        if (userId !== TEST_BYPASS_USER_ID) { setAiModelAvailability(false); setChatAvailability(false); }
    }
}

// wallet.py's two excel_-prefixed endpoints, deployed alongside the pre-existing
// /wallet/user_credit_balance on the same backend.

//const WALLET_BASE_URL = "https://transcriptanalyser.com/wallet";
const WALLET_BASE_URL = "https://transcriptanalyser.com/wallet";
const WALLET_DEDUCT_URL = `${WALLET_BASE_URL}/excel_deduct_model_cost`;
const WALLET_DEDUCT_CHAT_URL = `${WALLET_BASE_URL}/excel_deduct_chat_cost`;

// Pure usage tracking (login / download-data clicks) — SQL-only, never shown in the add-in's own
// UI, queried directly for usage analysis. Fire-and-forget on purpose: a logging failure must never
// delay or surface an error for the real action (login, download) it's just observing. The
// ut@gmail.com certification account's UserId ("test123") isn't numeric, so it'll fail this call's
// validation on the backend — harmless, since that account isn't a real user to track anyway.
function logUsageEvent(userId, eventType) {
    const numericUserId = Number(userId);
    if (!Number.isFinite(numericUserId)) return;
    fetch(`${WALLET_BASE_URL}/excel_log_event`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ user_id: numericUserId, event_type: eventType }),
    }).catch(() => { /* best-effort only */ });
}

// Enterprise accounts are keyed by enterprise_id in wallet_balance, not user_id — check_membership's
// response (already stored in full under localStorage["membership"] at login) carries enterprise_id
// whenever is_enterprise_member is true, confirmed live: {"is_enterprise_member":true,"enterprise_id":25,...}.
// Shared by deductModelCost and fetchWalletHistory so both branch the same way.
function getWalletIdentity() {
    const user = JSON.parse(localStorage.getItem("user") || "{}");
    const membership = JSON.parse(localStorage.getItem("membership") || "{}");
    const isEnterprise = membership.is_enterprise_member === true && membership.enterprise_id != null;
    return {
        userId: user.UserId || null,
        isEnterprise,
        enterpriseId: isEnterprise ? Number(membership.enterprise_id) : null,
    };
}

// Briefly shows "Cost for <Company> Model: X Rs." below the Download Data button so the user
// sees exactly what a deduction cost, then hides it again. Non-blocking, purely cosmetic.
function showWalletDeductBanner(companyName, costInr) {
    const el = document.getElementById("walletDeductBanner");
    if (!el || costInr == null) return;
    el.textContent = `Cost for ${companyName || "this"} Model: ${Math.round(Number(costInr))} Rs.`;
    el.style.display = "block";
    clearTimeout(showWalletDeductBanner._t);
    showWalletDeductBanner._t = setTimeout(() => {
        el.style.display = "none";
    }, 4500);
}

// Deducts the OpenRouter cost of a financial-model generation from the user's wallet, then
// refreshes the balance pill live (no page reload needed). Fire-and-forget from the caller —
// never blocks the "model complete" status on a wallet round-trip.
async function deductModelCost({ fincode, companyName, costUsd, usedFallback }) {
    try {
        const { userId, isEnterprise, enterpriseId } = getWalletIdentity();
        if (!userId) return;
        const res = await fetch(WALLET_DEDUCT_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                user_id: isEnterprise ? null : (Number(userId) || null),
                enterprise_id: enterpriseId,
                fincode: Number(fincode),
                cost_usd: costUsd,
                used_fallback: !!usedFallback,
            }),
        });
        if (!res.ok) {
            console.warn("[Wallet] excel_deduct_model_cost failed:", res.status, await res.text().catch(() => ""));
            return;
        }
        const data = await res.json().catch(() => null);
        showWalletDeductBanner(companyName, data?.cost_inr);
        await refreshWalletBalance();
    } catch (e) {
        console.warn("[Wallet] excel_deduct_model_cost request failed:", e);
    }
}

// Same idea as deductModelCost above, for DataGPT chat queries — deducts the SUMMED OpenRouter
// cost across every turn a single question took (tool-calling round trips included) and logs it
// against the question asked instead of a fincode (see excel_deduct_chat_cost/excel_chat_cost_log
// in backend/routers/wallet.py). Called once per chat turn, after the model's final answer is in
// (see send() below). Fire-and-forget from the caller, same as deductModelCost.
async function deductChatCost({ question, costUsd }) {
    try {
        const { userId, isEnterprise, enterpriseId } = getWalletIdentity();
        if (!userId) return;
        const res = await fetch(WALLET_DEDUCT_CHAT_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                user_id: isEnterprise ? null : (Number(userId) || null),
                enterprise_id: enterpriseId,
                question,
                cost_usd: costUsd,
            }),
        });
        if (!res.ok) {
            console.warn("[Wallet] excel_deduct_chat_cost failed:", res.status, await res.text().catch(() => ""));
            return;
        }
        await refreshWalletBalance();
    } catch (e) {
        console.warn("[Wallet] excel_deduct_chat_cost request failed:", e);
    }
}

// ── Wallet history ──
// Two logs feed the wallet: /wallet/excel_cost_history (AI financial models, keyed by fincode)
// and /wallet/excel_cost_history's DataGPT counterpart /wallet/excel_chat_cost_history (chat
// queries, keyed by the question text). They're separate endpoints because their rows have
// different shapes, so we fetch both in parallel, normalise each into a common {kind,label,...}
// record, then merge and sort by timestamp so the user sees one chronological spend list.
// Model rows additionally resolve their fincode to a company name via company_search (passing a
// fincode as searchtxt returns that company, same as passing a name), cached per-call so repeat
// fincodes don't refetch. Renders into every element matching `selector` — defaults to the full
// 20-entry list (.wallet-history-list, used by the in-panel history popup and, once clicked open,
// the Profile view's own wallet card); pass (5, ".wallet-history-preview") for the Profile view's
// upfront 5-entry teaser instead, so opening that view doesn't fetch 20 rows just to show 5.
async function fetchWalletHistory(limit = 20, selector = ".wallet-history-list") {
    const listEls = Array.from(document.querySelectorAll(selector));
    if (!listEls.length) return;
    const setAll = (html) => listEls.forEach(el => { el.innerHTML = html; });
    setAll("Loading…");

    const esc = (s) => String(s == null ? "" : s)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

    try {
        const { userId, isEnterprise, enterpriseId } = getWalletIdentity();
        if (!userId) { setAll("Not logged in."); return; }
        const qs = isEnterprise ? `enterprise_id=${enterpriseId}` : `user_id=${encodeURIComponent(userId)}`;

        // Each fetch is independently guarded: a missing/failing chat log shouldn't blank out the
        // model history (or vice versa) — we just show whichever side came back.
        const grab = async (path) => {
            try {
                const res = await fetch(`${WALLET_BASE_URL}/${path}?${qs}&limit=${limit}`);
                if (!res.ok) return null;
                const data = await res.json();
                return data.entries || [];
            } catch (e) {
                console.warn(`[Wallet] ${path} fetch failed:`, e);
                return null;
            }
        };
        const [modelEntries, chatEntries] = await Promise.all([
            grab("excel_cost_history"),
            grab("excel_chat_cost_history"),
        ]);

        if (modelEntries === null && chatEntries === null) { setAll("Couldn't load history."); return; }

        const merged = [
            ...(modelEntries || []).map(e => ({
                kind: "model",
                fincode: e.fincode,
                label: null,                       // filled in below once the fincode resolves
                costInr: e.cost_inr,
                createdUtc: e.created_utc,
                note: e.used_fallback ? "fallback model" : "",
            })),
            ...(chatEntries || []).map(e => ({
                kind: "chat",
                fincode: null,
                label: e.question || "(no question recorded)",
                costInr: e.cost_inr,
                createdUtc: e.created_utc,
                note: "",
            })),
        ];
        if (!merged.length) { setAll("No charges yet."); return; }

        // Both endpoints already return newest-first, but interleaving two lists needs a real sort.
        // Rows with no timestamp sink to the bottom rather than jumping to the top as epoch 0.
        merged.sort((a, b) => {
            const ta = a.createdUtc ? new Date(a.createdUtc).getTime() : -Infinity;
            const tb = b.createdUtc ? new Date(b.createdUtc).getTime() : -Infinity;
            return tb - ta;
        });
        // Each side was capped at `limit`, so the merge can hold up to 2×limit — re-cap so "last 5"
        // means the 5 most recent charges overall, not 5 of each.
        const entries = merged.slice(0, limit);

        const nameCache = new Map();
        const resolveCompanyName = async (fincode) => {
            if (nameCache.has(fincode)) return nameCache.get(fincode);
            let name = `Fincode ${fincode}`;
            try {
                const r = await fetch("https://transcriptanalyser.com/goindiastock/company_search", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ searchtxt: String(fincode) })
                });
                const d = await r.json();
                const match = d?.data?.[0];
                if (match?.label) name = match.label;
            } catch (e) { /* keep fincode fallback */ }
            nameCache.set(fincode, name);
            return name;
        };

        let rowsHtml = "";
        for (const entry of entries) {
            const isChat = entry.kind === "chat";
            const raw = isChat ? entry.label : await resolveCompanyName(entry.fincode);
            // Questions are free text and can run long; clamp the visible label but keep the full
            // text on hover. Company names are short enough to pass through untouched.
            const shown = raw.length > 90 ? `${raw.slice(0, 90)}…` : raw;
            const dt = entry.createdUtc ? new Date(entry.createdUtc) : null;
            const meta = [dt ? dt.toLocaleString() : "", entry.note].filter(Boolean).join(" · ");
            const tag = isChat
                ? `<span style="flex:none; font-size:9.5px; font-weight:700; letter-spacing:.03em; text-transform:uppercase; color:#5b7fa6; background:#eef3f8; border-radius:3px; padding:1px 5px;">DataGPT</span>`
                : `<span style="flex:none; font-size:9.5px; font-weight:700; letter-spacing:.03em; text-transform:uppercase; color:#7a6a3c; background:#f6f1e4; border-radius:3px; padding:1px 5px;">Model</span>`;
            rowsHtml += `
                <div style="padding:8px 0; border-bottom:1px solid #eee;">
                    <div style="display:flex; justify-content:space-between; gap:8px;">
                        <span style="display:flex; align-items:baseline; gap:6px; min-width:0;">
                            ${tag}
                            <span style="min-width:0; overflow-wrap:anywhere;" title="${esc(raw)}">${esc(shown)}</span>
                        </span>
                        <span style="font-weight:700; color:#173760; white-space:nowrap;">₹${Number(entry.costInr).toFixed(2)}</span>
                    </div>
                    <div style="font-size:10.5px; color:#999; margin-top:2px;">${esc(meta)}</div>
                </div>
            `;
        }
        setAll(rowsHtml);
    } catch (e) {
        console.warn("[Wallet] history fetch failed:", e);
        setAll("Couldn't load history.");
    }
}

// webpack scopes top-level functions inside its own module wrapper, not onto window — so
// taskpane.html's separate inline <script> can't see these by bare name even though a
// `typeof x === "function"` guard makes that failure silent instead of throwing. Explicitly
// exposing the functions taskpane.html actually calls across that boundary.
window.refreshWalletBalance = refreshWalletBalance;
window.fetchWalletHistory = fetchWalletHistory;
window.showAddinUI = showAddinUI;
window.updateLogoutBtnVisibility = updateLogoutBtnVisibility;

async function checkLocalSession() {
    const savedUser = localStorage.getItem("user");
    const savedMembership = localStorage.getItem("membership");

    if (savedUser && savedMembership) {
        const user = JSON.parse(savedUser);
        const membership = JSON.parse(savedMembership);

        try {
            // Check membership status from API
            const res = await fetch("https://transcriptanalyser.com/payment/check_membership", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ user_id: user.UserId })
            });

            if (!res.ok) throw new Error("Failed to verify membership");

            const updatedMembership = await res.json();

            // Only proceed if user has GIS Pro or Enterprise access
            if (updatedMembership.is_giin_pro === true || updatedMembership.is_admin === true || updatedMembership.is_enterprise_member === true) {
                // Keep stored membership in sync with the live API — this used to be fetched here
                // and discarded, leaving localStorage["membership"] permanently frozen at whatever
                // was captured at the original login (e.g. missing enterprise_id if that wasn't
                // part of the response back then, or just generally stale).
                localStorage.setItem("membership", JSON.stringify(updatedMembership));

                // Show member content
                document.getElementById("memberContent").style.display = "block";
                showAddinUI();
                updateLogoutBtnVisibility();

                // Update profile display
                const profileIcon = document.getElementById("profileIcon");
                const username = document.getElementById("username");
                if (user.FullName) {
                    username.textContent = user.FullName;
                    profileIcon.textContent = user.FullName
                        .split(" ")
                        .map(n => n[0])
                        .join("")
                        .toUpperCase();
                }
                refreshWalletBalance();

                return true; // Session valid
            } else {
                // Not a member → clear session
                localStorage.removeItem("user");
                localStorage.removeItem("membership");
            }
        } catch (err) {
            console.error("Error checking membership session:", err);
            localStorage.removeItem("user");
            localStorage.removeItem("membership");
        }
    }

    // No valid session → show login
    document.getElementById("loginScreen").style.display = "flex";
    return false;
}

// Usage example
window.addEventListener("DOMContentLoaded", () => {
    checkLocalSession();
    // Keeps the balance pill honest even when something else (another product, another
    // session) deducts from the same wallet_balance row — not just after a model this
    // add-in itself generated. refreshWalletBalance() already no-ops safely if not logged in.
    setInterval(refreshWalletBalance, 60000);
});



document.getElementById("logoutBtn").addEventListener("click", () => {
    localStorage.removeItem("user");
    localStorage.removeItem("membership");

    document.getElementById("memberContent").style.display = "none";
    document.getElementById("addinUI").style.display = "none";
    document.getElementById("logoutBtn").style.display = "none";
    document.getElementById("loginScreen").style.display = "flex";
});

let lastMatchedCompany = null;

async function fetchFinancialData() {
    try {
        let cellValue;

        await Excel.run(async (context) => {
            let sheet;

            // Case 1: Try to get "Key Financials" sheet
            try {
                sheet = context.workbook.worksheets.getItem("Key Financials");
                sheet.load("name");
                await context.sync();
            } catch (err) {
                console.warn("Key Financials sheet not found. First-time user?");
                document.getElementById("dataStatus").textContent = "Empty company data (Ignore if first time user)";
                sheet = null;
            }

            if (sheet) {
                // Try to read A1
                const range = sheet.getRange("A1");
                range.load("values");
                await context.sync();
                cellValue = range.values[0][0];
            }

            // Case 2: If no A1 value, fallback to Index sheet A4
            if (!cellValue) {
                let indexSheet;
                try {
                    indexSheet = context.workbook.worksheets.getItem("Index");
                    const fallbackRange = indexSheet.getRange("A4");
                    fallbackRange.load("values");
                    await context.sync();
                    cellValue = fallbackRange.values[0][0];
                } catch (err) {
                    console.warn("Index sheet or A4 not found. Cannot determine company.");
                }
            }
        });

        if (!cellValue) {
            document.getElementById("dataStatus").textContent = "Empty company data (Ignore if first time user)";
            document.getElementById("dataStatus").style.color = "orange";
            return;
        }
        // Step 2: Call company_search API
        const searchRes = await fetch("https://transcriptanalyser.com/goindiastock/company_search", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ searchtxt: cellValue })
        });
        const searchData = await searchRes.json();

        if (!searchData.data || !searchData.data[0] || !searchData.data[0].value)
            throw new Error("No company match in search");
        const companyValue = searchData.data[0].value;

        // Step 3: Get full company list
        const compRes = await fetch("https://transcriptanalyser.com/operational/companies_excel");
        const companies = await compRes.json();

        // Step 4: Match fincode
        const matchedCompany = companies.find(c => c.fincode.toString() === companyValue.toString());
        if (!matchedCompany) throw new Error("No matching company found in list");

        lastMatchedCompany = matchedCompany; // ✅ save globally

        const fincode = matchedCompany.fincode;
        const sectorType = matchedCompany.sector_type;
        // Step 5: Fetch Profit/Loss
        const plRes = await fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fincode, mode: "S", sector: sectorType, sheet: `QProfitLossIND` })
        });
        if (!plRes.ok) throw new Error("Failed to fetch Profit/Loss data");
        const plData = await plRes.json();

        // Step 6: Compare latest column in Quarterly Data sheet
        await Excel.run(async (context) => {
            try {
                const sheet = context.workbook.worksheets.getItem("Quarterly Data");

                // Load the used range and its displayed text
                const usedRange = sheet.getUsedRange();
                usedRange.load("text");
                await context.sync();

                // Row 5 is index 4
                const rowValues = usedRange.text[4];

                if (!rowValues) {
                    console.log("Row 5 is empty according to Excel API");
                    const statusDiv = document.getElementById("dataStatus");
                    statusDiv.textContent = "Quarterly Data sheet row 5 appears empty";
                    statusDiv.style.color = "orange";
                    return;
                }

                // Remove empty cells
                const headers = rowValues.map(h => h ? h.toString().trim() : null).filter(h => h);
                const latestSheetHeader = headers[headers.length - 1];

                // --- Step 1: Parse API headers and find the latest ---
                const apiHeaders = plData.column
                .map(c => c.header || c.accessorKey)
                .filter(h => {
                    if (!h) return false;
                    if (h === 'Parameter') return false;
                    // Skip non-date headers like "QProfitLossIND"
                    const match = /^[A-Z][a-z]{2}\d{4}$/.test(h); // e.g., Jun2025
                    return match;
                })
                .map(h => {
                    const monthStr = h.slice(0,3); // Jun
                    const year = parseInt(h.slice(3)); // 2025
                    const month = new Date(`${monthStr} 1, 2000`).getMonth();
                    return { month, year, original: h };
                });

                if (!apiHeaders.length) {
                    console.log("No API headers found");
                    const statusDiv = document.getElementById("dataStatus");
                    statusDiv.textContent = "No Profit/Loss data returned from API";
                    statusDiv.style.color = "orange";
                    return;
                }

                // Find the latest API header by date
                const latestApi = apiHeaders.reduce((a, b) => {
                    if (b.year > a.year) return b;
                    if (b.year === a.year && b.month > a.month) return b;
                    return a;
                });

                // --- Step 2: Parse Excel header ---
                function parseExcelHeader(header) {
                    const [monthStr, yearStr] = header.split("-");
                    const month = new Date(`${monthStr} 1, 2000`).getMonth();
                    const year = 2000 + parseInt(yearStr);
                    return { month, year };
                }

                const excelH = parseExcelHeader(latestSheetHeader);
                const apiH = latestApi;

                // --- Step 3: Compare and update UI ---
                const statusDiv = document.getElementById("dataStatus");
                const refreshBtn = document.getElementById("refreshStatusBtn");

                // Enable refresh only if new data available
                if (excelH.month !== apiH.month || excelH.year !== apiH.year) {
                    statusDiv.textContent = `New data available for ${matchedCompany.CompName}: ${apiH.original}`;
                    statusDiv.style.color = "green";

                    // Enable refresh button only when new data is available
                    refreshBtn.disabled = false;
                } else {
                    statusDiv.textContent = `Sheet data is up-to-date for ${matchedCompany.CompName}`;
                    statusDiv.style.color = "blue";

                    // Disable refresh button otherwise
                    refreshBtn.disabled = true;
                }

            } catch (err) {
                console.error("Excel.run error:", err);
                const statusDiv = document.getElementById("dataStatus");
                statusDiv.textContent = "Error reading Quarterly Data sheet";
                statusDiv.style.color = "red";
            }
        });




    } catch (err) {
        console.error(err);
        const statusDiv = document.getElementById("dataStatus");
        statusDiv.textContent = `Error: ${err.message}`;
        statusDiv.style.color = "red";
    }
}

// Initial load
Office.onReady(() => {
    fetchFinancialData();
});



// ── Ribbon view-switching ──────────────────────────────────────────────────────────────────
// The ribbon's Download Data / Ask DataGPT / Side Panel buttons are all plain ShowTaskpane
// actions (see manifest.xml) pointing at this same taskpane.html with a different ?view= query
// string. ShowTaskpane always re-navigates the pane to its target URL — confirmed even when the
// pane is already open — so this script runs fresh on every one of those clicks; the query
// string alone is enough state, no roamingSettings or live in-memory flag needed.
//
// "download" and "datagpt" both narrow #addinUI down to an ALLOWLIST of its own direct-child
// sections (verified against the actual DOM — every section below really is a direct child, not
// nested inside a shared wrapper), hiding everything else. "main" (the default — no ?view= at
// all, or an unrecognized value) leaves everything exactly as it already renders today.
const RIBBON_VIEW_ALLOWLIST = {
    datagpt: ["chatFeature"],
    download: [
        "companyDataControls", "companyDropdownWrap", "includeOptions", "buttons",
        "walletDeductBanner", "aiStatusBox", "warningMsg", "infoBox",
        "disclaimerModal", "staleModelModal",
    ],
    profile: ["profileSection", "profileExtras"],
    feedback: ["feedbackSection"],
};
// Several login-success paths below unconditionally reveal #logoutBtn once #addinUI itself
// becomes visible — this runs that reveal through the current ribbon view's allowlist instead,
// so a "download"/"datagpt" view stays free of it even after an async login check resolves
// later than applyRibbonView() already did.
function updateLogoutBtnVisibility() {
    const btn = document.getElementById("logoutBtn");
    if (!btn) return;
    btn.style.display = RIBBON_VIEW_ALLOWLIST[document.body.dataset.view] ? "none" : "block";
}

// Same problem as updateLogoutBtnVisibility above, on #addinUI itself: several login-success paths
// unconditionally set it to display:block once login is confirmed — which runs through this
// instead, since the "profile" view needs #addinUI to be a flex column (for its even gap between
// the profile/wallet/disclaimer cards, see the scoped CSS in taskpane.html) rather than plain block.
function showAddinUI() {
    const el = document.getElementById("addinUI");
    if (!el) return;
    el.style.display = document.body.dataset.view === "profile" ? "flex" : "block";
}

function applyRibbonView() {
    let view = "main";
    try { view = new URLSearchParams(location.search).get("view") || "main"; } catch (e) { /* malformed query string */ }
    document.body.dataset.view = view;
    const keep = RIBBON_VIEW_ALLOWLIST[view];
    if (!keep) return view; // "main" (or anything unrecognized) — show everything, as today

    const addinUI = document.getElementById("addinUI");
    if (!addinUI) return view;
    Array.from(addinUI.children).forEach(child => {
        if (!child.id || !keep.includes(child.id)) child.style.display = "none";
    });

    if (view === "download") {
        // Search bar reads first, data-source checkboxes sit below it — download view only,
        // the default "main" view keeps its original checkboxes-above-search order untouched.
        const dropdownWrap = document.getElementById("companyDropdownWrap");
        const dataControls = document.getElementById("companyDataControls");
        if (dropdownWrap && dataControls) dropdownWrap.insertAdjacentElement("afterend", dataControls);
        // companyDropdownWrap's own margin-top:12px (inline, set in HTML) exists to clear the
        // checkboxes that used to sit above it — now that it's the first visible element in this
        // view, that reserved gap is dead space; zero it (inline style, so only JS can override
        // it — a stylesheet rule can't win against an inline declaration).
        if (dropdownWrap) dropdownWrap.style.marginTop = "0";
    }

    // The ribbon's Profile button consolidates the old separate Wallet/Disclaimer/Profile buttons
    // into one view: #profileSection (the existing name/avatar pill, restyled into a small centered
    // card by the scoped CSS above) plus #profileExtras, a clickable wallet card and disclaimer
    // card (hidden by default so they don't leak into the main view). Each card opens the SAME
    // .dz-overlay modal the small in-panel wallet-history icon / disclaimer link already use
    // elsewhere in this file, so there's only one place that actually renders the full content.
    // refreshWalletBalance/fetchWalletHistory are declared further down but hoisted, so they're
    // safe to call here.
    if (view === "profile") {
        // #profileSection's markup sets position:absolute inline (it floats over the header in
        // every other view) — a stylesheet rule can never beat that, only re-setting the same
        // inline property can, same fix as companyDropdownWrap's margin above. Without this it
        // stays out of #addinUI's normal flow and visually overlaps #profileExtras below it.
        const profileSection = document.getElementById("profileSection");
        if (profileSection) profileSection.style.position = "static";

        const extras = document.getElementById("profileExtras");
        if (extras) extras.style.display = "flex"; // activates #profileExtras's flex-column gap
        if (typeof refreshWalletBalance === "function") refreshWalletBalance();
        if (typeof fetchWalletHistory === "function") fetchWalletHistory(5, ".wallet-history-preview");

        const walletCard = document.getElementById("profileWalletCard");
        const walletModal = document.getElementById("walletHistoryModal");
        if (walletCard && walletModal) {
            walletCard.addEventListener("click", () => {
                walletModal.style.display = "flex";
                if (typeof fetchWalletHistory === "function") fetchWalletHistory();
            });
        }
        const disclaimerCard = document.getElementById("profileDisclaimerCard");
        const disclaimerModal = document.getElementById("disclaimerInfoModal");
        if (disclaimerCard && disclaimerModal) {
            disclaimerCard.addEventListener("click", () => {
                disclaimerModal.style.display = "flex";
            });
        }
    }

    // Ribbon's Feedback view — a plain form (see #feedbackSection in taskpane.html), not a modal.
    // Username is read-only, auto-filled from the same localStorage "user" entry every other
    // profile/account display reads from; company is deliberately free text (not tied to the
    // company catalog), since feedback may be about something not in it, or general.
    if (view === "feedback") {
        const section = document.getElementById("feedbackSection");
        if (section) section.style.display = "block";

        const usernameEl = document.getElementById("feedbackUsername");
        let currentUser = {};
        try { currentUser = JSON.parse(localStorage.getItem("user") || "{}"); } catch (e) { /* not logged in */ }
        if (usernameEl) usernameEl.textContent = currentUser.FullName || "—";

        const submitBtn = document.getElementById("feedbackSubmit");
        const statusEl = document.getElementById("feedbackStatus");
        const companyEl = document.getElementById("feedbackCompany");
        const textEl = document.getElementById("feedbackText");
        if (submitBtn && statusEl && companyEl && textEl) {
            submitBtn.addEventListener("click", async () => {
                const feedback = textEl.value.trim();
                if (!feedback) {
                    statusEl.className = "fb-status err";
                    statusEl.textContent = "Please enter your feedback before submitting.";
                    statusEl.style.display = "block";
                    return;
                }
                submitBtn.disabled = true;
                statusEl.style.display = "none";
                try {
                    // Same one-or-the-other branching every wallet endpoint uses (excel_deduct_model_cost,
                    // excel_deduct_chat_cost) — an enterprise member's feedback should be attributed to
                    // enterprise_id, not their own user_id, consistent with how their usage is billed.
                    const { userId, isEnterprise, enterpriseId } = getWalletIdentity();
                    const res = await fetch(`${WALLET_BASE_URL}/excel_submit_feedback`, {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({
                            user_id: isEnterprise ? null : (Number(userId) || null),
                            enterprise_id: enterpriseId,
                            username: currentUser.FullName || "Unknown",
                            company: companyEl.value.trim() || "General",
                            feedback,
                        }),
                    });
                    if (!res.ok) throw new Error(`Server returned ${res.status}`);
                    statusEl.className = "fb-status ok";
                    statusEl.textContent = "Thanks — your feedback has been submitted.";
                    statusEl.style.display = "block";
                    companyEl.value = "";
                    textEl.value = "";
                } catch (e) {
                    console.warn("[Feedback] submit failed:", e);
                    statusEl.className = "fb-status err";
                    statusEl.textContent = "Couldn't submit feedback — please try again.";
                    statusEl.style.display = "block";
                } finally {
                    submitBtn.disabled = false;
                }
            });
        }
    }

    return view;
}

Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) return;

    const ribbonView = applyRibbonView();

    // Resume live current_price polling for whatever model (if any) already exists in this
    // workbook — e.g. the taskpane/workbook was just reopened. No-ops harmlessly (and self-stops)
    // if there's no model yet; handleBuildModel restarts this itself once a build finishes.
    startPriceRefreshTimer();

    const refreshBtn = document.getElementById("refreshBtn"); // download data
    const refreshStatusBtn = document.getElementById("refreshStatusBtn"); // new data refresh
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
        if (!fincode) return showWarning("Select a company first.");

        const url = `https://www.goindiastocks.com/companyinfo/${encodeURIComponent(fincode)}`;
        window.open(url, "_blank");
    };

    // Disclaimer gate — resolves true if the user accepts, false otherwise.
    // Consent is remembered after the first "I Agree" so regular users aren't
    // re-prompted every time; the text stays reachable via the "Disclaimer" link.
    const MODEL_DISCLAIMER_KEY = "goia_modelDisclaimerAgreed";
    function showDisclaimerGate() {
        if (localStorage.getItem(MODEL_DISCLAIMER_KEY)) return Promise.resolve(true);
        return new Promise((resolve) => {
            const modal = document.getElementById("disclaimerModal");
            const agree = document.getElementById("disclaimerAgree");
            const cancel = document.getElementById("disclaimerCancel");
            if (!modal || !agree || !cancel) return resolve(true); // gate missing — don't block
            const done = (result) => {
                modal.style.display = "none";
                agree.onclick = null; cancel.onclick = null;
                modal.onclick = null;
                resolve(result);
            };
            agree.onclick = () => { localStorage.setItem(MODEL_DISCLAIMER_KEY, "1"); done(true); };
            cancel.onclick = () => done(false);
            modal.onclick = (e) => { if (e.target === modal) done(false); };
            modal.style.display = "flex";
        });
    }

    refreshBtn.onclick = async () => {
        const activeType = document.querySelector(".data-tab.active")?.dataset.type || "company";
        if (activeType === "sector") { handleSectorRefresh(); return; }
        if (activeType === "compare") { handleCompare(); return; }

        // Company tab — financial data is always downloaded; the financial model is optional.
        const toggle = document.getElementById("dropdownToggle");
        if (!toggle.dataset.value) return;

        // If the workbook has an AI Financial Model built for a DIFFERENT company, downloading now
        // will overwrite Key Financials/Annual/Quarterly Data — and that model's historical cells
        // are live formula links into those sheets, so it would silently start showing the NEW
        // company's numbers under the OLD model's structure. Checked regardless of whether the AI
        // Model toggle is on for THIS download: even a same-download rebuild can fail after Key
        // Financials is already overwritten (e.g. the empty-sheet guard in handleBuildModel), which
        // would otherwise leave the old model corrupted with no warning ever shown.
        const existingModel = await getExistingModelOwner();
        if (existingModel && existingModel.fincode !== toggle.dataset.value) {
            const choice = await showStaleModelWarning(existingModel.companyName);
            if (choice === "cancel") return;
            await deleteExistingModel();
        }

        // Financial data always downloads; also build the model if the user asked for it.
        // NOTE: the data-source selection modal is DISABLED for now (each company has only
        // a few sources, so making the user tick boxes isn't worth the extra step). The
        // modal code is kept below, commented out. To re-enable, swap the block below for:
        //     if (doModel) { openModelModal(); return; }
        const doModel = document.getElementById("aiModelCheck").checked;

        // The AI financial model is gated by a disclaimer the user must accept first.
        if (doModel) {
            const agreed = await showDisclaimerGate();
            if (!agreed) return; // declined — don't download or build
        }

        try {
            const currentUser = JSON.parse(localStorage.getItem("user") || "{}");
            logUsageEvent(currentUser.UserId, "download_data");
        } catch (e) { /* logging is best-effort only */ }

        const originalText = refreshBtn.textContent;
        refreshBtn.disabled = true;
        try {
            refreshBtn.textContent = "Downloading data...";
            await handleRefresh();
            if (doModel) {
                refreshBtn.textContent = "Building model...";
                await handleBuildModel(); // no selection -> all sources included
            }
        } catch (err) {
            console.error("Download error:", err);
        } finally {
            refreshBtn.textContent = originalText;
            refreshBtn.disabled = false;
        }
    };

    /* --- Financial Model data-selection modal (DISABLED for now; kept for re-enabling) ---
    const modelModal = document.getElementById("modelDataModal");
    let opMetricsFetched = false; // whether the operational metric list has been loaded

    function openModelModal() {
        const toggle = document.getElementById("dropdownToggle");
        document.getElementById("modelModalCompany").textContent = toggle.value || "this company";

        // Reset operational sub-list (the company may have changed since last open).
        opMetricsFetched = false;
        const list = document.getElementById("opMetricList");
        list.style.display = "none";
        list.innerHTML = `<div class="fm-loading">Loading operational metrics…</div>`;
        const expandBtn = document.getElementById("opExpandBtn");
        expandBtn.setAttribute("aria-expanded", "false");
        expandBtn.textContent = "▸ choose";

        // Reset every source toggle to checked.
        modelModal.querySelectorAll(".fm-source-toggle").forEach(cb => { cb.checked = true; });
        modelModal.style.display = "flex";
    }
    function closeModelModal() { modelModal.style.display = "none"; }

    // Fetch the operational data for the selected company and list the actual METRIC
    // names (grouped under their section) so the user can pick exactly which metrics to
    // feed the model. Uses the same dashboard API the model itself reads.
    const esc = (s) => String(s)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

    async function loadOperationalMetrics() {
        const toggle = document.getElementById("dropdownToggle");
        const fincode = toggle.dataset.value;
        const list = document.getElementById("opMetricList");
        try {
            const res = await fetch("https://transcriptanalyser.com/pms/get_dashboard", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ fincode: parseInt(fincode) })
            });
            const data = await res.json();

            const groupsHtml = (data.table_data || []).map(section => {
                const metrics = (section.rows || []).filter(r => r && r.metric_name);
                if (!metrics.length) return "";
                const items = metrics.map(r => {
                    const unit = r.unit ? ` <span class="fm-unit">(${esc(r.unit)})</span>` : "";
                    return `<label class="fm-suboption"><input type="checkbox" class="fm-op-metric" value="${esc(r.metric_name)}" checked><span>${esc(r.metric_name)}${unit}</span></label>`;
                }).join("");
                return `<div class="fm-op-group">
                    <label class="fm-op-group-head"><input type="checkbox" class="fm-op-group-toggle" checked><span>${esc(section.section_name || "Section")}</span></label>
                    ${items}
                </div>`;
            }).filter(Boolean).join("");

            if (!groupsHtml) {
                list.innerHTML = `<div class="fm-loading">No operational metrics found for this company.</div>`;
                return;
            }
            list.innerHTML = groupsHtml;

            // Wire each section header to toggle all its metrics, and keep the header
            // in sync (checked / unchecked / indeterminate) as individual metrics change.
            list.querySelectorAll(".fm-op-group").forEach(group => {
                const head = group.querySelector(".fm-op-group-toggle");
                const metrics = Array.from(group.querySelectorAll(".fm-op-metric"));
                head.addEventListener("change", () => { metrics.forEach(m => { m.checked = head.checked; }); });
                metrics.forEach(m => m.addEventListener("change", () => {
                    const n = metrics.filter(x => x.checked).length;
                    head.checked = n === metrics.length;
                    head.indeterminate = n > 0 && n < metrics.length;
                }));
            });
        } catch (e) {
            console.warn("[Model modal] operational metrics fetch failed:", e);
            list.innerHTML = `<div class="fm-loading">Couldn't load operational metrics — all of them will be included.</div>`;
        }
    }

    if (modelModal) {
        document.getElementById("opExpandBtn").addEventListener("click", async (e) => {
            e.preventDefault();
            const list = document.getElementById("opMetricList");
            const btn = document.getElementById("opExpandBtn");
            const isOpen = list.style.display !== "none";
            if (isOpen) {
                list.style.display = "none";
                btn.setAttribute("aria-expanded", "false");
                btn.textContent = "▸ choose";
            } else {
                list.style.display = "block";
                btn.setAttribute("aria-expanded", "true");
                btn.textContent = "▾ hide";
                if (!opMetricsFetched) { opMetricsFetched = true; await loadOperationalMetrics(); }
            }
        });

        document.getElementById("modelModalClose").addEventListener("click", closeModelModal);
        document.getElementById("modelModalCancel").addEventListener("click", closeModelModal);
        modelModal.addEventListener("click", (e) => { if (e.target === modelModal) closeModelModal(); });

        document.getElementById("modelModalBuild").addEventListener("click", async () => {
            const srcChecked = (src) => {
                const cb = modelModal.querySelector(`.fm-source-toggle[data-source="${src}"]`);
                return cb ? cb.checked : false;
            };
            // Operational metric filter: null = all metrics; otherwise only the checked ones.
            let opMetrics = null;
            if (opMetricsFetched) {
                const boxes = Array.from(modelModal.querySelectorAll(".fm-op-metric"));
                if (boxes.length) opMetrics = boxes.filter(b => b.checked).map(b => b.value);
            }
            const selection = {
                operational: { include: srcChecked("operational"), metrics: opMetrics },
                earningCall: srcChecked("earningCall"),
                qna: srcChecked("qna"),
                broker: srcChecked("broker"),
                interview: srcChecked("interview"),
                orderBook: srcChecked("orderBook"),
            };
            closeModelModal();

            const originalText = refreshBtn.textContent;
            refreshBtn.disabled = true;
            try {
                refreshBtn.textContent = "Downloading data...";
                await handleRefresh();
                refreshBtn.textContent = "Building model...";
                await handleBuildModel(selection);
            } catch (err) {
                console.error("Build model error:", err);
            } finally {
                refreshBtn.textContent = originalText;
                refreshBtn.disabled = false;
            }
        });
    }
    --- end disabled modal block --- */
    refreshStatusBtn.onclick = () => {
        const toggle = document.getElementById("dropdownToggle");

        // If dropdown is empty but we have a lastMatchedCompany, populate it
        if ((!toggle.dataset.value || toggle.value === "Select Company") && lastMatchedCompany) {
            toggle.value = lastMatchedCompany.CompName;
            toggle.dataset.value = lastMatchedCompany.fincode;
            toggle.dataset.sector = lastMatchedCompany.sector_type;
        }

        // Now call the common handler
        handleRefresh();
    };

    // --- Operational Data Handler ---
    async function handleOperationalRefresh() {
        const toggle = document.getElementById("dropdownToggle");
        const fincode = toggle.dataset.value;
        const name = toggle.value;
        if (!fincode) return showWarning("Select a company first.");

        try {
            const res = await fetch("https://transcriptanalyser.com/pms/get_dashboard", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ fincode: parseInt(fincode) })
            });
            const data = await res.json();

            await Excel.run(async (context) => {
                const workbook = context.workbook;
                workbook.worksheets.load("items/name");
                await context.sync();

                const existingNames = workbook.worksheets.items.map(s => s.name);
                let sheet;
                if (existingNames.includes("Operational Data")) {
                    sheet = workbook.worksheets.getItem("Operational Data");
                    sheet.getUsedRange()?.clear();
                } else {
                    sheet = workbook.worksheets.add("Operational Data");
                }

                const nameCell = sheet.getRange("A1");
                nameCell.values = [[name]];
                nameCell.format.font.bold = true;
                nameCell.format.font.size = 14;
                nameCell.format.fill.color = "#bed1f8";
                sheet.getRange("A:A").format.columnWidth = 180;

                let row = 3;
                for (const section of (data.table_data || [])) {
                    sheet.getRange(`A${row}`).values = [[section.section_name]];
                    sheet.getRange(`A${row}`).format.font.bold = true;
                    sheet.getRange(`A${row}`).format.font.size = 13;
                    sheet.getRange(`A${row}`).format.fill.color = "#d9ead3";
                    row++;

                    const headers = ["Metric", "Unit", ...section.periods];
                    const lastCol = getExcelColumnLetter(headers.length);
                    const headerRange = sheet.getRange(`A${row}:${lastCol}${row}`);
                    headerRange.values = [headers];
                    headerRange.format.fill.color = "#e0e0e0";
                    headerRange.format.font.bold = true;
                    row++;

                    if (section.rows?.length > 0) {
                        const vals = section.rows.map(r => [
                            r.metric_name, r.unit, ...section.periods.map(p => r[p] ?? "")
                        ]);
                        sheet.getRange(`A${row}:${lastCol}${row + vals.length - 1}`).values = vals;
                        row += vals.length;
                    }
                    row += 2;
                }

                await context.sync();
            });
            console.log("✅ Operational data written to Excel");
        } catch (err) {
            console.error("❌ Error in operational refresh:", err);
            showWarning("Failed to fetch operational data. Check console for details.");
        }
    }

    // --- Sector Data Handler ---
    async function handleSectorRefresh() {
        const select = document.getElementById("sectorSelect");
        const sectorId = select.value;
        const sectorName = select.options[select.selectedIndex]?.text || "";
        if (!sectorId) return showWarning("Select a sector first.");

        try {
            const res = await fetch(
                `https://transcriptanalyser.com/sector_metrics/sector_metric_data2?sector_id=${sectorId}&page=1&limit=500`
            );
            const data = await res.json();

            await Excel.run(async (context) => {
                const workbook = context.workbook;
                workbook.worksheets.load("items/name");
                await context.sync();

                const existingNames = workbook.worksheets.items.map(s => s.name);
                let sheet;
                if (existingNames.includes("Sector Data")) {
                    sheet = workbook.worksheets.getItem("Sector Data");
                    sheet.getUsedRange()?.clear();
                } else {
                    sheet = workbook.worksheets.add("Sector Data");
                }

                const nameCell = sheet.getRange("A1");
                nameCell.values = [[`Sector: ${sectorName}`]];
                nameCell.format.font.bold = true;
                nameCell.format.font.size = 14;
                nameCell.format.fill.color = "#bed1f8";
                sheet.getRange("A:A").format.columnWidth = 220;

                let row = 3;
                for (const [category, metrics] of Object.entries(data)) {
                    if (!Array.isArray(metrics) || metrics.length === 0) continue;

                    sheet.getRange(`A${row}`).values = [[category]];
                    sheet.getRange(`A${row}`).format.font.bold = true;
                    sheet.getRange(`A${row}`).format.font.size = 13;
                    sheet.getRange(`A${row}`).format.fill.color = "#d9ead3";
                    row++;

                    const allDates = new Set();
                    metrics.forEach(m => (m.data_points || []).forEach(dp => allDates.add(dp.data_date)));
                    const sortedDates = Array.from(allDates).sort();

                    const headers = ["Metric", "Unit", "Frequency", ...sortedDates];
                    const lastCol = getExcelColumnLetter(headers.length);
                    const headerRange = sheet.getRange(`A${row}:${lastCol}${row}`);
                    headerRange.values = [headers];
                    headerRange.format.fill.color = "#e0e0e0";
                    headerRange.format.font.bold = true;
                    row++;

                    const vals = metrics.map(metric => {
                        const dateMap = {};
                        (metric.data_points || []).forEach(dp => { dateMap[dp.data_date] = dp.value; });
                        return [metric.metric_name, metric.unit, metric.frequency,
                                ...sortedDates.map(d => dateMap[d] ?? "")];
                    });
                    if (vals.length > 0) {
                        sheet.getRange(`A${row}:${lastCol}${row + vals.length - 1}`).values = vals;
                        row += vals.length;
                    }
                    row += 2;
                }

                await context.sync();
            });
            console.log("✅ Sector data written to Excel");
        } catch (err) {
            console.error("❌ Error in sector refresh:", err);
            showWarning("Failed to fetch sector data. Check console for details.");
        }
    }

    function showWarning(msg) {
        const el = document.getElementById("warningMsg");
        if (el) {
            el.textContent = msg;
            el.style.color = "red";
            el.style.display = "block";
            setTimeout(() => { el.style.display = "none"; }, 6000);
        }
        console.warn(msg);
    }

    // AI progress status box
    let _aiStatusTimer = null;
    function aiStatus(msg) {
        const box = document.getElementById("aiStatusBox");
        const txt = document.getElementById("aiStatusText");
        if (!box || !txt) return;
        if (msg === null) {
            box.style.display = "none";
            txt.textContent = "";
            if (_aiStatusTimer) { clearInterval(_aiStatusTimer); _aiStatusTimer = null; }
            return;
        }
        box.style.display = "";
        txt.textContent = msg;
    }
    function aiStatusCycle(messages, intervalMs) {
        if (_aiStatusTimer) clearInterval(_aiStatusTimer);
        let i = 0;
        aiStatus(messages[0]);
        _aiStatusTimer = setInterval(() => {
            i++;
            if (i < messages.length) aiStatus(messages[i]);
            else clearInterval(_aiStatusTimer);
        }, intervalMs);
    }

    // Renders the 5 automated model-quality checks (see the block near the end of
    // handleBuildModel's Excel.run) into the dismissible #modelChecksModal popup and opens it.
    // Also unhides #modelChecksBtn so the SAME results can be reopened later without rebuilding —
    // "temporarily" advisory-only per the current request, not a hard block on a failed check.
    function renderModelChecks(results) {
        // Model-quality-check diagnostics (see the automated checks in handleBuildModel) are an
        // internal debugging aid — raw pass/fail detail on historicals/linking/formula wiring meant
        // for reviewing the model-BUILDING logic itself, not something an end user building a
        // company's model needs to see. Restricted to the admin account; every other user's build
        // still runs the same checks (console warnings on failure are unaffected), it just never
        // shows the popup or the "View Model Checks" button.
        const MODEL_CHECKS_ADMIN_EMAIL = "pranav@goindiaadvisors.com";
        const modelChecksUser = JSON.parse(localStorage.getItem("user") || "{}");
        if (String(modelChecksUser.Email || "").trim().toLowerCase() !== MODEL_CHECKS_ADMIN_EMAIL) return;

        const listEl = document.getElementById("modelChecksList");
        const modal = document.getElementById("modelChecksModal");
        const btn = document.getElementById("modelChecksBtn");
        if (!listEl || !modal || !Array.isArray(results) || !results.length) return;
        const icon = (passed) => passed === true ? "✅" : passed === false ? "❌" : "➖";
        listEl.innerHTML = results.map(r => `
            <div style="display:flex; gap:8px; align-items:flex-start; padding:8px 0; border-bottom:1px solid #f0f0f0;">
                <span style="font-size:15px; line-height:1.4;">${icon(r.passed)}</span>
                <div>
                    <div style="font-weight:700; color:#173760; font-size:12.5px;">${r.label}</div>
                    <div style="font-size:11px; color:#666; margin-top:2px;">${r.detail}</div>
                </div>
            </div>
        `).join("");
        if (btn) btn.style.display = "inline-block";
        modal.style.display = "flex";
    }

    // --- AI Financial Model Builder ---
    async function handleBuildModel(selection) {
        // selection controls which optional sources are fed to the model.
        // Defaults to "everything" so the function still works if called without one.
        selection = selection || {
            operational: { include: true, metrics: null },
            earningCall: true, qna: true, broker: true, interview: true, orderBook: true
        };
        const OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions";
        const OPENROUTER_KEY = process.env.OPENROUTER_KEY; // injected at build time from .env — see webpack.config.js
        // Static workflow: schema-driven sheets that mostly transcribe numbers/formulas into a fixed
        // structure (no interpretive judgment needed) — cheap/fast model.
        // Dynamic workflow: the one call that must read qualitative sources (earnings call commentary,
        // Q&A, broker notes, management interviews) and derive judgment calls (growth/margin/capex
        // assumptions) — a stronger reasoning model.
        // All calls route through OpenRouter (a direct-to-Google API path was tried and reverted).
        const STATIC_MODEL = "google/gemini-2.5-flash";
        const DYNAMIC_MODEL = "openai/gpt-5.4";
        // Cross-PROVIDER last-resort fallback for every call (not just Assumptions). A same-provider
        // fallback (e.g. flash -> pro) is useless against a Google-infra-wide blip — observed in
        // practice: an "Upstream idle timeout" on gemini-2.5-pro and "JSON error injected into SSE
        const FALLBACK_MODEL = "openai/gpt-5.4";

        const GENERAL_DISCLAIMER = "General Disclaimer - The information on this website has been collected from certain public sources. GoIndia Advisors LLP believes that the information it uses comes from reliable sources, but does not guarantee the accuracy or completeness of this information, which is subject to change without notice, and nothing in this document shall be construed as such a guarantee. Employees involved in this service may hold positions in the companies mentioned in the services/information. We disclaim any liability arising from use of information contained on this website. Nothing herein shall constitute or be construed as an offering of financial instruments or as investment advice or recommendations by GoIndia Advisors LLP. The Site may include advertisements and links to external sites and co-branded pages which GoIndia Advisors LLP does not endorse and cannot accept any responsibility or liability for loss or damage suffered by the intended viewer. GoIndia Advisors LLP is unable to exercise control over the security or content of information passing over the network or via the Service, and GoIndia Advisors LLP hereby excludes all liability of any kind for the transmission or reception of infringing or unlawful information of whatever nature.";

        const toggle = document.getElementById("dropdownToggle");
        const fincode = toggle.dataset.value;
        const companyName = toggle.value;
        if (!fincode) return showWarning("Select a company first.");

        try {
            // Step 1: Read all available financial data from Excel
            aiStatus("Reading financial data from Excel...");
            let financialText = "";
            let annualText = "";
            let quarterlyText = "";

            const readSheet = (text, maxRows, maxCols) =>
                (text || []).slice(0, maxRows)
                    .map(row => row.slice(0, maxCols).join("\t"))
                    .filter(line => line.replace(/\t/g, "").trim() !== "")
                    .join("\n");

            // Annual Data / Quarterly Data can each hold up to 4 sub-tables per statement — IndAS
            // vs Detailed presentation, × Consolidated vs Standalone (see writeFinancialStatement/
            // formatTable in handleRefresh, which write each combination as its own titled block:
            // a "<Statement> (<IndAS|Detailed>)" header, then a "Consolidated"/"Standalone" title,
            // then a "Parameter | MonYYYY..." row, then the data). Dumping ALL of them into the LLM
            // prompt as raw text caused it to mix figures across presentations when transcribing a
            // row that autoLink couldn't confidently match on its own. The SHEET still keeps every
            // combination (for the user, and for autoLink/buildIndex's own "first occurrence wins,
            // Consolidated written first" rule) — only the PROMPT TEXT is trimmed to exactly ONE
            // sub-table per statement, by priority: IndAS Consolidated, then Detailed Consolidated,
            // then IndAS Standalone, then Detailed Standalone.
            const STATEMENT_BASES = ["Balance Sheet", "Cash Flows", "Detailed P&L", "Quarterly P&L"];
            const titleRe = new RegExp(`^(${STATEMENT_BASES.map(s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?:\\s*\\((IndAS|Detailed)\\))?$`);
            const presRe = /^(Consolidated|Standalone)$/;
            const rankOf = (format, presentation) => (format === "Detailed" ? 1 : 0) + (presentation === "Standalone" ? 2 : 0);
            const selectPreferredTables = (grid) => {
                if (!Array.isArray(grid) || !grid.length) return "";
                const titles = [];
                for (let r = 0; r < grid.length; r++) {
                    const a = String((grid[r] && grid[r][0]) || "").trim();
                    const m = a.match(titleRe);
                    if (m) titles.push({ row: r, base: m[1], format: m[2] || null });
                }
                const best = new Map(); // statement base -> { row, endRow, rank }
                for (let i = 0; i < titles.length; i++) {
                    const t = titles[i];
                    const nextRow = String((grid[t.row + 1] && grid[t.row + 1][0]) || "").trim();
                    const presMatch = nextRow.match(presRe);
                    if (!presMatch) continue; // "<presentation> not available for this company" or malformed — skip
                    const rank = rankOf(t.format, presMatch[1]);
                    const cur = best.get(t.base);
                    if (!cur || rank < cur.rank) {
                        const endRow = i + 1 < titles.length ? titles[i + 1].row : grid.length;
                        best.set(t.base, { row: t.row, endRow, rank });
                    }
                }
                const chunks = [];
                for (const base of STATEMENT_BASES) {
                    const b = best.get(base);
                    if (!b) continue;
                    const lines = grid.slice(b.row, b.endRow)
                        .map(row => row.slice(0, 30).join("\t"))
                        .filter(line => line.replace(/\t/g, "").trim() !== "");
                    if (lines.length) chunks.push(lines.join("\n"));
                }
                return chunks.join("\n\n");
            };

            const grids = {}; // sheetName -> full 2D text grid (for the link index)
            await Excel.run(async (context) => {
                const wb = context.workbook;
                wb.worksheets.load("items/name");
                await context.sync();
                const names = wb.worksheets.items.map(s => s.name);

                const ranges = {};
                for (const sheetName of ["Key Financials", "Annual Data", "Quarterly Data"]) {
                    if (names.includes(sheetName)) {
                        const rng = wb.worksheets.getItem(sheetName).getUsedRange();
                        rng.load("text");
                        ranges[sheetName] = rng;
                    }
                }
                await context.sync();

                for (const sn of Object.keys(ranges)) grids[sn] = ranges[sn].text || [];
                if (ranges["Key Financials"]) financialText = readSheet(ranges["Key Financials"].text, 80, 30);
                if (ranges["Annual Data"])    annualText    = selectPreferredTables(ranges["Annual Data"].text);
                if (ranges["Quarterly Data"]) quarterlyText = selectPreferredTables(ranges["Quarterly Data"].text);
            });

            if (!financialText) {
                aiStatus(null);
                return showWarning("Please download Financial data first — it's needed to build the model.");
            }

            // ── Build a label→cell link index from the downloaded sheets so historical cells can
            //    LIVE-LINK instead of being static. Each sheet may hold several sub-tables (e.g. Annual
            //    Data = Standalone + Consolidated × Balance Sheet / Cash Flow / Detailed P&L), each with
            //    its own "Parameter | MonYYYY…" header — so every match is SCOPED to its sub-table's row
            //    range, which avoids the duplicate-label collisions a whole-column lookup would hit. ──
            const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
            const normLabel = (s) => String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
            const parseNum = (s) => { const t = String(s == null ? "" : s).replace(/,/g, "").trim(); if (!t) return null; const n = parseFloat(t.replace("(", "-").replace(")", "")); return isFinite(n) ? n : null; };
            // Tags every row with which (statement, format, presentation) block it's in, using the
            // EXACT SAME title/presentation detection selectPreferredTables uses above (titleRe/
            // presRe/rankOf/STATEMENT_BASES) — needed so findEntryByLabel (below) can tell apart a
            // duplicate label that's a genuinely different presentation (Standalone vs
            // Consolidated — must NOT merge, that mixes bases) from a duplicate that's just a
            // DIFFERENT YEAR-CHUNK of the SAME sub-table (a vendor can emit several "Parameter"
            // blocks for the same Balance Sheet/Consolidated pair, one per handful of years — this
            // has actually happened: a company's "Gross Block"/"Capex" existed with real data, but
            // only in a block covering Mar-13..Mar-17, years the model never even asks about,
            // while a LATER block covering the actually-needed recent years was silently discarded
            // by "first occurrence wins" — MUST merge those).
            const tagBlocks = (grid) => {
                const tags = new Array(grid.length).fill(null);
                let cur = null;
                for (let r = 0; r < grid.length; r++) {
                    const a = String((grid[r] && grid[r][0]) || "").trim();
                    const titleMatch = a.match(titleRe);
                    if (titleMatch) { cur = { format: titleMatch[2] || null, rank: null }; continue; }
                    if (cur && cur.rank == null) {
                        const presMatch = a.match(presRe);
                        if (presMatch) cur.rank = rankOf(cur.format, presMatch[1]);
                    }
                    tags[r] = cur;
                }
                return tags;
            };
            const buildIndex = (grid, sheetName) => {
                const map = new Map();
                const rawList = []; // every occurrence, undeduped, tagged with its block's rank — see findEntryByLabel
                const byRow = new Map(); // 1-based Excel row -> {label, valByFY} — lets code look up "whatever is on row N" directly, without a label to search by (needed by findParentExcelRow/isPlausibleSubtotal below)
                if (!Array.isArray(grid) || !grid.length) return { map, rawList, byRow };
                const blockTags = tagBlocks(grid);
                const headers = []; // each "Parameter" header row + its FY->column map
                for (let r = 0; r < grid.length; r++) {
                    if (normLabel(grid[r] && grid[r][0]) === "parameter") {
                        const colByFY = {}, colIdxByFY = {}, rowArr = grid[r] || [];
                        for (let c = 1; c < rowArr.length; c++) {
                            const h = String(rowArr[c] || "").trim();
                            // FY2024 / Mar2024 / Mar-24 (a real vendor format — hyphenated, 2-digit
                            // year — that neither of the first two patterns matched at all, so this
                            // ENTIRE header's colByFY/colIdxByFY came back completely empty for any
                            // sheet using it: every row still got added to the index [it isn't
                            // blank], just with a valByFY of {} for every single row, indistinguishable
                            // downstream from "label not found" — this is what made buildCapexRows
                            // report "could not find 'Gross Block'/'Capital Expenditure'" even though
                            // both were plainly present with real numbers).
                            const m = h.match(/^FY(\d{4})E?$/) || h.match(/^[A-Za-z]{3}(\d{4})$/) || h.match(/^[A-Za-z]{3}-(\d{2})$/);
                            if (m) {
                                let fy = parseInt(m[1]);
                                if (fy < 100) fy += 2000; // "Mar-24" -> 2024 (2-digit-year vendor format)
                                colByFY[fy] = colLetter(c + 1); colIdxByFY[fy] = c;
                            }
                        }
                        headers.push({ hr: r, colByFY, colIdxByFY });
                    }
                }
                for (let i = 0; i < headers.length; i++) {
                    const h = headers[i];
                    const lastData = (i + 1 < headers.length ? headers[i + 1].hr : grid.length) - 1;
                    const r1 = h.hr + 2, r2 = lastData + 1; // 1-based Excel data-row range of this sub-table
                    for (let r = h.hr + 1; r <= lastData; r++) {
                        const rowArr = grid[r] || [];
                        const exact = String(rowArr[0] || "");
                        const label = exact.trim();
                        if (!label || normLabel(label) === "parameter") continue;
                        // skip title/blank rows (Standalone/Consolidated headings, spacers) — no values
                        if (!rowArr.slice(1).some(x => String(x || "").trim() !== "")) continue;
                        const valByFY = {};
                        for (const fy in h.colIdxByFY) { const v = parseNum(rowArr[h.colIdxByFY[fy]]); if (v !== null) valByFY[fy] = v; }
                        const tag = blockTags[r];
                        rawList.push({ label: exact, valByFY, _rank: (tag && tag.rank != null) ? tag.rank : 0 });
                        byRow.set(r + 1, { label: exact, valByFY });
                        // Keep the FIRST occurrence for a duplicate label, not the last, for the MAIN
                        // (live-linking) map. Annual/Quarterly Data write ALL Consolidated sub-tables
                        // before ANY Standalone ones (see the two sequential loops in handleRefresh —
                        // Consolidated first, Standalone second), and a Consolidated/Standalone pair
                        // for the same statement shares the IDENTICAL section header text with
                        // nothing in the sheet content itself to tell them apart — so "first
                        // occurrence" is the only thing that actually ties this to "prefer
                        // Consolidated", which is what DATA_BASIS_NOTE instructs the model to use
                        // consistently. An earlier version of this kept the LAST occurrence on the
                        // (verified wrong) assumption that Consolidated was written after Standalone
                        // — that silently linked historical cells to the Standalone figure instead,
                        // which is why linked cells could disagree with the model's own
                        // Consolidated-based historicals and, upstream, why cross-row sums (e.g. the
                        // Balance Sheet's assets vs liabilities+equity) could fail to tie out even
                        // though most individual rows "linked successfully". A single cell reference
                        // (row/colByFY) can only ever point at ONE physical row, so this map is
                        // intentionally NOT merged across year-chunks the way rawList is above —
                        // that's fine, because writeDataSheet's per-year linking already falls back
                        // to the plain historical number for any year the linked row's own colByFY
                        // doesn't cover (see the historical-columns loop), so a "wrong chunk" link
                        // here just means every year outside ITS OWN range harmlessly falls through
                        // to the (correctly merged) historical values findEntryByLabel provides.
                        if (!map.has(normLabel(label))) {
                            map.set(normLabel(label), { sheet: sheetName, label: exact, colByFY: h.colByFY, r1, r2, row: r + 1, valByFY });
                        }
                    }
                }
                return { map, rawList, byRow };
            };
            const { map: idxKF, rawList: rawKF, byRow: byRowKF } = buildIndex(grids["Key Financials"], "Key Financials");
            const { map: idxAnnual, rawList: rawAnnual, byRow: byRowAnnual } = buildIndex(grids["Annual Data"], "Annual Data");
            const resolveLink = (label) => idxKF.get(normLabel(label)) || idxAnnual.get(normLabel(label)) || null;
            const entryAtRow = (sheetName, excelRow) => {
                const byRow = sheetName === "Key Financials" ? byRowKF : (sheetName === "Annual Data" ? byRowAnnual : null);
                return byRow ? (byRow.get(excelRow) || null) : null;
            };

            // Looked up directly from the same Key Financials/Annual Data index autoLink uses,
            // rather than trusting the LLM to find/derive things itself. Strip a trailing/embedded
            // unit or qualifier annotation in parens (e.g. "Gross Block (Rs Cr)") before testing a
            // label against an ANCHORED pattern — without this, a fully-anchored regex silently
            // fails to match a label carrying an extra parenthetical qualifier.
            const stripUnit = (s) => String(s || "").replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
            // Scans EVERY raw occurrence (not idxKF/idxAnnual's deduped maps, which keep only the
            // first — right for live-linking a single cell, wrong here) matching the pattern,
            // takes the BEST-ranked (format,presentation) group among them (Consolidated/IndAS
            // priority — see rankOf/DATA_BASIS_NOTE), and MERGES valByFY across every occurrence in
            // THAT group — recovering a label that exists with real data, just split across several
            // "Parameter" blocks each covering a different handful of years (a real, observed
            // pattern; see the comment on rawList in buildIndex). Value-only merge — safe because
            // callers here (lookupHistoricalByLabel, the Total Assets cross-check) only ever need
            // numbers, never a live cell reference.
            const findEntryByLabel = (pattern) => {
                let bestRank = Infinity;
                const matches = [];
                for (const raw of [...rawKF, ...rawAnnual]) {
                    if (!(pattern.test(raw.label.trim()) || pattern.test(stripUnit(raw.label)))) continue;
                    if (raw._rank < bestRank) bestRank = raw._rank;
                    matches.push(raw);
                }
                if (!matches.length) return null;
                const chosen = matches.filter(m => m._rank === bestRank);
                const valByFY = {};
                for (const m of chosen) for (const fy in m.valByFY) if (valByFY[fy] === undefined) valByFY[fy] = m.valByFY[fy];
                return { label: chosen[0].label, valByFY };
            };
            // shares_out itself is now derived from Market Cap ÷ Price (see fetchSharesOutInputs
            // and where sharesOutInputsPromise is awaited/computed further below, alongside
            // basicInfo) — NOT from anything in idxKF/idxAnnual. A PAT ÷ EPS derivation was tried
            // first, but that's circular: the P&L's own EPS forecast formula is PAT ÷ shares_out, so
            // computing shares_out FROM PAT ÷ EPS makes the two rows depend on each other.

            // Resolve a model row to a data-sheet cell. An explicit "link" is trusted; otherwise the
            // row's label/key is tried but ACCEPTED only when a data value matches one of the row's
            // own historicals — so we link far more rows without risking a wrong auto-match.
            // (stripUnit itself is defined earlier, above findEntryByLabel — shared by both.)
            const valuesAgree = (a, b) => { if (a == null || b == null) return false; const m = Math.max(Math.abs(a), Math.abs(b), 1); return Math.abs(a - b) / m < 0.02; };
            const hit = (a, b) => a != null && b != null && Math.abs(a) > 0.5 && valuesAgree(a, b); // non-trivial agreement
            const autoLink = (item) => {
                if (!item) return null;
                // Explicit link — trusted, but only tried as the model wrote it (raw string). If the
                // model includes a unit suffix on "link" itself (e.g. "Net Revenue (Rs Cr)") while the
                // real data label has none, this raw exact-match fails and previously fell all the way
                // through to the label/key fast path below — which never re-tried "link" at all, only
                // "label"/"key". Retry it here with the same unit-stripping "label" already gets.
                if (item.link) {
                    const e = resolveLink(item.link) || resolveLink(stripUnit(item.link));
                    if (e) return e;
                }
                const hist = (Array.isArray(item.historical) ? item.historical : []).map(parseNum);
                // fast path: a label/key/link match confirmed by an agreeing value. "link" is included
                // here too (not just label/key) so a near-miss on the label text above — one that
                // resolves to SOME row but isn't confirmed by an agreeing value below — still gets a
                // value-gated second attempt instead of dropping straight to the expensive, higher-bar
                // (>=2 year) full-scan fallback further down.
                for (const cand of [
                    item.link && resolveLink(stripUnit(item.link)),
                    item.label && resolveLink(stripUnit(item.label)),
                    item.key && resolveLink(String(item.key).replace(/_/g, " ")),
                ]) {
                    if (cand) for (let i = 0; i < HIST.length; i++) if (hit(cand.valByFY[HIST[i]], hist[i])) return cand;
                }
                // value-based fallback: find the data row matching the MOST of this row's historicals,
                // ignoring the label — this links Annual/BS/CF items whose wording differs from the model.
                // Requires >=2 non-trivial matching years so coincidences don't create wrong links.
                if (hist.some(v => v != null)) {
                    let best = null, bestHits = 0;
                    for (const m of [idxKF, idxAnnual]) {
                        for (const e of m.values()) {
                            let hits = 0;
                            for (let i = 0; i < HIST.length; i++) if (hit(e.valByFY[HIST[i]], hist[i])) hits++;
                            if (hits > bestHits) { bestHits = hits; best = e; }
                        }
                    }
                    if (bestHits >= 2) return best;
                }
                return null;
            };
            const kfLabels = Array.from(new Set(Array.from(idxKF.values()).map(v => v.label.trim()))).filter(Boolean).slice(0, 60);
            const annualLabels = Array.from(new Set(Array.from(idxAnnual.values()).map(v => v.label.trim()))).filter(Boolean).slice(0, 150);
            const linkableList = Array.from(new Set([...kfLabels, ...annualLabels]));

            // Kicked off here (not awaited yet) so they run CONCURRENTLY with the operational data
            // fetch below rather than adding their own sequential delay — awaited just before the
            // Valuation call that actually needs it, further down.
            const latestPricePromise = fetchLatestPrice(fincode);
            const basicInfoPromise = fetchBasicInfo(fincode);
            const sharesOutInputsPromise = fetchSharesOutInputs(fincode);
            const valuationChartPromise = fetchValuationChart(fincode);

            // Step 1a: Fetch operational data directly from the dashboard API (no truncation).
            // Honour the user's selection: skip entirely if unchecked, and keep only the
            // chosen sections when a subset was picked in the modal.
            let operationalText = "";
            // Deduplicated, code-derived candidate segment names (section_name from the SAME
            // dashboard payload) — NOT a judgment call, just a raw list of the distinct section
            // labels the API actually returned (confirmed live, e.g. for a real conglomerate this
            // comes back as clean labels like "Airports (Adani Airport Holdings Ltd)", "Data Centers
            // (AdaniConneX)", "Mining Services" — genuine business verticals, not generic metric
            // categories). The SAME section_name repeats across multiple rows/metric groupings for
            // one segment (volumes vs. capacity vs. project counts, etc.), so this dedupes by name
            // rather than pushing one entry per section object. Feeds the "Segment Drivers" call
            // below as CANDIDATES only — that call still decides which of these are genuinely
            // distinct P&L segments (vs. a generic "Portfolio Overview"-style section, or two
            // candidates that are really the same business), the one judgment call that can't be
            // removed from the loop, mirroring holding_discount's own existing test.
            const operationalSegments = [];
            // Structured mirror of the SAME rows that get flattened into operationalText below —
            // kept alongside it so the segment-revenue recovery pass further down can read real
            // per-year figures out of the dashboard payload directly, instead of regex-scraping the
            // pipe-delimited text back apart (the text is built for an LLM to read, not to reparse).
            const operationalSections = [];
            if (selection.operational.include) {
                aiStatus("Fetching operational data...");
                const wantedMetrics = selection.operational.metrics; // null = all metrics
                try {
                    const opRes = await fetch("https://transcriptanalyser.com/pms/get_dashboard", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode: parseInt(fincode) })
                    });
                    if (opRes.ok) {
                        const opJson = await opRes.json();
                        const opLines = [];
                        const seenSegments = new Set();
                        for (const section of (opJson.table_data || [])) {
                            const rows = (section.rows || []).filter(row =>
                                !Array.isArray(wantedMetrics) || wantedMetrics.includes(row.metric_name));
                            if (!rows.length) continue; // skip sections with no selected metrics
                            const periods = section.periods || [];
                            opLines.push(`## ${section.section_name || "Section"}`);
                            opLines.push(["Metric", "Unit", ...periods].join(" | "));
                            for (const row of rows) {
                                opLines.push([row.metric_name, row.unit, ...periods.map(p => row[p] ?? "")].join(" | "));
                            }
                            opLines.push("");
                            operationalSections.push({
                                name: section.section_name || "Section",
                                periods,
                                rows: rows.map(row => ({
                                    metric: row.metric_name || "",
                                    unit: row.unit || "",
                                    byPeriod: Object.fromEntries(periods.map(p => [p, row[p]])),
                                })),
                            });
                            if (section.section_name && !seenSegments.has(section.section_name)) {
                                seenSegments.add(section.section_name);
                                operationalSegments.push(section.section_name);
                            }
                        }
                        operationalText = opLines.join("\n").trim();
                    }
                } catch (e) {
                    console.warn("[AI Model] Operational data fetch failed:", e);
                }
            }

            // Step 1b: Fetch qualitative & forward-looking sources from GoIndia APIs
            aiStatus("Fetching earning calls, broker reports, order book & interviews...");

            // Last 2 completed fiscal quarter_years (YYYYQ, Indian FY) — mirrors backend logic
            const lastQuarterYears = (cnt) => {
                const today = new Date();
                const m = today.getMonth() + 1, y = today.getFullYear();
                let fy, q;
                if (m <= 3) { fy = y; q = 4; }
                else if (m <= 6) { fy = y + 1; q = 1; }
                else if (m <= 9) { fy = y + 1; q = 2; }
                else { fy = y + 1; q = 3; }
                const out = [];
                for (let i = 0; i < cnt; i++) {
                    q -= 1;
                    if (q === 0) { q = 4; fy -= 1; }
                    out.push(fy * 10 + q);
                }
                return out;
            };
            const last2Q = lastQuarterYears(2);
            const fc = parseInt(fincode);

            // Some endpoints return a JSON-encoded string rather than plain text
            const readBody = async (res) => {
                if (!res || !res.ok) return "";
                const t = await res.text();
                if (t.startsWith('"') && t.endsWith('"')) {
                    try { return JSON.parse(t); } catch { return t; }
                }
                return t;
            };

            let earningCallText = "";
            let brokerReportText = "";
            let mgmtInterviewText = "";
            let orderBookText = "";
            let qnaText = "";
            try {
                const post = (url, body) => fetch(url, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(body)
                });
                const [ecRes, brRes, miRes, obRes, ...qnaResArr] = await Promise.all([
                    selection.earningCall ? post("https://goindiainvest.in/mcp/earning_call_summaries", { fincodes: [fc], quarter_years: last2Q }) : Promise.resolve(null),
                    selection.broker ? post("https://goindiainvest.in/mcp/broker_report_single_comp", { fincode: fc }) : Promise.resolve(null),
                    selection.interview ? post("https://goindiainvest.in/mcp/comp_management_interviews", { fincode: fc }) : Promise.resolve(null),
                    selection.orderBook ? post("https://goindiainvest.in/mcp/order_book_single_comp", { fincode: fc, quarter_list: last2Q.map(String) }) : Promise.resolve(null),
                    ...(selection.qna ? last2Q.map(qy => post("https://goindiainvest.in/mcp/fetch_transcript_ques_ans", { fincode: fc, quarter_year: qy })) : [])
                ]);
                earningCallText = await readBody(ecRes);
                brokerReportText = await readBody(brRes);
                mgmtInterviewText = await readBody(miRes);
                orderBookText = await readBody(obRes);
                const qnaParts = [];
                for (const r of qnaResArr) {
                    const t = await readBody(r);
                    if (t && t.trim()) qnaParts.push(t.trim());
                }
                qnaText = qnaParts.join("\n\n");
            } catch (e) {
                console.warn("[AI Model] Optional API data fetch failed:", e);
            }

            // ── Step 2: Period axis — 8 historical actuals + 5 forecast years ──
            // HIST_YEARS/FC_YEARS are the single source of truth for the model's time window —
            // every column array below is DERIVED from them (via colLetter) rather than hardcoded
            // letters, so this is the only place a future window-size change needs to happen. (An
            // earlier 5+5 version DID hardcode literal column letters throughout the sheet-writing
            // and formula-resolution code — that required touching ~30 separate spots to extend.)
            const HIST_YEARS = 8;
            const FC_YEARS = 5;
            // Latest actual fiscal year = max non-forecast FY found in the Key Financials header.
            const fyMatches = (financialText.match(/FY(\d{4})(?!E)/g) || [])
                .map(s => parseInt(s.slice(2))).filter(n => n >= 2000 && n <= 2100);
            const latestActualFY = fyMatches.length ? Math.max(...fyMatches) : new Date().getFullYear();
            const HIST = [], FCST = [];
            for (let i = HIST_YEARS - 1; i >= 0; i--) HIST.push(latestActualFY - i);
            for (let i = 1; i <= FC_YEARS; i++) FCST.push(latestActualFY + i);
            const histLabels = HIST.map(y => `FY${y}A`);
            const fcLabels = FCST.map(y => `FY${y}E`);
            const periodLabels = [...histLabels, ...fcLabels];            // HIST_YEARS+FC_YEARS column labels (B..)
            const HIST_COLS = Array.from({ length: HIST_YEARS }, (_, i) => colLetter(2 + i));
            const FC_COLS = Array.from({ length: FC_YEARS }, (_, i) => colLetter(2 + HIST_YEARS + i));
            const ALL_COLS = [...HIST_COLS, ...FC_COLS];
            const CAGR_COL = colLetter(2 + HIST_YEARS + FC_YEARS);
            const LAST_HIST_COL = HIST_COLS[HIST_COLS.length - 1];
            const LAST_FC_COL = FC_COLS[FC_COLS.length - 1];
            // Boundary index used by resolve()'s {N} placeholder: forecast-year-number = column
            // index (0-based within ALL_COLS) minus (HIST_YEARS-1), for columns at/after HIST_YEARS.
            const FC_BOUNDARY = HIST_YEARS;
            // Illustrative "[v1,v2,...]" placeholders used in the row-schema prompt text below —
            // generated from HIST_YEARS/FC_YEARS rather than hardcoded, so the schema example always
            // matches the actual number of values a row is expected to return.
            const histPlaceholders = Array.from({ length: HIST_YEARS }, (_, i) => `v${i + 1}`).join(",");
            const fcPlaceholders = Array.from({ length: FC_YEARS }, (_, i) => `v${i + 1}`).join(",");
            const CAGR_LABEL = `FY${String(latestActualFY).slice(2)}-${String(latestActualFY + 5).slice(2)} CAGR`;

            // ── Step 3: Per-sheet LLM calls (multiple calls; cost summed and printed) ──
            const SHEETS = {
                assum: "FM Assumptions", op: "FM Operational", pnl: "FM P&L",
                bs: "FM Balance Sheet", capex: "FM Capex & FCF", val: "FM Valuation & DCF", summary: "FM Summary"
            };
            const symbols = { assum: {}, op: {}, pnl: {}, bs: {}, capex: {}, val: {} }; // key -> Excel row
            let totalCost = 0, totalTokens = 0;
            let usedFallbackModel = false; // true if any call in this workflow had to use FALLBACK_MODEL
            const costRows = [];

            // Lenient JSON parse: strip fences/prose, isolate the outer object, repair truncation.
            const parseLenient = (content) => {
                let raw = (content || "").replace(/^```(?:json)?\s*/im, "").replace(/\s*```\s*$/im, "").trim();
                const s = raw.indexOf("{"), e = raw.lastIndexOf("}");
                if (s >= 0 && e > s) raw = raw.slice(s, e + 1); // drop any leading/trailing prose
                try { return JSON.parse(raw); } catch { /* fall through to repair */ }
                let rep = raw;
                const lc = rep.lastIndexOf(","); if (lc > rep.length - 50) rep = rep.slice(0, lc);
                const o = (rep.match(/[{[]/g) || []).length, c = (rep.match(/[}\]]/g) || []).length;
                for (let i = 0; i < o - c; i++) { const lo = Math.max(rep.lastIndexOf("{"), rep.lastIndexOf("[")); rep += rep[lo] === "[" ? "]" : "}"; }
                try { return JSON.parse(rep); } catch { return null; }
            };

            // Per-workflow cost/token subtotals, so the console breaks out static vs dynamic spend.
            const workflowTotals = {
                static: { cost: 0, tokens: 0, calls: 0 },
                dynamic: { cost: 0, tokens: 0, calls: 0 }
            };
            const workflowOf = (model) => model === DYNAMIC_MODEL ? "dynamic" : "static";
            const sleep = (ms) => new Promise(r => setTimeout(r, ms));

            const SYSTEM_PROMPT = "You are a senior equity research analyst building an institutional, driver-based financial model in Excel. Return ONLY raw JSON — no markdown, no prose, no code fences.";

            // Read an SSE (text/event-stream) body into an array of parsed "data:" JSON payloads.
            // A single corrupt line ("JSON error injected into SSE stream") is skipped, not fatal, so
            // a mid-stream provider hiccup doesn't lose everything already received.
            const readSSE = async (res, onChunk) => {
                if (!res.body || !res.body.getReader) return false; // caller falls back to res.json()
                const reader = res.body.getReader();
                const decoder = new TextDecoder();
                let buf = "";
                for (;;) {
                    const { done, value } = await reader.read();
                    if (done) break;
                    buf += decoder.decode(value, { stream: true });
                    const lines = buf.split("\n");
                    buf = lines.pop() || ""; // keep the last (possibly partial) line for next read
                    for (const raw of lines) {
                        const line = raw.trim();
                        if (!line.startsWith("data:")) continue;
                        const payload = line.slice(5).trim();
                        if (!payload || payload === "[DONE]") continue;
                        let chunk;
                        try { chunk = JSON.parse(payload); } catch { continue; }
                        onChunk(chunk);
                    }
                }
                return true;
            };

            // One attempt against one model. Returns { parsed } on success or { retryable } on failure
            // — never throws. Streams the response (SSE) rather than waiting for one giant JSON blob:
            // a non-streaming request has NOTHING flowing over the wire until the entire completion is
            // ready, which is exactly what trips an "upstream idle timeout" on a reasoning-heavy model
            // that can spend a long silent stretch "thinking" before its first token — streaming keeps
            // the connection alive with incremental chunks/keepalives throughout that stretch instead.
            // (A direct-to-Google path for Gemini models was tried and reverted — didn't show enough
            // benefit to justify the separate code path; everything goes through OpenRouter.)
            const callOnce = async (label, userPrompt, model, attemptLabel, maxTokens) => {
                const temperature = attemptLabel === "attempt 1" ? 0 : 0.3; // vary on retry to escape a bad generation
                try {
                    let content = "", finishReason = null, streamErr = null;
                    let completionTokens = null, callTotalTokens = 0, cost = null;

                    const res = await fetch(OPENROUTER_URL, {
                        method: "POST",
                        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${OPENROUTER_KEY}` },
                        body: JSON.stringify({
                            model,
                            messages: [
                                { role: "system", content: SYSTEM_PROMPT },
                                { role: "user", content: userPrompt }
                            ],
                            temperature,
                            max_tokens: maxTokens,
                            usage: { include: true },
                            stream: true
                        })
                    });
                    if (!res.ok) {
                        const t = await res.text();
                        console.warn(`[AI Model] ${label} (${model}) HTTP ${res.status} (${attemptLabel}): ${t.slice(0, 160)}`);
                        return { retryable: res.status === 429 || res.status >= 500 };
                    }
                    const onChunk = (chunk) => {
                        if (chunk.error) { streamErr = chunk.error; return; }
                        const c = chunk.choices?.[0];
                        if (c?.delta?.content) content += c.delta.content;
                        if (c?.finish_reason) finishReason = c.finish_reason;
                        if (chunk.usage) { // OpenRouter sends a final usage-bearing chunk before [DONE]
                            cost = Number(chunk.usage.cost) || 0;
                            completionTokens = chunk.usage.completion_tokens ?? completionTokens;
                            callTotalTokens = chunk.usage.total_tokens || callTotalTokens;
                        }
                    };
                    const streamed = await readSSE(res, onChunk);
                    if (!streamed) {
                        const data = await res.json();
                        const usage = data.usage || {};
                        cost = Number(usage.cost) || 0;
                        completionTokens = usage.completion_tokens;
                        callTotalTokens = usage.total_tokens || 0;
                        const c = data.choices?.[0];
                        content = c?.message?.content || "";
                        finishReason = c?.finish_reason;
                        streamErr = c?.error || data.error || null;
                    }

                    const wf = workflowOf(model);
                    if (attemptLabel === "fallback") usedFallbackModel = true;
                    totalCost += (cost || 0); totalTokens += callTotalTokens;
                    workflowTotals[wf].cost += (cost || 0); workflowTotals[wf].tokens += callTotalTokens; workflowTotals[wf].calls += 1;
                    costRows.push(`   • ${label} [${wf}/${model}]${attemptLabel !== "attempt 1" ? ` (${attemptLabel})` : ""}: ${cost != null ? "$" + cost.toFixed(5) : "n/a"}  (${callTotalTokens || "?"} tokens, ${completionTokens ?? "?"} completion / cap ${maxTokens})`);
                    const parsed = parseLenient(content);
                    if (parsed) return { parsed };
                    if (finishReason === "length") {
                        console.warn(`[AI Model] ${label} (${model}) hit the max_tokens cap (${maxTokens}) before finishing (${attemptLabel}) — output was truncated, not just malformed. Raise maxTokens for this call.`);
                    } else if (finishReason === "error" || streamErr) {
                        // A provider-side failure mid-generation — a distinct case from the model just
                        // writing bad JSON (e.g. Gemini SAFETY/RECITATION blocks, or an OpenRouter relay fault).
                        const upstreamErr = streamErr?.message || streamErr;
                        console.warn(`[AI Model] ${label} (${model}) upstream provider error (${attemptLabel})`
                            + (upstreamErr ? `: ${JSON.stringify(upstreamErr).slice(0, 300)}` : " (no further detail returned)")
                            + " — this is a provider-side fault, not a bad generation.");
                    } else {
                        console.warn(`[AI Model] ${label} (${model}) returned invalid JSON (${attemptLabel}, finish_reason=${finishReason}).`, content.slice(0, 400));
                    }
                    return { retryable: true };
                } catch (e) {
                    console.warn(`[AI Model] ${label} (${model}) call failed (${attemptLabel}):`, e.message);
                    return { retryable: true };
                }
            };

            // Resilient call: retries once on the same model (with a short backoff so transient
            // rate-limits/5xx have a chance to clear), then falls back to FALLBACK_MODEL — a DIFFERENT
            // provider — for every call, not just Assumptions. A same-provider fallback is worthless
            // against a provider-wide blip (both Gemini tiers have been observed failing in the same
            // run), and every sheet has downstream dependents, so none of them should be allowed to
            // come back completely empty just because their one provider had a bad moment. Never
            // throws — a fully-failed sheet returns {} so the rest of the model still builds.
            const callLLM = async (label, userPrompt, model, maxTokens = 32000) => {
                for (let attempt = 1; attempt <= 2; attempt++) {
                    if (attempt > 1) await sleep(700 * attempt);
                    const r = await callOnce(label, userPrompt, model, `attempt ${attempt}`, maxTokens);
                    if (r.parsed) return r.parsed;
                }
                if (model !== FALLBACK_MODEL) {
                    console.warn(`[AI Model] ${label}: ${model} failed twice — falling back to ${FALLBACK_MODEL}.`);
                    await sleep(700);
                    const r = await callOnce(label, userPrompt, FALLBACK_MODEL, "fallback", maxTokens);
                    if (r.parsed) return r.parsed;
                }
                return {}; // non-fatal fallback
            };

            // Real face value (see fetchBasicInfo) — still needed for its own "face_value" row,
            // independent of shares_out.
            aiStatus("Fetching face value...");
            const basicInfo = await basicInfoPromise;

            // shares_out is a genuine PER-YEAR figure — share count actually changes over time
            // (buybacks, ESOPs, rights issues, fresh issuance), so it's computed as Market Cap ÷
            // Price for EACH historical year individually, not one constant snapshot repeated
            // across every column (that was tried and reverted — freezing today's share count
            // across every historical year corrupts every per-share metric computed for a PAST
            // year, e.g. historical BVPS). A completely different relationship (Market Cap = Price
            // × Shares) than PAT/EPS, which breaks the circularity a PAT ÷ EPS derivation would
            // have created (the P&L's own EPS forecast formula is PAT ÷ shares_out).
            // NOTE: "estimate" comes back as the literal STRING "no" for actual years, not
            // null/blank — "no" is truthy in JS, so a plain "!r.estimate"/"r.estimate" check treats
            // EVERY row (both "yes" and "no") as an estimate row and rejects everything. Must
            // compare the actual string value instead.
            const isEstimateRow = (r) => String((r && r.estimate) || "").trim().toLowerCase() === "yes";
            const rawSharesOutInputRows = await sharesOutInputsPromise;
            // Defense in depth: the backend now casts these to real numbers, but coerce again here
            // via the same parseNum() used for financial-statement text elsewhere in this file —
            // cheap insurance against a string ("6176.02...") or a literal "-" placeholder slipping
            // through unnoticed again (exactly what broke this the first time).
            const sharesOutInputRows = rawSharesOutInputRows.map(r => r && {
                ...r, market_cap: parseNum(r.market_cap), price: parseNum(r.price),
            });
            const sharesOutByFY = {}; // fy -> { value, marketCap, price, source }
            const yearsNeedingEODPrice = [];
            for (const r of sharesOutInputRows) {
                if (!r || isEstimateRow(r) || r.year == null) continue;
                if (r.market_cap == null || r.market_cap <= 0) continue;
                if (r.price != null && r.price > 0) {
                    sharesOutByFY[r.year] = { value: r.market_cap / r.price, marketCap: r.market_cap, price: r.price, source: "sc_year_data" };
                } else {
                    yearsNeedingEODPrice.push(r); // has Market Cap but no Price for that SAME year
                }
            }
            // FALLBACK — sc_year_data sometimes has a Market Cap for a year but no Price for that
            // SAME year (a real, observed gap, not necessarily because the company wasn't trading
            // yet — e.g. a recent IPO legitimately has no pre-listing price, but a plain data gap
            // for an already-listed year is also possible). Recover the price from the live daily
            // EOD feed instead of leaving that year blank: find the closing price on the last
            // trading day on/before that fiscal year-end. A single bulk history fetch, only made
            // when actually needed, reused across every year missing a price.
            if (yearsNeedingEODPrice.length) {
                const series = await fetchAnnualClosingPrices(fincode);
                if (series.length) {
                    for (const r of yearsNeedingEODPrice) {
                        const derivedPrice = annualClosingPriceForFY(series, r.year);
                        if (derivedPrice != null && derivedPrice > 0) {
                            sharesOutByFY[r.year] = {
                                value: r.market_cap / derivedPrice, marketCap: r.market_cap, price: derivedPrice,
                                source: `sc_year_data Market Cap ÷ NSE EOD annual closing price (FY${r.year} year-end)`,
                            };
                        }
                    }
                }
            }
            const sharesOutFYs = Object.keys(sharesOutByFY).map(Number).filter(isFinite);
            const sharesOutComputedFY = sharesOutFYs.length ? Math.max(...sharesOutFYs) : null;
            const sharesOutComputed = sharesOutComputedFY != null ? sharesOutByFY[sharesOutComputedFY].value : null;
            const marketCapUsed = sharesOutComputedFY != null ? sharesOutByFY[sharesOutComputedFY].marketCap : null;
            const priceUsed = sharesOutComputedFY != null ? sharesOutByFY[sharesOutComputedFY].price : null;
            const priceSourceNote = sharesOutComputedFY != null ? sharesOutByFY[sharesOutComputedFY].source : null;
            if (!sharesOutFYs.length) {
                if (sharesOutInputRows.length) {
                    console.warn(`[AI Model] Fetched ${sharesOutInputRows.length} sc_year_data row(s) but none had a usable non-estimate Market Cap + Price pair (even after trying the live EOD price feed as a fallback) — could not compute any shares_out figures from Market Cap ÷ Price. Falling back to the model's own derivation (verify it manually).`);
                } else {
                    console.warn(`[AI Model] Could not fetch Market Cap / Price data (excel_shares_outstanding_inputs) — shares_out could not be computed from Market Cap ÷ Price and was left for the model to derive itself (verify it manually).`);
                }
            }

            // debtor_days/inventory_days/payable_days HISTORICALS also come from sc_year_data
            // (Receivable/Inventory/Payable Days columns — "receivable days" and "debtor days" are
            // the same thing) — real reported figures, not something the model should be deriving
            // or guessing for PAST years the way it's asked to for a forward-looking assumption.
            // Only non-estimate (actual) years are used. Forced onto the Assumptions rows further
            // below (see forceWorkingCapitalDaysHistory), which also anchors the FIRST forecast
            // year to the last real actual — the taper across the remaining forecast years stays
            // genuine LLM judgment, informed by whatever qualitative guidance exists.
            const wcDaysByFY = {};
            for (const r of sharesOutInputRows) {
                if (!r || isEstimateRow(r) || r.year == null) continue;
                const nonNeg = (v) => { const n = parseNum(v); return (n != null && n >= 0) ? n : null; };
                wcDaysByFY[r.year] = {
                    debtor_days: nonNeg(r.debtor_days),
                    inventory_days: nonNeg(r.inventory_days),
                    payable_days: nonNeg(r.payable_days),
                };
            }
            const wcDaysNote = Object.keys(wcDaysByFY).length
                ? `\nWORKING CAPITAL DAYS (real reported figures, from GIAStock sc_year_data) — use these EXACT values for the HISTORICAL years listed instead of deriving or guessing them yourself; for any historical year not listed here, derive it the usual way: ${Object.entries(wcDaysByFY).sort((a, b) => Number(b[0]) - Number(a[0])).map(([fy, d]) => `FY${fy}: debtor_days=${d.debtor_days ?? "?"}, inventory_days=${d.inventory_days ?? "?"}, payable_days=${d.payable_days ?? "?"}`).join("; ")}\n`
                : "";

            // Real historical trading-multiple distribution (see fetchValuationChart above). Several
            // overlapping figures come back (trailing vs forward-estimate, mean vs median, full-
            // period vs recent-window) — without an explicit hierarchy the model could just pick
            // whichever one happens to support whatever it already wants to conclude, which defeats
            // the point of giving it real data at all. So this note is explicit about: (1) which
            // figure is the analytically CORRECT one for this use (forward-basis, since
            // target_price_pe/target_price_ev_ebitda are both built off NEXT-YEAR (FWD1) EPS/EBITDA
            // — pairing that with a TRAILING multiple would be a basis mismatch), (2) that this is
            // the company's own recomputed trading history (market cap ÷ its own PAT/EBITDA), not a
            // peer-group average, whatever the vendor's "sector" tag might suggest, and (3) the
            // recent ~6-month window alongside the full-period stats, so a structural trend
            // (compressing/expanding multiple) is visible instead of hidden inside one static
            // average.
            const valChart = await valuationChartPromise;
            const fmtMult = (v) => (typeof v === "number" && isFinite(v)) ? v.toFixed(1) : "?";
            const valuationChartNote = valChart
                ? `\nHISTORICAL TRADING MULTIPLES (GIAStock valuation-chart) — this is this company's OWN trading history, recomputed daily as (market cap) ÷ (its own trailing or consensus-estimate PAT/EBITDA); it is NOT a peer-group/comparable-companies average${valChart.sector ? ` — the "${valChart.sector}" sector tag on this data is a generic exchange classification, not a curated peer basket, and should not be read as one` : ""}.\nPRIMARY REFERENCE for target_pe / target_ev_ebitda (forward-basis — Target Price on the Valuation sheet multiplies these against NEXT-YEAR (FWD1) EPS/EBITDA, so a trailing multiple would be a basis mismatch):\n  Forward P/E — full-period: mean ${fmtMult(valChart.avg_pe_est)}, median ${fmtMult(valChart.median_pe_est)}, ±1sd [${fmtMult(valChart.pe_est_minus_1sd)}, ${fmtMult(valChart.pe_est_add_1sd)}]${valChart.recentFrom ? ` | recent ~6mo (${valChart.recentFrom} to ${valChart.recentTo}): mean ${fmtMult(valChart.recent_avg_pe_est)}, median ${fmtMult(valChart.recent_median_pe_est)}` : ""}${valChart.latest_pe_est != null ? ` | latest (${valChart.latestDate}): ${fmtMult(valChart.latest_pe_est)}` : ""}\n  Forward EV/EBITDA — full-period: mean ${fmtMult(valChart.avg_ev_ebitda_est)}, median ${fmtMult(valChart.median_ev_ebitda_est)}, ±1sd [${fmtMult(valChart.ev_ebitda_est_minus_1sd)}, ${fmtMult(valChart.ev_ebitda_est_add_1sd)}]${valChart.recentFrom ? ` | recent ~6mo: mean ${fmtMult(valChart.recent_avg_ev_ebitda_est)}, median ${fmtMult(valChart.recent_median_ev_ebitda_est)}` : ""}${valChart.latest_ev_ebitda_est != null ? ` | latest: ${fmtMult(valChart.latest_ev_ebitda_est)}` : ""}\nCONTEXT ONLY (trailing — do NOT plug these into target_pe/target_ev_ebitda directly; use only to judge whether the stock is currently trading rich/cheap versus its own history):\n  Trailing P/E — full-period: mean ${fmtMult(valChart.avg_pe)}, median ${fmtMult(valChart.median_pe)}, ±1sd [${fmtMult(valChart.pe_minus_1sd)}, ${fmtMult(valChart.pe_add_1sd)}]${valChart.recentFrom ? ` | recent ~6mo: mean ${fmtMult(valChart.recent_avg_pe)}, median ${fmtMult(valChart.recent_median_pe)}` : ""}${valChart.latest_pe != null ? ` | latest: ${fmtMult(valChart.latest_pe)}` : ""}\n  Trailing EV/EBITDA — full-period: mean ${fmtMult(valChart.avg_ev_ebitda)}, median ${fmtMult(valChart.median_ev_ebitda)}, ±1sd [${fmtMult(valChart.ev_ebitda_minus_1sd)}, ${fmtMult(valChart.ev_ebitda_add_1sd)}]${valChart.recentFrom ? ` | recent ~6mo: mean ${fmtMult(valChart.recent_avg_ev_ebitda)}, median ${fmtMult(valChart.recent_median_ev_ebitda)}` : ""}${valChart.latest_ev_ebitda != null ? ` | latest: ${fmtMult(valChart.latest_ev_ebitda)}` : ""}\nHOW TO USE THIS — READ CAREFULLY, this is a range of what the market HAS PAID, not a target to default to matching: a persistently high own-history multiple can reflect a genuine, sustainable re-rating, or it can just as easily reflect momentum/speculative sentiment that a sober base-case forecast should NOT assume will simply continue — "the market has paid 80x+ before" is, on its own, NOT sufficient justification for underwriting that multiple going forward. Treat the full-period figures as the DEFAULT starting point, and treat anything above them (the recent window, the latest print) as requiring EXTRA justification, not as an upgrade to chase. Only move toward the recent-window level (or higher) if the qualitative sources name a SPECIFIC, CURRENTLY ACTIVE catalyst that would sustain a premium multiple — e.g. a named new segment about to scale with disclosed numbers, a concrete margin-accretive mix shift already underway and quantified by management. A generic "strong growth / large optionality" narrative is NOT sufficient evidence for a premium multiple — nearly every richly-valued stock has one; it's the reason valuation debates exist, not a decisive answer to them. Absent such a specific catalyst, pick a figure at or BELOW the full-period median — do not split the difference toward the recent level "just in case." If broker/analyst reports are available and discuss a forward multiple or valuation view, weigh that too, and prefer it over this stock's own trading history where the two disagree — a stock's own past multiple can run structurally rich for reasons (illiquidity, promoter-driven narrative, sector rotation) that a forward-looking analyst view will not simply repeat. Do NOT default to a generic textbook multiple (e.g. a flat ~15-20x) either — the point is to reason from evidence in both directions, not to anchor on any single number by default.\nIn "rationale": name the EXACT field you anchored to (e.g. "median_pe_est") and its value; state explicitly whether you are picking AT/BELOW the full-period figure (the default) or ABOVE it, and if above, name the SPECIFIC catalyst from the qualitative sources that justifies it — "high growth story" or "premium historically justified" is not an acceptable rationale on its own. Example of an acceptable rationale for staying conservative: "Anchored to the full-period median_pe_est of 68x rather than the recent_median_pe_est of 82x (${fmtMult(valChart.latest_pe_est)}x latest) — no specific near-term catalyst in the qualitative sources beyond general growth optimism, so did not underwrite the recent premium." Example of an acceptable rationale for going above it: "Anchored above the full-period median_pe_est of 68x, toward 78x — Q[X] earnings call cites the new [segment] scaling from ₹[X] Cr to a guided ₹[Y] Cr by FY[Z], a disclosed, currently-active margin-accretive shift, not a generic growth narrative."\n`
                : "";

            const basicInfoNote = `\nCOMPANY BASIC INFO: `
                + (sharesOutComputed ? `Shares outstanding = Market Cap ÷ Price for EACH year (share count genuinely changes over time — buybacks, ESOPs, rights issues, fresh issuance — so this is NOT one constant figure to repeat across every column). Real computed values are available for ${sharesOutFYs.length} historical year(s) and will be applied automatically to the shares_out row — you do NOT need to compute those years yourself. Most recent: FY${sharesOutComputedFY} = Market Cap (Rs ${marketCapUsed} Cr) ÷ Price (Rs ${priceUsed}, via ${priceSourceNote}) = ${sharesOutComputed.toFixed(4)} crore shares. For any historical year NOT covered by this, derive shares_out yourself the SAME way (Market Cap ÷ Price for that specific year) where you can, citing your source; only flat-carry from the nearest known year as a last resort. `
                    : `No reliable Market Cap/Price pairs were found to derive shares outstanding automatically — derive shares_out YOURSELF as Market Cap ÷ Price for EACH historical year (share count genuinely changes over time, so do NOT just repeat one constant figure across every column) and cite your source. Do NOT derive it as Share Capital ÷ Face Value (breaks down whenever face value has changed, e.g. a stock split/consolidation) and do NOT derive it as PAT ÷ EPS (this model's own EPS forecast formula is PAT ÷ shares_out, so that would make the two rows circular). `)
                + (basicInfo && basicInfo.faceValue ? `Face value = Rs ${basicInfo.faceValue} per share — this is a REAL, already-fetched figure, NOT something to look up or guess from the data yourself (do NOT default to a generic value like Rs 10 — that is frequently wrong). You MUST separately set face_value = EXACTLY ${basicInfo.faceValue} on its own row (key=face_value) and its "source" to "Live company data". This face value is ONLY for the face_value row — it is NOT an input to shares_out. ` : "")
                + "\n";

            // Data bundles fed to the calls (statement sheets get the numeric block;
            // the assumptions call also gets the qualitative sources).
            const DATA_BASIS_NOTE = `DATA BASIS — the figures below may include more than one presentation (e.g. an "IndAS" series and a more granular "Detailed" series, and standalone vs consolidated). YOU decide which to use: prefer CONSOLIDATED figures on a SINGLE consistent reporting basis across every year and every sheet; use the Detailed breakdown only for extra line-item granularity. Do not mix bases within a series. When more than one row could plausibly be "the" figure for a line item (e.g. two similarly-labelled EBIT/EBITDA rows from different presentations), deterministically prefer whichever appears FIRST in the Key Financials sheet — do not switch between presentations across a rerun of this same prompt, since that produces a different "historical actual" figure for the same real year each time, which should never happen for a reported number.\nLINKING — whenever a row's historicals come from a line item present in the data, set "link" to that item's EXACT label (from Key Financials OR the Detailed Annual Financials below) so its historical cells become LIVE references; ALSO include "historical" numbers as a fallback. If an item is not in the data, just give "historical" numbers. Do NOT link percentage / margin / ratio rows — provide those as DECIMAL fractions (0.069 for 6.9%, never 6.9) or compute them; only link absolute figures (currency amounts, volumes, counts, share counts).\nTOTALS vs LINE ITEMS — the source tables indent sub-items UNDERNEATH the line they roll up into (more leading spaces = deeper in the hierarchy); a line whose value equals the SUM of the more-indented rows immediately below it is a TOTAL/subtotal heading, not a real line item of its own — do NOT link or transcribe it AND separately link/transcribe the items inside it, that counts the same money twice. Example: "Other Current Assets  519" followed by indented "Trade Receivables  214", "Cash  122", "Loans  178", "Others  5" (214+122+178+5=519) — 519 is a heading, not its own line; use the 4 indented rows (or whichever of them you need) instead of the 519 total. This applies everywhere a hierarchy like this appears (Balance Sheet, Cash Flow, P&L alike), not only this example. SELF-CHECK: add up every line item you actually used on a sheet with a "Total" concept (e.g. Balance Sheet assets) — the sum MUST match the reported "Total Assets" (or equivalent) figure in the source data; if it's off, you've either double-counted a total-plus-its-children or missed a line, find and fix it before returning.\nFLOW vs STOCK — do not link a row to a similarly-worded label of the WRONG kind (e.g. a per-year "Depreciation" flow row must link a P&L-style annual depreciation label, never an "Accumulated/Cumulative Depreciation" balance label; a "Capex" flow row must never link a "Gross Block" balance label). If you are not sure a label matches the concept, prefer "historical" numbers you derive yourself over a mismatched link.\nROLL-FORWARD CHECK — for balance-sheet stock items you carry across years (gross block, accumulated depreciation, equity, debt), the LATEST historical year's linked/entered value should be consistent with prior-year value + this year's flow (e.g. Gross Block ≈ prior Gross Block + this year's Capex; Accumulated Depreciation ≈ prior Accumulated Depreciation + this year's Depreciation; Equity ≈ prior Equity + this year's PAT − Dividends). If the source data is stale for the latest year (identical to the prior year despite a non-zero flow that year — a common data-lag artifact), use the roll-forward-implied value for that one year instead of silently repeating the prior year's balance.`;
            const finBlock = `${DATA_BASIS_NOTE}${basicInfoNote}\n\nKEY FINANCIALS (downloaded sheet):\n${financialText}\n`
                + (annualText ? `\nDETAILED ANNUAL FINANCIALS — Balance Sheet, Cash Flow & Detailed P&L (these labels are ALSO linkable):\n${annualText}\n` : "")
                + (quarterlyText ? `\nQUARTERLY FINANCIALS:\n${quarterlyText}\n` : "")
                + (operationalText ? `\nOPERATIONAL DATA:\n${operationalText}\n` : "");
            // Cap each qualitative field before it goes into a prompt — these were going in completely
            // uncapped, unlike the AI chat tool's own fetches (which already cap at 2600 chars/field).
            // For a reasoning-heavy model, a longer input also means more "thinking" tokens spent
            // digesting it (billed the same as output), so this cuts cost on two fronts at once, not
            // just the raw prompt-token count.
            const QUAL_CAP = 3500;
            const capQual = (s) => {
                s = (s || "").trim();
                return s.length > QUAL_CAP ? s.slice(0, QUAL_CAP) + "\n…[truncated]" : s;
            };
            const qualBlock = (earningCallText ? `\nEARNING CALL INSIGHTS (last 2 quarters):\n${capQual(earningCallText)}\n` : "")
                + (qnaText ? `\nEARNINGS CALL Q&A (last 2 quarters):\n${capQual(qnaText)}\n` : "")
                + (brokerReportText ? `\nBROKER / ANALYST REPORTS:\n${capQual(brokerReportText)}\n` : "")
                + (mgmtInterviewText ? `\nMANAGEMENT INTERVIEWS:\n${capQual(mgmtInterviewText)}\n` : "")
                + (orderBookText ? `\nORDER BOOK (last 2 quarters):\n${capQual(orderBookText)}\n` : "");
            // Real, fetched EOD price (see fetchLatestPrice) — kicked off way earlier, alongside the
            // operational data fetch, so this await shouldn't add any noticeable delay here.
            // current_price is a SINGLE scalar fact (today's price, a snapshot) — it does not have
            // "a value for FY2023," any more than "today's date" does, so only the latest point is
            // ever needed. Used directly on the Valuation & DCF sheet (see currentPriceValueNote,
            // built below) rather than as a per-year Assumptions row.
            aiStatus("Fetching current market price...");
            const latestPrice = await latestPricePromise;
            // Embedded directly into the "Current price" row's own definition in the Valuation & DCF
            // prompt below, telling the model the exact scalar "value" to use — with a citation —
            // rather than asking it to estimate one from its own (necessarily stale) training data.
            // Empty when the fetch failed, in which case that row's prompt falls back to asking the
            // model for its own best estimate with a source citation instead.
            const currentPriceValueNote = latestPrice
                ? `"value":${latestPrice.price}, "source":"NSE EOD price feed, ${latestPrice.date}" — a REAL, already-fetched figure (NOT something to estimate; use it EXACTLY, no rounding to a "nicer" number), `
                : `"value":<your best estimate>, "source":"<citation>" — no live price was fetched, so estimate this from your own knowledge and cite the source, `;

            // Trimmed finBlock for the judgment-only Assumptions call — it's setting WACC/growth/
            // margin/capex assumptions from annual trend + qualitative commentary, not doing granular
            // quarter-by-quarter transcription, so it doesn't need the full Quarterly Data grid (the
            // single biggest block in finBlock, up to 120 rows x 30 cols). Built directly from the same
            // pieces as finBlock (rather than stripping quarterlyText back out of the assembled string)
            // so it can't accidentally mis-match a section header. The structural Assumptions call and
            // the other sheets still get the full finBlock.
            const finBlockJudgment = `${DATA_BASIS_NOTE}\n\nKEY FINANCIALS (downloaded sheet):\n${financialText}\n`
                + (annualText ? `\nDETAILED ANNUAL FINANCIALS — Balance Sheet, Cash Flow & Detailed P&L (these labels are ALSO linkable):\n${annualText}\n` : "")
                + (operationalText ? `\nOPERATIONAL DATA:\n${operationalText}\n` : "");

            const periodsNote = `PERIODS (fixed): historical ACTUAL years ${histLabels.join(", ")} then FORECAST years ${fcLabels.join(", ")}. Historicals come from the downloaded data; forecasts are formulas.`;

            const GRAMMAR = `FORECAST FORMULA GRAMMAR — write each forecast as ONE recurrence formula using only these placeholders (we resolve them to real Excel cells); never write raw cell addresses:
- {A:assumption_key}   -> value of that assumption (Assumptions sheet)
- {R:row_key}          -> another row on the SAME sheet, SAME (current) year
- {P:row_key}          -> another row on the SAME sheet, PREVIOUS year (use for roll-forwards / growth recurrences)
- {S:sheet.row_key}    -> a row on ANOTHER sheet, SAME year   (sheet = one of: op, pnl, bs, capex)
- {PS:sheet.row_key}   -> a row on ANOTHER sheet, PREVIOUS year
Plus standard Excel: + - * / ^ , ( ) and MIN, MAX, IF, IFERROR. The SAME formula is applied to every forecast year, so prefer {P:...} to chain off the prior year. Example: "{P:revenue}*(1+{A:revenue_growth})".`;

            const rowSchemaDoc = `Return JSON exactly: {"rows":[ ROW, ... ]}. Each ROW is EITHER a section divider {"section":"SECTION NAME"} (bold header, no data) OR a line item:
{"key":"snake_case_unique","label":"Human label (unit)","fmt":"#,##0" | "0.0%" | "0.00","cagr":true|false,
 "link":"<EXACT data label>"             (OPTIONAL — set to the line item's exact label from the data to LIVE-LINK its historicals; choose from the "Linkable labels" list below. ALSO fill "historical" as a fallback),
 "historical":[${histPlaceholders}]            (the ${HIST_YEARS} historical values ${histLabels.join("/")} read straight from the data; ALWAYS provide these — EXCEPT for a row whose own instructions below explicitly say to give ONLY a formula, e.g. "do NOT give this row historical or link". For those specific rows, omit BOTH "historical" AND "link" entirely — do not guess numbers just because this schema default says to always provide them; a guessed number there silently overrides the required formula and has caused real bugs (a working-capital-change row showing round guessed numbers, an operating-cash-flow row showing a plain PAT copy) instead of the correct cross-sheet figure),
 "formula":"<forecast recurrence per the FORMULA GRAMMAR>"   (applied to each forecast year; omit for rows that are historical-only)}
Order rows top-to-bottom as they should appear. Use "0.0%" for ratios/margins (express their historical values as decimals, e.g. 0.18). Reference other sheets/assumptions ONLY by keys listed below. Be detailed: 15-30 rows.\nLinkable labels (use the EXACT text in "link"): ${linkableList.join(", ")}.`;

            // Plan a sheet's rows -> assign Excel row numbers (data starts at row 5) + record keys.
            const DATA_START = 5;
            const planRows = (sheetKey, rows) => {
                let row = DATA_START;
                for (const item of rows) { item._row = row; if (item && item.key) symbols[sheetKey][item.key] = row; row++; }
            };

            // 3a — Assumptions: a MULTI-YEAR driver schedule (like a real analyst's working sheet),
            // done first because every other sheet references its keys via {A:key} (same-year column).
            aiStatus("AI (step 1/2): assumptions & driver schedule...");
            const assumSchema = `Return JSON {"rows":[ ROW, ... ]} where each ROW is a {"section":"SECTION NAME"} divider OR a driver line:
{"key":"snake_case_unique","label":"Human label incl. unit (e.g. 'Inventory days (days)')","fmt":"#,##0"|"0.0%"|"0.00"|"0",
 "input":true|false   (true = a HARDCODED assumption, shown in BLUE; false = computed/linked),
 "link":"<EXACT data label>"             (OPTIONAL — exact label from Key Financials OR the Detailed Annual Financials, to LIVE-LINK its historicals; e.g. share capital / reserves / debt / gross block are in the Detailed Annual Financials. ALSO fill "historical" as a fallback),
 "historical":[${histPlaceholders}]            (${HIST_YEARS} historical values ${histLabels.join("/")} read straight from the data; ALWAYS provide these),
 "forecast":[${fcPlaceholders}]              (${FC_YEARS} FORECAST input values ${fcLabels.join("/")}),
 "formula":"<recurrence per the FORMULA GRAMMAR>" (OPTIONAL alternative to "forecast" for computed driver rows — normally same-sheet {R:}/{P:} for driver-to-driver relationships; {S:sheet.key}/{PS:sheet.key} IS also valid here, e.g. a repayment schedule referencing {S:pnl.pat}. Do NOT put an ABSOLUTE RUPEE line item on THIS sheet if computing it correctly needs a figure that only exists on another sheet, like revenue — this sheet has no revenue row of its own, so "{R:revenue}"-style same-sheet references for it will silently fail and resolve to 0. Keep such figures as a RATIO driver here (e.g. capex as % of revenue) and let the sheet that actually has revenue compute the rupee amount),
 "source":"<short citation>"              (REQUIRED whenever "input":true AND the number is a market/judgment figure not directly implied by the downloaded historicals — e.g. beta, risk-free rate, equity risk premium, cost of debt, target multiples, terminal growth, or a forward margin/growth call. Keep it short and concrete, e.g. "NSE 3Y adjusted beta", "10Y G-Sec yield, Jul-2026", "Sector avg EV/EBITDA (peer comps)", "Mgmt guidance, Q1FY27 earnings call". Omit for rows that are plainly derived/linked from the data.),
 "rationale":"<1-2 sentence reasoning>"  (Assumptions (Judgment) call ONLY — REQUIRED on every row reflecting a genuine forward-looking judgment call, not a plain transcription. Explain the SPECIFIC evidence or reasoning behind this number or its year-by-year trajectory, e.g. "Q1FY27 call flagged capacity expansion in MP; growth tapers as expansion completes by FY29" or "Peer set (SAIL, JSW) trades 8-11x EV/EBITDA; used mid-point given diversification premium". This becomes a note on the cell — write it for a colleague reviewing the model, not as a formality. Not applicable to the Structural call's pure transcription rows.)}
The sheet is laid out ACROSS YEARS (columns ${periodLabels.join(", ")}). Other sheets read a driver via {A:key} from the SAME year column, so for CONSTANT global inputs repeat the same value in ALL ${HIST_YEARS} historical AND ${FC_YEARS} forecast cells. None of the constant global inputs may be 0 or blank. wacc MUST be at least 1.5 percentage points greater than terminal_growth (typical: wacc approx 0.11-0.14, terminal_growth approx 0.03-0.05) so the DCF terminal-value division never approaches zero.`;
            // Split Assumptions itself along the static/dynamic line, same as the rest of the model:
            // most of its rows (share capital, debt balances, fixed-asset roll-forward, investments,
            // working-capital balances) are the same mechanical transcription the STATIC sheets already
            // do — they don't need qualBlock or a reasoning-heavy model. Only the rows that set a
            // forward-looking rate/ratio/multiple (WACC inputs, growth/margin, capex-as-%-of-revenue,
            // working-capital DAYS) benefit from reading earnings-call commentary and genuinely warrant
            // the dynamic model. Splitting this way — by what needs judgment, not by row count — also
            // shrinks the dynamic call's output (was 35-60 rows citing all of Assumptions; now ~15-20)
            // and its input (trimmed finBlock, capped qualBlock), which directly reduces the "thinking"
            // token spend that was driving both the cost and the timeout risk on this call.
            // "Segment Drivers" — a THIRD, narrowly-scoped call running in true parallel with the two
            // below (not a sequential extra step: it needs operationalSegments, which is code-derived
            // from the SAME operational fetch already awaited above, and doesn't depend on either of
            // this Promise.all's other two outputs). Fires ONLY when there are at least 2 candidate
            // segment names to consider at all — the vast majority of single-business companies never
            // even attempt this call. Exists because asking the already-overloaded Judgment call above
            // to BOTH invent correct segment names from scratch AND size them proved unreliable in
            // practice (a real run on Adani Enterprises set holding_discount>0 — correctly recognizing
            // a conglomerate — but never emitted a single revenue_growth_<tag> pair, silently falling
            // back to the fully-blended model this whole SOTP path exists to avoid). Handing this call
            // a code-derived candidate list removes that specific failure mode: it only has to DECIDE
            // (from real candidates) and SIZE, never invent a name from raw text.
            const wantSegmentDriversCall = operationalSegments.length >= 2;
            const segmentDriversPrompt = wantSegmentDriversCall
                ? `${periodsNote}\n\n${GRAMMAR}\n\n${assumSchema}\n\n${operationalText ? `OPERATIONAL DATA:\n${operationalText}\n\n` : ""}${qualBlock}${valuationChartNote}\nCANDIDATE SEGMENTS (from ${companyName}'s own reported operational/business-segment breakdown): ${operationalSegments.map(s => `"${s}"`).join(", ")}.\nYour ONLY job is to decide which of these candidates are genuinely DISTINCT BUSINESSES an investor would value separately (the SAME test used elsewhere for a conglomerate holding-company discount: different segments of ONE coherent business — e.g. different product lines, or domestic vs export sales of the same business — do NOT count; only substantially UNRELATED businesses do, each better benchmarked against its own pure-play peers), and for EACH ONE you select, provide its own growth/margin/capex/valuation-multiple assumptions.\nIf FEWER THAN 2 candidates are genuinely distinct businesses (including if ALL of them are really facets of one business, or this turns out to be a single-business company after all), return {"rows":[]} — an empty array, nothing else. Do NOT force a segment split that isn't real just because candidates were listed above.\nFor each segment you DO select, use a short snake_case tag derived from its name (e.g. "Airports (Adani Airport Holdings Ltd)" -> "airports", "Data Centers (AdaniConneX)" -> "data_centers") and provide FOUR rows keyed EXACTLY: revenue_growth_<tag>, ebitda_margin_<tag>, capex_pct_revenue_<tag>, target_ev_ebitda_<tag> — all four together, no gaps. revenue_growth_<tag>, ebitda_margin_<tag>, and capex_pct_revenue_<tag> are decimal fractions (e.g. 0.12 for 12%), each needing its OWN genuine multi-year trajectory (a tapering growth curve, a margin path, easing capex intensity as a build-out completes) grounded in whatever the qualitative sources or operational KPIs above actually say about THAT SPECIFIC segment — do NOT just copy one flat number across all ${HIST_YEARS + FC_YEARS} year cells unless the evidence genuinely supports a flat path, and do NOT copy another segment's trajectory. target_ev_ebitda_<tag> is a multiple (e.g. 12.0), one considered judgment held constant across the years like the company-level target_ev_ebitda, anchored to THAT segment's own peer set — a mature commodity/trading segment should get a materially different multiple than an infrastructure or early-growth segment, not the same multiple as the rest of the company. EVERY row requires "source" AND "rationale" citing the SPECIFIC evidence behind it (an earnings-call figure, a stated capacity/ramp-up plan, a named peer's multiple) — "general growth optimism" is not sufficient.\nDo NOT provide the plain company-wide revenue_growth/ebitda_margin/capex_pct_revenue/target_ev_ebitda keys (no <tag> suffix) — a separate call already owns those; this call ONLY provides the <tag>-suffixed segment-specific versions, one full set of four per segment selected.`
                : "";
            const [structJson, judgJson, segDriversJson] = await Promise.all([
                callLLM("Assumptions (Structural)",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${assumSchema}\n\n${finBlock}\nBuild PART of a DETAILED ASSUMPTIONS / DRIVER SCHEDULE for ${companyName} — ONLY these sections (a separate call handles WACC/growth/margin/capex/working-capital-day assumptions, so do NOT include those here): SHARE CAPITAL (shares outstanding — key=shares_out; face value, as ITS OWN separate row — key=face_value; see COMPANY BASIC INFO above for the real fetched figures for both, if given), RESERVES & DEBT (reserves, secured/unsecured debt balances), FIXED ASSETS ROLL-FORWARD (gross block beginning/additions/ending, accumulated depreciation, net block, CWIP, depreciation as % of avg gross block — key=depreciation_pct_gross_block: give this as a PLAIN POSITIVE DECIMAL VALUE (NOT a formula), computed EXACTLY as this company's OWN annual depreciation charge / average reported gross block, using ITS actual reported figures — do NOT force it toward a generic "typical" range; the correct ratio varies enormously by industry (asset-heavy, long-asset-life businesses like refining, oil & gas, power, telecom towers legitimately run well under 3%, while asset-light services businesses can run well over 8% — trust this company's own historical arithmetic, not a rule of thumb), and repeat that SAME value in all ${HIST_YEARS + FC_YEARS} year cells. Do NOT compute it with a same-sheet formula that divides by a "gross block ending" roll-forward row — that roll-forward starts unanchored and can collapse to ~0, which makes this ratio 0 and SILENTLY ZEROES the entire forecast depreciation → EBIT → PAT chain on the P&L. It MUST be a non-zero positive decimal in every cell. Also make the "gross block ending" display row anchor correctly as beginning + additions using same-year {R:} references (e.g. "{R:gross_block_beginning}+{R:capex}"), never rolled off only the prior-year ending. CONTINUITY SELF-CHECK — before finalizing, multiply your depreciation_pct_gross_block by the projected average gross block for the FIRST forecast year and confirm the result is reasonably close to (not a multi-x jump from) the LAST historical year's actual depreciation charge, since gross block does not change overnight; if your ratio produces a first-forecast-year figure several times larger or smaller than the last actual, you have the wrong ratio — recompute it directly from the historical annual-charge ÷ average-gross-block arithmetic instead. The P&L sheet references this EXACT key to forecast the annual depreciation charge), INVESTMENTS, WORKING CAPITAL BALANCES (inventories, debtors, cash, loans & advances, payables, provisions — the ACTUAL reported rupee amounts; do NOT include the days/ratio assumptions for these, that's the other call). This is pure transcription/linking from the data, no forward-looking judgment needed here.\nHISTORICAL-ONLY ROWS — every row you produce in THIS ENTIRE CALL, in EVERY section (SHARE CAPITAL, RESERVES & DEBT, FIXED ASSETS ROLL-FORWARD, INVESTMENTS, and WORKING CAPITAL BALANCES alike — this explicitly includes gross_block_beginning, the additions/capex row, and accumulated_depreciation_beginning under FIXED ASSETS ROLL-FORWARD, not only the 3 sections that happen to be named in this sentence), EXCEPT shares_out and depreciation_pct_gross_block, is a REFERENCE/DISPLAY snapshot only — no other sheet reads it. Give these rows ONLY "link"/"historical" (numbers straight from the data). Do NOT give them a "forecast" array or a "formula", and above all do NOT invent a growth-rate key like "{A:long_term_debt_growth}" or a cross-sheet pull like "{S:pnl.dividends}"/"{S:pnl.revenue}"/"{S:pnl.cogs}"/"{S:capex.capex_total}" for them — none of those keys exist anywhere (the real ones are net_revenue, not revenue; capex, not capex_total; there is no separate "cf" sheet at all, cash flow lives on capex; and there is no "cogs" or "dividends" row exposed by any sheet), and referencing an invented one resolves to a visible #MISSING error. The Balance Sheet sheet is the ACTUAL, correctly-forecast source for every one of these line items (receivables, inventory, payables, provisions, borrowings, investments) — leaving these rows historical-only here causes them to display a sensible flat continuation automatically, which is the correct behavior for a reference snapshot.\nSHARES OUTSTANDING UNIT — shares_out MUST be expressed in CRORES OF SHARES, consistent with every other figure in this model being in Rs Crores (e.g. a company with 676 crore shares outstanding — 6.76 billion / 6,760 million shares — must show shares_out = 676, NOT 6760000000, NOT 6760, NOT 6.76). If the source data reports the share count in millions or as a raw absolute count, CONVERT it yourself (crore = raw_count / 1e7 = million_count / 10) before writing shares_out. Only ONE row may describe share count (key="shares_out") — do not also add a separate informational "shares outstanding" row.\nYou MUST include rows keyed EXACTLY, with a value in every one of the ${HIST_YEARS + FC_YEARS} year cells: shares_out (source citation REQUIRED), depreciation_pct_gross_block (no source citation needed, this comes from the fixed-asset roll-forward, not market data). You MUST ALSO include face_value as its OWN row (key=face_value, source citation REQUIRED) — do NOT fold it into shares_out or omit it; if COMPANY BASIC INFO above gave you a real fetched face value, use it EXACTLY (never default to a generic value like Rs 10 without checking — that is frequently wrong). Be thorough (20-30 rows including section headers).`,
                    STATIC_MODEL),
                callLLM("Assumptions (Judgment)",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${assumSchema}\n\n${finBlockJudgment}${qualBlock}${wcDaysNote}${valuationChartNote}\nYou are the equity analyst responsible for forecasting ${companyName}. You have the company's financial history, its qualitative disclosures (earnings calls, broker research, management commentary), and its recent operational data. Form a genuine, evidence-based view on where this business is headed, and translate that view into the specific rows below — read the qualitative sources the way an analyst actually would: weigh what management is guiding against what the numbers already show, note where guidance seems optimistic or conservative relative to the company's own recent trend, and let that judgment show up in the shape of your forecast (a tapering growth curve, a margin path reflecting a stated cost initiative, a WACC that reflects the company's actual risk profile). For every row that reflects a real judgment call — not a plain transcription — fill "rationale" with the specific evidence behind it: this becomes a note on the cell for whoever reviews the model, so write it for that reader, not as a formality.\nThis call covers ONLY these sections (a separate call handles share capital/debt/fixed-asset/investment/working-capital-balance transcription, so do NOT include those here): CAPEX (capex-as-%-of-revenue as a RATIO driver ONLY — key=capex_pct_revenue; do NOT put an absolute-rupee "capex" amount on this sheet — this sheet has no revenue row to multiply against, so a same-sheet reference for it would silently resolve to 0; the Capex & FCF sheet computes the actual rupee capex from this ratio via {S:pnl.net_revenue}), WORKING CAPITAL DRIVERS (debtor_days, inventory_days, payable_days, min_cash_pct_revenue — a minimum operating-cash buffer as a decimal fraction of revenue), OTHER INCOME, P&L DRIVERS (revenue growth, EBITDA margin, tax rate, dividend payout), VALUATION & WACC INPUTS, and HOLDING/CONGLOMERATE STRUCTURE (key=holding_discount — see below).\nHOLDING_DISCOUNT — a decimal fraction (e.g. 0.15 for a 15% discount) applied later to the DCF's equity value, reflecting the market's tendency to value a genuinely diversified conglomerate (e.g. a single listed entity spanning steel, power, oil, and other substantially UNRELATED businesses, each better benchmarked against its own pure-play peers) below the sum of its parts. Decide this from the Operational data's own segment breakdown and the qualitative sources — do NOT apply a discount just because a company has multiple product lines within ONE coherent business (e.g. different steel products, or domestic vs export sales of the same business), only when the segments are genuinely DIFFERENT businesses an investor would otherwise value separately. Set holding_discount to EXACTLY 0 (not blank) for a single-business company — this key must always have a value, never left out. Typical range when it does apply: 0.10-0.30. Requires "source" AND "rationale" explaining WHY (e.g. "Diversified steel/power/oil conglomerate — no pure-play peer spans all three segments" or "Single-business steel producer — no holding discount applicable"). NOTE: a separate, dedicated pass (not this call) may ALSO derive genuine segment-specific growth/margin/capex/multiple assumptions from the company's real reported segment breakdown when one exists — if that happens, this discount is layered on TOP of an already-correct sum-of-the-parts valuation, so it should reflect only a genuine RESIDUAL (illiquidity, governance/related-party complexity, cross-holding structure) rather than trying to correct for a blending distortion — set it using your own best judgment as described above regardless; you do not need to predict whether that separate pass will fire for this company.${brokerReportText ? `\nBROKER TARGET PRICE — the BROKER / ANALYST REPORTS section above contains real broker commentary. If it states an explicit numeric target price (if multiple are given, use their AVERAGE), include an ADDITIONAL row key=broker_target_price with that value (repeated across every cell, same convention as any other constant global input) and a "source" citation identifying the broker/report. If no explicit numeric target price is stated anywhere in that material, OMIT this row entirely — do NOT guess, estimate, or derive one yourself.\n` : ""}\nWORKING-CAPITAL DAYS — HISTORICALS ARE REAL DATA, NOT A GUESS — unlike most rows on this sheet, debtor_days/inventory_days/payable_days are REPORTED facts for past years (see WORKING CAPITAL DAYS above, if given) — treat them the same way you'd treat Revenue or PAT: transcribe the real historical figure, don't derive or estimate it. Give each of the three a genuine "historical" array for every year listed above, not a flat guess.\nWORKING-CAPITAL DAYS MUST ANCHOR TO THE LAST ACTUAL — debtor_days/inventory_days/payable_days each drive a Balance Sheet formula of days/365*revenue starting in the FIRST forecast year, but the Balance Sheet's LAST HISTORICAL year uses the REAL reported receivables/inventory/payables figure — a different basis. If your first forecast value ignores what the last reported balance sheet implies, that year's receivables/inventory/payables JUMPS discontinuously, producing an artificial one-off working-capital swing that distorts FCFF and the DCF — this has actually happened (a bogus ~₹69,000 Cr working-capital "change" flipped FCFF and value per share negative). Set each driver's FIRST forecast value to (very close to) its own LAST HISTORICAL day-count, THEN taper across the remaining years per the reasoning below. NOTE: this anchor is also enforced in code afterward as a safety net when live sc_year_data is available for this company — but get it right here too. Express percentages as decimals.\nOPERATING DRIVER PATHS — revenue_growth, ebitda_margin, capex_pct_revenue, debtor_days, inventory_days, and payable_days genuinely change year to year as a business matures or executes on stated plans. For each, decide from the qualitative evidence whether it should accelerate, decelerate, or hold steady, and let your 5-year forecast array reflect that reasoning (e.g. growth tapering from a higher near-term rate toward a sustainable long-run rate; margins drifting toward a normalized level; capex intensity easing as a build-out completes). A single flat value repeated across all 5 cells is essentially never the right answer for these six drivers — if you find yourself writing the same number five times, you haven't actually reasoned through the trajectory yet. WACC, beta, risk_free_rate, equity_risk_premium, cost_of_debt, terminal_growth, tax_rate, target_pe, and target_ev_ebitda are different: those are one considered judgment, held constant across the years, not a path.\nYour analysis needs to produce a considered view on: tax_rate, pat_retention, wacc, terminal_growth, risk_free_rate, equity_risk_premium, beta, cost_of_debt, target_pe (see HISTORICAL TRADING MULTIPLES above if given — anchor to this company's own real trading range, not a generic multiple), target_ev_ebitda (same — anchor to HISTORICAL TRADING MULTIPLES above if given), holding_discount (see above — 0 is a valid, common value, but the row and its citation are still required) — constant global inputs, so repeat the same value in every one of the ${HIST_YEARS + FC_YEARS} year cells, each with a "source" citation. You must ALSO produce: capex_pct_revenue (decimal fraction of revenue, e.g. 0.14 for 14%), debtor_days, inventory_days, payable_days (in DAYS, e.g. 45, not decimals), min_cash_pct_revenue (decimal fraction of revenue, e.g. 0.03 for 3% — set from the company's own historical average cash/revenue ratio), revenue_growth (see REVENUE GROWTH GROUNDING below), ebitda_margin (the TOTAL-COMPANY forward call, as decimals — the Operational Model references this EXACT key for the total and for any segment that a separate segment-driver pass (see NOTE above) did not size individually; one blended company-wide rate is the deliberate, honest simplification here, not a bug). Use these EXACT keys so the rest of the model can find your work.\nREVENUE GROWTH GROUNDING — the Key Financials table above includes the company's own actual historical "% Change YoY" revenue growth row for every reported year. Read that real trajectory before setting revenue_growth's forecast path: note the underlying trend once you look past one-off spikes or drops (e.g. a low base year mechanically inflating the next year's % change, or a demerger/acquisition step-change that doesn't represent organic growth), and anchor your forecast near that trend rather than a generic industry-average growth rate. If a forecast year deliberately departs from the historical trend (e.g. tapering down from an unsustainable recent spike, or stepping up on a stated capacity expansion), say so explicitly in "rationale" — cite the actual historical % Change YoY figures you're comparing against.\nMARGIN / TAX / PAYOUT GROUNDING — the same real-history-first approach applies to ebitda_margin, tax_rate, and pat_retention: the Key Financials table gives you the raw historical rows needed to compute each one's actual trend for every reported year — EBITDA and Revenue (or an EBITDA Margin % row, if the sheet already provides one directly) for ebitda_margin; PBT and Tax (or an effective tax rate row) for tax_rate; PAT and DPS/dividend (or a payout ratio row) for pat_retention = 1 − payout ratio. Work out that historical ratio yourself for each reported year, identify the real trend (stable, expanding, normalizing toward a stated statutory rate for tax_rate, etc.), and anchor your forecast path near it rather than picking a generic sector-typical figure. Where a forecast value deliberately departs from the historical trend (e.g. margin expansion from a stated cost initiative, tax_rate stepping toward the statutory rate as a tax holiday expires), say so explicitly in "rationale" — cite the actual historical figures/ratio you derived.\nCONSENSUS ESTIMATE COLUMNS — the Key Financials table above frequently extends a few years PAST the last actual year with columns labelled "FYxxxxE" (e.g. FY2027E, FY2028E) — these are the data vendor's OWN already-computed analyst/consensus estimates for Revenue, EBITDA, PAT, EPS etc., not something you need to derive from scratch, and not your own historical transcription. For every forecast year these columns cover, treat the implied consensus growth/margin path as a genuine data point to weigh — the same way you'd weigh a broker report — alongside the historical trend and the qualitative sources, rather than deriving revenue_growth/ebitda_margin/tax_rate/pat_retention as if this consensus view didn't exist. You are not required to match consensus exactly (your own read of the qualitative evidence may reasonably differ), but if your own forecast diverges MEANINGFULLY from what these columns imply for the same year, say so explicitly in "rationale" and give the specific reason (e.g. "Consensus FY2027E implies ~18% revenue growth; held to 22% given the qualitative sources' more specific guidance on the new segment's ramp-up" or "Broadly in line with the FY2027E consensus column implying a ~15.5% EBITDA margin"). Silently ignoring an available consensus figure that contradicts your own estimate is a bigger red flag than disagreeing with it for a stated reason.\nThree hard constraints, regardless of the judgment above — these aren't about realism, they break the model mechanically if violated:\n1. wacc MUST exceed terminal_growth by at least 1.5 percentage points (typical: wacc approx 0.11-0.14, terminal_growth approx 0.03-0.05) — otherwise the DCF terminal-value division approaches zero.\n2. min_cash_pct_revenue must be a positive decimal in every one of the ${HIST_YEARS + FC_YEARS} year cells, never 0 or blank — the Balance Sheet's revolver formula divides the whole forecast's cash-balancing logic on this key, and a 0/blank here silently disables the safeguard against negative forecast cash.\n3. Sanity-check scale before finalizing: if any per-share-driven figure is off by an order of magnitude from what the company's size implies, you have a unit mismatch — fix it. Also don't let a genuinely capex-heavy, debt-carrying company deleverage all the way to NET CASH within the forecast window unless the data clearly supports it (that unrealistically collapses forecast interest expense toward zero) — if the company is guiding continued heavy investment, keep capex_pct_revenue and payout high enough that net debt stays realistic.\nBe thorough (15-20 rows including section headers).`,
                    DYNAMIC_MODEL, 40000),
                // Skipped entirely (no network call) when there weren't even 2 candidate segments to
                // consider — Promise.resolve keeps this slot's shape identical either way, so the
                // destructuring below doesn't need its own branch.
                wantSegmentDriversCall
                    ? callLLM("Segment Drivers", segmentDriversPrompt, DYNAMIC_MODEL, 20000)
                    : Promise.resolve({ rows: [] }),
            ]);
            const assumRows = [
                ...(Array.isArray(structJson.rows) ? structJson.rows : []),
                ...(Array.isArray(judgJson.rows) ? judgJson.rows : []),
                ...(Array.isArray(segDriversJson.rows) ? segDriversJson.rows : [])
            ];
            // target_ev_ebitda is one of the mandatory "constant global input" keys the
            // Assumptions (Judgment) prompt already asks for explicitly — but compliance isn't
            // guaranteed, and unlike wacc/tax_rate/etc there's no purely mechanical formula to
            // force it with afterward (it's a genuine judgment call: the analyst's chosen target
            // multiple). If it's missing entirely, val.target_price_ev_ebitda has nothing to
            // reference and goes silently blank — along with the blended target_price averaging
            // it. Fall back to the company's own CURRENTLY reported EV/EBITDA (a defensible "hold
            // today's multiple" default, clearly labeled as a fallback) rather than leave it
            // blank. Inserted here, before the snapshot-table split below, so it flows through the
            // exact same row-planning logic as every other assumption.
            if (!assumRows.some(r => r && r.key === "target_ev_ebitda")) {
                const currentEvEbitda = findEntryByLabel(/^ev\s*\/?\s*ebitda$/i);
                const fy = currentEvEbitda
                    ? Object.keys(currentEvEbitda.valByFY).map(Number).filter(isFinite).sort((a, b) => b - a)[0]
                    : null;
                const val = fy != null ? currentEvEbitda.valByFY[fy] : null;
                if (typeof val === "number" && isFinite(val) && val > 0) {
                    assumRows.push({
                        key: "target_ev_ebitda", label: "Target EV/EBITDA (x)", fmt: "0.00", scalar: true, value: val,
                        source: `Fallback — model did not provide a target multiple; using this company's own current reported EV/EBITDA (FY${fy})`,
                    });
                } else {
                    console.warn("[AI Model] target_ev_ebitda was not provided by the model and no reported EV/EBITDA figure was found to derive a fallback from — val.target_price_ev_ebitda (and the blended target_price averaging it) will be left blank. Verify manually.");
                }
            }
            // ── Percent-row unit guard — whole-number percentage written where a DECIMAL was asked for ──
            // Every %-formatted row in this model is supposed to hold a DECIMAL FRACTION (0.22 for
            // 22%) — the row schema, the Assumptions prompt, and DATA_BASIS_NOTE all state that
            // explicitly ("0.069 for 6.9%, never 6.9"). The model still periodically writes the
            // whole-number form anyway, and since the cell carries a "0.0%" number format, Excel
            // then renders it ×100. Observed live: a revenue_growth the model had correctly reasoned
            // to 22% in its own rationale text was written as 22 and displayed as 2200%, which then
            // compounded through every forecast year of the model.
            // Detected PER ROW (the slip is systematic across a row, never a single stray cell) and
            // only above a threshold where the two readings can't be confused: 1.5 = 150%, which no
            // margin, tax rate, retention, WACC or terminal-growth figure can legitimately reach,
            // and which a sustained revenue-growth assumption essentially never does — while a
            // genuine decimal (≤1.0, i.e. ≤100%) is never touched. Deliberately conservative: it
            // would rather miss an exotic 120%-growth row than corrupt a correct one.
            const PERCENT_UNIT_THRESHOLD = 1.5;
            const normalizePercentRows = (sheetLabel, rows) => {
                for (const item of rows) {
                    if (!item || item.section || !/%/.test(item.fmt || "")) continue;
                    const arrays = [item.historical, item.forecast].filter(Array.isArray);
                    const hasScalar = typeof item.value === "number" && isFinite(item.value);
                    const values = arrays.flat().filter(v => typeof v === "number" && isFinite(v));
                    if (hasScalar) values.push(item.value);
                    if (!values.length) continue; // pure-formula row (roe/roce/margins) — nothing hardcoded to fix
                    const peak = Math.max(...values.map(Math.abs));
                    if (peak < PERCENT_UNIT_THRESHOLD) continue;
                    for (const arr of arrays) {
                        for (let i = 0; i < arr.length; i++) {
                            if (typeof arr[i] === "number" && isFinite(arr[i])) arr[i] = arr[i] / 100;
                        }
                    }
                    if (hasScalar) item.value = item.value / 100;
                    console.warn(`[AI Model] ${sheetLabel}.${item.key || item.label} looks like a WHOLE-NUMBER percentage on a %-formatted row (largest |value| was ${peak}) — divided every cell by 100 so it reads as the decimal fraction the schema requires. Left as-is, Excel's "%" format would have rendered it ×100 (e.g. 22 → "2200%") and fed that into every downstream forecast.`);
                }
            };
            normalizePercentRows("assum", assumRows);
            // ── SOTP (sum-of-the-parts) segment detection ──
            // Segment-specific driver rows (revenue_growth_<tag>, ebitda_margin_<tag>, etc.) now come
            // from the dedicated "Segment Drivers" call above (segDriversJson, merged into assumRows
            // already) — NOT from the Assumptions (Judgment) prompt, which no longer asks for them at
            // all (see its NOTE on holding_discount). This scan just picks up whatever tags actually
            // landed in assumRows, from whichever call(s) contributed rows to it, so nothing here
            // needed to change when the segment-driver source moved. Consumed by the Operational
            // Model / Capex / Valuation & DCF calls below. Every consumer of segmentTags/
            // isSotpCompany MUST fall back to today's single-blended-DCF behavior when isSotpCompany
            // is false — that fallback is what keeps every non-conglomerate company's model
            // byte-identical to before this existed.
            //
            // Requires BOTH a matching revenue_growth_<tag> AND ebitda_margin_<tag> pair before
            // trusting a tag as real — the Segment Drivers prompt asks for all four <tag> keys
            // together, but LLM compliance with "always provide X" instructions is never guaranteed
            // (see the target_ev_ebitda fallback immediately above, needed for exactly this reason),
            // and a half-present segment (growth key but no margin key, or vice versa) is worse than
            // no segment data at all — the downstream formula would resolve to #MISSING or silently 0.
            const SEGMENT_KEY_RE = /^revenue_growth_(.+)$/;
            const segmentTags = assumRows
                .filter(r => r && typeof r.key === "string" && SEGMENT_KEY_RE.test(r.key))
                .map(r => r.key.match(SEGMENT_KEY_RE)[1])
                .filter(tag => assumRows.some(r => r && r.key === `ebitda_margin_${tag}`));
            const holdingDiscountRow = assumRows.find(r => r && r.key === "holding_discount");
            const holdingDiscountVal = holdingDiscountRow
                ? Number((Array.isArray(holdingDiscountRow.forecast) && holdingDiscountRow.forecast[0])
                    ?? (Array.isArray(holdingDiscountRow.historical) && holdingDiscountRow.historical[0])
                    ?? holdingDiscountRow.value ?? 0)
                : 0;
            // Require BOTH signals to agree before treating this as a genuine SOTP company — either
            // one alone being "wrong" (holding_discount>0 but no usable segment keys emitted, or
            // segment keys present despite holding_discount=0) falls back to the safe default
            // instead of attempting a partial/inconsistent segment build.
            const isSotpCompany = segmentTags.length >= 2 && holdingDiscountVal > 0;
            if (holdingDiscountVal > 0 && segmentTags.length < 2) {
                console.warn(`[AI Model] holding_discount=${holdingDiscountVal} (genuine conglomerate) but fewer than 2 usable segment-keyed driver pairs were provided (found: ${segmentTags.join(", ") || "none"}) — falling back to the single blended DCF for this run.`);
            }
            if (isSotpCompany) console.log(`[AI Model] SOTP segments detected: ${segmentTags.join(", ")}`);
            // Per-segment fallbacks for capex_pct_revenue_<tag> and target_ev_ebitda_<tag> — unlike
            // revenue_growth_<tag>/ebitda_margin_<tag> (which GATE whether a tag counts as a real
            // segment at all, see segmentTags above), these two are only referenced by ONE row each
            // (buildSegmentCapexRows' capex_<tag>, and the Valuation sheet's sotp_ev_<tag>) — a
            // segment missing just one of these shouldn't take the whole segment (or, worse, the
            // whole capex/SOTP-EV SUM formula, which would inherit a #MISSING from any one term)
            // down with it. Falls back to the already-guaranteed company-wide blended value for
            // that ONE key on that ONE segment, same reasoning as the company-level target_ev_ebitda
            // fallback above.
            if (isSotpCompany) {
                const companyCapexPct = assumRows.find(r => r && r.key === "capex_pct_revenue");
                const companyTargetEvEbitda = assumRows.find(r => r && r.key === "target_ev_ebitda");
                const scalarLikeValue = (row) => row && ((row.scalar ? row.value : null)
                    ?? (Array.isArray(row.forecast) && row.forecast[0])
                    ?? (Array.isArray(row.historical) && row.historical[0]));
                for (const tag of segmentTags) {
                    const capexKey = `capex_pct_revenue_${tag}`;
                    if (!assumRows.some(r => r && r.key === capexKey)) {
                        const val = scalarLikeValue(companyCapexPct);
                        if (typeof val === "number" && isFinite(val)) {
                            assumRows.push({ key: capexKey, label: `Capex / Revenue — ${tag} (%)`, fmt: "0.0%", input: true, historical: new Array(HIST_YEARS).fill(val), forecast: new Array(FC_YEARS).fill(val), source: "Fallback — Segment Drivers call did not size this segment's own capex intensity; using the company-wide capex_pct_revenue instead" });
                            console.warn(`[AI Model] ${capexKey} was not provided — falling back to the company-wide capex_pct_revenue (${val}) for this segment.`);
                        } else {
                            console.warn(`[AI Model] ${capexKey} was not provided and no company-wide capex_pct_revenue fallback value was available either — this segment's capex row will show #MISSING.`);
                        }
                    }
                    const evKey = `target_ev_ebitda_${tag}`;
                    if (!assumRows.some(r => r && r.key === evKey)) {
                        const val = scalarLikeValue(companyTargetEvEbitda);
                        if (typeof val === "number" && isFinite(val)) {
                            assumRows.push({ key: evKey, label: `Target EV/EBITDA — ${tag} (x)`, fmt: "0.00", scalar: true, value: val, source: "Fallback — Segment Drivers call did not provide this segment's own target multiple; using the company-wide target_ev_ebitda instead" });
                            console.warn(`[AI Model] ${evKey} was not provided — falling back to the company-wide target_ev_ebitda (${val}) for this segment's SOTP EV term.`);
                        } else {
                            console.warn(`[AI Model] ${evKey} was not provided and no company-wide target_ev_ebitda fallback value was available either — this segment's SOTP EV term will show #MISSING.`);
                        }
                    }
                }
            }
            // Single-fact market/valuation inputs (tax rate, WACC, beta, target multiples, etc.)
            // don't have "a value per year" the way a driver schedule does — every forecast year
            // reads the SAME number via {A:key}. Pull these out of the main per-year grid entirely
            // (rather than just collapsing them to a scalar INLINE, which still left 9 blank
            // columns per row) and render them in their own small table below the main grid
            // instead — see the MARKET / VALUATION SNAPSHOT block further down.
            const MARKET_SNAPSHOT_KEYS = ["tax_rate", "pat_retention", "risk_free_rate", "equity_risk_premium", "beta", "cost_of_debt", "wacc", "terminal_growth", "target_pe", "target_ev_ebitda", "broker_target_price"];
            const snapshotRows = assumRows.filter(r => r && r.key && MARKET_SNAPSHOT_KEYS.includes(r.key));
            const mainAssumRowsRaw = assumRows.filter(r => !(r && r.key && MARKET_SNAPSHOT_KEYS.includes(r.key)));
            // Drop any section header left with nothing under it now that its rows were pulled out.
            const mainAssumRows = mainAssumRowsRaw.filter((r, i) => {
                if (!r || !r.section) return true;
                const next = mainAssumRowsRaw[i + 1];
                return next && !next.section;
            });
            planRows("assum", mainAssumRows);
            // Snapshot rows get their OWN sequential row numbers, positioned right after the main
            // grid ends (planRows can't be reused as-is — it always starts counting from
            // DATA_START, which would overlap the main grid).
            const assumSnapshotTitleRow = (mainAssumRows.length ? mainAssumRows[mainAssumRows.length - 1]._row : DATA_START - 1) + 2;
            { let r = assumSnapshotTitleRow + 1; for (const item of snapshotRows) { item._row = r; if (item.key) symbols.assum[item.key] = r; r++; } }
            const assumKeyList = assumRows.filter(r => r && r.key).map(r => `${r.key} = ${r.label || ""}`).join("\n");

            // Diagnostic only (no auto-correction — guessing the "right" scale could just substitute a
            // different wrong number silently). shares_out MUST be in crores of shares per the prompt;
            // a unit slip here (raw count / millions / billions) has been observed to wreck EPS, BVPS,
            // and every multiple built on them, including in HISTORICAL years.
            const sharesOutRow = assumRows.find(r => r && r.key === "shares_out");
            const sharesOutVal = sharesOutRow && (sharesOutRow.forecast?.[0] ?? sharesOutRow.historical?.[0]);
            if (typeof sharesOutVal === "number" && (sharesOutVal > 50000 || sharesOutVal < 0.5)) {
                console.warn(`[AI Model] shares_out = ${sharesOutVal} looks implausible for "crores of shares" (expected roughly 1-5000 for most listed companies) — likely a unit mismatch (raw share count, millions, or billions instead of crores), which will corrupt EPS/BVPS/every per-share multiple. Verify manually.`);
            }

            // depreciation_pct_gross_block must be a positive constant, company-specific (asset-heavy,
            // long-life businesses can legitimately sit well under 0.03; asset-light ones well over 0.08
            // — so this check only flags 0/blank/formula, never a specific numeric range). If the model
            // returned it as a FORMULA (dividing by a gross-block roll-forward) or as 0, it can resolve
            // to 0 at calc time and SILENTLY ZERO the whole forecast depreciation -> EBIT -> PAT chain.
            const depPctRow = assumRows.find(r => r && r.key === "depreciation_pct_gross_block");
            if (depPctRow) {
                const depPctVal = depPctRow.forecast?.[0] ?? depPctRow.historical?.[0];
                if (depPctRow.formula) {
                    console.warn(`[AI Model] depreciation_pct_gross_block was returned as a FORMULA ("${depPctRow.formula}") instead of a plain decimal — if it divides by a gross-block roll-forward that collapses to 0, it zeroes ALL forecast depreciation/EBIT/PAT. Verify the P&L forecast depreciation is non-zero and continuous with the last historical year.`);
                } else if (!(typeof depPctVal === "number" && depPctVal > 0)) {
                    console.warn(`[AI Model] depreciation_pct_gross_block = ${depPctVal} (should be a positive decimal, company-specific) — a 0/blank here silently zeroes the entire forecast depreciation/EBIT/PAT chain. Verify manually.`);
                }
            }

            // min_cash_pct_revenue funds the Balance Sheet's short_term_borrowings revolver, which is
            // what keeps forecast cash from ever going negative. A 0/blank here disables that revolver.
            const minCashRow = assumRows.find(r => r && r.key === "min_cash_pct_revenue");
            if (minCashRow) {
                const minCashVal = minCashRow.forecast?.[0] ?? minCashRow.historical?.[0];
                if (!(typeof minCashVal === "number" && minCashVal > 0)) {
                    console.warn(`[AI Model] min_cash_pct_revenue = ${minCashVal} (should be a small positive decimal, e.g. 0.02-0.06) — this funds the Balance Sheet's revolver that prevents forecast cash from going negative; a 0/blank here disables that safeguard. Verify the Balance Sheet's forecast cash is non-negative.`);
                }
            } else {
                console.warn(`[AI Model] min_cash_pct_revenue was not returned by the Assumptions call — the Balance Sheet's revolver formula references {A:min_cash_pct_revenue}, which will resolve to 0/#MISSING without it, disabling the safeguard against negative forecast cash. Verify the Balance Sheet's forecast cash is non-negative.`);
            }

            // debtor_days/inventory_days/payable_days are real REPORTED figures for historical
            // years (see wcDaysByFY above) — force their "historical" arrays to the actual
            // sc_year_data values rather than trust the model's own transcription/derivation, and
            // anchor the FIRST forecast year to the last real actual (the exact rule the prompt's
            // own "WORKING-CAPITAL DAYS MUST ANCHOR" paragraph asks for, guaranteed here regardless
            // of prompt compliance — this is what prevents receivables/inventory/payables from
            // jumping discontinuously in year one, per the ~₹69,000 Cr bogus-swing incident already
            // documented in that prompt section). Years 2-5 of the forecast taper are left as the
            // model's own judgment call, informed by qualitative guidance — genuine forecasting,
            // not something code should be dictating.
            const forceWorkingCapitalDaysHistory = (key, labelMatch) => {
                const row = (symbols.assum && symbols.assum[key] && assumRows.find(r => r && r._row === symbols.assum[key]))
                    || assumRows.find(r => r && !r.section && r.key === key)
                    || assumRows.find(r => r && !r.section && r.label && labelMatch.test(r.label));
                if (!row) {
                    console.warn(`[AI Model] Could not locate ${key} on the Assumptions sheet to force its historicals from sc_year_data — row not found.`);
                    return;
                }
                const fetched = HIST.map(fy => (wcDaysByFY[fy] ? wcDaysByFY[fy][key] : null));
                if (fetched.every(v => v == null)) {
                    console.warn(`[AI Model] No sc_year_data ${key} figures found for any historical year — left as the model's own transcription/derivation. Verify manually.`);
                    return;
                }
                // MERGE, don't replace: only overwrite years sc_year_data actually covers, keeping
                // the model's own (prompt-instructed) derivation for any year it doesn't (e.g. an
                // early year missing from sc_year_data) — a flat null/gap there was previously
                // rendering as a blank or confusing cell instead of the model's real fallback value.
                const existing = Array.isArray(row.historical) ? row.historical.slice(0, HIST_YEARS) : new Array(HIST_YEARS).fill(null);
                const merged = fetched.map((v, i) => (v != null ? v : (existing[i] != null ? existing[i] : null)));
                row.historical = merged;
                row.source = "sc_year_data (Receivable/Inventory/Payable Days)";
                // These labels ("Inventory Days" etc.) are common enough to ALSO appear (often on a
                // different basis/timing) in Key Financials/Annual Data — autoLink's own label/value
                // matching would happily rediscover that OTHER row and LIVE-LINK the whole row to
                // it, silently overriding these authoritative sc_year_data figures with a different
                // one for every year the other row happens to cover (this has actually happened — a
                // forced FY value got overridden this way). sc_year_data is the source of truth for
                // these 3 keys now, so disable auto-relinking entirely, not just the explicit "link".
                delete row.link;
                row.noAutoLink = true;
                const lastActual = [...merged].reverse().find(v => v != null);
                if (lastActual != null && Array.isArray(row.forecast) && row.forecast.length) {
                    row.forecast[0] = lastActual;
                }
            };
            forceWorkingCapitalDaysHistory("debtor_days", /debtor\s*days|receivable\s*days/i);
            forceWorkingCapitalDaysHistory("inventory_days", /inventory\s*days/i);
            forceWorkingCapitalDaysHistory("payable_days", /payable\s*days/i);

            // Canonical cross-sheet keys. Because the remaining five sheets are generated IN
            // PARALLEL (they never see each other's output), they agree on this fixed vocabulary
            // so {S:sheet.key} references still resolve at write time. Each sheet MUST emit its
            // own canonical rows; missing refs fall back to 0 (and are logged).
            const CANON = `CANONICAL CROSS-SHEET KEYS — the sheets are built in parallel, so to stay linked you MUST use these EXACT "key" values for these standard rows (you may add extra rows with your own keys too). Reference them on OTHER sheets via {S:sheet.key} (same year) or {PS:sheet.key} (previous year):
- op:    net_revenue, ebitda, volume
- pnl:   net_revenue, ebitda, depreciation, ebit, interest, tax, pat, eps
- bs:    equity, net_debt, capital_employed, net_working_capital, total_assets, total_liabilities_equity
- capex: capex, ocf, fcf, gross_block, depreciation, wc_change
Examples: P&L revenue formula "{S:op.net_revenue}"; Capex OCF "{S:pnl.ebitda}-{S:pnl.interest}-{S:pnl.tax}"; Balance-sheet equity "{P:equity}+{S:pnl.pat}*{A:pat_retention}"; Capex change-in-working-capital "{S:bs.net_working_capital}-{PS:bs.net_working_capital}".
SINGLE SOURCE OF TRUTH — pnl.depreciation is the ONLY place the annual depreciation CHARGE is derived; every other sheet needing it (including the Capex & FCF sheet's own "depreciation" row) must reference "{S:pnl.depreciation}" rather than re-deriving or re-linking it, so the figure can never disagree across sheets. Likewise capex.gross_block is the ONLY place gross block is rolled forward; the Balance Sheet references it via "{S:capex.gross_block}" rather than tracking its own separate copy.
EXACT KEYS ONLY — when you write {S:sheet.key} or {PS:sheet.key}, "key" MUST be one of the EXACT canonical keys listed above, or an exact key from the ASSUMPTION KEYS list — copied literally, character-for-character. Do NOT invent, rename, pluralize, or "improve" a key name (e.g. writing "pat_attributable_to_owners" or "net_income" instead of "pat"; "capex_total" instead of "capex") — the other sheets are generated in PARALLEL, blind to each other, so a key you invent yourself has NO row anywhere and silently resolves to 0, corrupting every formula built on it (this has actually happened — an invented "capex.capex_total" and "pnl.pat_attributable_to_owners" both resolved to 0 and broke an equity roll-forward and a gross-block roll-forward). If you think a more precise figure would be more correct (e.g. adjusting PAT for minority interest before an equity roll-forward), you may NOT invent a new cross-sheet key for it — use the canonical key as given, or compute the adjustment entirely from rows/assumptions that exist on YOUR OWN sheet. Do not create your own duplicate of a row that is already canonical on another sheet either (e.g. a second "gross_block" roll-forward on the Balance Sheet) — reference the other sheet's canonical row per SINGLE SOURCE OF TRUTH above.
PURE-FORMULA ROWS — when this prompt gives you the COMPLETE formula for a row (rather than asking you to design it), give ONLY "formula" — do NOT also attach "historical" or "link" to the same row. The base row schema says to always provide "historical", but for a row whose formula is dictated in full, attaching your own guessed historical numbers on top will SILENTLY OVERRIDE the formula for the historical columns (formulas only ever populate forecast columns unless historical/link are both absent) — this has caused a working-capital-change row to show round guessed numbers (-10,000/-15,000/...) instead of the real balance-sheet-derived figure, and an operating-cash-flow row to show plain PAT instead of EBITDA-interest-tax, in both cases silently overriding a correct dictated formula.`;

            // Extra grammar for the Valuation & DCF sheet.
            const VAL_GRAMMAR = `VALUATION-SHEET EXTRA GRAMMAR:
- {N}             -> forecast year number 1..5 (e.g. discount factor "1/(1+{A:wacc})^{N}").
- {SUM:key} or {SUM:sheet.key}   -> SUM of that row's 5 forecast years (e.g. sum of PV of FCFF).
- {LAST:key} or {LAST:sheet.key} -> that row's LAST (5th, terminal) forecast-year value — a FUTURE, projected figure (e.g. the terminal-year working capital level used inside the terminal value formula).
- {CUR:key} or {CUR:sheet.key}   -> that row's CURRENT figure — the LAST HISTORICAL/ACTUAL year, i.e. today, not a forecast year. The EV bridge's net debt MUST use this, NOT {LAST:}: Enterprise Value from a DCF is a value AS OF TODAY (future cash flows discounted back to the present), so subtracting a 5-years-out PROJECTED net debt to reach today's equity value is a mismatch — it must be TODAY'S actual net debt.
- {FWD1:key} or {FWD1:sheet.key} -> that row's FIRST (near-term/NTM) forecast-year value — use this, NOT {LAST:}, for relative-valuation target prices: {LAST:} is 5 years out and undiscounted, so multiplying it by a multiple wildly overstates fair value versus the DCF.
- {SCALAR:key} or {SCALAR:sheet.key} -> a SCALAR row's single value, from a PER-YEAR (non-scalar) row — current_price specifically: it's a single today's-price fact given to you directly (see below), not something with its own value per year, so a per-year ratio (PER, P/BV, EV/EBITDA) that needs it MUST use {SCALAR:current_price}, never {A:current_price} (current_price is NOT on the Assumptions sheet).
- SCALAR row {"scalar":true,...}: a single output value (NOT a per-year series); give a "formula" (using {A:},{R:other_scalar},{SUM:},{LAST:},{CUR:},{FWD1:}) or a plain "value". All scalar values live in one column, so {R:another_scalar_key} works between scalars (e.g. DCF upside referencing current_price: "{R:value_per_share}/{R:current_price}-1").
- Per-year DCF flow rows: set "forecast_only":true (don't compute historicals) AND still give a "formula" — "forecast_only" ONLY means "skip the historical columns"; it does NOT mean "give plain forecast VALUES instead of a formula." Every row in this section (FCFF, discount factor, PV of FCFF, etc.) is a MECHANICAL calculation with exactly one correct formula, never a judgment call — so it must never be a bare "forecast" array of guessed or placeholder numbers. In particular, the discount factor MUST use "1/(1+{A:wacc})^{N}" verbatim as its "formula" — writing it as a flat forecast array (e.g. [1,1,1,1,1]) silently disables ALL discounting and has actually happened, producing a DCF where every "present value" simply equals the raw undiscounted cash flow.`;

            // Looks up a label in Key Financials/Annual Data and returns its historical values
            // aligned to HIST (one entry per historical year, null for any year outside the fetched
            // window), plus the exact label text to use as a trusted "link". Used by
            // buildCapexRows below to resolve capex/gross_block deterministically in code, the same
            // way autoLink already does for LLM-proposed links — just without needing an LLM call to
            // propose a candidate label first.
            const lookupHistoricalByLabel = (pattern) => {
                const entry = findEntryByLabel(pattern);
                if (!entry) return null;
                const historical = HIST.map((fy) => (entry.valByFY[fy] != null ? entry.valByFY[fy] : null));
                if (historical.every((v) => v == null)) return null; // matched a label, but no usable year in our window
                return { link: entry.label, historical };
            };

            // ── Deterministic per-segment revenue/EBITDA lookup from the operational dashboard ──
            // Same spirit as lookupHistoricalByLabel above (resolve real historicals in code rather
            // than trust an LLM transcription), but against operationalSections instead of Key
            // Financials — segment revenue only exists in the dashboard's own segment disclosures,
            // never as a financial-statement line item.
            //
            // The candidate set has to be filtered hard before matching, because a single segment
            // carries many rows whose label contains "revenue" but which are NOT its rupee revenue
            // (verified against real Piramal Pharma data): percentage splits ("CDMO Revenue by
            // Customer Type (FY26) — Generics", unit "%"), foreign-currency figures ("Revenue from
            // ADC services", unit "US$Mn"), and quarterly prints ("CHG Revenue (Quarterly)",
            // periods "Q1 FY26"). Requiring unit=Cr + a non-quarterly label + clean annual FY
            // period labels narrows it to the handful of genuine segment revenue rows.
            const normSeg = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
            // Annual fiscal-year column labels ONLY. Deliberately strict: the same sections also
            // carry "Q1 FY26", "9MFY25", "H2 FY26", "FY24-26", "Mar'24", "31-Mar-25" and
            // "5 Years Ago", none of which are a full-year figure this model can use.
            const parseAnnualFY = (label) => {
                const m = /^fy\s*(\d{2}|\d{4})$/i.exec(String(label || "").trim());
                if (!m) return null;
                const y = parseInt(m[1], 10);
                return y < 100 ? 2000 + y : y;
            };
            // Strips the boilerplate off a metric label so "CDMO Revenue — None" reduces to just
            // "cdmo" — the segment's own name — which is what gets compared against the tag. The
            // trailing "— None"/"— Total" is how this feed marks an unsplit total row.
            const revenueSubject = (metric) => normSeg(metric)
                .replace(/\bquarterly\b/g, " ")
                .replace(/\brevenues?\b/g, " ")
                .replace(/\b(none|total)\b/g, " ")
                .replace(/\s+/g, " ").trim();
            const segmentAnnualSeries = (tag, kind) => {
                if (!operationalSections.length) return null;
                const tagNorm = normSeg(tag);
                if (!tagNorm) return null;
                const tagWords = tagNorm.split(" ").filter(Boolean);
                const kindRe = kind === "ebitda" ? /\bebitda\b/i : /\brevenues?\b/i;
                let best = null;
                for (const section of operationalSections) {
                    const fyByPeriod = new Map();
                    for (const p of section.periods) {
                        const fy = parseAnnualFY(p);
                        if (fy != null) fyByPeriod.set(p, fy);
                    }
                    if (!fyByPeriod.size) continue; // quarterly/derived-period section — not usable here
                    const sectionNorm = normSeg(section.name);
                    for (const row of section.rows) {
                        if (!kindRe.test(row.metric)) continue;
                        if (/\bquarterly\b/i.test(row.metric)) continue;
                        // Rupee-crore only — "%" splits and "US$Mn" figures are a different unit and
                        // would be silently wrong if written into a Rs-Crore model.
                        if (!/^(rs\.?\s*)?cr\.?$/i.test(String(row.unit || "").trim())) continue;
                        const subject = revenueSubject(row.metric);
                        // Tiering matters: a segment's own total ("CHG Revenue — None" -> subject
                        // "chg") must outrank both a sibling brand row in the same section ("Power
                        // Brands Revenue", subject "power brands") and any partial-label collision.
                        let score = 0;
                        if (subject === tagNorm) score = 100;                                   // exact: "chg" === "chg"
                        else if (tagWords.every(w => subject.split(" ").includes(w))) score = 70; // every tag word present as a word
                        else if (tagWords.every(w => sectionNorm.includes(w))) score = 40;        // tag identifies the SECTION, label differs
                        else continue;
                        // Prefer the row covering the most years this model actually needs.
                        const values = HIST.map((fy) => {
                            let v = null;
                            for (const [p, pfy] of fyByPeriod) {
                                if (pfy !== fy) continue;
                                const raw = row.byPeriod[p];
                                const n = parseNum(raw);
                                if (n != null) { v = n; break; }
                            }
                            return v;
                        });
                        const covered = values.filter(v => v != null && v !== 0).length;
                        if (!covered) continue;
                        const rank = score * 1000 + covered;
                        if (!best || rank > best.rank) best = { rank, values, covered, label: row.metric, section: section.name };
                    }
                }
                return best;
            };

            // Capex & FCF, generated entirely in code — first sheet converted this way. Every one of
            // its forecast formulas was ALREADY being dictated to the LLM verbatim (see the prompt
            // text this replaced), with no judgment call involved anywhere in it; only 2 of its 8
            // rows (capex, gross_block) ever needed real historical data, and even those are resolved
            // the exact same deterministic way autoLink already resolves LLM-proposed links. Asking an
            // LLM to transcribe formulas it was already being handed complete added cost, latency, and
            // an entire class of risk (hallucinated key names, mistyped formulas) for zero benefit.
            // Operational/P&L/Balance Sheet/Valuation still go through the LLM for now — their
            // historical data spans far more sector-specific terminology (a bank's "Advances" vs a
            // manufacturer's "Receivables," etc.) than capex/gross block do, which still benefits from
            // an LLM's flexible semantic matching rather than a fixed, hand-maintained label list.
            const buildCapexRows = () => {
                const capexLink = lookupHistoricalByLabel(/capital\s*expenditure|purchase of (fixed|tangible) assets|^capex$/i);
                const gbLink = lookupHistoricalByLabel(/gross\s*block|gross\s*fixed\s*assets/i);
                if (!capexLink) console.warn("[AI Model] Capex & FCF (code-generated): could not find a 'Capital Expenditure' label in Key Financials/Annual Data — capex historicals left blank; the forecast still computes correctly from capex_pct_revenue regardless.");
                if (!gbLink) console.warn("[AI Model] Capex & FCF (code-generated): could not find a 'Gross Block' label in Key Financials/Annual Data — gross_block historicals left blank, so its roll-forward forecast will start from 0 in the first forecast year until this is fixed manually.");
                return [
                    { section: "CAPEX & FREE CASH FLOW" },
                    {
                        key: "capex", label: "Total Capex (Rs Cr)", fmt: "#,##0", cagr: true,
                        link: capexLink ? capexLink.link : null,
                        historical: capexLink ? capexLink.historical : new Array(HIST_YEARS).fill(null),
                        formula: "{A:capex_pct_revenue}*{S:pnl.net_revenue}",
                    },
                    { key: "capex_pct_revenue_actual", label: "Capex / Revenue (%)", fmt: "0.0%", formula: "{R:capex}/{S:pnl.net_revenue}" },
                    { key: "ocf", label: "Operating Cash Flow (Rs Cr)", fmt: "#,##0", formula: "{S:pnl.ebitda}-{S:pnl.interest}-{S:pnl.tax}" },
                    { key: "wc_change", label: "Change in Net Working Capital (Rs Cr)", fmt: "#,##0", formula: "{S:bs.net_working_capital}-{PS:bs.net_working_capital}" },
                    { key: "fcf", label: "Free Cash Flow (levered) (Rs Cr)", fmt: "#,##0", cagr: true, formula: "{R:ocf}-{R:capex}" },
                    {
                        key: "gross_block", label: "Gross Block (Rs Cr)", fmt: "#,##0", cagr: true,
                        link: gbLink ? gbLink.link : null,
                        historical: gbLink ? gbLink.historical : new Array(HIST_YEARS).fill(null),
                        formula: "{P:gross_block}+{R:capex}",
                    },
                    { key: "depreciation", label: "Depreciation (Rs Cr)", fmt: "#,##0", formula: "{S:pnl.depreciation}" },
                ];
            };

            // SOTP variant of buildCapexRows above — ONLY called when isSotpCompany AND the
            // Operational Model call actually delivered every expected segment revenue row (see
            // opDeliveredSegments at the call site below). buildCapexRows itself is completely
            // untouched, still the code path every non-SOTP company (and any SOTP company whose
            // Operational output didn't come back as expected) runs through unchanged.
            // Allocates capex PER SEGMENT off that segment's own capex_pct_revenue_<tag> against
            // its own op.<tag>_revenue — replacing the single capex_pct_revenue applied to total
            // consolidated revenue, which smeared a capex-heavy growth segment's intensity evenly
            // across a mature segment that doesn't need anywhere near that much reinvestment. Total
            // capex is DICTATED as the sum of the segment rows (never computed independently), so it
            // can never silently diverge from the sum of its own parts — same principle as the
            // Operational Model's net_revenue/ebitda totals above.
            const buildSegmentCapexRows = () => {
                const capexLink = lookupHistoricalByLabel(/capital\s*expenditure|purchase of (fixed|tangible) assets|^capex$/i);
                const gbLink = lookupHistoricalByLabel(/gross\s*block|gross\s*fixed\s*assets/i);
                if (!capexLink) console.warn("[AI Model] Capex & FCF (code-generated, SOTP): could not find a 'Capital Expenditure' label in Key Financials/Annual Data — capex historicals left blank; the forecast still computes correctly from the segment capex_pct_revenue_<tag> rows regardless.");
                if (!gbLink) console.warn("[AI Model] Capex & FCF (code-generated, SOTP): could not find a 'Gross Block' label in Key Financials/Annual Data — gross_block historicals left blank, so its roll-forward forecast will start from 0 in the first forecast year until this is fixed manually.");

                const capexSegKey = (tag) => `capex_${tag}`;
                const allTags = [...segmentTags, "other_segments"];
                const segmentCapexRows = allTags.map(tag => ({
                    key: capexSegKey(tag),
                    label: `Capex — ${tag === "other_segments" ? "Other Segments" : tag} (Rs Cr)`,
                    fmt: "#,##0",
                    // "other_segments" has no dedicated capex_pct_revenue_<tag> key (the Assumptions
                    // call only tags genuinely distinct segments, not the catch-all) — it uses the
                    // plain blended capex_pct_revenue key instead, same convention as its revenue/
                    // EBITDA rows on the Operational Model sheet.
                    formula: tag === "other_segments"
                        ? `{A:capex_pct_revenue}*{S:op.other_segments_revenue}`
                        : `{A:capex_pct_revenue_${tag}}*{S:op.${tag}_revenue}`,
                }));
                const capexSumFormula = allTags.map(t => `{R:${capexSegKey(t)}}`).join("+");

                return [
                    { section: "CAPEX & FREE CASH FLOW" },
                    ...segmentCapexRows,
                    {
                        key: "capex", label: "Total Capex (Rs Cr)", fmt: "#,##0", cagr: true,
                        link: capexLink ? capexLink.link : null,
                        historical: capexLink ? capexLink.historical : new Array(HIST_YEARS).fill(null),
                        formula: capexSumFormula,
                    },
                    { key: "capex_pct_revenue_actual", label: "Capex / Revenue (%)", fmt: "0.0%", formula: "{R:capex}/{S:pnl.net_revenue}" },
                    { key: "ocf", label: "Operating Cash Flow (Rs Cr)", fmt: "#,##0", formula: "{S:pnl.ebitda}-{S:pnl.interest}-{S:pnl.tax}" },
                    { key: "wc_change", label: "Change in Net Working Capital (Rs Cr)", fmt: "#,##0", formula: "{S:bs.net_working_capital}-{PS:bs.net_working_capital}" },
                    { key: "fcf", label: "Free Cash Flow (levered) (Rs Cr)", fmt: "#,##0", cagr: true, formula: "{R:ocf}-{R:capex}" },
                    {
                        key: "gross_block", label: "Gross Block (Rs Cr)", fmt: "#,##0", cagr: true,
                        link: gbLink ? gbLink.link : null,
                        historical: gbLink ? gbLink.historical : new Array(HIST_YEARS).fill(null),
                        formula: "{P:gross_block}+{R:capex}",
                    },
                    { key: "depreciation", label: "Depreciation (Rs Cr)", fmt: "#,##0", formula: "{S:pnl.depreciation}" },
                ];
            };

            // SOTP-aware Operational Model prompt — ONLY for a genuine multi-segment conglomerate
            // (isSotpCompany, see the segment-detection block above). Each detected segment gets its
            // OWN revenue_growth_<tag>/ebitda_margin_<tag> keys (already resolved by the Assumptions
            // (Judgment) call) instead of every segment sharing one blended rate, and net_revenue/
            // ebitda are DICTATED as the exact sum of the segment rows (never left for the model to
            // compute independently) so the total can never silently diverge from the sum of its own
            // parts. Non-SOTP companies (the common case) get operationalModelPrompt assigned the
            // EXACT SAME string this call has always sent — see the `:` branch below, copied
            // verbatim — so nothing about their model changes.
            const segRevKey = (tag) => `${tag}_revenue`;
            const segEbitdaKey = (tag) => `${tag}_ebitda`;
            const sotpRevenueSumFormula = [...segmentTags.map(segRevKey), "other_segments_revenue"].map(k => `{R:${k}}`).join("+");
            const sotpEbitdaSumFormula = [...segmentTags.map(segEbitdaKey), "other_segments_ebitda"].map(k => `{R:${k}}`).join("+");
            const sotpMandatoryKeys = ["net_revenue", "ebitda", ...segmentTags.flatMap(t => [segRevKey(t), segEbitdaKey(t)]), "other_segments_revenue", "other_segments_ebitda"].join(", ");
            const operationalModelPrompt = isSotpCompany
                ? `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild the OPERATIONAL MODEL for ${companyName} as a SUM-OF-THE-PARTS segment breakdown — this company was determined to be a genuine multi-segment conglomerate (holding_discount > 0), so each segment below is forecast off its OWN dedicated growth/margin assumption instead of one company-wide blended rate.\nSEGMENTS AND THEIR DEDICATED KEYS — create ONE segment revenue/EBITDA row pair for EACH of these tags, using EXACTLY these row keys and formulas (do not rename, do not invent your own segment keys):\n${segmentTags.map(t => `- Segment "${t}": revenue row key="${segRevKey(t)}" (HISTORICAL: real reported figures for this segment from the OPERATIONAL DATA/segment disclosures above; FORECAST formula EXACTLY "{P:${segRevKey(t)}}*(1+{A:revenue_growth_${t}})"). EBITDA row key="${segEbitdaKey(t)}" (FORECAST formula EXACTLY "{R:${segRevKey(t)}}*{A:ebitda_margin_${t}}").`).join("\n")}\nOTHER SEGMENTS CATCH-ALL — the tagged segments above may not cover 100% of consolidated revenue (a residual/unallocated/corporate segment, or a minor business the Assumptions call didn't tag). You MUST ALSO include exactly ONE catch-all pair: revenue row key="other_segments_revenue" (HISTORICAL = total consolidated revenue for that year MINUS the sum of the tagged segments' revenue that same year — a small or near-zero number is fine if the tagged segments cover nearly everything; FORECAST formula EXACTLY "{P:other_segments_revenue}*(1+{A:revenue_growth})" — this ONE catch-all row uses the BLENDED company-wide keys, unlike every tagged segment above), EBITDA row key="other_segments_ebitda" (FORECAST formula EXACTLY "{R:other_segments_revenue}*{A:ebitda_margin}").\nTOTAL ROWS — the canonical TOTAL rows MUST use these EXACT dictated formulas, which SUM the segment rows above (do NOT invent your own, do NOT grow the total independently off the blended revenue_growth/ebitda_margin — that would let the total silently diverge from the sum of its own parts): net_revenue forecast = "${sotpRevenueSumFormula}", ebitda forecast = "${sotpEbitdaSumFormula}".\nOTHER DISCLOSURE ROWS — capacity, utilisation, volumes/throughput, realisations/ASP per segment, and cost-per-unit remain informational KPIs that do NOT feed revenue or EBITDA, same as always: each one's forecast MUST be a FLAT carry-forward of its last actual, written EXACTLY as "{P:<this row's own key>}". Do NOT attach a growth rate, and do NOT reference an invented assumption key — the ONLY {A:...} keys you may use are those in the ASSUMPTION KEYS list above. If you have no real driver for a KPI, a flat carry-forward IS the correct, honest forecast.\nNO BLANK FORECASTS — every assumption you reference via {A:...} must itself have a value in ALL ${HIST_YEARS + FC_YEARS} year columns, never left blank — a blank/missing forecast cell is read as ZERO by Excel and silently zeroes out everything downstream that multiplies by it.\nYou MUST include rows keyed exactly: ${sotpMandatoryKeys}.`
                : `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild the OPERATIONAL MODEL for ${companyName}, including a per-segment/product-line breakdown where the company reports one (e.g. by business segment). Capacity, utilisation, volumes/throughput, realisations/ASP per segment, and cost-per-unit are DISCLOSURE rows — informational KPIs that do NOT feed revenue or EBITDA. Each disclosure row's forecast MUST be a FLAT carry-forward of its last actual, written EXACTLY as "{P:<this row's own key>}". Do NOT attach a growth rate, and above all do NOT reference an invented assumption key such as "{A:jio_arpu_growth}" or "{A:o2c_production_growth}" — those keys do not exist (the ONLY {A:...} keys you may use are those in the ASSUMPTION KEYS list), so they render as a #MISSING error. If you have no real driver for a KPI, a flat carry-forward IS the correct, honest forecast.\nFORECASTING SEGMENT/TOTAL REVENUE AND EBITDA — do NOT compute forecast revenue by multiplying a forecast "volume" row by a forecast "realisation/ASP" row: those two drivers' units have repeatedly ended up mismatched (observed real failures: a 10x, a 100x, and a 1000x understatement on three different segments of the same company in one run, because volume and realisation were each forecast independently and their product no longer matched revenue's own actual unit scale). Instead, forecast each segment's (and the total's) revenue as a GROWTH-RATE recurrence anchored to its OWN historical actual level, which is unit-safe by construction since it only ever multiplies a real linked/historical figure by a dimensionless growth rate — e.g. segment revenue: "{P:segment_key}*(1+{A:revenue_growth})". Forecast segment EBITDA the same way, off a margin assumption applied to that segment's OWN revenue row: "{R:segment_key_revenue}*{A:ebitda_margin}". The canonical TOTAL rows MUST use these EXACT dictated formulas (do NOT invent your own): net_revenue forecast = "{P:net_revenue}*(1+{A:revenue_growth})", and ebitda forecast = "{R:net_revenue}*{A:ebitda_margin}" (revenue times the margin assumption — this equals the sum of the segments, since the SAME company-wide margin is applied to every segment). NEVER write a growth-style recurrence for the total EBITDA row such as "{P:ebitda}*(1+{A:ebitda_growth})" — there is NO "ebitda_growth" assumption key, so it resolves to a #MISSING error that poisons pnl.ebitda, the DCF, and every sheet pulling op.ebitda.\nUSE THE EXACT ASSUMPTION KEYS "revenue_growth" AND "ebitda_margin" FROM THE ASSUMPTION KEYS LIST ABOVE for every segment revenue/EBITDA row. The ONLY growth/margin keys that exist are revenue_growth and ebitda_margin; do NOT invent "segment_x_growth", "ebitda_growth", or any other {A:...} key — an invented key exists nowhere and resolves to a #MISSING error. Applying the SAME company-wide growth/margin to every segment is the correct, deliberate simplification here.\nNO BLANK FORECASTS — every assumption you reference via {A:...} (growth rates, EBITDA margins, etc.) must itself have a value in ALL ${HIST_YEARS + FC_YEARS} year columns, never left blank — a blank/missing forecast cell is read as ZERO by Excel and silently zeroes out everything downstream that multiplies by it (e.g. an EBITDA margin left blank makes segment EBITDA compute to 0 even though revenue is fine).\nYou MUST include rows keyed exactly: net_revenue, ebitda (and volume if applicable, informational only — not used to derive net_revenue).`;

            // SOTP cross-check section for the Valuation & DCF prompt — ADDITIVE only: every row and
            // formula in the existing DCF/target-price/WACC sections stays completely untouched for
            // EVERY company (including SOTP ones — the blended DCF's value_per_share is left exactly
            // as it's always been computed, so nothing existing changes meaning). For a genuine
            // conglomerate this ONLY splices in one new section (3b) pricing each segment separately
            // by its own target_ev_ebitda_<tag> multiple against its own forward EBITDA, summed to a
            // sotp_value_per_share the user can compare against the blended DCF figure — this is the
            // properly-priced sum-of-the-parts number Adani-style companies were missing. Nothing
            // downstream (Summary Dashboard, sensitivity table) reads these new sotp_* keys, so if
            // the model doesn't fully comply, the worst case is just a missing cross-check row — no
            // cascading #MISSING risk into the DCF chain, unlike the Operational/Capex changes above.
            const sotpSegEvKey = (tag) => `sotp_ev_${tag}`;
            const sotpEvSumFormula = [...segmentTags.map(sotpSegEvKey), sotpSegEvKey("other_segments")].map(k => `{R:${k}}`).join("+");
            const sotpValuationSection = isSotpCompany
                ? `\n(3b) "SUM-OF-THE-PARTS VALUATION (segment cross-check)" — this company has genuinely distinct segments (see SEGMENT-SPECIFIC DRIVERS on Assumptions). The blended DCF above already reflects segment-level growth/margin/capex under the hood (via the Operational Model and Capex sheets), but THIS section prices the segments the way an equity analyst actually would for a conglomerate — separately, by segment-appropriate multiple — as a cross-check against the DCF's value_per_share, NOT a replacement for it. SCALAR rows:\n${segmentTags.map(t => `Segment EV — ${t} key=${sotpSegEvKey(t)} ("{FWD1:op.${segEbitdaKey(t)}}*{A:target_ev_ebitda_${t}}", fmt "#,##0" — this segment's own NEXT-YEAR EBITDA times ITS OWN target multiple, never the blended target_ev_ebitda).`).join("\n")}\nSegment EV — Other Segments key=${sotpSegEvKey("other_segments")} ("{FWD1:op.other_segments_ebitda}*{A:target_ev_ebitda}", fmt "#,##0" — this ONE catch-all uses the blended target_ev_ebitda, matching its Operational/Capex rows).\nSum-of-parts enterprise value key=sotp_ev ("${sotpEvSumFormula}", fmt "#,##0" — the SUM of every segment EV row above; do NOT compute this any other way).\nSum-of-parts equity value key=sotp_equity_value ("{R:sotp_ev}+{R:non_operating_assets}-{R:net_debt_last}" — reuse the SAME non_operating_assets/net_debt_last rows from section (3) above, do not recompute them).\nSum-of-parts value per share key=sotp_value_per_share ("{R:sotp_equity_value}/{CUR:assum.shares_out}", emphasis:"highlight").\nSum-of-parts upside key=sotp_upside ("{R:sotp_value_per_share}/{R:current_price}-1", fmt 0.0%).`
                : "";
            const valuationPrompt = `${periodsNote}\n\n${GRAMMAR}\n\n${VAL_GRAMMAR}\n\n${CANON}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild a detailed VALUATION & DCF sheet for ${companyName}. Return JSON {"rows":[...]} with these sections (use {"section":"NAME"} dividers):\n(1) "RELATIVE VALUATION" — per-year live ratios: EPS key=eps ("{S:pnl.eps}"), BVPS key=bvps ("{S:bs.equity}/{A:shares_out}"), PER ("{SCALAR:current_price}/{R:eps}", fmt 0.0), P/BV ("{SCALAR:current_price}/{R:bvps}", 0.0), EV/EBITDA ("(({SCALAR:current_price}*{A:shares_out})+{S:bs.net_debt})/{S:pnl.ebitda}", 0.0), RoE ("{S:pnl.pat}/{S:bs.equity}", 0.0%), RoCE ("{S:pnl.ebit}/{S:bs.capital_employed}", 0.0%), Net debt/EBITDA ("{S:bs.net_debt}/{S:pnl.ebitda}", 0.00). current_price MUST be referenced via {SCALAR:current_price} here (a per-year row reading a single fixed fact), NEVER {A:current_price} — current_price is a scalar row on THIS sheet (see section 3), not a per-year row on Assumptions.\n(2) "DCF — FREE CASH FLOW TO FIRM" (every row "forecast_only":true): EBIT key=ebit ("{S:pnl.ebit}"), NOPAT key=nopat ("{R:ebit}*(1-{A:tax_rate})"), Add depreciation key=dep ("{S:pnl.depreciation}"), Less change in working capital key=wc_change ("{S:capex.wc_change}" — pull the SAME figure the Capex & FCF sheet computes; do NOT leave this without a formula or default it to zero), FCFF key=fcff ("{R:nopat}+{R:dep}-{R:wc_change}-{S:capex.capex}"), Discount factor key=discount_factor ("1/(1+{A:wacc})^{N}", fmt 0.000), PV of FCFF key=pv_fcff ("{R:fcff}*{R:discount_factor}").\n(3) "DCF VALUATION" — SCALAR rows: Sum of PV key=sum_pv ("{SUM:pv_fcff}"), Terminal value key=tv — use this EXACT dictated formula, copied verbatim (do not redesign it): "MAX(0,({LAST:nopat}-{A:terminal_growth}*{LAST:bs.net_working_capital})*(1+{A:terminal_growth})/({A:wacc}-{A:terminal_growth}))". Do NOT build this off {LAST:fcff} (the raw final explicit-year FCFF) — for a still-investing, capex-heavy company the LAST explicit forecast year's FCFF can itself be depressed or even negative (heavy capex, a working-capital step-up, etc.), and extrapolating THAT into a perpetuity produces a nonsensical negative terminal value (this has actually happened — a real run's terminal FCFF was negative, making equity value collapse to a large negative per-share figure). The dictated formula instead uses a NORMALIZED steady-state terminal cash flow: terminal NOPAT (which excludes capex/working-capital entirely, so it stays sensible even in a buildout year) less a working-capital investment that grows at the terminal growth rate (g × the terminal year's own net working capital LEVEL, not the raw year-over-year change row, which is more robust to any residual basis discontinuity in an early forecast year) — implicitly assuming steady-state maintenance capex equals depreciation, the standard textbook simplification — wrapped in MAX(0,...) so a structurally unprofitable terminal NOPAT can never produce a negative terminal value. PV of terminal value key=pv_tv ("{R:tv}/(1+{A:wacc})^5"), Enterprise value key=ev ("{R:sum_pv}+{R:pv_tv}"), Less net debt key=net_debt_last — use this EXACT dictated formula: "{CUR:bs.net_debt}" (the LAST HISTORICAL/ACTUAL year's net debt — today's real, reported figure — NOT {LAST:bs.net_debt}, which is 5 years OUT, a projection. Enterprise Value from a DCF is a value AS OF TODAY (all future cash flows discounted back to the present), so the EV-to-equity bridge must subtract TODAY'S actual net debt, not some future projected net debt — mixing a present-day EV with a future net debt double-counts/mismatches the time basis and distorts equity value), Add back non-operating assets key=non_operating_assets — use this EXACT dictated formula: "{CUR:bs.cwip}+{CUR:bs.investments}" (fmt "#,##0"). This row is REQUIRED and must never be omitted or set to zero. Rationale: FCFF above charges 100% of capex as a cash outflow, but NOPAT is built off EBIT, which earns NOTHING on either of these two asset classes — CWIP is still under construction (not commissioned, so zero revenue/EBIT during the whole forecast even though its cost has already been fully deducted from FCFF), and long-term investments in associates/JVs return via OTHER INCOME, which sits BELOW EBIT and therefore never enters NOPAT or FCFF at all. Both are textbook non-operating assets: the correct EV-to-equity bridge is "EV + non-operating assets − net debt". Omitting them makes the model pay for these assets in full and then value them at exactly zero, which systematically understates value per share for capital-intensive companies (a real run left ~₹28,800 Cr of CWIP — 11% of total assets — valued at nil). Use {CUR:} (the last HISTORICAL/actual year) for the same time-basis reason net_debt_last does: a DCF enterprise value is a value AS OF TODAY, so every other item in the bridge must be today's real reported balance, never a projection), Equity value key=equity_value ("{R:ev}+{R:non_operating_assets}-{R:net_debt_last}"), Holding company discount key=holding_discount_pct ("{A:holding_discount}", fmt 0.0% — a DISPLAY row showing the discount being applied below, so it's visible rather than buried inside the next row's formula), Equity value (post holding discount) key=equity_value_discounted ("{R:equity_value}*(1-{A:holding_discount})" — for a single-business company holding_discount is 0, so this equals equity_value unchanged), Value per share key=value_per_share ("{R:equity_value_discounted}/{CUR:assum.shares_out}", emphasis:"highlight" — MUST use {CUR:assum.shares_out} (today's/latest actual share count), NEVER {A:shares_out}: this is a SCALAR row (one fixed column), and shares_out is now a genuine PER-YEAR row (share count changes over time), so {A:shares_out} here would silently pick up whichever column this scalar row happens to sit in — likely the OLDEST historical year — instead of the current count — MUST reference equity_value_discounted, NOT the pre-discount equity_value, so the conglomerate discount actually reaches the headline number), Current price key=current_price (${currentPriceValueNote}a single fact — today's actual trading price, NOT a per-year series; every other row in THIS sheet that needs it references it via {SCALAR:current_price} or, from another scalar row, {R:current_price} — current_price itself must NEVER use {A:...} since it does not live on the Assumptions sheet), DCF upside key=upside ("{R:value_per_share}/{R:current_price}-1", fmt 0.0%).${sotpValuationSection}\n(4) "TARGET PRICE (relative, NEAR-TERM FY+1 basis)" — SCALAR rows, built off the FIRST forecast year via {FWD1:} (NOT {LAST:}, which is 5 years out and undiscounted — using it overstates the target price versus the DCF): P/E-based key=target_price_pe ("{FWD1:pnl.eps}*{A:target_pe}"), EV/EBITDA-based key=target_price_ev_ebitda ("(({FWD1:pnl.ebitda}*{A:target_ev_ebitda})-{FWD1:bs.net_debt})/{CUR:assum.shares_out}" — {CUR:}, not {A:}: this is a SCALAR row, and shares_out is a PER-YEAR row now, so {A:shares_out} would silently resolve to whatever column this scalar row sits in rather than the current share count), Blended target price key=target_price (average of the two, emphasis:"highlight"), Upside to target ("{R:target_price}/{R:current_price}-1", fmt 0.0%). Use these EXACT keys (target_price_pe, target_price_ev_ebitda) — not a paraphrase — every other row in this schema dictates a "key" for exactly this reason: a row with no fixed key gives you no reliable way to reference or verify it afterward.\n(5) "WACC BUILD-UP" — SCALAR rows: risk free ("{A:risk_free_rate}", 0.0%), equity risk premium ("{A:equity_risk_premium}", 0.0%), beta ("{A:beta}", 0.00), cost of equity ("{A:risk_free_rate}+{A:beta}*{A:equity_risk_premium}", 0.0%), cost of debt ("{A:cost_of_debt}", 0.0%), tax rate ("{A:tax_rate}", 0.0%), WACC ("{A:wacc}", 0.0%).\nYou MUST emit scalar keys: current_price, value_per_share, target_price, target_price_pe, target_price_ev_ebitda, upside (plus helpers sum_pv, pv_tv, tv, ev, net_debt_last, non_operating_assets, equity_value, holding_discount_pct, equity_value_discounted)${isSotpCompany ? ", plus (this is a SOTP company) sotp_ev, sotp_equity_value, sotp_value_per_share, sotp_upside" : ""}. Use fmt "#,##0" for currency rows. Be thorough.`;

            // 3b–3g — generate the remaining sheets IN PARALLEL.
            aiStatus("AI (step 2/2): building all model sheets in parallel...");
            const [opJson, pnlJson, bsJson, valJson, sumJson] = await Promise.all([
                callLLM("Operational Model", operationalModelPrompt, STATIC_MODEL),
                callLLM("P&L Model",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild a DETAILED P&L MODEL for ${companyName}.\nCANONICAL EBITDA — key=ebitda, formula = "{S:op.ebitda}" (pulled DIRECTLY from Operational; this is the single source of truth, per CANON above). Do NOT compute EBITDA as net_revenue minus the expense lines below — that makes EBITDA depend on four fragile disclosure rows instead of the one already-correct canonical figure, and if any one of those breaks, EBITDA silently breaks with it.\nDISCLOSURE expense build (informational detail only, does NOT feed EBITDA): cost of materials/COGS, power & fuel, employee cost, other expenses (for EACH of these four: FORECAST formula = "{P:<this row's own key>}/{P:net_revenue}*{R:net_revenue}" — holds that expense's ratio-to-revenue constant at its last historical/prior-year level. Use {P:net_revenue}, NOT {PS:pnl.net_revenue} — net_revenue is a row on THIS SAME sheet, so it's addressed the same-sheet way ({P:}/{R:}), never with a sheet-qualified {PS:sheet.key}/{S:sheet.key} reference to your own sheet. Do NOT invent a new assumption key like "cogs_pct_revenue" for these — none exists in the ASSUMPTION KEYS list above, and referencing one that doesn't exist resolves to a visible #MISSING error, not a usable number. This ratio-preserving formula needs no assumption key at all), total expenditure (key=total_expenditure, formula = "{R:net_revenue}-{R:ebitda}" — a DERIVED memo line computed from the two authoritative figures, NOT a sum of the four disclosure lines above, so it can never disagree with EBITDA even if a disclosure line is off), EBITDA margin (0.0%), depreciation (key=depreciation — the ANNUAL depreciation CHARGE for the year; if the data has both a P&L-style annual figure and a Balance-Sheet-style "Accumulated/Cumulative Depreciation" balance, link the ANNUAL one — never the cumulative one. PLAUSIBILITY CHECK: annual depreciation is typically 3-10% of average gross block per year, and should NOT be a monotonically-increasing-by-similar-amounts series across historical years (that pattern is a signature of an accumulated/cumulative balance, not an annual charge) — if your linked figure fails this check, you linked the WRONG (cumulative) row; either find the correct annual-charge label, or — if only the cumulative balance is available in the data — compute the annual figure yourself as this year's cumulative balance MINUS last year's (do this for every historical year using the raw source numbers, then put the resulting 5 numbers directly in "historical", not a link). This row's FORECAST formula is REQUIRED and must NOT be left blank — a blank forecast is read as 0 by Excel, which silently makes EBIT equal EBITDA and inflates every line below it. Tie it to the gross block build using the EXACT key "depreciation_pct_gross_block" from the ASSUMPTION KEYS list: "{A:depreciation_pct_gross_block}*({S:capex.gross_block}+{PS:capex.gross_block})/2"), EBIT, EBIT margin, other income, finance cost (key=interest — charge it on OPENING (prior-year) net debt ONLY, e.g. "MAX(0,{A:cost_of_debt}*{PS:bs.net_debt})" — do NOT average current-year and prior-year net debt: current-year net debt depends on the cash plug, which depends on retained earnings, which depends on PAT, which depends on this interest figure, so including current-year net debt here creates a genuine circular reference that silently evaluates to 0 across the whole PBT/tax/PAT/EPS chain once it closes the loop. Opening-balance-only breaks the loop by construction, since the prior year's net debt is already fully resolved before this year is computed. Also floor it at zero so it can never go negative even if net debt turns negative/net-cash in later forecast years), PBT, exceptional items, tax, effective tax rate (0.0%), PAT, PAT margin, minority interest where relevant (key it minority_interest and do NOT leave it a flat carry-forward — give it a forecast formula that grows it with earnings, holding it at a constant share of PAT: "{P:minority_interest}/{P:pat}*{R:pat}"), EPS (key=eps — do NOT leave it a flat carry-forward either: give it the formula "{R:pat}/{A:shares_out}" for EVERY column, historical and forecast alike, so it always ties out to whatever PAT and shares_out actually resolve to instead of just repeating last year's number — a row with real historicals but no forecast formula gets flat-carried forward automatically, which is a fallback for genuine gaps, not a substitute for giving this row its own well-known formula), DPS (link/enter the ACTUAL reported dividend per share for EACH historical year — do NOT leave historical DPS at 0 when the company actually paid a dividend; forecast DPS from the dividend-payout assumption so it stays continuous with the reported history). Pull revenue & EBITDA from operational via {S:op.net_revenue} / {S:op.ebitda}. You MUST include rows keyed exactly: net_revenue, ebitda, depreciation, ebit, interest, tax, pat, eps — and NONE of these forecast formulas may be left blank/omitted.`,
                    STATIC_MODEL),
                callLLM("Balance Sheet",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild a DETAILED BALANCE SHEET & RETURNS sheet for ${companyName}.\nASSETS (use these EXACT keys so the cash plug below can reference them): receivables (key=receivables — HISTORICAL: link/enter from the data; FORECAST formula = "{A:debtor_days}/365*{S:pnl.net_revenue}" — do NOT leave this as historical-only, it must scale with revenue), inventory (key=inventory — HISTORICAL: link/enter from the data; FORECAST formula = "{A:inventory_days}/365*{S:pnl.net_revenue}"), other_current_assets (HISTORICAL: link/enter; FORECAST formula = "{P:other_current_assets}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}" — scales flat with revenue since there's no dedicated days assumption for it), net_fixed_assets (key=net_fixed_assets — HISTORICAL: link/enter the REAL reported Net Fixed Assets figure from the data, exactly like every other asset row above — do NOT leave this historical-derived-from-formula; FORECAST formula = "{S:capex.gross_block}-{R:acc_depreciation}" — do NOT track your own separate gross block, capex.gross_block is the single source of truth, and this formula applies ONLY to the forecast columns, never the historical ones), acc_depreciation (key=acc_depreciation — a STOCK/balance, NOT the same thing as pnl.depreciation which is the annual CHARGE. HISTORICAL: do NOT try to independently look up or link a separate "accumulated depreciation" line from the data — unlike gross block and net block, it is often not cleanly reported as one unambiguous figure, and guessing at it has actually produced a large understatement in a real run (accumulated depreciation understated by over ₹220,000 Cr for a company with a ~₹18 lakh Cr gross block), which overstated the very next forecast year's Net Fixed Assets by the same amount — a phantom multi-lakh-crore asset increase that then forces the short_term_borrowings revolver into an enormous, implausible draw to fund it, even though the revolver formula itself, the capex driver, and the working-capital formulas were all completely correct. Instead, DERIVE each historical year's value directly as (that year's capex.gross_block − that year's net_fixed_assets, both already reliably known from the two rows above) and put these 5 COMPUTED numbers directly in "historical" — do not attempt a "link" for this row. FORECAST formula = "{P:acc_depreciation}+{S:pnl.depreciation}" (this only rolls forward correctly because the historical base is properly derived, not guessed). SELF-CHECK before finalizing: gross_block(last historical year) − acc_depreciation(last historical year) MUST equal net_fixed_assets(last historical year) EXACTLY — if it does not, you have not derived acc_depreciation correctly and every forecast year's Net Fixed Assets will be wrong), cwip, investments (key=investments — this row means LONG-TERM/NON-CURRENT investments specifically. Many companies report investments TWICE — a "Current Investments" line (a child of Current/Other Current Assets) and a SEPARATE "Long Term Investments"/"Non-current Investments" line elsewhere — link THIS row to the non-current one; linking the current-investments child instead has actually happened and, combined with also linking other_current_assets to ITS parent total, silently double-counts the same rupees into Total Assets twice over), other_non_current_assets (do NOT freeze these flat across the forecast — give EACH a revenue-scaling forecast formula that holds its ratio-to-revenue constant at the last actual level: "{P:<this row's own key>}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", so the balance sheet grows coherently instead of showing an identical number in every forecast year), and cash (key=cash — historicals from the data as usual, but see BALANCING below for the forecast).\nTotal assets: key=total_assets, formula/sum = "{R:cash}+{R:receivables}+{R:inventory}+{R:other_current_assets}+{R:net_fixed_assets}+{R:cwip}+{R:investments}+{R:other_non_current_assets}".\nLIABILITIES & EQUITY: share_capital (key=share_capital, flat/no formula is acceptable — no fresh issuance is being modeled), reserves (roll forward "{P:equity}+{S:pnl.pat}*{A:pat_retention}"), shareholders' equity (key=equity), long_term_borrowings (key=long_term_borrowings, flat/no formula is acceptable — no repayment/drawdown schedule is being modeled), short_term_borrowings (key=short_term_borrowings — THIS IS THE REVOLVER/FUNDING PLUG, not a flat carry: HISTORICAL — link/enter the real reported figure as usual. FORECAST — use this EXACT dictated formula, copied verbatim (do not shorten or redesign it): "{P:short_term_borrowings}+MAX(0,{A:min_cash_pct_revenue}*{S:pnl.net_revenue}-({R:share_capital}+{R:reserves}+{R:long_term_borrowings}+{P:short_term_borrowings}+{R:payables}+{R:provisions}+{R:deferred_tax_liabilities}+{R:other_non_current_liabilities}+{R:other_current_liabilities}-{R:receivables}-{R:inventory}-{R:other_current_assets}-{R:net_fixed_assets}-{R:cwip}-{R:investments}-{R:other_non_current_assets}))" — this holds short_term_borrowings at last year's level UNLESS funding the rest of the balance sheet at last year's borrowing level would push cash below the minimum buffer, in which case it draws exactly enough extra short-term debt to keep cash at that minimum. This is what prevents the cash plug below from ever going negative — do NOT simplify this formula away or replace it with a flat carry-forward, and do NOT change the cash formula to "fix" a negative value directly; the fix belongs on THIS row, not on cash. If this draw comes out implausibly large (e.g. short-term borrowings jumping to several times their historical level in a single year), that is almost always a symptom of an unrealistic upstream working-capital or capex assumption — most commonly a debtor/inventory/payable-days driver whose FIRST forecast value was not anchored to the company's own last reported balance sheet (see the working-capital-days anchoring requirement in the Assumptions prompt), which manufactures a large artificial one-off funding gap. Do not "fix" a large draw by capping or overriding this formula — the correct fix is upstream, in the days/capex assumptions), payables (key=payables — HISTORICAL: link/enter from the data; FORECAST formula = "{A:payable_days}/365*{S:pnl.net_revenue}" — do NOT leave this as historical-only, it must scale with revenue), provisions (HISTORICAL: link/enter; FORECAST formula = "{P:provisions}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}"), deferred_tax_liabilities, other_non_current_liabilities, other_current_liabilities (do NOT freeze these flat — give EACH the same revenue-scaling forecast formula "{P:<this row's own key>}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}" so they scale with the business rather than repeating an identical number every forecast year).\nNONE of receivables/inventory/other_current_assets/payables/provisions may be left without a forecast formula — a flat/frozen working-capital line while revenue grows is a modeling error, not a simplification.\nTotal liabilities & equity: key=total_liabilities_equity, formula/sum of every liabilities & equity row above (via {R:} references, mirroring total_assets).\nPlus: net working capital (key=net_working_capital = "{R:receivables}+{R:inventory}+{R:other_current_assets}-{R:payables}-{R:provisions}-{R:other_current_liabilities}" — include other_current_liabilities here, not just payables/provisions: it's the liability-side counterpart to other_current_assets, which IS already included, so omitting it understates the liabilities subtracted and overstates Net Working Capital; the Capex & FCF sheet's change-in-WC pulls this), net debt (key=net_debt, formula = "{R:long_term_borrowings}+{R:short_term_borrowings}-{R:cash}" — do NOT leave this without a formula, it must move with the computed cash plug, not stay flat), capital employed (key=capital_employed), working-capital days, and returns: ROE (key=roe, formula = "{S:pnl.pat}/{R:equity}", fmt "0.0%" — this MUST be a "formula", never a plain "forecast" array of guessed numbers; a bare zero here has actually happened when the formula was omitted), ROCE (key=roce, formula = "{S:pnl.ebit}/{R:capital_employed}", fmt "0.0%" — same requirement, must be a formula), net debt/EBITDA (key=net_debt_ebitda, formula = "{R:net_debt}/{S:pnl.ebitda}", fmt "0.00").\nHISTORICAL BALANCE (the 5 ACTUAL-year columns) — a REPORTED balance sheet ALWAYS balances, so your historical columns MUST balance too, not only the forecast columns: for EVERY historical year the summed asset rows must equal the summed liabilities-&-equity rows (balance_check ≈ 0 in the actual years as well). Historical cash is NOT a plug — take it from the data. To make the actuals tie out: (a) transcribe EVERY reported line from the downloaded Balance Sheet; (b) NEVER double-count — cash, current investments and non-current investments are DISTINCT lines (include each exactly once) and reserves are counted once inside equity, not again as a separate asset/liability; (c) cross-check your historical total_assets against the reported "Total Assets" figure in the downloaded data — if it differs by more than ~1% you have mislinked or double-counted a line, so find and fix it before returning (a real prior run overstated total assets by ~18%); (d) reconcile any small residual through the low-materiality catch-all rows (other_current_assets / other_non_current_assets on the asset side, other_current_liabilities / other_non_current_liabilities on the liabilities side) so the historical balance_check lands at ~0; (e) EVERY historical year must reflect that SPECIFIC year's own reported figures — do NOT copy or repeat the prior year's balance-sheet numbers into the next year because a distinct source figure was hard to find (this has actually happened: the latest two historical years showed IDENTICAL balance-sheet values even though the P&L for those same years was correctly distinct). Before returning, spot-check total_assets, cash, and equity across adjacent historical years — if any two adjacent years show the SAME value for several rows at once, you have accidentally duplicated one year's column into another; go back to the source data and locate that specific year's real figures instead. Historical net_debt (borrowings − cash) must also match the company's REPORTED net debt — if it comes out roughly double the reported figure you have linked the wrong borrowings or cash line; fix it.\nBALANCING — cash's FORECAST formula (columns G-K only; historicals still come from the data) MUST be the plug: "{R:total_liabilities_equity}-{R:receivables}-{R:inventory}-{R:other_current_assets}-{R:net_fixed_assets}-{R:cwip}-{R:investments}-{R:other_non_current_assets}" (total liabilities & equity minus every OTHER asset row) — this makes total assets equal total liabilities & equity by construction in every forecast year. Do NOT drive forecast cash off a %-of-sales ratio; that breaks the balance. Because short_term_borrowings above already includes the revolver top-up whenever needed, this cash plug should NEVER compute to a negative number in any forecast year — a negative result here means the short_term_borrowings revolver formula was not applied exactly as dictated; re-check that row before returning, since a company literally cannot hold negative cash. Add a final check row "Balance Check (Assets - Liab.&Equity, should be ~0)" key=balance_check, formula "{R:total_assets}-{R:total_liabilities_equity}".\nYou MUST include rows keyed exactly: equity, net_debt, net_debt_ebitda, capital_employed, net_working_capital, total_assets, total_liabilities_equity, cash, receivables, inventory, other_current_assets, net_fixed_assets, acc_depreciation, cwip, investments, other_non_current_assets, payables, provisions, share_capital, long_term_borrowings, short_term_borrowings, roe, roce.`,
                    STATIC_MODEL),
                callLLM("Valuation & DCF", valuationPrompt,
                    STATIC_MODEL),
                callLLM("Summary Dashboard",
                    `Build a SUMMARY DASHBOARD for ${companyName} that links the most decision-relevant lines from the other sheets, referenced ONLY by these canonical keys:\nop.net_revenue, op.ebitda, op.volume\npnl.net_revenue, pnl.ebitda, pnl.depreciation, pnl.ebit, pnl.interest, pnl.tax, pnl.pat, pnl.eps\nbs.equity, bs.net_debt, bs.capital_employed\ncapex.capex, capex.ocf, capex.fcf, capex.gross_block\nReturn JSON: {"rows":[ {"section":"SECTION NAME"} OR {"label":"Display label","ref":"sheet.canonical_key","fmt":"#,##0"|"0.00","cagr":true|false} ]}. Group into OPERATIONAL, P&L, BALANCE SHEET, CAPEX & FCF. Pick ~15-22 lines. EVERY ref points to an ABSOLUTE canonical value (rupee levels, EPS, or share/volume counts) — there are NO margin/ratio/percentage canonical keys available here, so do NOT add EBITDA-margin, ROE, ROCE, or any other "%" row: applying a percentage format to an absolute-rupee figure renders a nonsensical multi-million-percent number (a real failure produced "11046000%"). Use ONLY fmt "#,##0" for rupee amounts or "0.00" for EPS / x-multiples — NEVER a "0.0%" percentage format on this sheet.`,
                    STATIC_MODEL)
            ]);

            const opRows = Array.isArray(opJson.rows) ? opJson.rows : [];
            // ── Segment revenue/EBITDA zero-fill recovery ──────────────────────────────────────
            // The SOTP prompt tells the model to populate each <tag>_revenue / <tag>_ebitda row from
            // the OPERATIONAL DATA block and states outright that "zero is never an acceptable
            // historical value here" — but nothing enforced it, and the model periodically returns 0
            // for every year of every segment even when the source plainly carries the figures
            // (confirmed on Piramal Pharma: CHG Revenue FY23-FY26 = 2286/2449/2633/2703 and PCH
            // Revenue = 874/985/1093/1274 are both right there under unit "Cr", yet chg_revenue,
            // pch_revenue and cdmo_revenue all came back as straight zeros).
            //
            // All-zero segments are worse than merely missing: net_revenue's forecast is dictated as
            // the SUM of the segment rows, so zeros there silently collapse the entire forecast, and
            // the fixOtherSegments reconciliation immediately below dumps 100% of revenue into the
            // catch-all — producing an "Other Segments" line that just restates Net Revenue while the
            // real segments read 0. Hence this runs BEFORE that reconciliation, so it has genuine
            // per-segment values to subtract rather than reconciling against zeros.
            const zeroSegmentFlags = [];
            if (isSotpCompany) {
                const opRowByKey = (key) => opRows.find(r => r && r.key === key);
                const isAllBlank = (row) => !row || !Array.isArray(row.historical)
                    || !row.historical.some(v => typeof v === "number" && isFinite(v) && v !== 0);
                for (const tag of segmentTags) {
                    for (const kind of ["revenue", "ebitda"]) {
                        const row = opRowByKey(`${tag}_${kind}`);
                        if (!row || !isAllBlank(row)) continue;
                        // A segment with no revenue at all is the real failure; a segment that HAS
                        // revenue but no separately-disclosed EBITDA is normal (very few companies
                        // break EBITDA out by segment), so don't flag that as an error — the
                        // ebitda_margin_<tag> assumption already drives it off the revenue row.
                        const revenueRow = opRowByKey(`${tag}_revenue`);
                        const revenueMissingToo = isAllBlank(revenueRow);
                        const found = segmentAnnualSeries(tag, kind);
                        if (found) {
                            row.historical = found.values;
                            // Never let autoLink's label/value matching re-point these at an unrelated
                            // financial-statement row afterwards — segment revenue has no legitimate
                            // counterpart there, and a previous run mislinked one to a Balance Sheet
                            // debt line. Same rationale as forceMechanicalValue's own noAutoLink use.
                            row.noAutoLink = true;
                            delete row.link;
                            console.warn(`[AI Model] ${tag}_${kind} came back all-zero from the model — recovered ${found.covered} year(s) deterministically from the operational dashboard row "${found.label}" (section "${found.section}"): ${found.values.map(v => v == null ? "–" : Math.round(v)).join(", ")}.`);
                        } else if (kind === "revenue" || !revenueMissingToo) {
                            if (kind === "revenue") {
                                zeroSegmentFlags.push(tag);
                                console.warn(`[AI Model] ${tag}_revenue is all-zero and no annual Rs-Crore revenue row for it could be found in the operational dashboard — this segment will contribute nothing to net_revenue and its share will fall entirely into Other Segments. Flagged on the Summary Dashboard.`);
                            } else {
                                console.log(`[AI Model] ${tag}_ebitda has no separately-disclosed figures (normal — most companies don't break EBITDA out by segment); it stays driven by {A:ebitda_margin_${tag}} off the segment's own revenue row.`);
                            }
                        }
                    }
                }
            }
            // ── other_segments_* historicals: recompute deterministically ──
            // (percent-row unit guard for these sheets runs further below, once every row array
            //  for them has been parsed — see the normalizePercentRows calls after bsRows)
            // other_segments_revenue is defined as "total minus the tagged segments" — per-year
            // arithmetic across 8 columns, which the model has gotten wrong in a specific, repeatable
            // way: its FY(n) cell held FY(n+1)'s revenue, a clean one-year-forward shift. Because
            // net_revenue's FORECAST is dictated as the SUM of the segment rows, that shifted base
            // then compounded through every forecast year and roughly doubled the total.
            // Recomputed here in code instead — the same reasoning buildCapexRows already applies to
            // formulas the model was being handed complete.
            // The per-year total is taken from the company's own REPORTED figure via findEntryByLabel
            // (the authoritative downloaded-data lookup), NOT from the model's net_revenue row:
            // net_revenue is frequently link-only ({"link":"Net Revenue"} with the optional
            // "historical" array omitted), so keying off that row's historicals would silently skip
            // the correction on exactly the runs that need it. Falls back to the model's own row only
            // if the label lookup finds nothing, and leaves the year untouched if neither resolves.
            if (isSotpCompany) {
                const findOpRow = (key) => opRows.find(r => r && r.key === key);
                const histAt = (row, i) => (row && Array.isArray(row.historical) && typeof row.historical[i] === "number" && isFinite(row.historical[i])) ? row.historical[i] : null;
                const fixOtherSegments = (totalKey, otherKey, tagSuffix, labelPattern) => {
                    const otherRow = findOpRow(otherKey);
                    if (!otherRow) return;
                    const totalRow = findOpRow(totalKey);
                    const reportedEntry = findEntryByLabel(labelPattern);
                    const tagRows = segmentTags.map(t => findOpRow(`${t}_${tagSuffix}`));
                    let anchored = 0;
                    const fixed = HIST.map((fy, i) => {
                        const rep = reportedEntry ? reportedEntry.valByFY[fy] : null;
                        const total = (typeof rep === "number" && isFinite(rep)) ? rep : histAt(totalRow, i);
                        if (total == null) return histAt(otherRow, i); // no authoritative total this year — leave the model's own figure
                        anchored++;
                        return total - tagRows.reduce((s, r) => s + (histAt(r, i) ?? 0), 0);
                    });
                    otherRow.historical = fixed;
                    console.log(`[AI Model] Recomputed ${otherKey} historicals in code as (reported ${totalKey} − tagged segments), per year — ${anchored}/${HIST.length} year(s) anchored to a real reported figure.`);
                };
                fixOtherSegments("net_revenue", "other_segments_revenue", "revenue", /^(net\s*revenue|total\s*revenue|revenue\s*from\s*operations|net\s*sales|total\s*income|revenue)$/i);
                fixOtherSegments("ebitda", "other_segments_ebitda", "ebitda", /^ebitda$/i);
            }
            const pnlRows = Array.isArray(pnlJson.rows) ? pnlJson.rows : [];
            const bsRows = Array.isArray(bsJson.rows) ? bsJson.rows : [];
            // Re-verify against the Operational call's ACTUAL output before trusting the segment
            // capex path — isSotpCompany only confirms the Assumptions call provided usable
            // segment-driver keys, not that the Operational Model call went on to comply with the
            // SOTP prompt's dictated row keys (LLM compliance with "use exactly this key" has
            // failed before — see the target_ev_ebitda fallback earlier in this file). Falls back
            // to the ordinary buildCapexRows() — identical to every non-SOTP company's path — the
            // moment either check fails, rather than referencing a segment revenue row that turns
            // out not to exist and cascading a #MISSING error into FCFF and the DCF.
            const opDeliveredSegments = isSotpCompany
                && segmentTags.every(t => opRows.some(r => r && r.key === `${t}_revenue`))
                && opRows.some(r => r && r.key === "other_segments_revenue");
            if (isSotpCompany && !opDeliveredSegments) {
                console.warn("[AI Model] isSotpCompany was true but the Operational Model call did not deliver all the dictated segment/other_segments revenue rows — falling back to the single blended Capex & FCF sheet for this run.");
            }
            const capexRows = opDeliveredSegments ? buildSegmentCapexRows() : buildCapexRows(); // code-generated, no LLM call (see buildCapexRows/buildSegmentCapexRows above)
            const valRows = Array.isArray(valJson.rows) ? valJson.rows : [];
            const summaryRows = Array.isArray(sumJson.rows) ? sumJson.rows : [];
            // Same whole-number-percentage guard the Assumptions sheet already ran (see
            // normalizePercentRows above) — applied to the remaining sheets now that each one's rows
            // are parsed. Most %-rows here are pure formulas and are skipped untouched; this only
            // catches a row where the model hardcoded its own percent VALUES (e.g. an EBITDA-margin
            // or effective-tax-rate history typed as 18 instead of 0.18), which renders ×100 for
            // exactly the same reason. Runs BEFORE planRows so nothing downstream sees the bad units.
            normalizePercentRows("op", opRows);
            normalizePercentRows("pnl", pnlRows);
            normalizePercentRows("bs", bsRows);
            normalizePercentRows("val", valRows);
            planRows("op", opRows);
            planRows("pnl", pnlRows);
            // Balance Sheet ASSETS vs LIABILITIES & EQUITY section dividers — the row schema lets
            // the model add its own {"section":...} rows, but compliance on this specific pair isn't
            // guaranteed (has been observed missing entirely, leaving the sheet as one long
            // undivided list). Insert them deterministically instead of trusting the model to
            // remember, right before the FIRST row on each side (identified by canonical key, same
            // ones the post-write balance diagnostic further below uses) — but only if the model
            // didn't already provide an equivalent header nearby, so a compliant run doesn't get a
            // duplicate. Must run BEFORE planRows("bs", ...) assigns row numbers, since inserting
            // rows afterward would shift every _row it already handed out.
            {
                const ASSET_KEYS_FOR_SECTIONS = ["cash", "receivables", "inventory", "other_current_assets", "net_fixed_assets", "acc_depreciation", "cwip", "investments", "other_non_current_assets"];
                const LIAB_KEYS_FOR_SECTIONS = ["share_capital", "reserves", "equity", "long_term_borrowings", "short_term_borrowings", "payables", "provisions", "deferred_tax_liabilities", "other_non_current_liabilities", "other_current_liabilities"];
                const hasNearbySection = (idx, re) => {
                    for (let i = Math.max(0, idx - 2); i < idx; i++) {
                        if (bsRows[i] && bsRows[i].section && re.test(bsRows[i].section)) return true;
                    }
                    return false;
                };
                const assetIdx = bsRows.findIndex(r => r && !r.section && ASSET_KEYS_FOR_SECTIONS.includes(r.key));
                const liabIdx = bsRows.findIndex(r => r && !r.section && LIAB_KEYS_FOR_SECTIONS.includes(r.key));
                const inserts = [];
                if (assetIdx >= 0 && !hasNearbySection(assetIdx, /asset/i)) inserts.push({ idx: assetIdx, section: "ASSETS" });
                if (liabIdx >= 0 && !hasNearbySection(liabIdx, /liabilit/i)) inserts.push({ idx: liabIdx, section: "LIABILITIES & EQUITY" });
                inserts.sort((a, b) => b.idx - a.idx); // insert from the end so earlier indices stay valid
                for (const { idx, section } of inserts) bsRows.splice(idx, 0, { section });
            }
            planRows("bs", bsRows);
            planRows("capex", capexRows);
            // Ensure the fully-mechanical DCF/target-price chain always exists as REAL rows —
            // forceMechanicalFormula further below can only fix a WRONG formula on a row that
            // exists; if the model omits a row from its own JSON output entirely (a genuine
            // compliance gap, not a naming mismatch — this has happened even with the key already
            // dictated in the prompt, e.g. equity_value_discounted going missing), any OTHER row
            // referencing it by key (e.g. value_per_share's "{R:equity_value_discounted}") has
            // nothing to resolve to, and the reference-integrity gate later strips that row's
            // formula entirely rather than write a #MISSING — leaving it silently blank. Insert a
            // synthetic row with the dictated formula for any of these that's missing, BEFORE
            // planRows assigns row numbers, so every one is guaranteed a real Excel row this build.
            {
                const REQUIRED_VAL_ROWS = [
                    { key: "sum_pv", label: "Sum of PV of FCFF (Rs Cr)", fmt: "#,##0", formula: "{SUM:pv_fcff}" },
                    { key: "tv", label: "Terminal Value (Rs Cr)", fmt: "#,##0", formula: "MAX(0,({LAST:nopat}-{A:terminal_growth}*{LAST:bs.net_working_capital})*(1+{A:terminal_growth})/({A:wacc}-{A:terminal_growth}))" },
                    { key: "pv_tv", label: "PV of Terminal Value (Rs Cr)", fmt: "#,##0", formula: "{R:tv}/(1+{A:wacc})^5" },
                    { key: "ev", label: "Enterprise Value (Rs Cr)", fmt: "#,##0", formula: "{R:sum_pv}+{R:pv_tv}" },
                    { key: "net_debt_last", label: "Less: Net Debt (Rs Cr)", fmt: "#,##0", formula: "{CUR:bs.net_debt}" },
                    { key: "non_operating_assets", label: "Add: Non-operating Assets — CWIP + Investments (Rs Cr)", fmt: "#,##0", formula: "{CUR:bs.cwip}+{CUR:bs.investments}" },
                    { key: "equity_value", label: "Equity Value (Rs Cr)", fmt: "#,##0", formula: "{R:ev}+{R:non_operating_assets}-{R:net_debt_last}" },
                    { key: "equity_value_discounted", label: "Equity Value, post holding discount (Rs Cr)", fmt: "#,##0", formula: "{R:equity_value}*(1-{A:holding_discount})" },
                    { key: "value_per_share", label: "Value per Share (DCF)", fmt: "#,##0", formula: "{R:equity_value_discounted}/{CUR:assum.shares_out}", emphasis: "highlight" },
                    { key: "target_price_pe", label: "Target Price (P/E based)", fmt: "#,##0", formula: "{FWD1:pnl.eps}*{A:target_pe}" },
                    { key: "target_price_ev_ebitda", label: "Target Price (EV/EBITDA based)", fmt: "#,##0", formula: "(({FWD1:pnl.ebitda}*{A:target_ev_ebitda})-{FWD1:bs.net_debt})/{CUR:assum.shares_out}" },
                    { key: "target_price", label: "Blended Target Price", fmt: "#,##0", formula: "({R:target_price_pe}+{R:target_price_ev_ebitda})/2", emphasis: "highlight" },
                    { key: "upside", label: "DCF Upside to Current Price", fmt: "0.0%", formula: "{R:value_per_share}/{R:current_price}-1" },
                ];
                const missingKeys = [];
                for (const spec of REQUIRED_VAL_ROWS) {
                    if (valRows.some(r => r && r.key === spec.key)) continue;
                    valRows.push({ key: spec.key, label: spec.label, fmt: spec.fmt, scalar: true, formula: spec.formula, emphasis: spec.emphasis });
                    missingKeys.push(spec.key);
                }
                if (missingKeys.length) console.warn(`[AI Model] val.${missingKeys.join(", val.")} ${missingKeys.length === 1 ? "was" : "were"} missing from the model's own output entirely (not just a wrong formula) — inserted with the dictated formula so rows referencing ${missingKeys.length === 1 ? "it" : "them"} don't silently go blank.`, missingKeys);
            }
            // WACC BUILD-UP (section 5) is PURE DISPLAY — every value it needs already exists as a
            // correctly-computed scalar on the Assumptions sheet (risk_free_rate, equity_risk_premium,
            // beta, cost_of_debt, tax_rate, wacc itself). There is zero judgment left for the model to
            // exercise here, yet — unlike every other scalar section on this sheet — the prompt never
            // dictates a "key" for any of its 7 rows, so the model is free to invent its own
            // (key/label pairs observed varying run to run: "tax_rate_wacc"/"wacc_final" one run,
            // something else the next), and just as often omits "scalar":true, the formula, or the
            // row entirely. Forcing each row's FORMULA after the fact (as done for every other
            // section) still depends on first FINDING that row by key-or-label — and that lookup has
            // kept failing for most of these 7 rows even with a forced formula ready to apply, because
            // there's often no row to find at all. Rather than keep chasing whatever the model did or
            // didn't produce, generate this section entirely in code — replacing the model's own rows
            // in-place if it emitted a "WACC BUILD-UP" section header, or appending a fresh one if it
            // omitted the section altogether — so every value is guaranteed present, keyed
            // predictably, and correct, with no dependency on the model at all for this section.
            {
                const WACC_ROWS = [
                    { key: "risk_free", label: "Risk-free Rate", fmt: "0.0%", formula: "{A:risk_free_rate}" },
                    { key: "equity_risk_premium", label: "Equity Risk Premium", fmt: "0.0%", formula: "{A:equity_risk_premium}" },
                    { key: "beta", label: "Beta", fmt: "0.00", formula: "{A:beta}" },
                    { key: "cost_of_equity", label: "Cost of Equity", fmt: "0.0%", formula: "{A:risk_free_rate}+{A:beta}*{A:equity_risk_premium}" },
                    { key: "cost_of_debt", label: "Cost of Debt", fmt: "0.0%", formula: "{A:cost_of_debt}" },
                    { key: "wacc_tax_rate", label: "Tax Rate", fmt: "0.0%", formula: "{A:tax_rate}" },
                    { key: "wacc_final", label: "WACC", fmt: "0.0%", formula: "{A:wacc}" },
                ].map(r => ({ ...r, scalar: true }));
                const sectionIdx = valRows.findIndex(r => r && r.section && /wacc.{0,4}build.?.?up/i.test(r.section));
                if (sectionIdx >= 0) {
                    let endIdx = sectionIdx + 1;
                    while (endIdx < valRows.length && !(valRows[endIdx] && valRows[endIdx].section)) endIdx++;
                    valRows.splice(sectionIdx + 1, endIdx - (sectionIdx + 1), ...WACC_ROWS);
                    console.warn(`[AI Model] Replaced the model's own "WACC BUILD-UP" rows with a deterministic version referencing the Assumptions sheet directly — this section is pure display of already-computed assumption values, and the model has been observed omitting or breaking these specific rows unpredictably (no dictated key for any of them, unlike every other section on this sheet).`);
                } else {
                    valRows.push({ section: "WACC BUILD-UP" }, ...WACC_ROWS);
                    console.warn(`[AI Model] "WACC BUILD-UP" section was missing entirely from the model's output — appended a deterministic version referencing the Assumptions sheet directly.`);
                }
            }
            planRows("val", valRows);

            // ── Guard: abort BEFORE writing anything if a critical sheet came back essentially
            // empty. callLLM() never throws — a fully-failed call (timeout, exhausted retries,
            // exhausted cross-provider fallback) returns {} so one bad call doesn't block the rest
            // of the model — but that means nothing downstream previously distinguished "this sheet
            // has real content" from "this sheet quietly died," and a blank FM sheet would still get
            // written and reported as a success. Better to fail the whole build with a clear message
            // than ship a model with a blank sheet that looks legitimate.
            const emptySheets = Object.entries({ assum: assumRows, op: opRows, pnl: pnlRows, bs: bsRows, capex: capexRows, val: valRows, summary: summaryRows })
                .filter(([, rows]) => (rows || []).filter(r => r && !r.section).length === 0)
                .map(([key]) => SHEETS[key]);
            if (emptySheets.length) {
                aiStatus(null);
                showWarning(`AI Financial Model build failed — ${emptySheets.join(", ")} came back empty, likely an OpenRouter timeout or outage. Nothing was written to Excel; please try again.`);
                return;
            }

            // ── Double-counting guard: "total row + its own children, both linked" ──
            // The Balance Sheet prompt's TOTALS vs LINE ITEMS rule (see DATA_BASIS_NOTE) tells the
            // model not to link a SUBTOTAL row AND separately link the indented children that sum
            // into it — but that's a judgment call the model can get wrong, and it has: "Other
            // Current Assets" linked as the PARENT total while Cash/Receivables were ALSO linked
            // separately to its own children, double-counting them into Total Assets (confirmed by
            // tracing FM's Total Assets against the reported figure — the excess matched the
            // double-counted lines almost exactly, and forecast years balanced perfectly since they
            // never touch Annual Data at all, only the historical LINKING was ever exposed to this).
            // Rather than trust the model to self-police this rule correctly every time, detect and
            // correct it deterministically: for each "catch-all" row (the ones the prompt already
            // treats as low-materiality residual buckets), check whether its OWN linked source is a
            // SUBTOTAL whose indented-children subtree contains another BS row's ALREADY-linked
            // source — if so, subtract the overlap out of the catch-all's historicals.
            const getSubtreeExcelRows = (sheetName, startExcelRow) => {
                const grid = grids[sheetName];
                const rows = new Set();
                if (!Array.isArray(grid)) return rows;
                const startIdx = startExcelRow - 1;
                if (startIdx < 0 || startIdx >= grid.length) return rows;
                const startRaw = String((grid[startIdx] && grid[startIdx][0]) || "");
                const startIndent = (startRaw.match(/^(\s*)/) || ["", ""])[1].length;
                for (let r = startIdx + 1; r < grid.length; r++) {
                    const raw = String((grid[r] && grid[r][0]) || "");
                    const trimmed = raw.trim();
                    if (!trimmed) break; // blank separator row — end of this sub-table/section
                    if (normLabel(trimmed) === "parameter") break; // next sub-table's header
                    const indent = (raw.match(/^(\s*)/) || ["", ""])[1].length;
                    if (indent <= startIndent) break; // sibling or higher-level row — subtree ends
                    rows.add(r + 1); // back to 1-based Excel row number
                }
                return rows;
            };
            // Same idea as getSubtreeExcelRows, but ONE level deep only — the immediate children,
            // not every deeper descendant. Needed for isPlausibleSubtotal below: summing the FULL
            // subtree there double-counts a grandchild breakdown on top of its own parent line
            // (e.g. "Long Term Investment" plus its own nested ":Quoted"/":Unquoted" split both
            // being summed), which can push the sum well past the parent's real total in years
            // where that nested breakdown is large — an intermittent false negative that has
            // actually caused promotion to silently succeed in one run and fail in the next for the
            // exact same company/row, depending only on how large that particular year's nested
            // breakdown happened to be.
            const getDirectChildExcelRows = (sheetName, startExcelRow) => {
                const grid = grids[sheetName];
                const rows = new Set();
                if (!Array.isArray(grid)) return rows;
                const startIdx = startExcelRow - 1;
                if (startIdx < 0 || startIdx >= grid.length) return rows;
                const startRaw = String((grid[startIdx] && grid[startIdx][0]) || "");
                const startIndent = (startRaw.match(/^(\s*)/) || ["", ""])[1].length;
                let directIndent = null;
                for (let r = startIdx + 1; r < grid.length; r++) {
                    const raw = String((grid[r] && grid[r][0]) || "");
                    const trimmed = raw.trim();
                    if (!trimmed) break;
                    if (normLabel(trimmed) === "parameter") break;
                    const indent = (raw.match(/^(\s*)/) || ["", ""])[1].length;
                    if (indent <= startIndent) break;
                    if (directIndent == null) directIndent = indent; // first row inside = the direct-child depth
                    if (indent === directIndent) rows.add(r + 1); // deeper rows are grandchildren — skip
                }
                return rows;
            };
            // Walk UP from a child row to the nearest LESS-indented row above it (its immediate
            // parent in the sheet's hierarchy) within the same sub-table. Returns null if the row
            // is already top-level, or if a blank/"Parameter" boundary is hit first (left the
            // sub-table without finding a parent).
            const findParentExcelRow = (sheetName, childExcelRow) => {
                const grid = grids[sheetName];
                if (!Array.isArray(grid)) return null;
                const childIdx = childExcelRow - 1;
                if (childIdx < 0 || childIdx >= grid.length) return null;
                const childRaw = String((grid[childIdx] && grid[childIdx][0]) || "");
                const childIndent = (childRaw.match(/^(\s*)/) || ["", ""])[1].length;
                if (childIndent === 0) return null;
                for (let r = childIdx - 1; r >= 0; r--) {
                    const raw = String((grid[r] && grid[r][0]) || "");
                    const trimmed = raw.trim();
                    if (!trimmed) return null;
                    if (normLabel(trimmed) === "parameter") return null;
                    const indent = (raw.match(/^(\s*)/) || ["", ""])[1].length;
                    if (indent < childIndent) return r + 1;
                }
                return null;
            };
            // Confirms a candidate parent row is a GENUINE subtotal — i.e. its DIRECT children's
            // values actually sum to it — rather than just "some less-indented row above it"
            // (indentation alone doesn't guarantee a total/child relationship). Deliberately uses
            // getDirectChildExcelRows here (one level deep), NOT the full descendant subtree a
            // caller might pass as "childRows" for overlap-detection purposes elsewhere — summing
            // every deeper descendant would double-count a grandchild breakdown on top of its own
            // parent line (see getDirectChildExcelRows's comment).
            const isPlausibleSubtotal = (sheetName, parentRow) => {
                const parentEntry = entryAtRow(sheetName, parentRow);
                const childEntries = [...getDirectChildExcelRows(sheetName, parentRow)].map(r => entryAtRow(sheetName, r)).filter(Boolean);
                if (!parentEntry || !childEntries.length) return false;
                let matched = 0, total = 0;
                for (const fy in parentEntry.valByFY) {
                    total++;
                    const parentVal = parentEntry.valByFY[fy];
                    const childSum = childEntries.reduce((s, e) => s + (e.valByFY[fy] || 0), 0);
                    if (Math.abs(childSum - parentVal) / Math.max(Math.abs(parentVal), 1) < 0.05) matched++;
                }
                return total > 0 && matched / total >= 0.5;
            };
            // These 4 catch-all keys always correspond to a predictable, well-known label in the
            // source data — unlike every other Balance Sheet row, they don't need the model's OWN
            // "link" field to be findable. Relying on it anyway means this whole guard silently
            // never engages for a row the model happened to omit "link" on (giving only its own
            // guessed "historical" numbers instead) — observed happening specifically to
            // other_current_assets/other_current_liabilities while their "Non-current" counterparts
            // came through linked correctly in the same run. Look these up directly as a fallback
            // whenever the model's own link is missing or doesn't resolve.
            const CATCHALL_LABEL_FALLBACK = {
                other_current_assets: /^other\s+current\s+assets?$/i,
                other_non_current_assets: /^other\s+non-?\s*current\s+assets?$/i,
                other_current_liabilities: /^other\s+current\s+liabilities$/i,
                other_non_current_liabilities: /^other\s+non-?\s*current\s+liabilities$/i,
            };
            const findMapEntryByLabelRegex = (regex) => {
                for (const m of [idxKF, idxAnnual]) {
                    for (const e of m.values()) {
                        if (regex.test(e.label.trim())) return e;
                    }
                }
                return null;
            };
            // These 4 catch-all rows don't correspond to any single real row in the source data —
            // "current assets, EXCEPT the ones that already have their own FM row" isn't a concept
            // the vendor's table expresses directly, so whatever the model links THIS row to is
            // inherently ambiguous (a fat parent subtotal that already includes cash/receivables on
            // the assets side; a plain sibling with no overlap at all on the liabilities side; a
            // leaf almost-identically labelled to its own parent on the non-current side). Trying to
            // repair whatever the model happened to link THIS row to is fixing a symptom. Instead,
            // anchor off a DIFFERENT, unambiguous sibling row that's already reliably linked — none
            // of these four have a near-duplicate label anywhere near them — and use ITS immediate
            // parent as this catch-all's true home, regardless of what (if anything) the catch-all's
            // own link says. Verified by hand against real source data for all four: inventory's
            // parent is "Current Assets" (correct target for other_current_assets); investments'
            // (non-current) parent is "Other Non-current Assets"; payables' parent is "Current
            // Liabilities"; deferred_tax_liabilities' parent is "Other Non-Current Liabilities".
            const CATCHALL_ANCHORS = {
                other_current_assets: "inventory",
                other_non_current_assets: "investments",
                other_current_liabilities: "payables",
                other_non_current_liabilities: "deferred_tax_liabilities",
            };
            const resolveItemLink = (it) => it && it.link && (resolveLink(it.link) || resolveLink(stripUnit(it.link)));
            const CATCHALL_BS_KEYS = ["other_current_assets", "other_non_current_assets", "other_current_liabilities", "other_non_current_liabilities"];
            for (const key of CATCHALL_BS_KEYS) {
                const row = symbols.bs && symbols.bs[key];
                const item = row && bsRows.find(r => r && r._row === row);
                if (!item) continue;

                let linkEntry = null, subtreeRows = null, baseHistorical = null;

                // PRIMARY — anchor off a reliably-linked sibling's parent (see comment above).
                const anchorKey = CATCHALL_ANCHORS[key];
                const anchorRow = anchorKey && symbols.bs && symbols.bs[anchorKey];
                const anchorItem = anchorRow && bsRows.find(r => r && r._row === anchorRow);
                const anchorEntry = resolveItemLink(anchorItem);
                if (anchorEntry) {
                    const parentRow = findParentExcelRow(anchorEntry.sheet, anchorEntry.row);
                    const parentEntry = parentRow && entryAtRow(anchorEntry.sheet, parentRow);
                    const parentSubtree = parentEntry && getSubtreeExcelRows(anchorEntry.sheet, parentRow);
                    if (parentEntry && parentSubtree && parentSubtree.size && isPlausibleSubtotal(anchorEntry.sheet, parentRow)) {
                        linkEntry = { sheet: anchorEntry.sheet, label: parentEntry.label, row: parentRow, valByFY: parentEntry.valByFY };
                        subtreeRows = parentSubtree;
                        baseHistorical = HIST.map(fy => (parentEntry.valByFY[fy] != null ? parentEntry.valByFY[fy] : null));
                    }
                }

                // FALLBACK — the anchor row itself wasn't linked this run (or its parent didn't pass
                // the plausible-subtotal check): fall back to starting from the catch-all's OWN link
                // (direct, or via the label regex below when the model omitted "link" entirely) and
                // climbing to its immediate parent when that parent captures more already-linked
                // rows than the catch-all's own level does. Capped at exactly ONE hop, and never
                // allowed to reach the sheet's absolute top-level section (ASSETS / EQUITY AND
                // LIABILITIES) — climbing that far would make the residual balloon into "everything
                // else on this side of the balance sheet."
                if (!linkEntry) {
                    linkEntry = (item.link && resolveItemLink(item))
                        || (CATCHALL_LABEL_FALLBACK[key] && findMapEntryByLabelRegex(CATCHALL_LABEL_FALLBACK[key]));
                    if (!linkEntry) continue;
                    subtreeRows = getSubtreeExcelRows(linkEntry.sheet, linkEntry.row);
                    baseHistorical = HIST.map((fy, i) => {
                        if (linkEntry.valByFY && linkEntry.valByFY[fy] != null) return linkEntry.valByFY[fy];
                        return (Array.isArray(item.historical) && item.historical[i] != null) ? item.historical[i] : null;
                    });
                    if (baseHistorical.every(v => v == null)) baseHistorical = null;

                    const countOverlaps = (subtree) => {
                        let n = 0;
                        for (const other of bsRows) {
                            if (!other || other === item || other.section || !other.link) continue;
                            const otherEntry = resolveLink(other.link) || resolveLink(stripUnit(other.link));
                            if (otherEntry && otherEntry.sheet === linkEntry.sheet && subtree.has(otherEntry.row)) n++;
                        }
                        return n;
                    };
                    const parentRow = findParentExcelRow(linkEntry.sheet, linkEntry.row);
                    if (parentRow) {
                        const parentEntry = entryAtRow(linkEntry.sheet, parentRow);
                        if (parentEntry && !/^(assets|equity\s*(and|&)\s*liabilities)$/i.test(normLabel(parentEntry.label))) {
                            const parentSubtree = getSubtreeExcelRows(linkEntry.sheet, parentRow);
                            if (parentSubtree.size && isPlausibleSubtotal(linkEntry.sheet, parentRow)) {
                                const ownOverlapCount = subtreeRows.size ? countOverlaps(subtreeRows) : -1; // -1 forces re-basing below when there's no subtree at all (the leaf case)
                                const parentOverlapCount = countOverlaps(parentSubtree);
                                if (parentOverlapCount > ownOverlapCount) {
                                    console.warn(`[AI Model] bs.${key} was linked to "${linkEntry.label}" — re-based on its parent "${parentEntry.label}" so sibling line items claimed elsewhere on this sheet (or with no dedicated FM row of their own) aren't double-counted or dropped.`);
                                    linkEntry = { sheet: linkEntry.sheet, label: parentEntry.label, row: parentRow, valByFY: parentEntry.valByFY };
                                    subtreeRows = parentSubtree;
                                    baseHistorical = HIST.map(fy => (parentEntry.valByFY[fy] != null ? parentEntry.valByFY[fy] : null));
                                }
                            }
                        }
                    }
                }
                if (!subtreeRows.size || !baseHistorical) continue; // no usable parent found either way — nothing to fix

                const overlapping = [];
                for (const other of bsRows) {
                    if (!other || other === item || other.section || !other.link) continue;
                    const otherEntry = resolveLink(other.link) || resolveLink(stripUnit(other.link));
                    if (otherEntry && otherEntry.sheet === linkEntry.sheet && subtreeRows.has(otherEntry.row)) {
                        overlapping.push({ label: other.key || other.label, entry: otherEntry });
                    }
                }
                // Always recompute from baseHistorical (the link's own authoritative per-year
                // values, not the model's transcription — see above) and write it, even with zero
                // overlaps: baseHistorical can differ from whatever the model originally gave this
                // row regardless of whether any overlap was found, so there's no "nothing changed,
                // skip" case left to detect.
                const adjusted = HIST.map((fy, i) => {
                    const total = baseHistorical[i];
                    if (total == null) return total;
                    let sub = 0;
                    for (const o of overlapping) { const v = o.entry.valByFY[fy]; if (v != null) sub += v; }
                    return total - sub;
                });
                if (overlapping.length) {
                    console.warn(`[AI Model] bs.${key} was linked to "${linkEntry.label}", which is a SUBTOTAL that already includes ${overlapping.map(o => o.label).join(", ")} — those are ALSO linked separately elsewhere on this sheet, double-counting them into Total Assets/Liabilities. Adjusted bs.${key}'s historicals to exclude the overlap.`);
                }
                item.historical = adjusted;
                delete item.link; // no longer a direct cell link — now a derived/adjusted figure
                // A derived figure like this (total minus overlap) can easily land within 2 years'
                // rounding tolerance of some UNRELATED row elsewhere in the data purely by
                // coincidence — and without this flag, writeDataSheet still calls autoLink(item) on
                // every row lacking "link" (its value-based fallback doesn't require one), which
                // would silently re-link the WHOLE row (all 8 years) to that unrelated match,
                // discarding this carefully-computed adjustment for every year, not just the
                // coincidentally-matching ones. Same rationale as forceMechanicalValue's noAutoLink
                // — a deliberately-computed authoritative value must never be second-guessed by
                // label/value coincidence matching afterward.
                item.noAutoLink = true;
                item.source = `"${linkEntry.label}"${overlapping.length ? ` minus ${overlapping.map(o => o.label).join("+")}` : ""} (avoids double-counting / dropped line items)`;
            }

            // Deterministic label-anchoring for the P&L's 3 most critical absolute figures — same
            // rationale as CATCHALL_ANCHORS/CATCHALL_LABEL_FALLBACK above: Revenue/EBITDA/PAT always
            // correspond to a predictable, well-known label in Key Financials, so don't leave
            // correctness dependent on the model's OWN "historical" transcription of a WIDE table (13+
            // columns: 8 historical + 5 forecast, sometimes more with a CAGR column). A systematic
            // one-column-forward misread has actually been observed here — 7 straight years of
            // Revenue/PAT reading one FY ahead of their own column label (confirmed independently by an
            // adjacent, correctly-aligned PAT Margin row landing on the right column while Revenue/PAT
            // didn't) — because autoLink's value-confirmation gate requires the model's own array to
            // AGREE with a candidate row for at least one year, and a fully-shifted array essentially
            // never coincidentally agrees, so it falls through to being written as-is: wrong, static,
            // unlinked numbers with no live formula to audit against. Scoped to just these 3 keys
            // (not the whole P&L) because their vendor labels are consistently standardized across
            // companies — other rows like depreciation/interest/tax have far less consistent labelling
            // and a blind regex match risks grabbing the wrong row (e.g. "Interest Coverage" instead of
            // "Interest Expense").
            const PNL_CANONICAL_LABELS = {
                net_revenue: /^(net\s*revenue|total\s*revenue|revenue\s*from\s*operations|net\s*sales|total\s*income|revenue)$/i,
                ebitda: /^ebitda$/i,
                pat: /^(pat|net\s*profit|profit\s*after\s*tax)$/i,
            };
            for (const key of Object.keys(PNL_CANONICAL_LABELS)) {
                const row = symbols.pnl && symbols.pnl[key];
                const item = row && pnlRows.find(r => r && r._row === row);
                if (!item) continue;
                const entry = findMapEntryByLabelRegex(PNL_CANONICAL_LABELS[key]);
                if (!entry) continue;
                const linkedHist = HIST.map(fy => (entry.valByFY[fy] != null ? entry.valByFY[fy] : null));
                if (linkedHist.every(v => v == null)) continue; // nothing usable found in the source data — leave the model's own data alone
                const ownHist = Array.isArray(item.historical) ? item.historical : [];
                const disagreed = HIST.some((fy, i) => {
                    const own = ownHist[i], linked = linkedHist[i];
                    if (own == null || linked == null) return false;
                    return Math.abs(Number(own) - linked) / Math.max(Math.abs(linked), 1) > 0.02;
                });
                if (disagreed) {
                    console.warn(`[AI Model] pnl.${key} — re-anchored historicals to the deterministically-parsed "${entry.label}" row (source data), overriding the model's own transcription which disagreed by more than 2% in at least one year (a wide-table column-misalignment misread has been observed for this exact row before).`);
                }
                item.link = entry.label; // exact-label fast path in autoLink — resolves to a live formula, no value-gate needed
                item.historical = linkedHist;
                item.noAutoLink = false;
            }

            // Force a handful of PURELY MECHANICAL rows to their single correct formula,
            // regardless of what the model actually wrote. A discount factor is ALWAYS
            // 1/(1+WACC)^N; ROE/ROCE are ALWAYS PAT/equity and EBIT/capital-employed — none of
            // these involve any judgment call, so there's no legitimate reason for a different
            // formula to appear. In practice "forecast_only":true rows like these have been
            // observed getting misread as "give plain forecast VALUES instead of a formula" (a
            // flat [1,1,1,1,1] placeholder for the discount factor, a bare [0,0,0,0,0] for
            // ROE/ROCE) — which silently breaks every downstream DCF/returns figure without ever
            // surfacing as a visible error (discount factor stuck at 1.0 means nothing gets
            // discounted; ROE/ROCE read 0.0% every year). Locate each row by its dictated key
            // first, falling back to a label match (mirrors reconcileCanonKeys below) in case the
            // model used a different key name, and overwrite its formula unconditionally.
            const forceMechanicalFormula = (sheetKey, rows, key, labelMatch, formula, opts = {}) => {
                let row = symbols[sheetKey] && symbols[sheetKey][key];
                let item = row && rows.find(r => r && r._row === row);
                if (!item) {
                    item = rows.find(r => r && !r.section && r._row && r.label && labelMatch.test(r.label));
                    if (item) {
                        symbols[sheetKey][key] = item._row;
                        console.warn(`[AI Model] ${sheetKey}.${key} located via label match ("${item.label}") instead of its dictated key — the model used a different key name.`);
                    }
                }
                if (!item) {
                    console.warn(`[AI Model] Could not locate ${sheetKey}.${key} to force its formula — row not found by key or label; it may be missing from the sheet entirely.`);
                    return;
                }
                item.formula = formula;
                if (!opts.keepHistorical) { delete item.forecast; delete item.historical; delete item.link; }
                delete item.value;
                if (opts.forecastOnly) item.forecast_only = true; // DCF-style row — never guess this either way
                // A scalar row (single column-B value, e.g. value_per_share/target price) needs
                // "scalar":true set for writeDataSheet to render it correctly — if the row came
                // back broken specifically because the model omitted its formula, it may well have
                // omitted this flag too, and a scalar row has no flat-carry-forward fallback the
                // way a per-year row does, so getting this wrong means a genuinely blank cell.
                if (opts.scalar) item.scalar = true;
            };
            // Same idea as forceMechanicalFormula, but for a row that should hold a fixed NUMBER
            // (repeated across every year cell) rather than a formula — current_price specifically:
            // even though the prompt already hands the Judgment call the real fetched price as a
            // given fact, this guarantees the sheet shows that exact figure regardless of whether the
            // model transcribed it faithfully, rounded it, or ignored the instruction outright.
            const findForcedRow = (sheetKey, rows, key, labelMatch) => {
                let row = symbols[sheetKey] && symbols[sheetKey][key];
                let item = row && rows.find(r => r && r._row === row);
                if (!item) {
                    item = rows.find(r => r && !r.section && r._row && r.label && labelMatch.test(r.label));
                    if (item) {
                        symbols[sheetKey][key] = item._row;
                        console.warn(`[AI Model] ${sheetKey}.${key} located via label match ("${item.label}") instead of its dictated key — the model used a different key name.`);
                    }
                }
                if (!item) {
                    console.warn(`[AI Model] Could not locate ${sheetKey}.${key} to force its value — row not found by key or label; it may be missing from the sheet entirely.`);
                }
                return item;
            };
            // TEMP DIAGNOSTIC (see the PROVENANCE REPORT further below) — logs what the model
            // ORIGINALLY produced for a row right before code overwrites it, so a real mismatch
            // (e.g. the model using a different figure than the one it was explicitly given) is
            // visible in the console rather than silently replaced. Remove alongside the report once
            // the current round of hallucination/mismatch bugs is tracked down.
            const logForcedCorrection = (sheetKey, key, item, newValueDesc) => {
                const before = item.scalar ? item.value
                    : (Array.isArray(item.historical) ? item.historical[0] : undefined);
                console.log(`[AI Model][PROVENANCE] ${sheetKey}.${key}: model originally had "${before}" -> forced to ${newValueDesc}`);
            };
            const forceMechanicalValue = (sheetKey, rows, key, labelMatch, value, sourceNote) => {
                const item = findForcedRow(sheetKey, rows, key, labelMatch);
                if (!item) return;
                logForcedCorrection(sheetKey, key, item, value);
                item.historical = new Array(HIST_YEARS).fill(value);
                item.forecast = new Array(FC_YEARS).fill(value);
                delete item.formula;
                delete item.link;
                delete item.value;
                // A code-forced value is, by definition, authoritative — never let autoLink's own
                // label/value-matching heuristic second-guess it later by live-linking this row to
                // some UNRELATED cell that happens to share a similar label or a coincidentally
                // agreeing value for 2+ years (this has actually happened — face_value ended up
                // live-linked to a completely unrelated "Net Debt/EBITDA" row this way).
                item.noAutoLink = true;
                if (sourceNote) item.source = sourceNote;
            };
            // Same idea, but for a SCALAR row ("scalar":true — a single value in column B, not a
            // per-year series) — current_price on the Valuation & DCF sheet specifically: it's one
            // fact, today's price, not something with a value per year (see currentPriceValueNote).
            const forceMechanicalScalar = (sheetKey, rows, key, labelMatch, value, sourceNote) => {
                const item = findForcedRow(sheetKey, rows, key, labelMatch);
                if (!item) return;
                logForcedCorrection(sheetKey, key, item, value);
                item.scalar = true;
                item.value = value;
                delete item.formula;
                delete item.historical;
                delete item.forecast;
                delete item.link;
                item.noAutoLink = true; // see forceMechanicalValue above — same rationale
                if (sourceNote) item.source = sourceNote;
            };
            // shares_out is a genuine PER-YEAR figure (share count changes over time — buybacks,
            // ESOPs, rights issues), NOT a constant to freeze into one cell — force each
            // historical year's own computed Market-Cap÷Price value (see sharesOutByFY above),
            // and flat-carry the LATEST known share count across the forecast years (no further
            // dilution/buyback modeled — same convention already used for share_capital/
            // long_term_borrowings elsewhere on this sheet).
            {
                const item = findForcedRow("assum", assumRows, "shares_out", /shares?\s*out(standing)?/i);
                if (item) {
                    const historical = HIST.map(fy => (sharesOutByFY[fy] ? sharesOutByFY[fy].value : null));
                    if (historical.some(v => v != null)) {
                        const lastKnown = [...historical].reverse().find(v => v != null);
                        item.historical = historical;
                        item.forecast = new Array(FC_YEARS).fill(lastKnown);
                        item.scalar = false;
                        delete item.value;
                        delete item.formula;
                        delete item.link;
                        item.noAutoLink = true; // authoritative sourced data — see forceMechanicalValue's rationale
                        item.source = "sc_year_data (Market Cap / Annual Closing Price, per year)";
                    }
                }
            }
            if (basicInfo && basicInfo.faceValue) {
                forceMechanicalValue("assum", assumRows, "face_value", /face\s*value/i, basicInfo.faceValue, "Live company data (transcriptanalyser.com)");
            } else {
                // FALLBACK — the live face-value fetch failed/returned nothing usable for this
                // company. Rather than leave the LLM to guess (which has actually mislinked this
                // row to a completely unrelated metric before — e.g. Net Debt/EBITDA — via
                // autoLink's own value-matching heuristic), derive it deterministically: Face
                // Value = Share Capital ÷ Number of Shares, using the SAME latest-year share count
                // already computed above (sharesOutByFY, Market Cap ÷ Price) — Share Capital
                // (Rs Cr) ÷ Shares Outstanding (Cr) = Rs per share, exactly the definition of face
                // value. Picks the most recent year where BOTH figures are actually available,
                // not necessarily the very latest sharesOutByFY year.
                const shareCapitalEntry = findEntryByLabel(/^(equity\s*)?(paid.?up\s*)?share\s*capital$/i);
                const faceValueFY = shareCapitalEntry
                    ? Object.keys(sharesOutByFY).map(Number).filter(isFinite)
                        .filter(fy => shareCapitalEntry.valByFY[fy] != null)
                        .sort((a, b) => b - a)[0]
                    : null;
                if (faceValueFY != null) {
                    const derivedFaceValue = shareCapitalEntry.valByFY[faceValueFY] / sharesOutByFY[faceValueFY].value;
                    if (isFinite(derivedFaceValue) && derivedFaceValue > 0) {
                        forceMechanicalValue("assum", assumRows, "face_value", /face\s*value/i, derivedFaceValue, `Share Capital ÷ Number of Shares, FY${faceValueFY}`);
                    }
                } else {
                    console.warn("[AI Model] Could not fetch a live face value AND could not derive one from Share Capital ÷ Number of Shares (missing Share Capital label or share count) — face_value was left for the model to derive itself. Verify manually.");
                }
            }
            if (latestPrice) {
                forceMechanicalScalar("val", valRows, "current_price", /current\s*price/i, latestPrice.price, `NSE EOD price feed, ${latestPrice.date}`);
            }
            // wacc/beta/target_pe/target_ev_ebitda are single, unchanging JUDGMENT facts too — the
            // Assumptions row schema (assumSchema) only knows one row shape (a per-year series), so
            // the model still writes these as HIST_YEARS+FC_YEARS identical repeated cells same as
            // every other "constant global input." Rather than teach the schema a second row shape
            // just for these keys (more prompt-compliance risk), collapse each into a genuine
            // one-cell scalar here, keeping whatever value/source the model actually gave it — this
            // doesn't need forceMechanicalScalar (which OVERRIDES with a code-supplied value); it
            // just reshapes the row the model already produced.
            const collapseToScalar = (sheetKey, rows, key, labelMatch) => {
                const item = findForcedRow(sheetKey, rows, key, labelMatch);
                if (!item) return;
                const value = Array.isArray(item.forecast) && item.forecast.length ? item.forecast[0]
                    : (Array.isArray(item.historical) && item.historical.length ? item.historical[item.historical.length - 1]
                        : (typeof item.value === "number" ? item.value : null));
                if (typeof value !== "number" || !isFinite(value)) {
                    console.warn(`[AI Model] ${sheetKey}.${key} had no usable numeric value to collapse into a scalar — left as a per-year row.`);
                    return;
                }
                item.scalar = true;
                item.value = value;
                delete item.formula;
                delete item.historical;
                delete item.forecast;
                delete item.link;
            };
            collapseToScalar("assum", assumRows, "wacc", /\bwacc\b/i);
            collapseToScalar("assum", assumRows, "beta", /\bbeta\b/i);
            collapseToScalar("assum", assumRows, "target_pe", /target\s*p\/?e/i);
            collapseToScalar("assum", assumRows, "target_ev_ebitda", /target\s*ev\s*\/?\s*ebitda/i);
            collapseToScalar("assum", assumRows, "tax_rate", /tax\s*rate/i);
            collapseToScalar("assum", assumRows, "pat_retention", /pat\s*retention/i);
            collapseToScalar("assum", assumRows, "risk_free_rate", /risk.?free/i);
            collapseToScalar("assum", assumRows, "equity_risk_premium", /equity\s*risk\s*premium/i);
            collapseToScalar("assum", assumRows, "cost_of_debt", /cost\s*of\s*debt/i);
            collapseToScalar("assum", assumRows, "terminal_growth", /terminal\s*growth/i);
            // Only present when a broker report was fed AND the model found an explicit target
            // price in it (see the conditional BROKER TARGET PRICE instruction above) — no-ops
            // silently otherwise, same as every other collapseToScalar call when its row is absent.
            collapseToScalar("assum", assumRows, "broker_target_price", /broker.*target|target.*broker/i);
            // Reasonableness guardrail on target_pe/target_ev_ebitda — an aggressively high multiple
            // can still be "real" (this company's own genuine historical trading range, per
            // valuationChartNote above) yet imply an earnings/EBITDA yield below the business's own
            // cost of capital. That's sometimes a legitimate call (market pricing in strong growth)
            // but is exactly the kind of thing worth a visible flag rather than silently flowing
            // through into the target price unexamined. Diagnostic only — logs a warning for whoever
            // reviews the model; does not override the model's genuine judgment call.
            (function checkTargetMultipleYield() {
                const peItem = findForcedRow("assum", assumRows, "target_pe", /target\s*p\/?e/i);
                const evItem = findForcedRow("assum", assumRows, "target_ev_ebitda", /target\s*ev\s*\/?\s*ebitda/i);
                const waccItem = findForcedRow("assum", assumRows, "wacc", /\bwacc\b/i);
                const rfItem = findForcedRow("assum", assumRows, "risk_free_rate", /risk.?free/i);
                const wacc = waccItem && typeof waccItem.value === "number" ? waccItem.value : null;
                const riskFree = rfItem && typeof rfItem.value === "number" ? rfItem.value : null;
                const pct = (v) => (v * 100).toFixed(1) + "%";
                if (peItem && typeof peItem.value === "number" && peItem.value > 0) {
                    const earningsYield = 1 / peItem.value;
                    if (wacc !== null && earningsYield < wacc) {
                        console.warn(`[AI Model] target_pe=${peItem.value.toFixed(1)}x implies an earnings yield of ${pct(earningsYield)} — BELOW the model's own WACC of ${pct(wacc)}. May be a genuine call (market pricing in strong growth) but is worth a manual sanity check before trusting the target price.`);
                    } else if (riskFree !== null && earningsYield < riskFree) {
                        console.warn(`[AI Model] target_pe=${peItem.value.toFixed(1)}x implies an earnings yield of ${pct(earningsYield)} — below even the risk-free rate of ${pct(riskFree)}. Verify this is intentional before trusting the target price.`);
                    }
                }
                if (evItem && typeof evItem.value === "number" && evItem.value > 0 && wacc !== null) {
                    const ebitdaYield = 1 / evItem.value;
                    if (ebitdaYield < wacc) {
                        console.warn(`[AI Model] target_ev_ebitda=${evItem.value.toFixed(1)}x implies an EBITDA yield of ${pct(ebitdaYield)} — BELOW the model's own WACC of ${pct(wacc)}. Worth a manual sanity check before trusting the target price.`);
                    }
                }
            })();
            // {A:key} in resolve() below needs to know WHICH assumption rows are now scalar (so it
            // can always point at their single column-B cell, regardless of which year column the
            // referencing formula is being written in) — every dictated formula elsewhere in this
            // file still just writes "{A:wacc}" etc. unchanged, so this has to be transparent at
            // resolve time rather than requiring every formula string to be rewritten.
            const scalarAssumRows = new Set(assumRows.filter(r => r && r.scalar && r._row).map(r => r._row));
            // discount_factor is a per-year DCF flow row (forecast columns only) — force that flag
            // too, in case the model's row omitted it, so historicals aren't computed for it.
            forceMechanicalFormula("val", valRows, "discount_factor", /discount\s*factor/i, "1/(1+{A:wacc})^{N}", { forecastOnly: true });
            // Terminal value MUST be built off a NORMALIZED steady-state terminal cash flow, never
            // the raw final-year explicit-period FCFF — for a capex-heavy company still investing at
            // the end of the forecast window, that raw figure can itself be negative, which would
            // extrapolate into a nonsensical negative terminal value (and from there, a large
            // negative equity value) even though the underlying business is genuinely profitable.
            // This is 100% mechanical (a standard textbook normalization, no company-specific
            // judgment involved), so force it unconditionally rather than trust compliance on a row
            // that has already been observed drifting from its dictated formula.
            forceMechanicalFormula("val", valRows, "tv", /terminal value/i,
                "MAX(0,({LAST:nopat}-{A:terminal_growth}*{LAST:bs.net_working_capital})*(1+{A:terminal_growth})/({A:wacc}-{A:terminal_growth}))");
            // The EV-to-equity bridge must subtract TODAY'S net debt, not a 5-years-out PROJECTED
            // figure — Enterprise Value from a DCF is a value AS OF TODAY (future cash flows
            // discounted back to the present), so pairing it with a future net debt mismatches the
            // time basis. {LAST:bs.net_debt} (final forecast year) has been the dictated formula
            // here; force {CUR:bs.net_debt} (the last HISTORICAL/actual year) instead — pure
            // mechanical fact, zero judgment, same rationale as tv/discount_factor above.
            forceMechanicalFormula("val", valRows, "net_debt_last", /^(less\s*)?net\s*debt$/i, "{CUR:bs.net_debt}");
            // NON-OPERATING ASSETS ADD-BACK — the single biggest systematic understatement in this
            // whole DCF, and the reason value_per_share came back implausibly low for capex-heavy
            // companies specifically. FCFF (nopat + dep − wc_change − capex) charges 100% of capex
            // as a cash outflow, but NOPAT is built off EBIT, which by construction earns nothing
            // on two large asset classes the company is paying for:
            //   • CWIP — assets still UNDER CONSTRUCTION. Not commissioned, so zero revenue and zero
            //     EBIT contribution during the forecast, yet every rupee of their cost has already
            //     been deducted from FCFF. For a long-gestation infra/energy build-out this is
            //     enormous: a real run showed ~₹28,800 Cr of CWIP (11% of total assets) fully
            //     charged against cash flow while contributing nothing to the terminal NOPAT the
            //     terminal value is extrapolated from.
            //   • investments — long-term/non-consolidated investments in associates & JVs. Their
            //     returns show up in OTHER INCOME, which sits BELOW EBIT (ebit = ebitda − dep;
            //     other_income is only added at PBT), so it never enters NOPAT and therefore never
            //     enters FCFF at all.
            // Both are textbook non-operating assets: the standard EV→equity bridge is
            // "EV + non-operating assets − net debt", not "EV − net debt". Omitting them means the
            // model pays for these assets in full and then values them at exactly zero — a
            // one-directional leak that scales with how capital-intensive the company is, which is
            // precisely why the shortfall looked "consistent" rather than random. Valued at the
            // LAST HISTORICAL year ({CUR:}) for the same time-basis reason net_debt_last uses it:
            // a DCF enterprise value is a value AS OF TODAY, so everything in the bridge around it
            // must be today's actual reported balance, never a 5-years-out projection.
            forceMechanicalFormula("val", valRows, "non_operating_assets", /non.?operating\s*assets/i, "{CUR:bs.cwip}+{CUR:bs.investments}", { scalar: true });
            forceMechanicalFormula("val", valRows, "equity_value", /^equity\s*value$/i, "{R:ev}+{R:non_operating_assets}-{R:net_debt_last}", { scalar: true });
            // value_per_share and the two "TARGET PRICE" rows are pure mechanical calculations
            // (zero judgment, exactly one correct formula each) that had actually been coming back
            // BLANK — the model sometimes omits "formula" entirely for a scalar row instead of a
            // guessed number, which leaves writeDataSheet nothing to write at all (unlike a
            // per-year row, a scalar row has no flat-carry-forward fallback to fall back on).
            // Force all three, same as tv/discount_factor/net_debt_last above.
            forceMechanicalFormula("val", valRows, "value_per_share", /value\s*per\s*share/i, "{R:equity_value_discounted}/{CUR:assum.shares_out}", { scalar: true });
            // Neither TARGET PRICE row has a dictated "key" in the prompt (only "target_price",
            // the blended average, does) — located by label instead. The label regexes require
            // "based"/"target" nearby specifically to avoid matching section (1)'s bare per-year
            // "EV/EBITDA" ratio row, which appears earlier in the row list and would otherwise be
            // found first by a looser pattern.
            forceMechanicalFormula("val", valRows, "target_price_pe", /p\/?e.{0,15}based|based.{0,15}p\/?e/i, "{FWD1:pnl.eps}*{A:target_pe}", { scalar: true });
            forceMechanicalFormula("val", valRows, "target_price_ev_ebitda", /ev.?\/?.?ebitda.{0,15}based|based.{0,15}ev.?\/?.?ebitda/i, "(({FWD1:pnl.ebitda}*{A:target_ev_ebitda})-{FWD1:bs.net_debt})/{CUR:assum.shares_out}", { scalar: true });
            // WACC BUILD-UP (section 5) is the one part of this sheet that was never given the
            // same code-level guarantee as everything else here — every other scalar row (tv,
            // discount_factor, net_debt_last, value_per_share, both target-price rows) is
            // force-verified above, but these 7 rows existed only as prose in the prompt with no
            // dictated "key" for any of them, so the model is free to invent its own key/label
            // (e.g. "tax_rate_wacc", "wacc_final" have both been observed) and can just as easily
            // omit "scalar":true or the formula entirely. That's exactly the risk profile that has
            // repeatedly caused silent blanks/zeros elsewhere on this sheet (value_per_share,
            // target_price_pe/ev_ebitda) — a broken WACC BUILD-UP table is worse than most, since
            // it actively misleads a reviewer into thinking WACC itself is malfunctioning even when
            // the real DCF math (which DOES reference {A:wacc} correctly via the force-verified
            // rows above) is fine. Patterns use [\s_]?/.? rather than \s* so they still match if a
            // row's label comes back as raw key text (underscore-joined) instead of prose — see
            // displayLabel's own fallback for why that happens.
            forceMechanicalFormula("val", valRows, "risk_free", /risk.?free/i, "{A:risk_free_rate}", { scalar: true });
            forceMechanicalFormula("val", valRows, "equity_risk_premium", /equity.?risk.?premium/i, "{A:equity_risk_premium}", { scalar: true });
            forceMechanicalFormula("val", valRows, "beta", /beta/i, "{A:beta}", { scalar: true });
            forceMechanicalFormula("val", valRows, "cost_of_equity", /cost.?of.?equity/i, "{A:risk_free_rate}+{A:beta}*{A:equity_risk_premium}", { scalar: true });
            forceMechanicalFormula("val", valRows, "cost_of_debt", /cost.?of.?debt/i, "{A:cost_of_debt}", { scalar: true });
            forceMechanicalFormula("val", valRows, "wacc_tax_rate", /tax.?rate/i, "{A:tax_rate}", { scalar: true });
            forceMechanicalFormula("val", valRows, "wacc_final", /wacc/i, "{A:wacc}", { scalar: true });
            // short_term_borrowings (the revolver/funding plug) and cash (the balancing plug) are
            // BOTH pure mechanical formulas over a FIXED set of guaranteed/required canonical rows —
            // no company-specific judgment involved, so both are forced the same way as discount_factor/
            // tv/roe/roce above. This has real teeth: a run has been observed writing SOME other
            // (well-formed, ref-gate-clean) formula for short_term_borrowings instead of the dictated
            // revolver recurrence, silently disabling the safeguard and letting forecast cash go
            // deeply negative every year even though nothing here ever showed up as a #MISSING error.
            // keepHistorical:true on both — their HISTORICAL columns are real reported figures from
            // the data, not a plug, and must be left untouched; only the FORECAST columns are forced.
            forceMechanicalFormula("bs", bsRows, "short_term_borrowings", /short.?term\s*borrowing/i,
                "{P:short_term_borrowings}+MAX(0,{A:min_cash_pct_revenue}*{S:pnl.net_revenue}-({R:share_capital}+{R:reserves}+{R:long_term_borrowings}+{P:short_term_borrowings}+{R:payables}+{R:provisions}+{R:deferred_tax_liabilities}+{R:other_non_current_liabilities}+{R:other_current_liabilities}-{R:receivables}-{R:inventory}-{R:other_current_assets}-{R:net_fixed_assets}-{R:cwip}-{R:investments}-{R:other_non_current_assets}))",
                { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "cash", /^cash\b/i,
                "{R:total_liabilities_equity}-{R:receivables}-{R:inventory}-{R:other_current_assets}-{R:net_fixed_assets}-{R:cwip}-{R:investments}-{R:other_non_current_assets}",
                { keepHistorical: true });
            // roe/roce are ordinary Balance Sheet rows (historicals ARE meaningful here) — do NOT
            // force forecast_only for these two, unlike discount_factor above.
            forceMechanicalFormula("bs", bsRows, "roe", /^roe$|return on equity/i, "{S:pnl.pat}/{R:equity}");
            forceMechanicalFormula("bs", bsRows, "roce", /^roce$|return on capital employed/i, "{S:pnl.ebit}/{R:capital_employed}");

            // acc_depreciation's historical values MUST equal gross_block − net_fixed_assets for
            // EVERY historical year — Net Fixed Assets is only ever correct in the forecast if this
            // identity holds at the boundary, since net_fixed_assets' forecast formula is
            // "gross_block − acc_depreciation" and its first forecast year rolls forward directly off
            // the LAST historical year's acc_depreciation value. Trusting the model to independently
            // source or compute this figure has actually failed: a real run understated accumulated
            // depreciation by ~₹220,000+ Cr (gross block and net block ARE cleanly reported, but a
            // separate standalone "accumulated depreciation" line often is not, so the model's own
            // guess drifted) — which silently overstated Net Fixed Assets by the same amount in the
            // very first forecast year, manufacturing a phantom multi-lakh-crore asset increase that
            // then forced the short_term_borrowings revolver into an enormous, implausible draw to
            // fund it (even though the revolver formula, capex driver, and working-capital formulas
            // were all completely correct). This is pure arithmetic with zero judgment involved once
            // gross_block and net_fixed_assets are known, so compute it deterministically here rather
            // than trust the model's own figure for this one row.
            (() => {
                const gbRow = symbols.capex && symbols.capex.gross_block;
                const gbItem = gbRow && capexRows.find(r => r && r._row === gbRow);
                const nfaRow = symbols.bs && symbols.bs.net_fixed_assets;
                const nfaItem = nfaRow && bsRows.find(r => r && r._row === nfaRow);
                const adRow = symbols.bs && symbols.bs.acc_depreciation;
                const adItem = adRow && bsRows.find(r => r && r._row === adRow);
                if (!gbItem || !nfaItem || !adItem) {
                    console.warn("[AI Model] Could not locate capex.gross_block / bs.net_fixed_assets / bs.acc_depreciation to derive accumulated depreciation — verify Net Fixed Assets manually.");
                    return;
                }
                if (!Array.isArray(gbItem.historical) || !Array.isArray(nfaItem.historical)) {
                    console.warn("[AI Model] bs.net_fixed_assets or capex.gross_block is missing historical figures — could not derive bs.acc_depreciation; verify Net Fixed Assets ties out at the FY forecast boundary manually.");
                    return;
                }
                // NOTE: this used to be a hardcoded [0,1,2,3,4] (5 years) — a stale leftover from
                // before the historical window was extended to HIST_YEARS (8). With only 5
                // entries, adItem.historical came up short for the 3 MOST RECENT historical years
                // (indices 5-7, since HIST is chronologically ascending) — those years, and every
                // forecast year rolling forward from them, wrote as BLANK cells instead of a real
                // number (this has actually happened — "Less: Accumulated Depreciation" looked
                // like an unrelated linking failure, but net_fixed_assets and acc_depreciation
                // going blank for the last few historical years plus every forecast year was
                // really just this array being too short).
                const derived = Array.from({ length: HIST_YEARS }, (_, i) => {
                    const gb = Number(gbItem.historical[i]), nfa = Number(nfaItem.historical[i]);
                    return (isFinite(gb) && isFinite(nfa)) ? gb - nfa : null;
                });
                if (derived.some(v => v === null)) {
                    console.warn("[AI Model] bs.net_fixed_assets or capex.gross_block historicals were incomplete — could not derive all " + HIST_YEARS + " years of bs.acc_depreciation; verify Net Fixed Assets ties out at the FY forecast boundary manually.");
                    return;
                }
                adItem.historical = derived;
                delete adItem.link; // a derived figure, not a linkable data-sheet cell
                console.log(`[AI Model] Derived bs.acc_depreciation historicals as gross_block − net_fixed_assets (guarantees Net Fixed Assets ties out at the forecast boundary):`, derived.map(v => Math.round(v)));
            })();

            // Every remaining Balance Sheet forecast formula below is, per the prompt above,
            // already fully dictated with zero company-specific judgment — pure mechanical scaling
            // off revenue/days assumptions or a fixed accounting identity. Force each one exactly as
            // specified rather than trust transcription, mirroring short_term_borrowings/cash/roe/
            // roce above (this has real teeth: a flat/frozen working-capital line while revenue
            // grows, or a bare 0 for a row the model omitted a formula for, is a modeling error that
            // otherwise ships silently — see the P&L EPS bug). keepHistorical:true wherever the
            // HISTORICAL column is meant to stay actual linked/entered data (only the FORECAST
            // formula is being forced); omitted wherever the row is a pure derived aggregate with no
            // independent historical source of its own — total_assets, net_debt, etc. are always
            // exactly the sum/difference of their components, in every year including the actuals,
            // not a separately reported figure to transcribe.
            forceMechanicalFormula("bs", bsRows, "receivables", /receivable|debtor/i,
                "{A:debtor_days}/365*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "inventory", /inventory/i,
                "{A:inventory_days}/365*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "other_current_assets", /other current assets/i,
                "{P:other_current_assets}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "net_fixed_assets", /net fixed assets/i,
                "{S:capex.gross_block}-{R:acc_depreciation}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "acc_depreciation", /accumulated depreciation|acc\.?\s*depreciation/i,
                "{P:acc_depreciation}+{S:pnl.depreciation}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "cwip", /capital work.?in.?progress|\bcwip\b/i,
                "{P:cwip}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "investments", /investments/i,
                "{P:investments}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "other_non_current_assets", /other non.?current assets/i,
                "{P:other_non_current_assets}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "payables", /payables|creditors/i,
                "{A:payable_days}/365*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "provisions", /provisions/i,
                "{P:provisions}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "deferred_tax_liabilities", /deferred tax/i,
                "{P:deferred_tax_liabilities}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "other_non_current_liabilities", /other non.?current liabilities/i,
                "{P:other_non_current_liabilities}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "other_current_liabilities", /other current liabilities/i,
                "{P:other_current_liabilities}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", { keepHistorical: true });
            forceMechanicalFormula("bs", bsRows, "reserves", /reserves/i,
                "{P:equity}+{S:pnl.pat}*{A:pat_retention}", { keepHistorical: true });

            // Pure derived aggregates — accounting identities with no independent historical source
            // of their own — forced for ALL columns, not just the forecast.
            forceMechanicalFormula("bs", bsRows, "equity", /^(shareholders'?\s*)?equity$/i,
                "{R:share_capital}+{R:reserves}");
            forceMechanicalFormula("bs", bsRows, "total_assets", /total assets/i,
                "{R:cash}+{R:receivables}+{R:inventory}+{R:other_current_assets}+{R:net_fixed_assets}+{R:cwip}+{R:investments}+{R:other_non_current_assets}");
            forceMechanicalFormula("bs", bsRows, "total_liabilities_equity", /total liabilities.*equity|total liab/i,
                "{R:share_capital}+{R:reserves}+{R:long_term_borrowings}+{R:short_term_borrowings}+{R:payables}+{R:provisions}+{R:deferred_tax_liabilities}+{R:other_non_current_liabilities}+{R:other_current_liabilities}");
            // Must mirror EVERY current-asset/current-liability row this model actually tracks, not
            // just the two "obvious" liability lines — omitting other_current_liabilities (a real,
            // often large current liability, since it's the catch-all counterpart to
            // other_current_assets on the asset side) understates the liabilities side, overstating
            // Net Working Capital. That matters beyond just this row's own display value: the DCF's
            // terminal value formula subtracts terminal_growth × {LAST:bs.net_working_capital} from
            // terminal NOPAT, so an inflated NWC here directly drags down Terminal Value (and from
            // there, Value per Share) — this was traced as a real, material contributor to a
            // suspiciously low DCF output, not just a cosmetic display issue.
            forceMechanicalFormula("bs", bsRows, "net_working_capital", /net working capital/i,
                "{R:receivables}+{R:inventory}+{R:other_current_assets}-{R:payables}-{R:provisions}-{R:other_current_liabilities}");
            forceMechanicalFormula("bs", bsRows, "net_debt", /net debt/i,
                "{R:long_term_borrowings}+{R:short_term_borrowings}-{R:cash}");
            // Capital employed = non-current liabilities + equity (equivalently, total assets minus
            // current liabilities) — the standard ROCE denominator. Deliberately excludes
            // short_term_borrowings: that row is a funding plug/revolver, not invested capital.
            forceMechanicalFormula("bs", bsRows, "capital_employed", /capital employed/i,
                "{R:equity}+{R:long_term_borrowings}+{R:deferred_tax_liabilities}+{R:other_non_current_liabilities}");
            forceMechanicalFormula("bs", bsRows, "net_debt_ebitda", /net debt.*ebitda/i,
                "{R:net_debt}/{S:pnl.ebitda}");
            forceMechanicalFormula("bs", bsRows, "balance_check", /balance check/i,
                "{R:total_assets}-{R:total_liabilities_equity}");

            // ── Tie the HISTORICAL balance sheet to the source's own REPORTED totals ──
            // Even with every catch-all correctly derived from its true parent subtotal (see the
            // double-counting guard above), the historical columns can still fail to balance —
            // because the SOURCE DATA ITSELF doesn't add up. This vendor's top-level "ASSETS" row
            // is the company's real reported total, but the four asset categories displayed beneath
            // it (Tangible & Intangible / CWIP / Other Non-current / Current) sum to LESS than it,
            // and the difference is never broken out as a row anywhere in the table — items like
            // goodwill and intangible-assets-under-development simply have no slot in this schema.
            // The same happens on the other side (e.g. "Share warrants and outstanding" sits under
            // Share Capital as a child but is NOT part of its total, so it vanishes from any
            // category-based sum). Verified against real data across all 8 historical years: the
            // model's own rows summed EXACTLY to the vendor's displayed-category sum on both sides,
            // and the leftover imbalance matched the vendor's own top-level shortfall to the rupee
            // — i.e. every remaining discrepancy originated in the source, not in the linking or
            // derivation logic.
            // A REPORTED balance sheet always balances, and both of this vendor's top-level totals
            // carry the identical figure, so anchoring each side's residual catch-all to its OWN
            // reported total makes the historical columns balance BY CONSTRUCTION — and is honest
            // rather than a fudge: an "Other … Assets" row is exactly where assets the source never
            // itemised belong. Runs AFTER every other Balance Sheet row is final, since it's
            // computed as (reported total − every other row on that side).
            (() => {
                // Scoped to Annual Data only (where the Balance Sheet lives) — Key Financials can
                // carry a similarly-named row on a different basis, and findEntryByLabel's own
                // KF-first scan order would silently prefer it. Still merges across year-chunks the
                // same way findEntryByLabel does, so a label split over several "Parameter" blocks
                // is recovered rather than half-missing.
                const findAnnualEntryByLabel = (pattern) => {
                    let bestRank = Infinity;
                    const matches = [];
                    for (const raw of rawAnnual) {
                        if (!(pattern.test(raw.label.trim()) || pattern.test(stripUnit(raw.label)))) continue;
                        if (raw._rank < bestRank) bestRank = raw._rank;
                        matches.push(raw);
                    }
                    if (!matches.length) return null;
                    const chosen = matches.filter(m => m._rank === bestRank);
                    const valByFY = {};
                    for (const m of chosen) for (const fy in m.valByFY) if (valByFY[fy] === undefined) valByFY[fy] = m.valByFY[fy];
                    return { label: chosen[0].label, valByFY };
                };
                const itemFor = (k) => { const r = symbols.bs && symbols.bs[k]; return r ? bsRows.find(x => x && x._row === r) : null; };
                // Mirrors writeDataSheet's own historical-cell logic exactly: a row autoLink
                // resolves gets a LIVE link for every year that link's column map covers, and falls
                // back to the row's plain "historical" number only for years it doesn't — so the
                // figure computed here is the one the sheet will actually display.
                const histValues = (item) => {
                    if (!item) return null;
                    const fmt = item.fmt || "#,##0";
                    const le = (/%/.test(fmt) || item.noAutoLink) ? null : autoLink(item);
                    return HIST.map((fy, i) => {
                        if (le && le.colByFY && le.colByFY[fy] != null && le.valByFY[fy] != null) return le.valByFY[fy];
                        const h = Array.isArray(item.historical) ? item.historical[i] : null;
                        return (typeof h === "number" && isFinite(h)) ? h : null;
                    });
                };
                const SIDES = [
                    {
                        label: "assets",
                        totalRe: /^(assets|total\s*assets)$/i,
                        keys: ["cash", "receivables", "inventory", "other_current_assets", "net_fixed_assets", "cwip", "investments", "other_non_current_assets"],
                        absorber: "other_non_current_assets",
                    },
                    {
                        label: "liabilities & equity",
                        totalRe: /^(equity\s*(and|&)\s*liabilities|total\s*liabilities\s*(and|&)\s*equity)$/i,
                        keys: ["share_capital", "reserves", "long_term_borrowings", "short_term_borrowings", "payables", "provisions", "deferred_tax_liabilities", "other_non_current_liabilities", "other_current_liabilities"],
                        absorber: "other_non_current_liabilities",
                    },
                ];
                for (const side of SIDES) {
                    const reported = findAnnualEntryByLabel(side.totalRe);
                    const absorber = itemFor(side.absorber);
                    if (!reported || !absorber) {
                        console.warn(`[AI Model] Could not tie the historical ${side.label} to the source's reported total (${!reported ? "reported total row not found in Annual Data" : "bs." + side.absorber + " row missing"}) — the historical balance check may not reach ~0.`);
                        continue;
                    }
                    const valsByKey = {};
                    for (const k of side.keys) valsByKey[k] = histValues(itemFor(k));
                    const absorberVals = valsByKey[side.absorber];
                    if (!absorberVals) continue;
                    const adjusted = absorberVals.slice();
                    const bigResiduals = [];
                    HIST.forEach((fy, i) => {
                        const rep = reported.valByFY[fy];
                        if (rep == null) return;
                        let sum = 0, complete = true;
                        for (const k of side.keys) {
                            const v = valsByKey[k] && valsByKey[k][i];
                            if (v == null) { if (k !== side.absorber) complete = false; continue; }
                            sum += v;
                        }
                        if (!complete) return; // a component is unknown this year — a "residual" here would be meaningless
                        const residual = rep - sum;
                        if (!isFinite(residual)) return;
                        adjusted[i] = (absorberVals[i] || 0) + residual;
                        if (Math.abs(residual) / Math.max(Math.abs(rep), 1) > 0.02) bigResiduals.push(`FY${fy}: ${Math.round(residual).toLocaleString()}`);
                    });
                    absorber.historical = adjusted;
                    delete absorber.link; // now a derived residual, not a single source cell
                    absorber.noAutoLink = true; // never let autoLink's value-coincidence matching override this
                    absorber.source = `Reported ${side.label} total minus every other ${side.label} line on this sheet (absorbs items the source reports in its total but never itemises)`;
                    if (bigResiduals.length) {
                        console.warn(`[AI Model] bs.${side.absorber} absorbed a residual >2% of the reported ${side.label} total so the historical columns tie out: ${bigResiduals.join(", ")}. This is what the SOURCE reports in its own top-level total but never breaks out as a row (typically goodwill / intangible assets under development on the asset side, share warrants on the equity side) — sanity-check that it looks plausible for this company.`);
                    }
                }
            })();

            // Cross-check: do the ACTUAL historical figures used sum to the reported Total Assets?
            // A general symptom of the "total row + its own indented children both linked"
            // double-counting bug described in the prompt's TOTALS vs LINE ITEMS section above
            // (e.g. linking BOTH a rolled-up "Other Current Assets" total AND its own "Trade
            // Receivables"/"Cash" children counts the same money twice). This can't fully detect or
            // fix the bug on its own (that needs the source's real parent/child hierarchy, which
            // isn't preserved this far downstream) — it's a concrete, automatic red flag using the
            // SAME historical figures actually written to the sheet, mirroring the self-check the
            // prompt already asks the model to do by hand.
            (() => {
                const ASSET_KEYS = ["cash", "receivables", "inventory", "other_current_assets", "net_fixed_assets", "cwip", "investments", "other_non_current_assets"];
                const items = ASSET_KEYS.map(k => {
                    const row = symbols.bs && symbols.bs[k];
                    return row && bsRows.find(r => r && r._row === row);
                });
                if (items.some(it => !it || !Array.isArray(it.historical))) return; // one or more asset rows missing/unlinked — nothing reliable to check
                const reportedTotalAssets = findEntryByLabel(/^total\s*assets$/i);
                if (!reportedTotalAssets) return;
                const mismatches = [];
                HIST.forEach((fy, i) => {
                    const reported = reportedTotalAssets.valByFY[fy];
                    if (reported == null) return;
                    const summed = items.reduce((s, it) => s + (Number(it.historical[i]) || 0), 0);
                    const diffPct = Math.abs(summed - reported) / Math.max(Math.abs(reported), 1);
                    if (diffPct > 0.02) mismatches.push(`FY${fy}: summed asset lines = ${Math.round(summed)}, reported Total Assets = ${Math.round(reported)} (${(diffPct * 100).toFixed(1)}% off)`);
                });
                if (mismatches.length) {
                    console.warn(`[AI Model] Balance Sheet historical asset lines don't sum to the reported Total Assets for ${mismatches.length} year(s) — likely a total-row-plus-its-own-children double-count (see TOTALS vs LINE ITEMS in the prompt) or a missing/mislinked line item. Verify manually:`, mismatches);
                }
            })();

            // Canonical-key repair pass. If a generated sheet used a near-match
            // instead of the required CANON key, alias it before formulas resolve.
            const CANON_KEYS = {
                op: ["net_revenue", "ebitda"],
                pnl: ["net_revenue", "ebitda", "depreciation", "ebit", "interest", "tax", "pat", "eps"],
                bs: ["equity", "net_debt", "capital_employed", "net_working_capital", "total_assets", "total_liabilities_equity"],
                capex: ["capex", "ocf", "fcf", "gross_block", "depreciation", "wc_change"]
            };
            // Narrower list used only to force the CAGR column (see below): excludes rows that are
            // period CHANGES rather than levels (wc_change, net_working_capital can be near-zero or
            // flip sign, making a 5-year CAGR meaningless/undefined) — those keep whatever "cagr" the
            // model itself set, if any.
            const CAGR_KEYS = {
                op: CANON_KEYS.op,
                pnl: CANON_KEYS.pnl,
                bs: ["equity", "net_debt", "capital_employed", "total_assets", "total_liabilities_equity"],
                capex: ["capex", "ocf", "fcf", "gross_block", "depreciation"]
            };
            const reconcileCanonKeys = (sheetKey, rows) => {
                const need = CANON_KEYS[sheetKey] || [];
                for (const k of need) {
                    if (symbols[sheetKey][k]) continue;
                    const target = k.replace(/_/g, "");
                    const match = rows.find(r => r && !r.section && r._row && (
                        (r.key && r.key.replace(/_/g, "").toLowerCase() === target) ||
                        (r.label && r.label.replace(/[^a-z]/gi, "").toLowerCase().includes(target))
                    ));
                    if (match) {
                        symbols[sheetKey][k] = match._row;
                        console.warn(`[AI Model] Aliased ${sheetKey}.${k} -> row ${match._row} ("${match.label || match.key}")`);
                    } else {
                        console.warn(`[AI Model] Canonical key ${sheetKey}.${k} not found on any row - {S:${sheetKey}.${k}} will resolve to 0.`);
                    }
                }
            };
            reconcileCanonKeys("op", opRows);
            reconcileCanonKeys("pnl", pnlRows);
            reconcileCanonKeys("bs", bsRows);
            reconcileCanonKeys("capex", capexRows);

            // Force CAGR on canonical absolute-value rows (revenue, EBITDA, PAT, EPS, equity, capex,
            // FCF, etc.) regardless of whether the model remembered to set "cagr":true on them — the
            // model is inconsistent about this flag since the prompt leaves it to its own judgment,
            // but these rows always warrant a growth column. Non-canonical rows still rely on
            // whatever "cagr" the model set for them.
            const forceCanonCagr = (sheetKey, rows) => {
                for (const k of CAGR_KEYS[sheetKey] || []) {
                    const row = symbols[sheetKey][k];
                    const item = row && rows.find(r => r && r._row === row);
                    if (item) item.cagr = true;
                }
            };
            forceCanonCagr("op", opRows);
            forceCanonCagr("pnl", pnlRows);
            forceCanonCagr("bs", bsRows);
            forceCanonCagr("capex", capexRows);

            // These rows are explicitly instructed to be PURE cross-sheet formulas spanning every
            // column (not just forecast) — e.g. change in working capital should equal the Balance
            // Sheet's own delta, not a guessed number. Despite that instruction, the model has been
            // observed attaching its own "historical" anyway (round-number guesses for wc_change,
            // or PAT copied in as a stand-in for OCF) — which silently overrides the formula for
            // historical columns since writeDataSheet prefers "historical" when present. Strip it
            // deterministically rather than trust the model to omit it as asked.
            const FORMULA_ONLY_KEYS = { capex: ["wc_change", "ocf", "depreciation"] };
            const forceFormulaOnly = (sheetKey, rows) => {
                for (const k of FORMULA_ONLY_KEYS[sheetKey] || []) {
                    const row = symbols[sheetKey][k];
                    const item = row && rows.find(r => r && r._row === row);
                    if (item && item.formula && (item.historical || item.link)) {
                        console.warn(`[AI Model] ${sheetKey}.${k} had a guessed "historical"/"link" alongside its required formula — discarding it so every column (including historicals) computes from the formula instead.`);
                        delete item.historical;
                        delete item.link;
                    }
                }
            };
            forceFormulaOnly("capex", capexRows);

            // ── Driver-fallback pass (Tier 1, fully deterministic — no LLM) ─────────────────────
            // Many rows across every sheet legitimately have real historicals but no forecast
            // formula — either by design (the Assumptions sheet's reserves/debt/WC-balance rows are
            // meant to be a flat REFERENCE snapshot, since the Balance Sheet sheet does the actual
            // forecasting) or by omission (a sheet forgot one of its dictated WC/asset formulas
            // despite the prompt requiring it). Today every one of these falls through to
            // writeDataSheet's flat-carry-forward safety net — a real "true zero" is worse, but a
            // frozen level is still economically wrong for anything that should scale with the
            // business, and it's what made the Assumptions-sheet snapshot rows look "broken" in
            // review even though nothing downstream reads them. Before accepting a flat carry, try
            // two purely mechanical repairs, in order:
            //   1a. NAMED DRIVER MATCH — if the row's key/label matches a recognised WC/asset
            //       category (receivables/inventory/payables/cash/capex) AND the matching
            //       Assumptions-sheet driver (debtor_days, inventory_days, payable_days,
            //       min_cash_pct_revenue, capex_pct_revenue) actually exists, wire the row to that
            //       driver directly — the same days-based/%-of-revenue formula already dictated
            //       for the Balance Sheet's own equivalents.
            //   1b. SELF-RATIO FALLBACK — otherwise, hold the row's OWN last historical ratio-to-
            //       revenue constant ("{P:key}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}"), the exact
            //       same safe pattern already dictated for every "other_*" catch-all row elsewhere
            //       in these prompts. Requires the row to have a "key" (needed to self-reference).
            // Financing/capital-structure rows (borrowings, share capital, reserves/equity) are
            // deliberately EXCLUDED from both — those have no natural revenue relationship and are
            // meant to stay flat absent a real repayment/issuance schedule (see the Balance Sheet
            // prompt's "flat/no formula is acceptable" rows). Both repairs only ever ASSIGN a
            // formula string in the SAME {A:}/{R:}/{P:}/{S:}/{PS:} grammar every other row uses, so
            // whatever they produce still passes through the reference-integrity gate below before
            // anything is written — this pass cannot itself introduce a bad reference undetected.
            const DRIVER_EXCLUDE_RE = /borrowing|debenture|share_capital|other_equity|^capital$|^equity$|^reserves$/i;
            const DRIVER_PATTERNS = [
                { re: /receivable|debtor/i, need: "debtor_days", formula: "{A:debtor_days}/365*{S:pnl.net_revenue}" },
                { re: /inventor/i, need: "inventory_days", formula: "{A:inventory_days}/365*{S:pnl.net_revenue}" },
                { re: /payable|creditor/i, need: "payable_days", formula: "{A:payable_days}/365*{S:pnl.net_revenue}" },
                { re: /cash/i, need: "min_cash_pct_revenue", formula: "{A:min_cash_pct_revenue}*{S:pnl.net_revenue}" },
                // CWIP is deliberately NOT matched here. It reads like a capex row but is a STOCK
                // (the balance of assets still under construction), not the annual capex FLOW — and
                // this pattern would hand it capex_pct_revenue*revenue, i.e. the IDENTICAL formula
                // the real capex row gets. That has actually happened: the Assumptions sheet showed
                // "Capital Work in Progress" and "Capex (Additions to Gross Block)" as the exact same
                // number every forecast year, while the same capex ALSO rolled into gross block in
                // full — capitalising the same spend twice across the two sheets. CWIP keeps its own
                // revenue-scaling roll-forward instead (see the forceMechanicalFormula for bs.cwip).
                { re: /capex/i, need: "capex_pct_revenue", formula: "{A:capex_pct_revenue}*{S:pnl.net_revenue}" },
            ];
            const driverFallbackApplied = [];
            const driverFallbackResidual = []; // rows neither repair could reach -> Tier 2 candidates
            const runDriverFallback = (sheetKey, rows) => {
                for (const item of rows) {
                    if (!item || item.section || item.formula || Array.isArray(item.forecast) || item.forecast_only) continue;
                    const hasHist = !!item.link || (Array.isArray(item.historical) && item.historical.some(v => v !== null && v !== undefined && v !== ""));
                    if (!hasHist) continue; // nothing to drive forward — writeDataSheet's own logic handles this (e.g. blank row)
                    const idKey = item.key || "";
                    if (DRIVER_EXCLUDE_RE.test(idKey)) continue; // financing/capital — intentionally stays flat

                    if (idKey) {
                        const haystack = `${idKey} ${item.label || ""}`;
                        const pattern = DRIVER_PATTERNS.find(p => p.re.test(haystack) && symbols.assum[p.need]);
                        if (pattern) {
                            item.formula = pattern.formula;
                            driverFallbackApplied.push(`${sheetKey}.${idKey} -> {A:${pattern.need}}`);
                            continue;
                        }
                        if (symbols.pnl && symbols.pnl.net_revenue) {
                            item.formula = `{P:${idKey}}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}`;
                            driverFallbackApplied.push(`${sheetKey}.${idKey} -> revenue-ratio (own historical ratio held constant)`);
                            continue;
                        }
                    }
                    driverFallbackResidual.push({ sheetKey, item }); // no key, or net_revenue unavailable
                }
            };
            const ALL_ROWS_BY_SHEET = { assum: assumRows, op: opRows, pnl: pnlRows, bs: bsRows, capex: capexRows, val: valRows };
            for (const sk of Object.keys(ALL_ROWS_BY_SHEET)) runDriverFallback(sk, ALL_ROWS_BY_SHEET[sk]);
            if (driverFallbackApplied.length) console.log(`[AI Model] Driver-fallback wired ${driverFallbackApplied.length} row(s) that had historicals but no forecast formula, using a matched driver instead of a flat carry-forward:`, driverFallbackApplied);

            // ── Residual growth-rate fallback (Tier 2, LLM-assisted, heavily constrained) ───────
            // Whatever's LEFT after Tier 1 (rows with no "key" to self-reference, or the rare case
            // where pnl.net_revenue itself is unavailable) is genuinely driver-less — there's no
            // mechanical repair left to try. Rather than accept a flat carry-forward outright, make
            // ONE lightweight batched Flash call and let it suggest a growth rate — but its output
            // is constrained to a SINGLE bounded decimal per row (never a free-floating level
            // value), which is then hard-clamped in code regardless of what it returns and applied
            // as literal computed "forecast" VALUES (not a formula) — so it needs no {A:}/{R:}/{S:}
            // placeholder at all and cannot introduce a bad reference, a circular reference, or a
            // #MISSING. A parse/call failure just leaves these rows on the existing flat-carry
            // safety net — this is a pure enhancement, never a new way for the model build to fail.
            if (driverFallbackResidual.length) {
                try {
                    const residualDesc = driverFallbackResidual.map(({ sheetKey, item }, i) =>
                        `${i}: sheet=${sheetKey}, label="${item.label || item.key || "unnamed"}", historical=[${(Array.isArray(item.historical) ? item.historical : []).map(v => (v === null || v === undefined || v === "") ? "?" : v).join(",")}]`
                    ).join("\n");
                    const tier2Prompt = `For each financial-model line item below (real historical actuals, no forecast driver could be matched), suggest ONE constant year-over-year growth rate to compound its LAST historical value forward for 5 forecast years. Base it on the row's own historical trend where a trend is visible in the numbers given; otherwise use a conservative low-single-digit rate. Do NOT invent or return any absolute rupee figures — ONLY a growth rate decimal.\n\n${residualDesc}\n\nReturn ONLY JSON: {"rates":[{"i":0,"growth_rate":n}, ...]} — exactly one entry per row index above, growth_rate as a decimal (e.g. 0.05 for 5%, -0.02 for -2%), bounded within -0.3 to 0.5.`;
                    const tier2Result = await callLLM("Residual driver-less rows (Tier 2 fallback)", tier2Prompt, STATIC_MODEL, 2000);
                    const rates = Array.isArray(tier2Result.rates) ? tier2Result.rates : [];
                    let tier2Applied = 0;
                    const tier2Log = [];
                    for (const r of rates) {
                        const entry = driverFallbackResidual[r && r.i];
                        if (!entry) continue;
                        const { sheetKey, item } = entry;
                        let g = Number(r.growth_rate);
                        if (!isFinite(g)) continue;
                        g = Math.max(-0.3, Math.min(0.5, g)); // hard clamp regardless of what the model said
                        const histArr = Array.isArray(item.historical) ? item.historical : [];
                        const lastHist = [...histArr].reverse().find(v => typeof v === "number");
                        if (lastHist == null) continue; // nothing to compound from
                        let v = lastHist;
                        item.forecast = [0, 1, 2, 3, 4].map(() => (v = v * (1 + g)));
                        tier2Applied++;
                        tier2Log.push(`${sheetKey}.${item.key || item.label} -> ${(g * 100).toFixed(1)}%/yr`);
                    }
                    if (tier2Applied) console.log(`[AI Model] Tier-2 (${STATIC_MODEL}) growth-rate fallback wired ${tier2Applied} residual row(s) with no deterministic driver match:`, tier2Log);
                } catch (e) {
                    console.warn("[AI Model] Tier-2 driver-less-row fallback call failed (non-fatal — those rows remain on flat carry-forward):", e.message);
                }
            }

            // ── Reference-integrity gate — runs after every sheet is planned and canonical
            //    keys are reconciled, but BEFORE anything is written. The sheets are generated
            //    in parallel, blind to each other, so a sheet can reference an {A:}/{S:}/{PS:}/
            //    {R:}/{P:} key that NO sheet actually emitted (a hallucinated key). Until now
            //    such a reference produced a "#MISSING" sentinel that still LANDED in the cell —
            //    and on a canonical row (e.g. op.ebitda) it poisoned every downstream sheet.
            //    Here we catch those references deterministically, BEFORE write, and repair them:
            //      • a canonical row with a known-safe rewrite gets the correct dictated formula;
            //      • any other offending row has its formula stripped, so writeDataSheet's
            //        flat-carry-forward fallback fills it with the last actual — an honest number,
            //        never a "#MISSING" string. No extra LLM calls; deterministic; cannot loop.
            const REF_RE = /\{(SUM|LAST|FWD1|CUR|SCALAR|PS|S|A|R|P):([^}]+)\}/g;
            const refResolves = (sheetKey, kind, body) => {
                if (kind === "A") return !!symbols.assum[body];
                if (kind === "R" || kind === "P") return !!(symbols[sheetKey] && symbols[sheetKey][body]);
                if (kind === "SUM" || kind === "LAST" || kind === "FWD1" || kind === "CUR" || kind === "SCALAR") {
                    const dot = body.indexOf(".");
                    if (dot > 0) { const sk = body.slice(0, dot), rk = body.slice(dot + 1); return !!(SHEETS[sk] && symbols[sk] && symbols[sk][rk]); }
                    return !!(symbols[sheetKey] && symbols[sheetKey][body]);
                }
                const dot = body.indexOf(".");
                const sk = body.slice(0, dot), rk = body.slice(dot + 1);
                return !!(SHEETS[sk] && symbols[sk] && symbols[sk][rk]);
            };
            const badRefsIn = (formula, sheetKey) => {
                const bad = [];
                if (!formula) return bad;
                REF_RE.lastIndex = 0;
                let m;
                while ((m = REF_RE.exec(String(formula))) !== null) {
                    if (!refResolves(sheetKey, m[1], m[2])) bad.push(`${m[1]}:${m[2]}`);
                }
                return bad;
            };
            // Known-safe canonical rewrites, built only from keys guaranteed to exist. op.ebitda
            // is the recurring offender (the model invents {A:ebitda_growth}); revenue × margin is
            // identical to summing the segments since one company-wide margin applies to all.
            const SAFE_REWRITE = {
                op: { ebitda: "{R:net_revenue}*{A:ebitda_margin}", net_revenue: "{P:net_revenue}*(1+{A:revenue_growth})" }
            };
            // Cross-sheet key ALIASES — a sheet generated blind to another sheet's exact row keys
            // sometimes guesses a plausible-but-wrong name (e.g. "pnl.revenue" instead of the real
            // "pnl.net_revenue"). Only list a mapping here once its target is VERIFIED against this
            // file's own CANON_KEYS/CANON text — a wrong guess on the right-hand side would just
            // trade one silent bug for another. Deliberately does NOT include "pnl.cogs" or
            // "pnl.dividends": neither has a canonical/mandated key anywhere to alias to, so those
            // fall through to the flatten path instead of guessing.
            const S_KEY_ALIAS = {
                "pnl.revenue": "pnl.net_revenue",
                "pnl.depreciation_charge": "pnl.depreciation",
                "capex.capex_total": "capex.capex",
                "capex.total_capex": "capex.capex",
                "cf.fcf": "capex.fcf" // there is no "cf" sheet at all — FCF lives on capex
            };
            const tryAliasRewrite = (formula, sheetKey) => {
                let changed = false;
                const out = String(formula).replace(/\{(SUM|LAST|FWD1|PS|S):([^}]+)\}/g, (m, kind, body) => {
                    const alias = S_KEY_ALIAS[body];
                    if (alias && refResolves(sheetKey, kind, alias)) { changed = true; return `{${kind}:${alias}}`; }
                    return m;
                });
                return changed ? out : null;
            };
            const gateAliased = [], gateRepaired = [], gateFlattened = [];
            const runRefGate = (sheetKey, rows) => {
                for (const item of rows) {
                    if (!item || item.section || !item.formula) continue;
                    if (!badRefsIn(item.formula, sheetKey).length) continue;
                    const aliased = tryAliasRewrite(item.formula, sheetKey);
                    if (aliased && !badRefsIn(aliased, sheetKey).length) {
                        item.formula = aliased;
                        gateAliased.push(`${sheetKey}.${item.key || item.label || item._row}`);
                        continue;
                    }
                    const canonKey = SAFE_REWRITE[sheetKey] &&
                        Object.keys(SAFE_REWRITE[sheetKey]).find(k => symbols[sheetKey][k] === item._row);
                    if (canonKey && !badRefsIn(SAFE_REWRITE[sheetKey][canonKey], sheetKey).length) {
                        item.formula = SAFE_REWRITE[sheetKey][canonKey];
                        gateRepaired.push(`${sheetKey}.${canonKey}`);
                        continue;
                    }
                    // No safe rewrite -> drop the formula so the flat-carry-forward fallback
                    // in writeDataSheet fills it with the last actual instead of a #MISSING.
                    delete item.formula;
                    gateFlattened.push(`${sheetKey}.${item.key || item.label || item._row}`);
                }
            };
            runRefGate("op", opRows);
            runRefGate("pnl", pnlRows);
            runRefGate("bs", bsRows);
            runRefGate("capex", capexRows);
            runRefGate("val", valRows);
            runRefGate("assum", assumRows);
            if (gateAliased.length) console.warn(`[AI Model] Ref-gate aliased ${gateAliased.length} row(s) that referenced a plausible-but-wrong key name to the real one:`, gateAliased);
            if (gateRepaired.length) console.warn(`[AI Model] Ref-gate repaired ${gateRepaired.length} canonical row(s) that referenced a missing key, via a safe rewrite:`, gateRepaired);
            if (gateFlattened.length) console.warn(`[AI Model] Ref-gate stripped ${gateFlattened.length} row(s) referencing a missing key — flat-carried the last actual instead of writing #MISSING (review; usually a hallucinated {A:...} growth key on a disclosure KPI row):`, gateFlattened);

            // ── TEMP DIAGNOSTIC: MODEL INPUTS REPORT — companion to the PROVENANCE REPORT right
            // below: prints exactly what was fed to the model, so a suspicious output value (e.g. an
            // inventory/payable-days figure that looks made up) can be checked directly against the
            // real source text instead of just trusted at face value. Remove alongside the
            // provenance report once the current round of hallucination/mismatch bugs is tracked
            // down. Grouped/collapsed in the console (click to expand) since some of these blocks are
            // long — expand only the ones you actually need to check.
            try {
                console.groupCollapsed(`[AI Model][INPUTS] Everything fed to the model for ${companyName} (fincode ${fincode}) — click to expand`);
                console.groupCollapsed("Key Financials (downloaded sheet)"); console.log(financialText || "(empty)"); console.groupEnd();
                if (annualText) { console.groupCollapsed("Annual Data"); console.log(annualText); console.groupEnd(); }
                if (quarterlyText) { console.groupCollapsed("Quarterly Data"); console.log(quarterlyText); console.groupEnd(); }
                if (operationalText) { console.groupCollapsed("Operational Data"); console.log(operationalText); console.groupEnd(); }
                if (earningCallText) { console.groupCollapsed("Earning Call Insights"); console.log(earningCallText); console.groupEnd(); }
                if (qnaText) { console.groupCollapsed("Earnings Call Q&A"); console.log(qnaText); console.groupEnd(); }
                if (brokerReportText) { console.groupCollapsed("Broker / Analyst Reports"); console.log(brokerReportText); console.groupEnd(); }
                if (mgmtInterviewText) { console.groupCollapsed("Management Interviews"); console.log(mgmtInterviewText); console.groupEnd(); }
                if (orderBookText) { console.groupCollapsed("Order Book"); console.log(orderBookText); console.groupEnd(); }
                // These specific facts were fetched live and given to the model DIRECTLY (not left for
                // it to find/derive itself) — compare these exact numbers against what actually ended
                // up in the sheet (see the PROVENANCE REPORT and any "model originally had X -> forced
                // to Y" lines above) to tell "the model ignored a given fact" apart from "the model
                // correctly derived this itself from the data blocks above."
                console.log("Given facts (fetched live, told to the model directly):", {
                    current_price: latestPrice,
                    face_value: (basicInfo && basicInfo.faceValue) ?? null,
                    diluted_shares_outstanding_fetched_but_NOT_used_for_shares_out: (basicInfo && basicInfo.dilutedShares) ?? null,
                    market_cap_used_for_shares_out_calc: marketCapUsed,
                    price_used_for_shares_out_calc: priceUsed,
                    shares_out_computed: sharesOutComputed,
                    shares_out_computed_fy: sharesOutComputedFY,
                });
                console.groupEnd();
            } catch (e) {
                console.warn("[AI Model][INPUTS] Report failed to build:", e.message);
            }

            // ── TEMP DIAGNOSTIC: DATA PROVENANCE REPORT — remove once the current round of
            // hallucination/mismatch bugs (inventory_days/payable_days looking made up, face_value
            // not matching what it was given) is tracked down. For every row, reports whether its
            // value is a verified LIVE LINK to real downloaded data, a self-reported "source"
            // citation (the model claiming responsibility for a judgment call), or NEITHER — a row
            // with no link and no citation is the closest signal available to "this may have been
            // invented," though it's not proof either way (a row can cite a source and still be
            // wrong, or some keys are explicitly told they don't need a citation — debtor/inventory/
            // payable_days among them — so use the printed VALUE itself, not just this flag, to
            // judge plausibility). This can't catch a row that silently contradicts an EARLIER
            // instruction either — that's what logForcedCorrection above is for instead (it already
            // caught exactly that for face_value/shares_out/current_price).
            try {
                const buildProvenanceRows = (sheetKey, rows) => (rows || []).filter(r => r && !r.section).map((item) => {
                    const fmt = item.fmt || "#,##0";
                    const linked = !/%/.test(fmt) && !!autoLink(item);
                    const firstVal = item.scalar ? item.value
                        : (Array.isArray(item.historical) ? item.historical[0]
                            : (Array.isArray(item.forecast) ? item.forecast[0] : undefined));
                    return {
                        sheet: sheetKey,
                        key: item.key || "",
                        label: item.label || "",
                        value: firstVal,
                        source: item.source || "",
                        provenance: linked ? "LINKED (real cell)" : (item.source ? "SELF-REPORTED (cited)" : "NO LINK / NO CITATION"),
                    };
                });
                const allProvenance = [
                    ...buildProvenanceRows("assum", assumRows),
                    ...buildProvenanceRows("op", opRows),
                    ...buildProvenanceRows("pnl", pnlRows),
                    ...buildProvenanceRows("bs", bsRows),
                    ...buildProvenanceRows("capex", capexRows),
                    ...buildProvenanceRows("val", valRows),
                ];
                console.log("[AI Model][PROVENANCE] Data provenance for every row across all sheets — review VALUE for plausibility, not just the flag (some keys are legitimately never cited):");
                console.table(allProvenance);
                const unexplained = allProvenance.filter(r => r.provenance === "NO LINK / NO CITATION");
                if (unexplained.length) {
                    console.warn(`[AI Model][PROVENANCE] ${unexplained.length} row(s) have neither a live data link nor a source citation — review these first for possible hallucination:`, unexplained.map(r => `${r.sheet}.${r.key || r.label} = ${r.value}`));
                }
            } catch (e) {
                console.warn("[AI Model][PROVENANCE] Report failed to build:", e.message);
            }

            aiStatus("Writing model to Excel...");

            // Populated after the model is written and recalculated (see the 5-check block near
            // the end of the Excel.run below) — rendered as a dismissible popup once this whole
            // function returns. Declared out here so it survives past the Excel.run closure.
            let modelCheckResults = [];

            // ── Step 4: Write the interlinked, multi-sheet model into Excel ──
            await Excel.run(async (context) => {
                const wb = context.workbook;
                wb.worksheets.load("items/name");
                await context.sync();
                const existing = wb.worksheets.items.map(s => s.name);

                // Clean rebuild: drop prior model sheets (and the old single-sheet model).
                const FM_NAMES = Object.values(SHEETS);
                for (const nm of [...FM_NAMES, "Financial Model"]) {
                    if (existing.includes(nm)) wb.worksheets.getItem(nm).delete();
                }
                await context.sync();

                // Create the model sheets, grouped via a common green tab colour + "FM " prefix.
                const TAB_COLOR = "#38761d";
                const MODEL_FONT = "Arial"; // standard modelling convention, not the Excel-365 default (Aptos Narrow)
                const made = {};
                for (const key of ["assum", "op", "pnl", "bs", "capex", "val", "summary"]) {
                    const sh = wb.worksheets.add(SHEETS[key]);
                    sh.tabColor = TAB_COLOR;
                    sh.getRange("A1:M300").format.font.name = MODEL_FONT;
                    made[key] = sh;
                }
                await context.sync();

                // Standard accounting convention: parentheses for negatives, "-" for zero (instead of
                // a bare minus sign / blank cell). Applied in CODE rather than left to the model's own
                // "fmt" choice, since — like the CAGR flag — the model is inconsistent about honouring
                // formatting conventions if only asked for in the prompt; wrapping every fmt string
                // here guarantees it regardless of what the model emits. No currency symbol is added
                // since every sheet is denominated in Rs Crores, not $.
                const normalizeFmt = (fmt) => {
                    const f = String(fmt || "#,##0").trim();
                    return f.includes(";") ? f : `${f};(${f});"-"`;
                };

                // Fallback display label for a row whose own "label" is missing, or IS its raw
                // snake_case key (the row schema only dictates a "key" for some rows, leaving
                // "label" to the model's own judgment for the rest — e.g. section (1) of the
                // Valuation sheet dictates key=eps/key=bvps but not a key for PER/P-BV/EV-EBITDA/
                // RoE/RoCE, and the model has been observed writing the raw key text itself as the
                // "label" too, e.g. "p_bv" or "net_debt_ebitda" instead of "P/BV (x)"). Converts
                // snake_case to Title Case, keeping recognised finance acronyms upper-cased.
                const FINANCE_ACRONYMS = new Set(["eps", "bvps", "per", "pe", "pb", "ev", "ebit", "ebitda", "roe", "roce", "dcf", "fcff", "fcf", "ocf", "nopat", "wacc", "pat", "pbt", "cagr", "cwip", "tv", "pv", "kpi", "ntm", "fy", "eq"]);
                const humanizeLabel = (raw) => String(raw || "")
                    .replace(/_/g, " ")
                    .split(" ")
                    .filter(Boolean)
                    .map(w => (FINANCE_ACRONYMS.has(w.toLowerCase()) ? w.toUpperCase() : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()))
                    .join(" ");
                // A "key-like" label is bare lowercase/underscore text with no spaces, punctuation,
                // or unit annotation — the kind of thing that's clearly a key, not prose meant to be
                // read (a real label always has at least a space or a unit like "(Rs Cr)"/"(x)").
                const displayLabel = (item) => {
                    const lbl = item && item.label;
                    if (lbl && !/^[a-z][a-z0-9_]*$/.test(String(lbl).trim())) return lbl;
                    const source = lbl || item.key;
                    return source ? humanizeLabel(source) : "";
                };

                // ── Formula placeholder resolver ({A}/{R}/{P}/{S}/{PS}) -> real Excel addresses ──
                const colAt = (i) => ALL_COLS[i];
                const prevColAt = (i) => ALL_COLS[i - 1];
                const refMissing = [];
                const missingForecast = [];
                // Diagnostic only — counts how many historical cells per sheet ended up LIVE-LINKED
                // to Key Financials/Annual Data vs written as plain hardcoded numbers (because
                // autoLink couldn't confidently match a label/value), so a run can actually be
                // measured instead of guessed at. Percentage/margin rows are deliberately never
                // linked (see writeDataSheet below) and assumption "judgment" inputs (WACC, growth,
                // multiples, etc.) are deliberately hardcoded by design — both are tracked separately
                // from genuine linking misses so they don't inflate the "hardcoded" count.
                const linkStats = {};
                const recordLinkStat = (sheetKey, kind) => {
                    const s = linkStats[sheetKey] || (linkStats[sheetKey] = { linked: 0, hardcoded: 0, excluded: 0 });
                    s[kind]++;
                };
                let curSheetKey = null;
                let curRowKey = null; // set per-row in writeDataSheet, so a bad reference logs WHERE it was written, not just the raw template
                const resolve = (template, ci) => {
                    const col = colAt(ci), pcol = prevColAt(ci) || col;
                    const nIdx = ci >= FC_BOUNDARY ? (ci - (FC_BOUNDARY - 1)) : (ci + 1); // {N} = forecast year number (1..FC_YEARS)
                    let bad = false, badRef = null;
                    // {SUM:..}/{LAST:..} accept "key" (same sheet) or "sheet.key"; resolve over forecast cols.
                    const aggRowOf = (body) => {
                        const dot = body.indexOf(".");
                        if (dot > 0) {
                            const sk = body.slice(0, dot), rk = body.slice(dot + 1);
                            if (!SHEETS[sk] || !symbols[sk] || !symbols[sk][rk]) return null;
                            return { pfx: `'${SHEETS[sk]}'!`, row: symbols[sk][rk] };
                        }
                        const row = symbols[curSheetKey] && symbols[curSheetKey][body];
                        return row ? { pfx: "", row } : null;
                    };
                    // A hallucinated/invented key (never emitted by any sheet — the model made it up,
                    // e.g. a per-segment assumption key that doesn't exist) MUST NOT resolve to a
                    // plausible-looking "0": F9*(1+0) computes cleanly to a wrong-but-normal-looking
                    // number, so the blanket IFERROR below never even engages — there's no error to
                    // catch. Substituting NA() instead forces an actual Excel error to propagate from
                    // that exact spot, so it can be surfaced distinctly (see the "bad" branch below)
                    // instead of silently vanishing into the sheet as a believable zero.
                    const markBad = (ref) => { bad = true; badRef = badRef || ref; return "NA()"; };
                    const out = String(template)
                        .replace(/\{N\}/g, String(nIdx))
                        .replace(/\{(SUM|LAST|FWD1|CUR|SCALAR|PS|S|A|R|P):([^}]+)\}/g, (m, kind, body) => {
                            if (kind === "A") {
                                const row = symbols.assum[body];
                                if (!row) return markBad(`A:${body}`);
                                // A handful of assumption rows (wacc, beta, target multiples — but
                                // NOT shares_out, which is a genuine per-year figure since share
                                // count changes over time) were collapsed into single-cell scalars —
                                // see scalarAssumRows above — so every reference to one of THOSE, no
                                // matter which year column is currently being written, must point
                                // at that one fixed column-B cell, unlike an ordinary multi-year row.
                                if (scalarAssumRows.has(row)) return `'${SHEETS.assum}'!B${row}`;
                                // Assumptions is a multi-year schedule — read the SAME year column.
                                return `'${SHEETS.assum}'!${col}${row}`;
                            }
                            if (kind === "R" || kind === "P") {
                                const row = symbols[curSheetKey] && symbols[curSheetKey][body];
                                if (!row) return markBad(`${kind}:${body}`);
                                return `${kind === "P" ? pcol : col}${row}`;
                            }
                            if (kind === "SUM" || kind === "LAST" || kind === "FWD1" || kind === "CUR") {
                                const a = aggRowOf(body);
                                if (!a) return markBad(`${kind}:${body}`);
                                const lastFcCol = FC_COLS[FC_COLS.length - 1];
                                if (kind === "SUM") return `SUM(${a.pfx}${FC_COLS[0]}${a.row}:${lastFcCol}${a.row})`;
                                // CUR = last HISTORICAL/actual year (today); FWD1 = first forecast
                                // year (near-term/NTM); LAST = final forecast year (5 years out).
                                if (kind === "CUR") return `${a.pfx}${LAST_HIST_COL}${a.row}`;
                                return `${a.pfx}${kind === "FWD1" ? FC_COLS[0] : lastFcCol}${a.row}`;
                            }
                            if (kind === "SCALAR") {
                                // A scalar ("scalar":true) row holds its ONE value in column B only —
                                // not one-per-year like normal rows — so a per-year row referencing it
                                // (e.g. PER = current_price ÷ that year's EPS) needs a FIXED column B
                                // reference regardless of which year column is currently being
                                // written, unlike {R:}/{A:} which always follow the current column.
                                const a = aggRowOf(body);
                                if (!a) return markBad(`${kind}:${body}`);
                                return `${a.pfx}B${a.row}`;
                            }
                            // S / PS — cross-sheet, same / previous year
                            const dot = body.indexOf(".");
                            const sk = body.slice(0, dot), rk = body.slice(dot + 1);
                            const row = symbols[sk] && symbols[sk][rk];
                            if (!row || !SHEETS[sk]) return markBad(`${kind}:${body}`);
                            return `'${SHEETS[sk]}'!${kind === "PS" ? pcol : col}${row}`;
                        });
                    if (bad) refMissing.push(`${curSheetKey}.${curRowKey} -> "${template}"`);
                    // Ordinary runtime errors (#DIV/0!, #REF!, etc. from an otherwise-valid formula,
                    // e.g. dividing by a legitimately-zero denominator) still guard to a quiet 0 — that
                    // class of error is expected and fine to hide. A hallucinated-key error is
                    // different: it's a real bug, not a benign edge case, so it gets a visible marker
                    // instead of blending back into the sheet as an indistinguishable zero.
                    return bad ? `=IFERROR(${out},"#MISSING:${badRef}")` : `=IFERROR(${out},0)`;
                };

                // ── Styling helpers ──
                const titleBar = (sh, text) => {
                    sh.getRange("A1").values = [[text]];
                    const rng = sh.getRange(`A1:${CAGR_COL}1`);
                    rng.format.fill.color = "#173760"; rng.format.font.color = "#ffffff";
                    rng.format.font.bold = true; rng.format.font.size = 13;
                    sh.getRange("A:A").format.columnWidth = 230;
                    for (const col of [...ALL_COLS, CAGR_COL]) sh.getRange(`${col}:${col}`).format.columnWidth = 74;
                };
                const periodHeader = (sh) => {
                    const hdr = sh.getRange(`B3:${CAGR_COL}3`);
                    hdr.values = [[...periodLabels, CAGR_LABEL]];
                    hdr.format.font.bold = true; hdr.format.font.size = 9; hdr.format.font.color = "#173760";
                    for (let i = 0; i < ALL_COLS.length; i++) {
                        sh.getRange(`${ALL_COLS[i]}3`).format.fill.color = i < HIST_YEARS ? "#e8edf3" : "#d0d9f0";
                    }
                    sh.getRange(`${CAGR_COL}3`).format.fill.color = "#fff2cc";
                };

                // ── Write a period-grid sheet (assum / op / pnl / bs / capex / val) ──
                // opts.blue colours hardcoded numeric inputs blue; opts.unitLabel overrides the title unit.
                const writeDataSheet = (sheetKey, rows, opts = {}) => {
                    curSheetKey = sheetKey;
                    const sh = made[sheetKey];
                    titleBar(sh, `${companyName} — ${SHEETS[sheetKey].replace(/^FM /, "")}  |  ${opts.unitLabel || "Rs Crores"}`);
                    periodHeader(sh);
                    for (const item of rows) {
                        const rr = item._row;
                        curRowKey = item.key || item.label || rr;
                        if (item.section) {
                            sh.getRange(`A${rr}`).values = [[item.section]];
                            const rng = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                            rng.format.font.bold = true; rng.format.font.color = "#173760"; rng.format.fill.color = "#e8edf3";
                            continue;
                        }
                        sh.getRange(`A${rr}`).values = [[displayLabel(item)]];
                        const fmt = item.fmt || "#,##0";

                        // Scalar row: a single value/formula in column B (valuation outputs, WACC inputs).
                        if (item.scalar) {
                            const b = sh.getRange(`B${rr}`);
                            if (item.formula) b.formulas = [[resolve(item.formula, 0)]];
                            else if (typeof item.value === "number" && isFinite(item.value)) b.values = [[item.value]];
                            b.numberFormat = [[normalizeFmt(fmt)]];
                            if (item.emphasis) {
                                sh.getRange(`A${rr}`).format.font.bold = true;
                                b.format.font.bold = true;
                                if (item.emphasis === "highlight") sh.getRange(`A${rr}:B${rr}`).format.fill.color = "#d9ead3";
                            } else if (rr % 2 === 0) {
                                sh.getRange(`A${rr}:B${rr}`).format.fill.color = "#f9f9f9";
                            }
                            continue;
                        }

                        // Historical columns (B..F): LIVE-LINK via the index when the label resolves
                        // (scoped INDEX/MATCH into Key Financials OR Annual Data), with the AI's numbers
                        // as a per-cell fallback; else plain numbers; else derive.
                        // Never link percentage/margin rows: the data stores them as percent NUMBERS
                        // (e.g. 6.9) but our cells use a "%" format (which ×100s the value -> 690%).
                        // Use the AI's decimals / a derived formula for those instead. Also never
                        // relink a row explicitly marked noAutoLink (e.g. debtor_days/inventory_days/
                        // payable_days, forced from sc_year_data) — autoLink's label/value matching
                        // would happily rediscover an unrelated same-labeled row in Key Financials/
                        // Annual Data and silently override the authoritative sourced figures with a
                        // different, wrong one (this has actually happened).
                        const linkEntry = (/%/.test(fmt) || item.noAutoLink) ? null : autoLink(item);
                        if (linkEntry) {
                            recordLinkStat(sheetKey, "linked");
                            const histNums = Array.isArray(item.historical) ? item.historical : [];
                            for (let i = 0; i < HIST_YEARS; i++) {
                                const col = linkEntry.colByFY[HIST[i]];
                                const cell = sh.getRange(`${HIST_COLS[i]}${rr}`);
                                if (col) {
                                    cell.formulas = [[`='${linkEntry.sheet}'!${col}${linkEntry.row}`]];
                                } else if (histNums[i] !== null && histNums[i] !== undefined && histNums[i] !== "") {
                                    cell.values = [[histNums[i]]];
                                }
                            }
                        } else if (Array.isArray(item.historical)) {
                            // Assumptions rows the model itself flagged "input":true (WACC/growth/
                            // multiple-style judgment calls) and percentage/margin rows are
                            // deliberately never linked by design — don't count those toward
                            // "hardcoded when it should have linked".
                            recordLinkStat(sheetKey, (item.input === true || /%/.test(fmt)) ? "excluded" : "hardcoded");
                            const vals = item.historical.slice(0, HIST_YEARS);
                            while (vals.length < HIST_YEARS) vals.push("");
                            sh.getRange(`B${rr}:${LAST_HIST_COL}${rr}`).values = [vals.map(v => (v === null || v === undefined) ? "" : v)];
                            if (opts.blue) sh.getRange(`B${rr}:${LAST_HIST_COL}${rr}`).format.font.color = "#1155CC";
                        } else if (item.formula && !item.forecast_only) {
                            // derived row (e.g. a margin) — compute historicals too
                            for (let i = 0; i < HIST_YEARS; i++) sh.getRange(`${HIST_COLS[i]}${rr}`).formulas = [[resolve(item.formula, i)]];
                        }

                        // Forecast columns: recurrence formula.
                        if (item.formula) {
                            for (let i = 0; i < FC_YEARS; i++) sh.getRange(`${FC_COLS[i]}${rr}`).formulas = [[resolve(item.formula, HIST_YEARS + i)]];
                        } else if (Array.isArray(item.forecast)) {
                            const vals = item.forecast.slice(0, FC_YEARS);
                            while (vals.length < FC_YEARS) vals.push("");
                            sh.getRange(`${FC_COLS[0]}${rr}:${LAST_FC_COL}${rr}`).values = [vals.map(v => (v === null || v === undefined) ? "" : v)];
                            if (opts.blue) sh.getRange(`${FC_COLS[0]}${rr}:${LAST_FC_COL}${rr}`).format.font.color = "#1155CC";
                        } else if (!item.forecast_only && (linkEntry || Array.isArray(item.historical))) {
                            // The model gave this row real historicals but NEITHER a forecast "formula"
                            // NOR "forecast" values — left as-is, Excel reads the blank forecast cells as
                            // 0, which silently zeroes out everything downstream (e.g. an EBITDA margin
                            // row defaulting to 0% makes segment EBITDA compute to 0 even though revenue
                            // is fine, or depreciation defaulting to 0 makes EBIT = EBITDA). Flat-carry-
                            // forward the last historical value instead of a silent zero, and flag it —
                            // this is a real gap that should get a real forecast, not just a duplicate.
                            for (let i = 0; i < FC_YEARS; i++) sh.getRange(`${FC_COLS[i]}${rr}`).formulas = [[`=IFERROR(${i === 0 ? LAST_HIST_COL : FC_COLS[i - 1]}${rr},0)`]];
                            // Rows matching DRIVER_EXCLUDE_RE (financing/capital-structure items —
                            // borrowings, share capital, reserves/equity) are INTENTIONALLY flat by
                            // design (no repayment/issuance schedule modeled — see the Balance Sheet
                            // prompt), the same list the driver-fallback pass already uses to skip
                            // trying to wire them to a revenue driver. Flat-carrying them is the
                            // CORRECT behaviour here, not a gap — don't clutter the review list with
                            // rows that were never broken; only log genuinely-unresolved ones.
                            if (!DRIVER_EXCLUDE_RE.test(item.key || "")) {
                                missingForecast.push(`${sheetKey}.${item.key || item.label || rr}`);
                            }
                        }

                        sh.getRange(`B${rr}:${LAST_FC_COL}${rr}`).numberFormat = [new Array(ALL_COLS.length).fill(normalizeFmt(fmt))];
                        if (item.cagr) {
                            sh.getRange(`${CAGR_COL}${rr}`).formulas = [[`=IFERROR((${LAST_FC_COL}${rr}/${LAST_HIST_COL}${rr})^(1/${FC_YEARS})-1,"")`]];
                            sh.getRange(`${CAGR_COL}${rr}`).numberFormat = [[normalizeFmt("0.0%")]];
                        }
                    }
                };

                // ── Assumptions: multi-year driver schedule (blue = input you can flex) ──
                writeDataSheet("assum", mainAssumRows, { blue: true, unitLabel: "in row's stated unit" });
                made.assum.getRange("A2").values = [["Blue = hardcoded input you can flex · Black = formula / linked actual · each driver feeds the other sheets via the same year column"]];
                made.assum.getRange("A2").format.font.size = 9;
                made.assum.getRange("A2").format.font.color = "#777777";

                // ── Market / Valuation Snapshot — single-fact inputs (tax rate, WACC, beta,
                // target multiples, etc.) pulled out of the main per-year grid above (see
                // MARKET_SNAPSHOT_KEYS) and rendered as their own small Label|Value table instead,
                // since none of them have "a value per year" — every forecast year already reads
                // the SAME number via {A:key} regardless of where the row physically sits.
                if (snapshotRows.length) {
                    const sh = made.assum;
                    sh.getRange(`A${assumSnapshotTitleRow}`).values = [["MARKET / VALUATION SNAPSHOT  (single fact — not a per-year series)"]];
                    const titleRng = sh.getRange(`A${assumSnapshotTitleRow}:B${assumSnapshotTitleRow}`);
                    titleRng.format.font.bold = true; titleRng.format.font.color = "#173760"; titleRng.format.fill.color = "#e8edf3";
                    for (const item of snapshotRows) {
                        const rr = item._row;
                        sh.getRange(`A${rr}`).values = [[displayLabel(item)]];
                        const b = sh.getRange(`B${rr}`);
                        const val = (typeof item.value === "number" && isFinite(item.value)) ? item.value : null;
                        if (val != null) b.values = [[val]];
                        b.numberFormat = [[normalizeFmt(item.fmt || "#,##0")]];
                    }
                }

                // Source citations AND the Assumptions (Judgment) call's per-row "rationale" (the
                // analyst's stated reasoning for a judgment-driven number/trajectory — see that
                // call's prompt), as a single Excel cell comment on the row's OWN LABEL cell (column
                // A) — not a data cell. This used to sit on the first forecast-year data cell (column
                // G for a per-year row, B for a scalar row), which reads oddly: a note attached to
                // one specific year's number looks like it's ABOUT that year, when it's actually
                // about the row/assumption as a whole. Column A is where every row's own name already
                // lives (see displayLabel above), so that's the natural, unambiguous place for a note
                // about the row. Combined into ONE comment string rather than two separate
                // wb.comments.add() calls on the same cell — the Comments API only supports one
                // comment per cell, and a second call there would throw. Best-effort: the Comments
                // API needs a reasonably current Excel client, and a citation/rationale is only as
                // good as the model's own text — failures here must never abort the rest of the
                // model build.
                // The hard 300-char cap this used to have was well below what Excel's Comments API
                // actually supports (tens of thousands of characters) — it existed as an arbitrary
                // safety margin, but became a real problem once the Judgment call's prompt started
                // REQUIRING a specific, evidence-cited rationale (see valuationChartNote/REVENUE
                // GROWTH GROUNDING/MARGIN-TAX-PAYOUT GROUNDING above) rather than a one-line note —
                // that text now routinely exceeds 300 characters and was being sliced off mid-word,
                // sometimes mid-sentence, which reads as broken rather than intentionally trimmed.
                // Raised the cap substantially and truncate at the nearest word boundary (with a
                // trailing "…" so a cut note is visibly a cut note, not a sentence that just stops).
                const truncateNote = (s, max) => {
                    if (s.length <= max) return s;
                    let cut = s.slice(0, max);
                    const lastSpace = cut.lastIndexOf(" ");
                    if (lastSpace > max * 0.6) cut = cut.slice(0, lastSpace); // don't back up more than ~40% of the budget looking for a space
                    return cut.trim() + "…";
                };
                try {
                    let sourced = 0;
                    for (const item of assumRows) {
                        if (!item || !item.key || !item._row) continue;
                        const parts = [];
                        if (item.rationale) parts.push(String(item.rationale));
                        if (item.source) parts.push(`Source: ${item.source}`);
                        if (!parts.length) continue;
                        wb.comments.add(made.assum.getRange(`A${item._row}`), truncateNote(parts.join(" — "), 1200));
                        sourced++;
                    }
                    await context.sync();
                    if (sourced) console.log(`[AI Model] Added ${sourced} source/rationale note(s) to the Assumptions sheet.`);
                } catch (e) {
                    console.warn("[AI Model] Could not attach source comments (Comments API unavailable on this Excel client?):", e.message);
                }

                writeDataSheet("op", opRows);
                writeDataSheet("pnl", pnlRows);
                writeDataSheet("bs", bsRows);
                writeDataSheet("capex", capexRows);
                writeDataSheet("val", valRows);

                // ── Sensitivity: Value per Share vs WACC / Terminal Growth ──
                // Code-generated, not LLM-generated — WACC and terminal growth are the only two DCF
                // inputs that affect PURELY the discounting math, not the underlying projected cash
                // flows, so this grid is mechanically exact: it recomputes sum-of-PV-FCFF, terminal
                // value, and value per share for a 5x5 grid of WACC/growth combinations, reusing the
                // SAME fixed FCFF/NOPAT/net-working-capital/net-debt/shares-out figures the base case
                // already computed. It deliberately does NOT vary revenue growth or margin — doing
                // that correctly would mean rebuilding the entire operating model (Operational/P&L/
                // Balance Sheet), not just re-deriving this sheet's own formulas.
                {
                    const fcffRow = symbols.val && symbols.val.fcff;
                    const nopatRow = symbols.val && symbols.val.nopat;
                    const nwcRow = symbols.bs && symbols.bs.net_working_capital;
                    const netDebtRow = symbols.val && symbols.val.net_debt_last;
                    const sharesOutRow = symbols.assum && symbols.assum.shares_out;
                    const holdingDiscRow = symbols.assum && symbols.assum.holding_discount;
                    const waccRow = symbols.assum && symbols.assum.wacc;
                    const tgRow = symbols.assum && symbols.assum.terminal_growth;

                    if (fcffRow && nopatRow && nwcRow && netDebtRow && sharesOutRow && waccRow && tgRow) {
                        const valSh = made.val;
                        const lastUsedRow = Math.max(...valRows.filter(r => r && r._row).map(r => r._row));
                        const startRow = lastUsedRow + 3;
                        const noteRow = startRow + 1;
                        const headerRow = startRow + 3;
                        const gridStartRow = headerRow + 1;
                        const steps = [-0.01, -0.005, 0, 0.005, 0.01]; // -1%, -0.5%, base, +0.5%, +1%
                        const gridCols = ["C", "D", "E", "F", "G"]; // 5 terminal-growth columns (B holds the WACC row headers)

                        // wacc is a single-cell SCALAR (column B only — see collapseToScalar
                        // above), so it's read from column B directly. shares_out is a genuine
                        // PER-YEAR row (share count changes over time), so it's read from the LAST
                        // HISTORICAL/actual column — today's real share count — same {CUR:}
                        // convention used in the dictated formulas (see value_per_share). terminal
                        // growth and holding discount are still ordinary constant-global-input rows
                        // that repeat the SAME value in every year column, so reading them from the
                        // first forecast column is as valid as any other column.
                        const waccCell = `'${SHEETS.assum}'!B${waccRow}`;
                        const tgCell = `'${SHEETS.assum}'!${FC_COLS[0]}${tgRow}`;
                        const sharesOutCell = `'${SHEETS.assum}'!${LAST_HIST_COL}${sharesOutRow}`;
                        const holdingDiscCell = holdingDiscRow ? `'${SHEETS.assum}'!${FC_COLS[0]}${holdingDiscRow}` : null;
                        const fcffRange = `'${SHEETS.val}'!${FC_COLS[0]}${fcffRow}:${LAST_FC_COL}${fcffRow}`;
                        const nopatLastCell = `'${SHEETS.val}'!${LAST_FC_COL}${nopatRow}`;
                        const nwcLastCell = `'${SHEETS.bs}'!${LAST_FC_COL}${nwcRow}`;
                        // net_debt_last is ALSO a scalar row (column B only, like wacc above) —
                        // this previously read LAST_FC_COL, a blank cell for a scalar row, which
                        // silently zeroed net debt out of every sensitivity-grid value.
                        const netDebtLastCell = `'${SHEETS.val}'!B${netDebtRow}`;
                        const fcYearExponents = Array.from({ length: FC_YEARS }, (_, i) => i + 1).join(",");

                        valSh.getRange(`A${startRow}`).values = [["SENSITIVITY — Value per Share (WACC vs Terminal Growth)"]];
                        valSh.getRange(`A${startRow}:G${startRow}`).format.font.bold = true;
                        valSh.getRange(`A${startRow}`).format.font.color = "#173760";
                        valSh.getRange(`A${startRow}:G${startRow}`).format.fill.color = "#e8edf3";

                        valSh.getRange(`A${noteRow}`).values = [["Rows vary WACC, columns vary terminal growth, both ± the base case in 0.5pp steps. Holds every other assumption (revenue growth, margins, capex, etc.) at its base case — only WACC/terminal growth affect the discounting math directly."]];
                        valSh.getRange(`A${noteRow}`).format.font.size = 9;
                        valSh.getRange(`A${noteRow}`).format.font.color = "#777777";

                        valSh.getRange(`B${headerRow}`).values = [["WACC ▼ / Term. Growth ►"]];
                        valSh.getRange(`B${headerRow}`).format.font.bold = true;
                        valSh.getRange(`B${headerRow}`).format.font.size = 9;
                        for (let c = 0; c < steps.length; c++) {
                            const cell = valSh.getRange(`${gridCols[c]}${headerRow}`);
                            cell.formulas = [[`=${tgCell}${steps[c] >= 0 ? "+" : ""}${steps[c]}`]];
                            cell.numberFormat = [["0.0%"]];
                            cell.format.font.bold = true;
                            cell.format.fill.color = "#d0d9f0";
                        }

                        for (let r = 0; r < steps.length; r++) {
                            const rowNum = gridStartRow + r;
                            const waccHeaderCell = valSh.getRange(`B${rowNum}`);
                            waccHeaderCell.formulas = [[`=${waccCell}${steps[r] >= 0 ? "+" : ""}${steps[r]}`]];
                            waccHeaderCell.numberFormat = [["0.0%"]];
                            waccHeaderCell.format.font.bold = true;
                            waccHeaderCell.format.fill.color = "#d0d9f0";

                            for (let c = 0; c < steps.length; c++) {
                                const waccRef = `$B${rowNum}`;
                                const tgRef = `${gridCols[c]}$${headerRow}`;
                                const discountPart = holdingDiscCell ? `*(1-${holdingDiscCell})` : "";
                                const formula = `=((SUMPRODUCT(${fcffRange},1/(1+${waccRef})^{${fcYearExponents}})`
                                    + `+MAX(0,(${nopatLastCell}-${tgRef}*${nwcLastCell})*(1+${tgRef})/(${waccRef}-${tgRef}))/(1+${waccRef})^${FC_YEARS}`
                                    + `-${netDebtLastCell})${discountPart})/${sharesOutCell}`;
                                const cell = valSh.getRange(`${gridCols[c]}${rowNum}`);
                                cell.formulas = [[formula]];
                                cell.numberFormat = [["#,##0"]];
                            }
                        }
                    } else {
                        console.warn("[AI Model] Skipped WACC/Terminal-Growth sensitivity table — one or more required rows (fcff, nopat, net_working_capital, net_debt_last, shares_out, wacc, terminal_growth) were missing from the model.");
                    }
                }

                // ── Summary Dashboard (pure cross-sheet links) ──
                {
                    const sh = made.summary;
                    titleBar(sh, `${companyName} — Summary Dashboard  |  Rs Crores`);
                    periodHeader(sh);
                    let rr = 5;
                    for (const item of summaryRows) {
                        if (item.section) {
                            sh.getRange(`A${rr}`).values = [[item.section]];
                            const rng = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                            rng.format.font.bold = true; rng.format.font.color = "#173760"; rng.format.fill.color = "#e8edf3";
                            rr++; continue;
                        }
                        const dot = String(item.ref || "").indexOf(".");
                        const sk = dot > 0 ? item.ref.slice(0, dot) : "";
                        const rk = dot > 0 ? item.ref.slice(dot + 1) : "";
                        const srcRow = symbols[sk] && symbols[sk][rk];
                        sh.getRange(`A${rr}`).values = [[item.label || rk || ""]];
                        if (srcRow && SHEETS[sk]) {
                            for (let i = 0; i < ALL_COLS.length; i++) {
                                sh.getRange(`${ALL_COLS[i]}${rr}`).formulas = [[`='${SHEETS[sk]}'!${ALL_COLS[i]}${srcRow}`]];
                            }
                            sh.getRange(`B${rr}:${LAST_FC_COL}${rr}`).numberFormat = [new Array(ALL_COLS.length).fill(normalizeFmt(item.fmt || "#,##0"))];
                            // Same canonical-key override as the source sheets: force CAGR for
                            // absolute-value metrics regardless of the model's own "cagr" flag.
                            if (item.cagr || (CAGR_KEYS[sk] || []).includes(rk)) {
                                sh.getRange(`${CAGR_COL}${rr}`).formulas = [[`=IFERROR((${LAST_FC_COL}${rr}/${LAST_HIST_COL}${rr})^(1/${FC_YEARS})-1,"")`]];
                                sh.getRange(`${CAGR_COL}${rr}`).numberFormat = [[normalizeFmt("0.0%")]];
                            }
                        }
                        if (rr % 2 === 0) sh.getRange(`A${rr}:${CAGR_COL}${rr}`).format.fill.color = "#f9f9f9";
                        rr++;
                    }
                    // ── SOTP fallback flag (shown on the dashboard) — visible, not just a console
                    // warning, whenever this company was expected to get segment-aware modeling but
                    // didn't fully get it: either the Segment Drivers call never landed 2+ usable
                    // segment pairs despite holding_discount>0 (see the console.warn near
                    // isSotpCompany above), or it did but the Operational Model call's output didn't
                    // deliver the dictated segment revenue rows (see opDeliveredSegments below).
                    // Someone opening this file has no way to see a browser console warning — this
                    // is the only place they'd ever learn the SOTP treatment was expected but silently
                    // fell back to the single blended DCF for this run.
                    if (holdingDiscountVal > 0 && !(isSotpCompany && opDeliveredSegments)) {
                        rr += 2;
                        const reason = segmentTags.length < 2
                            ? "the segment-driver assumptions were never produced"
                            : "the Operational Model did not deliver the expected per-segment revenue rows";
                        sh.getRange(`A${rr}`).values = [[`⚠  SOTP SEGMENT MODELING EXPECTED BUT DID NOT ENGAGE`]];
                        const sotpHdr = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                        sotpHdr.format.font.bold = true; sotpHdr.format.font.size = 10;
                        sotpHdr.format.font.color = "#8a4b00"; sotpHdr.format.fill.color = "#fff2cc";
                        rr++;
                        sh.getRange(`A${rr}`).values = [[`This company was judged a genuine multi-segment conglomerate (holding_discount = ${(holdingDiscountVal * 100).toFixed(0)}%), which should route the model to segment-specific growth/margin/capex/valuation assumptions instead of one blended company-wide rate — but ${reason}, so this run fell back to the single blended DCF. The numbers below are not wrong, but they blend distinct businesses into one average the way every pre-SOTP model did; re-running the build may pick up the segment treatment.`]];
                        const sotpNote = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                        sotpNote.format.font.size = 9; sotpNote.format.font.color = "#5a4a30";
                        sotpNote.format.fill.color = "#fff8e5"; sotpNote.format.wrapText = true;
                        sh.getRange(`A${rr}`).format.rowHeight = 60;
                    }
                    // ── Zero-revenue segment flag ──
                    // A segment the model left at zero AND that couldn't be recovered from the
                    // operational dashboard contributes nothing to net_revenue, and its real share of
                    // the business silently lands in "Other Segments" instead. That's invisible in
                    // the numbers themselves (the total still ties), so it gets said out loud here —
                    // a console warning is no use to whoever opens the delivered file.
                    if (zeroSegmentFlags.length) {
                        rr += 2;
                        sh.getRange(`A${rr}`).values = [[`⚠  ${zeroSegmentFlags.length} SEGMENT${zeroSegmentFlags.length === 1 ? "" : "S"} HAVE NO REVENUE DATA`]];
                        const zHdr = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                        zHdr.format.font.bold = true; zHdr.format.font.size = 10;
                        zHdr.format.font.color = "#8a4b00"; zHdr.format.fill.color = "#fff2cc";
                        rr++;
                        sh.getRange(`A${rr}`).values = [[`The AI returned zero historical revenue for ${zeroSegmentFlags.join(", ")}, and no annual Rs-Crore revenue figure for ${zeroSegmentFlags.length === 1 ? "it" : "them"} could be found in this company's own operational disclosures either. ${zeroSegmentFlags.length === 1 ? "This segment contributes" : "These segments contribute"} nothing to the modelled Net Revenue, and ${zeroSegmentFlags.length === 1 ? "its" : "their"} real share of the business is absorbed into "Other Segments" instead — so the total still ties to reported revenue, but the segment split on the Operational sheet understates ${zeroSegmentFlags.length === 1 ? "this segment" : "these segments"} and overstates Other Segments. Review the Operational sheet's segment revenue rows before relying on the sum-of-the-parts valuation.`]];
                        const zNote = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                        zNote.format.font.size = 9; zNote.format.font.color = "#5a4a30";
                        zNote.format.fill.color = "#fff8e5"; zNote.format.wrapText = true;
                        sh.getRange(`A${rr}`).format.rowHeight = 60;
                    }
                    // ── Disclaimers (shown on the dashboard) ──
                    rr += 2;
                    sh.getRange(`A${rr}`).values = [["⚠  AI-COMPILED — PLEASE DOUBLE-CHECK"]];
                    const aiHdr = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                    aiHdr.format.font.bold = true; aiHdr.format.font.size = 10;
                    aiHdr.format.font.color = "#8a4b00"; aiHdr.format.fill.color = "#fff2cc";
                    rr++;
                    sh.getRange(`A${rr}`).values = [["This financial model has been compiled by AI from public data sources. The figures, formulas and projections are estimates and may contain errors — please double-check every number before relying on it. It is not investment advice."]];
                    const aiNote = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                    aiNote.format.font.size = 9; aiNote.format.font.color = "#5a4a30";
                    aiNote.format.fill.color = "#fff8e5"; aiNote.format.wrapText = true;
                    sh.getRange(`A${rr}`).format.rowHeight = 48;
                    rr += 2;
                    sh.getRange(`A${rr}`).values = [["General Disclaimer"]];
                    sh.getRange(`A${rr}`).format.font.bold = true; sh.getRange(`A${rr}`).format.font.size = 9; sh.getRange(`A${rr}`).format.font.color = "#555555";
                    rr++;
                    sh.getRange(`A${rr}`).values = [[GENERAL_DISCLAIMER]];
                    const disc = sh.getRange(`A${rr}:${CAGR_COL}${rr}`);
                    disc.format.font.size = 8; disc.format.font.color = "#888888"; disc.format.wrapText = true;
                    sh.getRange(`A${rr}`).format.rowHeight = 200;
                }

                await context.sync();
                // Force a full recalculation regardless of the workbook's calculation mode. If it's
                // set to Manual (common in large models, and possibly inherited from however this
                // workbook was created), freshly-written cross-sheet formulas — e.g. the Summary
                // Dashboard's direct ='FM P&L'!H23-style references — display 0/stale ("-" per our
                // accounting number format) until something triggers a recalc, even though the
                // source cell already holds the correct value. Don't rely on the user pressing F9.
                context.workbook.application.calculate(Excel.CalculationType.full);
                await context.sync();
                made.summary.activate();
                await context.sync();
                if (refMissing.length) {
                    // The SAME broken template repeats once per year-column it was written into (5-10x) —
                    // dedupe so the console shows each DISTINCT hallucinated/mistyped reference once.
                    const distinct = Array.from(new Set(refMissing));
                    console.warn(`[AI Model] ${refMissing.length} formula placeholder(s) referenced unknown keys, set to 0 (${distinct.length} distinct broken formula(s) — usually a made-up key name instead of the canonical one):`, distinct);
                }
                if (missingForecast.length) {
                    console.warn(`[AI Model] ${missingForecast.length} row(s) had historicals but no forecast formula/values — defaulted to a flat carry-forward of the last historical instead of a silent zero (review these manually):`, missingForecast);
                }
                // Diagnostic: how many historical cells per sheet ended up live-linked to Key
                // Financials/Annual Data vs written as plain hardcoded numbers because autoLink
                // couldn't confidently match them (label text didn't match, or too few years agreed
                // within tolerance) — "excluded" rows are deliberately-hardcoded judgment inputs and
                // percentage/margin rows, not a linking failure, so they're broken out separately.
                const linkTotals = Object.entries(linkStats);
                if (linkTotals.length) {
                    const totalHardcoded = linkTotals.reduce((sum, [, s]) => sum + s.hardcoded, 0);
                    console.log(`[AI Model] Historical cell linking by sheet (linked / hardcoded / deliberately-excluded):`,
                        Object.fromEntries(linkTotals.map(([k, s]) => [SHEETS[k] || k, `${s.linked} / ${s.hardcoded} / ${s.excluded}`])));
                    if (totalHardcoded) {
                        console.warn(`[AI Model] ${totalHardcoded} historical cell(s) across the model could not be confidently matched to a Key Financials/Annual Data row and were written as plain numbers instead of a live link — see the per-sheet breakdown above.`);
                    }
                }

                // Verify the balance-sheet revolver actually held: forecast cash should never compute
                // negative, and the balance check should stay near 0 (including historicals). Read back
                // the ACTUAL Excel-computed post-recalc values rather than trusting the formula text.
                const cashRow = symbols.bs && symbols.bs.cash;
                const balRow = symbols.bs && symbols.bs.balance_check;
                const taRow = symbols.bs && symbols.bs.total_assets;
                // Everything else needed for the 5 automated model-quality checks below — batched
                // into this SAME load/sync pass rather than a separate one.
                const epsRow = symbols.pnl && symbols.pnl.eps;
                const revenueRow = symbols.pnl && symbols.pnl.net_revenue;
                const patRow = symbols.pnl && symbols.pnl.pat;
                const sharesOutRow = symbols.assum && symbols.assum.shares_out;
                const upsideRow = symbols.val && symbols.val.upside;
                const brokerRow = symbols.assum && symbols.assum.broker_target_price;
                const targetPriceRow = symbols.val && symbols.val.target_price;
                const valuePerShareRow = symbols.val && symbols.val.value_per_share;
                let cashRange, balRange, taRange, epsRange, revenueRange, patRange, sharesOutCurCell, upsideCell, brokerCell, targetPriceCell, valuePerShareCell;
                if (cashRow) { cashRange = made.bs.getRange(`B${cashRow}:${LAST_FC_COL}${cashRow}`); cashRange.load("values"); }
                if (balRow) { balRange = made.bs.getRange(`B${balRow}:${LAST_FC_COL}${balRow}`); balRange.load("values"); }
                if (taRow) { taRange = made.bs.getRange(`B${taRow}:${LAST_FC_COL}${taRow}`); taRange.load("values"); }
                if (epsRow) { epsRange = made.pnl.getRange(`B${epsRow}:${LAST_HIST_COL}${epsRow}`); epsRange.load("values"); }
                if (revenueRow) { revenueRange = made.pnl.getRange(`B${revenueRow}:${LAST_HIST_COL}${revenueRow}`); revenueRange.load("values"); }
                if (patRow) { patRange = made.pnl.getRange(`B${patRow}:${LAST_HIST_COL}${patRow}`); patRange.load("values"); }
                if (sharesOutRow) { sharesOutCurCell = made.assum.getRange(`${LAST_HIST_COL}${sharesOutRow}`); sharesOutCurCell.load("values"); }
                if (upsideRow) { upsideCell = made.val.getRange(`B${upsideRow}`); upsideCell.load("values"); }
                if (brokerRow) { brokerCell = made.assum.getRange(`B${brokerRow}`); brokerCell.load("values"); }
                if (targetPriceRow) { targetPriceCell = made.val.getRange(`B${targetPriceRow}`); targetPriceCell.load("values"); }
                if (valuePerShareRow) { valuePerShareCell = made.val.getRange(`B${valuePerShareRow}`); valuePerShareCell.load("values"); }
                if (cashRow || balRow || taRow || epsRow || revenueRow || patRow || sharesOutRow || upsideRow || brokerRow || targetPriceRow || valuePerShareRow) await context.sync();

                let negatives = [];
                if (cashRange) {
                    const cashVals = cashRange.values[0];
                    negatives = cashVals.map((v, i) => ({ v, i })).filter(x => typeof x.v === "number" && x.v < 0);
                    if (negatives.length) {
                        console.warn(`[AI Model] Balance Sheet forecast CASH is NEGATIVE in ${negatives.length} column(s) — (${negatives.map(x => `${periodLabels[x.i]}: ${Math.round(x.v).toLocaleString()}`).join(", ")}) — a company cannot hold negative cash; the short_term_borrowings revolver formula should prevent this. Check that row's formula on the Balance Sheet sheet.`);
                    }
                }
                let offCols = [];
                if (balRange && taRange) {
                    const balVals = balRange.values[0];
                    const taVals = taRange.values[0];
                    for (let i = 0; i < balVals.length; i++) {
                        const b = balVals[i], ta = taVals[i];
                        if (typeof b === "number" && typeof ta === "number" && ta > 0 && Math.abs(b) / ta > 0.005) {
                            offCols.push(`${periodLabels[i]}: ${Math.round(b).toLocaleString()}`);
                        }
                    }
                    if (offCols.length) {
                        console.warn(`[AI Model] Balance Sheet does not balance (Assets − Liab.&Equity, >0.5% of total assets) in ${offCols.length} column(s): ${offCols.join(", ")}. Review the Balance Sheet sheet's historical links/formulas.`);
                    }
                }

                // Diagnostic-only, read back EVERY row that feeds total_assets/total_liabilities_equity
                // (see their forced sum-of-components formulas above) for the HISTORICAL columns, as
                // ACTUALLY WRITTEN to Excel — only when an imbalance was just detected above. The
                // balance_check figure alone only says "off by how much", not which row is wrong; the
                // pre-write "summed asset lines vs reported Total Assets" cross-check further up this
                // file also can't help here since it silently bails the moment any one component row
                // lacks a plain "historical" array (e.g. a link-only row) — this reads back the REAL
                // post-recalc Excel values instead, for every component on both sides, so the specific
                // offending row(s) can be identified directly rather than guessed at.
                if (offCols.length) {
                    const ASSET_KEYS = ["cash", "receivables", "inventory", "other_current_assets", "net_fixed_assets", "cwip", "investments", "other_non_current_assets"];
                    const LIAB_KEYS = ["share_capital", "reserves", "long_term_borrowings", "short_term_borrowings", "payables", "provisions", "deferred_tax_liabilities", "other_non_current_liabilities", "other_current_liabilities"];
                    const breakdownRows = [];
                    for (const [side, keys] of [["ASSET", ASSET_KEYS], ["LIAB+EQ", LIAB_KEYS]]) {
                        for (const k of keys) {
                            const row = symbols.bs && symbols.bs[k];
                            const item = row && bsRows.find(r => r && r._row === row);
                            if (!row) { breakdownRows.push({ side, key: k, row: null, item: null, rng: null }); continue; }
                            const rng = made.bs.getRange(`B${row}:${LAST_HIST_COL}${row}`);
                            rng.load("values");
                            breakdownRows.push({ side, key: k, row, item, rng });
                        }
                    }
                    await context.sync();
                    // A row can show "linked" overall (autoLink found SOME match) while still
                    // falling back to a plain hardcoded number for individual years outside that
                    // match's own colByFY range (year-chunks in Annual Data aren't merged for a
                    // single live-link target — see the long comment on buildIndex's "map" above).
                    // Those fallback years are the LLM's own guess, not a real linked cell, and are
                    // the most likely place a stale/wrong figure hides — mark each YEAR individually
                    // rather than just the row as a whole.
                    const table = breakdownRows.map(({ side, key, row, item, rng }) => {
                        if (!row) return { side, key, row: "NOT FOUND" };
                        const vals = rng.values[0];
                        const fmt = (item && item.fmt) || "#,##0";
                        const linkEntry = (item && !/%/.test(fmt) && !item.noAutoLink) ? autoLink(item) : null;
                        return {
                            side, key, row,
                            ...Object.fromEntries(HIST.map((fy, i) => {
                                const v = vals[i];
                                const isLive = !!(linkEntry && linkEntry.colByFY && linkEntry.colByFY[fy]);
                                return [`FY${fy}`, typeof v === "number" ? `${Math.round(v)}${isLive ? "" : " (H)"}` : v];
                            })),
                        };
                    });
                    console.log("[AI Model] Balance Sheet component breakdown (every row feeding total_assets/total_liabilities_equity, HISTORICAL columns, as actually written to Excel) — '(H)' marks a year that fell back to a hardcoded/guessed number instead of a live link into Annual Data; those are the most likely place a wrong figure hides. Compare ASSET rows vs LIAB+EQ rows per year to find which specific row is causing the imbalance reported above:");
                    console.table(table);
                }

                // ── 5 automated model-quality checks — TEMPORARY: shown as a dismissible popup
                // (see renderModelChecks/#modelChecksModal) rather than blocking the build outright,
                // pending a decision on hard-blocking a failed model from being published. ──
                modelCheckResults = [];

                // 1. Balance sheet ties (Assets = Liab.&Equity in every year) AND forecast cash
                // never goes negative — reuses the reads/checks just above. "Cash flow ties to the
                // balance sheet" doesn't have a literal, independently-derived closing-cash figure
                // to cross-check against in this model (no full indirect cash-flow statement is
                // built) — a revolver-driven negative-cash failure is the most direct, honest proxy
                // available for that half of the check.
                {
                    const pass = balRange && taRange && cashRange ? (offCols.length === 0 && negatives.length === 0) : null;
                    const negDetail = negatives.map(x => `${periodLabels[x.i]}: ${Math.round(x.v).toLocaleString()}`);
                    modelCheckResults.push({
                        label: "Balance sheet ties, and forecast cash never goes negative",
                        passed: pass,
                        // List the SPECIFIC years, not just a count — "8 year(s)" on its own tells
                        // you neither which years nor how far off, which is the whole point of
                        // surfacing this at all.
                        detail: pass == null ? "Could not verify — required rows not found"
                            : pass ? "Assets = Liabilities & Equity in every year; cash stays non-negative"
                                : [offCols.length ? `Doesn't balance (Assets − Liab.&Equity) in ${offCols.length} year(s): ${offCols.join(", ")}` : null,
                                   negatives.length ? `Cash negative in ${negatives.length} year(s): ${negDetail.join(", ")}` : null].filter(Boolean).join(" | "),
                    });
                }

                // 2. Modelled EPS ({R:pat}/{A:shares_out}, every historical year) vs the ACTUAL
                // reported EPS from the downloaded data — catches share-count, minority-interest,
                // and PAT-attribution errors in one line.
                {
                    const reportedEpsEntry = findEntryByLabel(/^(basic\s*|diluted\s*)?(eps|earnings\s*per\s*share)$/i);
                    let pass = null, detail = "Could not verify — EPS row or reported EPS figure not found";
                    if (epsRange && reportedEpsEntry) {
                        const epsVals = epsRange.values[0];
                        const mismatches = [];
                        let checked = 0;
                        HIST.forEach((fy, i) => {
                            // The OLDEST historical year (HIST[0]) sits at the edge of the requested
                            // window — the source data is frequently incomplete or on a different
                            // basis for that one year (e.g. a presentation change straddling it),
                            // which shows up here as a spurious EPS mismatch that has nothing to do
                            // with the model's own PAT/shares_out arithmetic. Not "2019" specifically
                            // — HIST[0] is whichever year is oldest in the window, which shifts every
                            // year as latestActualFY advances.
                            if (i === 0) return;
                            const reported = reportedEpsEntry.valByFY[fy];
                            const modelled = epsVals[i];
                            if (reported == null || typeof modelled !== "number") return;
                            checked++;
                            const diffPct = Math.abs(modelled - reported) / Math.max(Math.abs(reported), 0.01);
                            if (diffPct > 0.05) mismatches.push(`FY${fy}: modelled ${modelled.toFixed(2)} vs reported ${reported.toFixed(2)}`);
                        });
                        if (checked > 0) {
                            pass = mismatches.length === 0;
                            detail = pass ? `Modelled EPS matches reported EPS (±5%) in all ${checked} historical year(s) checked (oldest year FY${HIST[0]} excluded — data for that edge year is frequently incomplete)` : mismatches.join("; ");
                        }
                    }
                    modelCheckResults.push({ label: "Modelled EPS = reported EPS, for every actual year", passed: pass, detail });
                }

                // 2b/2c. Modelled Revenue and PAT (the two figures everything else on the P&L —
                // growth rates, margins, EPS, the DCF, target price — is ultimately built on top of)
                // vs the ACTUAL reported figures. Same rationale/tolerance/oldest-year exclusion as
                // the EPS check above. Added specifically because a systematic one-column-forward
                // transcription shift (the model reading a wide 13+-column table and silently landing
                // one FY off for 7 straight years) has been observed here — it would NOT be caught by
                // the EPS check alone if PAT and shares_out happened to both carry a self-consistent
                // shift, so these need their own independent check against reported data.
                const checkReportedVsModelled = (label, range, labelRegex) => {
                    const reportedEntry = findEntryByLabel(labelRegex);
                    let pass = null, detail = `Could not verify — row or reported ${label} figure not found`;
                    if (range && reportedEntry) {
                        const vals = range.values[0];
                        const mismatches = [];
                        let checked = 0;
                        HIST.forEach((fy, i) => {
                            if (i === 0) return; // see EPS check above — oldest year is frequently incomplete/on a different basis
                            const reported = reportedEntry.valByFY[fy];
                            const modelled = vals[i];
                            if (reported == null || typeof modelled !== "number") return;
                            checked++;
                            const diffPct = Math.abs(modelled - reported) / Math.max(Math.abs(reported), 1);
                            if (diffPct > 0.02) mismatches.push(`FY${fy}: modelled ${Math.round(modelled).toLocaleString()} vs reported ${Math.round(reported).toLocaleString()}`);
                        });
                        if (checked > 0) {
                            pass = mismatches.length === 0;
                            detail = pass ? `Modelled ${label} matches reported ${label} (±2%) in all ${checked} historical year(s) checked (oldest year FY${HIST[0]} excluded)` : mismatches.join("; ");
                        }
                    }
                    modelCheckResults.push({ label: `Modelled ${label} = reported ${label}, for every actual year`, passed: pass, detail });
                };
                checkReportedVsModelled("Revenue", revenueRange, /^(net\s*revenue|total\s*revenue|revenue\s*from\s*operations|net\s*sales|total\s*income|revenue)$/i);
                checkReportedVsModelled("PAT", patRange, /^(pat|net\s*profit|profit\s*after\s*tax)$/i);

                // 3. Market Cap ÷ Price = the share count actually used in the DCF — verifies the
                // {CUR:assum.shares_out} write actually took effect as intended (catches wiring bugs like
                // an autoLink override or a resolver pointing at the wrong column, not just a data
                // problem).
                {
                    let pass = null, detail = "Could not verify — shares_out or Market Cap/Price data not available";
                    if (sharesOutCurCell && sharesOutComputed != null) {
                        const writtenVal = sharesOutCurCell.values[0][0];
                        if (typeof writtenVal === "number") {
                            const diffPct = Math.abs(writtenVal - sharesOutComputed) / Math.max(Math.abs(sharesOutComputed), 0.01);
                            pass = diffPct < 0.01;
                            detail = pass
                                ? `DCF share count (${writtenVal.toFixed(2)} Cr) matches Market Cap ÷ Price (${sharesOutComputed.toFixed(2)} Cr, FY${sharesOutComputedFY})`
                                : `DCF share count (${writtenVal}) does not match Market Cap ÷ Price (${sharesOutComputed.toFixed(2)} Cr, FY${sharesOutComputedFY})`;
                        }
                    }
                    modelCheckResults.push({ label: "Market Cap ÷ Price = share count used in the DCF", passed: pass, detail });
                }

                // 4. Implied DCF upside must not exceed 100% — a hard-block candidate once this
                // becomes enforced rather than advisory.
                {
                    let pass = null, detail = "Could not verify — upside row not found";
                    if (upsideCell) {
                        const upsideVal = upsideCell.values[0][0];
                        if (typeof upsideVal === "number") {
                            pass = upsideVal <= 1.0;
                            detail = `DCF implied upside = ${(upsideVal * 100).toFixed(1)}%` + (pass ? "" : " — exceeds 100%, review before publishing");
                        }
                    }
                    modelCheckResults.push({ label: "Implied DCF upside ≤ 100%", passed: pass, detail });
                }

                // 5. If a broker report was fed AND the model extracted an explicit broker target
                // price from it (see the conditional BROKER TARGET PRICE prompt instruction), the
                // model's own blended target price and DCF value/share must be positive and within
                // 2x (i.e. not more than ±100%) of that broker figure.
                {
                    let pass = null, detail = "No broker report was fed into this build";
                    if (brokerReportText) {
                        if (!brokerCell) {
                            detail = "Broker report was fed, but no explicit target price was found in it";
                        } else {
                            const brokerVal = brokerCell.values[0][0];
                            const targetVal = targetPriceCell ? targetPriceCell.values[0][0] : null;
                            const vpsVal = valuePerShareCell ? valuePerShareCell.values[0][0] : null;
                            if (typeof brokerVal === "number" && brokerVal > 0) {
                                const issues = [];
                                if (typeof targetVal === "number" && targetVal < 0) issues.push("blended target price is negative");
                                if (typeof vpsVal === "number" && vpsVal < 0) issues.push("DCF value per share is negative");
                                const checkRatio = (label, val) => {
                                    if (typeof val !== "number") return;
                                    const ratio = val / brokerVal;
                                    if (ratio > 2 || ratio < 0.5) issues.push(`${label} (Rs ${val.toFixed(2)}) differs from broker target (Rs ${brokerVal.toFixed(2)}) by more than 100%`);
                                };
                                checkRatio("Blended target price", targetVal);
                                checkRatio("DCF value per share", vpsVal);
                                pass = issues.length === 0;
                                detail = pass ? `Model's targets are within 100% of broker consensus (Rs ${brokerVal.toFixed(2)})` : issues.join("; ");
                            } else {
                                detail = "Broker report was fed, but no explicit target price was found in it";
                            }
                        }
                    }
                    modelCheckResults.push({ label: "Model targets vs broker consensus target price", passed: pass, detail });
                }

                // Tag these FM sheets with the company they were just built for (see FM_FINCODE_PROPERTY
                // above) — refreshBtn.onclick checks this before any future download to warn if it would
                // silently invalidate this model. Overwrite-safe: clear any existing property first,
                // since CustomPropertyCollection.add() throws if the key already exists.
                const oldFincodeProp = wb.properties.custom.getItemOrNullObject(FM_FINCODE_PROPERTY);
                const oldCompanyProp = wb.properties.custom.getItemOrNullObject(FM_COMPANY_PROPERTY);
                await context.sync();
                if (!oldFincodeProp.isNullObject) oldFincodeProp.delete();
                if (!oldCompanyProp.isNullObject) oldCompanyProp.delete();
                wb.properties.custom.add(FM_FINCODE_PROPERTY, String(fincode));
                wb.properties.custom.add(FM_COMPANY_PROPERTY, String(companyName));
                await context.sync();
            });

            // Dismissible popup summarizing the 5 automated model-quality checks computed above —
            // see renderModelChecks/#modelChecksModal. Also leaves a "View Model Checks" button
            // behind so this can be reopened at any time, not just right after the build.
            renderModelChecks(modelCheckResults);

            // (Re)start live current_price polling for the company just built — the workbook's
            // stored fincode may have just changed (a different company than whatever this timer
            // was previously polling for).
            startPriceRefreshTimer();

            // ── OpenRouter cost summary for this company (per-call + per-workflow + total) ──
            console.log(
                `[AI Model] 💰 TOTAL cost for ${companyName}: $${totalCost.toFixed(5)}  (${totalTokens} tokens across ${costRows.length} calls)\n`
                + `   Static  [${STATIC_MODEL}]: $${workflowTotals.static.cost.toFixed(5)}  (${workflowTotals.static.tokens} tokens, ${workflowTotals.static.calls} calls)\n`
                + `   Dynamic [${DYNAMIC_MODEL}]: $${workflowTotals.dynamic.cost.toFixed(5)}  (${workflowTotals.dynamic.tokens} tokens, ${workflowTotals.dynamic.calls} calls)\n`
                + costRows.join("\n")
            );
            deductModelCost({ fincode, companyName, costUsd: totalCost, usedFallback: usedFallbackModel });

            aiStatus(null);
            console.log("✅ AI Financial Model (multi-sheet, linked) written to Excel");
        } catch (err) {
            aiStatus(null);
            console.error("❌ handleBuildModel error:", err);
            showWarning(`Failed to build model: ${err.message}`);
        }
    }

    // --- AI Company Comparison ---
    async function handleCompare() {
        const OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions";
        const OPENROUTER_KEY = process.env.OPENROUTER_KEY; // injected at build time from .env — see webpack.config.js
        const AI_MODEL = "anthropic/claude-opus-4.8";

        const selected = window.__compareSelected || [];
        if (selected.length < 2) return showWarning("Select at least 2 companies to compare.");
        if (selected.length > 5) return showWarning("Maximum 5 companies for comparison.");

        const btn = document.getElementById("refreshBtn");
        const originalText = btn.textContent;
        btn.disabled = true;

        try {
            // Step 1: Fetch financials + operational data for each company in parallel
            btn.textContent = "Fetching data...";
            aiStatus(`Fetching financials, earning calls, broker reports & interviews for ${selected.length} companies...`);
            const companyTexts = await Promise.all(selected.map(async (comp) => {
                // Key financials — also picks up the Consolidated/Standalone fallback fix (see
                // fetchKeyFinancials). NOTE: this previously read finData?.data, but actuals_forwards
                // actually returns the row array under "value" (matching how handleRefresh's Key
                // Financials fetch already reads it) — "data" doesn't exist on this endpoint's
                // response at all, so finRows was silently empty for every company here before this.
                const finData = await fetchKeyFinancials(comp.fincode, comp.sector_type);
                const finRows = finData?.value || [];
                let finText = "(No financial data)";
                if (finRows.length) {
                    const keys = Object.keys(finRows[0]).filter(k => k !== "child" && k !== "parent_id");
                    finText = keys.join("\t") + "\n" + finRows.slice(0, 30).map(r => keys.map(k => r[k] ?? "").join("\t")).join("\n");
                }

                // Operational data
                let opText = "";
                try {
                    const opResp = await fetch("https://transcriptanalyser.com/pms/get_dashboard", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode: comp.fincode })
                    });
                    const opData = await opResp.json();
                    const categories = opData?.data || [];
                    if (categories.length) {
                        const lines = [];
                        categories.slice(0, 8).forEach(cat => {
                            lines.push(`[${cat.category || ""}]`);
                            (cat.metrics || []).slice(0, 10).forEach(m => {
                                const pts = (m.data_points || []).slice(-4).map(dp => `${dp.data_date}:${dp.value}`).join(", ");
                                lines.push(`  ${m.metric_name} (${m.unit || ""}) — ${pts}`);
                            });
                        });
                        opText = lines.join("\n");
                    }
                } catch (e) { /* operational data is optional */ }

                // Earning call insights (transcript summary + sentiment)
                let ecText = "";
                try {
                    const ecRes = await fetch("https://goindiainvest.in/mcp/earning_call_summaries", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincodes: [comp.fincode] })
                    });
                    if (ecRes.ok) ecText = await ecRes.text();
                } catch (e) { /* earning call data is optional */ }

                // Broker/analyst reports (recommendations, price targets)
                let brText = "";
                try {
                    const brRes = await fetch("https://goindiainvest.in/mcp/broker_report_single_comp", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode: comp.fincode })
                    });
                    if (brRes.ok) brText = await brRes.text();
                } catch (e) { /* broker data is optional */ }

                // Management interviews (strategic direction, capex plans)
                let miText = "";
                try {
                    const miRes = await fetch("https://goindiainvest.in/mcp/comp_management_interviews", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode: comp.fincode })
                    });
                    if (miRes.ok) miText = await miRes.text();
                } catch (e) { /* interview data is optional */ }

                return { name: comp.name, finText, opText, ecText, brText, miText };
            }));

            const companyNames = selected.map(c => c.name);
            const companyCount = companyNames.length;

            // Step 2: Build comparison prompt
            btn.textContent = "AI comparing...";
            aiStatusCycle([
                "Sending financials, earning calls, broker reports & interviews to AI...",
                "AI is analysing financial metrics & management sentiment...",
                "Comparing growth trajectories, analyst views & management strategy...",
                "Evaluating profitability & margins...",
                "Assessing operational efficiency...",
                "Analysing leverage & liquidity...",
                "Incorporating broker recommendations & price targets...",
                "Computing relative valuations...",
                "Drafting strengths & weaknesses...",
                "Finalising comparison..."
            ], 3500);

            let dataBlock = "";
            companyTexts.forEach(ct => {
                dataBlock += `\n=== ${ct.name} — KEY FINANCIALS ===\n${ct.finText}\n`;
                if (ct.opText) dataBlock += `\n=== ${ct.name} — OPERATIONAL DATA ===\n${ct.opText}\n`;
                if (ct.ecText) dataBlock += `\n=== ${ct.name} — EARNING CALL INSIGHTS ===\n${ct.ecText}\n`;
                if (ct.brText) dataBlock += `\n=== ${ct.name} — BROKER/ANALYST REPORTS ===\n${ct.brText}\n`;
                if (ct.miText) dataBlock += `\n=== ${ct.name} — MANAGEMENT INTERVIEWS ===\n${ct.miText}\n`;
            });

            const prompt = `Compare the following ${companyCount} Indian companies using all available data below — financials, operational metrics, management commentary from earning calls, professional analyst reports, and management interview insights. Analyse strengths, weaknesses, relative valuation, management quality, strategic direction, and operational performance.

${dataBlock}

Return ONLY a valid JSON object — no markdown, no explanation, no code fences. Each "values" array must have exactly ${companyCount} elements, one per company in this order: ${companyNames.join(", ")}.
{
  "companies":[${companyNames.map(n => `"${n}"`).join(",")}],
  "comparison_summary":"3-4 sentence comparative analysis",
  "verdict":"1-2 sentence conclusion on most attractive company",
  "overview":[
    {"label":"Revenue (Cr)","values":[n]},
    {"label":"PAT (Cr)","values":[n]},
    {"label":"Market Cap (Cr)","values":[n]},
    {"label":"Net Worth (Cr)","values":[n]},
    {"label":"Total Debt (Cr)","values":[n]}
  ],
  "growth":[
    {"label":"Revenue Growth 3Y CAGR (%)","values":[n]},
    {"label":"PAT Growth 3Y CAGR (%)","values":[n]},
    {"label":"EPS Growth 3Y CAGR (%)","values":[n]}
  ],
  "profitability":[
    {"label":"Gross Margin (%)","values":[n]},
    {"label":"EBITDA Margin (%)","values":[n]},
    {"label":"PAT Margin (%)","values":[n]},
    {"label":"ROE (%)","values":[n]},
    {"label":"ROCE (%)","values":[n]}
  ],
  "operational":[
    {"label":"metric_name","values":[n]}
  ],
  "efficiency":[
    {"label":"Asset Turnover (x)","values":[n]},
    {"label":"Receivable Days","values":[n]},
    {"label":"Inventory Days","values":[n]},
    {"label":"Payable Days","values":[n]}
  ],
  "leverage":[
    {"label":"Debt/Equity (x)","values":[n]},
    {"label":"Interest Coverage (x)","values":[n]},
    {"label":"Current Ratio (x)","values":[n]}
  ],
  "valuation":[
    {"label":"P/E (x)","values":[n]},
    {"label":"P/B (x)","values":[n]},
    {"label":"EV/EBITDA (x)","values":[n]},
    {"label":"Dividend Yield (%)","values":[n]}
  ],
  "analyst_consensus":[
    {"label":"Recommendation","values":["Buy/Hold/Sell"]},
    {"label":"Target Price (₹)","values":[n]},
    {"label":"Upside/Downside (%)","values":[n]}
  ],
  "management_sentiment":[
    {"label":"Overall Sentiment","values":["Positive/Neutral/Negative"]},
    {"label":"Growth Outlook","values":["text"]},
    {"label":"Key Guidance","values":["text"]}
  ],
  "strengths_weaknesses":[
    {"company":"Name","strengths":["s1","s2"],"weaknesses":["w1","w2"]}
  ]
}`;

            // Step 3: Call OpenRouter
            const aiRes = await fetch(OPENROUTER_URL, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${OPENROUTER_KEY}`
                },
                body: JSON.stringify({
                    model: AI_MODEL,
                    messages: [
                        { role: "system", content: "You are a senior equity research analyst at a top-tier investment bank. Return raw JSON only — no markdown, no text, no code fences." },
                        { role: "user", content: prompt }
                    ],
                    temperature: 0,
                    max_tokens: 32000
                })
            });
            if (!aiRes.ok) {
                const errText = await aiRes.text();
                throw new Error(`OpenRouter ${aiRes.status}: ${errText.slice(0, 300)}`);
            }

            const aiData = await aiRes.json();
            const finishReason = aiData.choices?.[0]?.finish_reason || "";
            let raw = aiData.choices?.[0]?.message?.content || "";
            console.log("[Compare] raw response:", raw.slice(0, 600));
            console.log("[Compare] finish_reason:", finishReason);
            raw = raw.replace(/^```(?:json)?\s*/im, "").replace(/\s*```$/im, "").trim();

            let model;
            try {
                model = JSON.parse(raw);
            } catch (parseErr) {
                let repaired = raw;
                const opens = (repaired.match(/[{[]/g) || []).length;
                const closes = (repaired.match(/[}\]]/g) || []).length;
                const lastGoodComma = repaired.lastIndexOf(",");
                if (lastGoodComma > repaired.length - 50) repaired = repaired.slice(0, lastGoodComma);
                for (let i = 0; i < opens - closes; i++) {
                    const lastOpen = Math.max(repaired.lastIndexOf("{"), repaired.lastIndexOf("["));
                    repaired += repaired[lastOpen] === "[" ? "]" : "}";
                }
                try {
                    model = JSON.parse(repaired);
                    console.warn("[Compare] Repaired truncated JSON. Some sections may be incomplete.");
                } catch (e2) {
                    console.error("[Compare] JSON parse failed. Raw:", raw.slice(0, 1000));
                    throw new Error("AI response was too long and got cut off. Try again.");
                }
            }
            console.log("[Compare] parsed OK. Keys:", Object.keys(model).join(", "));

            // Step 4: Write comparison sheet
            btn.textContent = "Writing sheet...";
            aiStatus("Writing comparison to Excel...");
            await Excel.run(async (context) => {
                const wb = context.workbook;
                wb.worksheets.load("items/name");
                await context.sync();

                const names = wb.worksheets.items.map(s => s.name);
                if (names.includes("Company Comparison")) {
                    wb.worksheets.getItem("Company Comparison").delete();
                }
                const sheet = wb.worksheets.add("Company Comparison");
                await context.sync();

                const cos = model.companies || companyNames;
                const colCount = cos.length;
                const lastCol = String.fromCharCode(65 + colCount); // B,C,D... for data cols; +1 for label col A

                sheet.getRange("A:A").format.columnWidth = 200;
                for (let i = 0; i < colCount; i++) {
                    sheet.getRange(`${String.fromCharCode(66 + i)}:${String.fromCharCode(66 + i)}`).format.columnWidth = 110;
                }

                const darkHdr = (row, text) => {
                    sheet.getRange(`A${row}`).values = [[text]];
                    const rng = sheet.getRange(`A${row}:${lastCol}${row}`);
                    rng.format.font.bold = true;
                    rng.format.font.size = 11;
                    rng.format.font.color = "#ffffff";
                    rng.format.fill.color = "#173760";
                };
                const companyHdrs = (row) => {
                    const vals = cos.slice(0, colCount);
                    const rng = sheet.getRange(`B${row}:${lastCol}${row}`);
                    rng.values = [vals];
                    rng.format.font.bold = true;
                    rng.format.font.color = "#173760";
                    rng.format.fill.color = "#d0d9f0";
                };
                const metricRow = (row, label, values, fmt) => {
                    sheet.getRange(`A${row}`).values = [[label]];
                    sheet.getRange(`A${row}`).format.font.size = 10;
                    const safe = Array.isArray(values) ? values.slice(0, colCount).map(x => x ?? "") : [];
                    while (safe.length < colCount) safe.push("");
                    if (safe.some(v => v !== "")) {
                        const rng = sheet.getRange(`B${row}:${lastCol}${row}`);
                        rng.values = [safe];
                        rng.format.font.size = 10;
                        if (fmt) rng.format.numberFormat = fmt;
                    }
                    if (row % 2 === 0) sheet.getRange(`A${row}:${lastCol}${row}`).format.fill.color = "#f9f9f9";
                };
                const writeSection = (startRow, sectionName, items, fmt) => {
                    darkHdr(startRow, sectionName);
                    companyHdrs(startRow);
                    let row = startRow + 1;
                    (items || []).forEach(item => {
                        metricRow(row, item.label, item.values, fmt);
                        row++;
                    });
                    return row + 1;
                };

                let r = 1;

                // Title
                sheet.getRange(`A${r}`).values = [[`Company Comparison — AI Analysis`]];
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.bold = true;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.size = 14;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.color = "#ffffff";
                sheet.getRange(`A${r}:${lastCol}${r}`).format.fill.color = "#173760";
                sheet.getRange(`A${r}`).format.rowHeight = 30;
                r++;

                sheet.getRange(`A${r}`).values = [[`${cos.join(" vs ")}  |  INR Crores  |  AI-generated — verify before use`]];
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.size = 9;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.color = "#777777";
                r += 2;

                // Comparison Summary
                darkHdr(r, "COMPARISON SUMMARY");
                r++;
                sheet.getRange(`A${r}`).values = [[model.comparison_summary || ""]];
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.size = 10;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.wrapText = true;
                sheet.getRange(`A${r}`).format.rowHeight = 54;
                r += 2;

                // Verdict
                darkHdr(r, "VERDICT");
                r++;
                sheet.getRange(`A${r}`).values = [[model.verdict || ""]];
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.size = 10;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.bold = true;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.fill.color = "#d9ead3";
                sheet.getRange(`A${r}:${lastCol}${r}`).format.wrapText = true;
                sheet.getRange(`A${r}`).format.rowHeight = 36;
                r += 2;

                // Metric sections
                r = writeSection(r, "OVERVIEW", model.overview, "#,##0");
                r = writeSection(r, "GROWTH", model.growth, "0.0");
                r = writeSection(r, "PROFITABILITY", model.profitability, "0.0");
                if (model.operational && model.operational.length) {
                    r = writeSection(r, "OPERATIONAL METRICS", model.operational, "0.0");
                }
                r = writeSection(r, "EFFICIENCY", model.efficiency, "0.0");
                r = writeSection(r, "LEVERAGE", model.leverage, "0.00");
                r = writeSection(r, "VALUATION", model.valuation, "0.0");
                if (model.analyst_consensus && model.analyst_consensus.length) {
                    r = writeSection(r, "ANALYST CONSENSUS", model.analyst_consensus, "@");
                }
                if (model.management_sentiment && model.management_sentiment.length) {
                    r = writeSection(r, "MANAGEMENT SENTIMENT", model.management_sentiment, "@");
                }

                // Strengths & Weaknesses
                darkHdr(r, "STRENGTHS & WEAKNESSES");
                companyHdrs(r);
                r++;
                const sw = model.strengths_weaknesses || [];
                const maxItems = Math.max(...sw.map(c => Math.max((c.strengths || []).length, (c.weaknesses || []).length)), 0);

                sheet.getRange(`A${r}`).values = [["Strengths"]];
                sheet.getRange(`A${r}`).format.font.bold = true;
                sheet.getRange(`A${r}`).format.font.size = 10;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.fill.color = "#e8edf3";
                r++;
                for (let si = 0; si < Math.max(2, maxItems); si++) {
                    const vals = cos.map((_, ci) => {
                        const c = sw.find(s => s.company === cos[ci]) || sw[ci] || {};
                        return (c.strengths || [])[si] || "";
                    });
                    metricRow(r, `  ${si + 1}.`, vals, "@");
                    r++;
                }
                r++;

                sheet.getRange(`A${r}`).values = [["Weaknesses"]];
                sheet.getRange(`A${r}`).format.font.bold = true;
                sheet.getRange(`A${r}`).format.font.size = 10;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.fill.color = "#e8edf3";
                r++;
                for (let wi = 0; wi < Math.max(2, maxItems); wi++) {
                    const vals = cos.map((_, ci) => {
                        const c = sw.find(s => s.company === cos[ci]) || sw[ci] || {};
                        return (c.weaknesses || [])[wi] || "";
                    });
                    metricRow(r, `  ${wi + 1}.`, vals, "@");
                    r++;
                }
                r++;

                // Disclaimer
                sheet.getRange(`A${r}`).values = [["DISCLAIMER: This comparison is AI-generated for informational purposes only and does not constitute investment advice. Metrics are estimates based on available data. Conduct your own due diligence before making investment decisions."]];
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.size = 8;
                sheet.getRange(`A${r}:${lastCol}${r}`).format.font.color = "#999999";
                sheet.getRange(`A${r}:${lastCol}${r}`).format.wrapText = true;
                sheet.getRange(`A${r}`).format.rowHeight = 40;

                await context.sync();
                sheet.activate();
            });

            aiStatus(null);
            console.log("✅ Company Comparison written to Excel");
        } catch (err) {
            aiStatus(null);
            console.error("❌ handleCompare error:", err);
            showWarning(`Failed to compare: ${err.message}`);
        } finally {
            btn.textContent = originalText;
            btn.disabled = false;
        }
    }

    // Attach Refresh button
    async function handleRefresh() {
        const toggle = document.getElementById("dropdownToggle");
        const fincode = toggle.dataset.value;
        const name = toggle.value;

        if (!fincode) return showWarning("Select a company first.");

        const sectorType = toggle.dataset.sector;

        // Read selected modes
        const indASChecked = document.getElementById("indASCheck").checked;
        const detailedChecked = document.getElementById("detailedCheck").checked;
        const operationalChecked = document.getElementById("operationalCheck").checked;
        if (!indASChecked && !detailedChecked && !operationalChecked) return showWarning("Please select at least one option (IndAS, Detailed, or Operational).");

        // Operational is independent of the IndAS/Detailed financial-statement flow below — fire
        // it in parallel rather than gating it on (or being gated by) that flow. It writes its own
        // "Operational Data" sheet and handles its own errors, so nothing more to do with it here.
        if (operationalChecked) handleOperationalRefresh().catch(e => console.error("Operational fetch failed:", e));

        const modes = [];
        if (indASChecked) modes.push({ label: "IndAS", suffix: "" });
        if (detailedChecked) modes.push({ label: "Detailed", suffix: "IND" });
        if (modes.length === 0) return; // Operational-only selection — nothing else to do

        try {
            // Key Financials — fetch once (independent of the IndAS/Detailed mode toggles below,
            // NOT independent of Consolidated/Standalone: see fetchKeyFinancials for why this can no
            // longer just always request mode:"C").
            const keyfData = await fetchKeyFinancials(fincode, sectorType);

            // PROBE — some companies (e.g. no subsidiaries) have ONLY a standalone (S) entry in
            // basic_info and no consolidated (C) one at all. Rather than fire 4-8 Consolidated
            // calls per format that are guaranteed to come back empty, fetch ONE cheap Consolidated
            // sheet up front (base/IndAS ProfitLoss) and use it to decide whether to bother with
            // Consolidated at all for the rest of this download.
            let hasConsolidated = true;
            let probePlCData = null;
            let probeSucceeded = false;
            try {
                const probeResp = await fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ fincode, mode: "C", sector: sectorType, sheet: "ProfitLoss" })
                });
                probePlCData = await probeResp.json();
                hasConsolidated = !isConsolidatedEmpty(probePlCData);
                probeSucceeded = true;
            } catch (e) {
                console.warn("Consolidated-availability probe failed — defaulting to fetching both Consolidated and Standalone.", e);
            }

            // Fetch mode-dependent data for each selected mode in parallel. When the probe found no
            // Consolidated data, skip every "C" fetch below entirely (cashC/qplC/bsC/plC) — only
            // Standalone is fetched, and the Consolidated fields are left null so the sheet-writer
            // knows to omit those sections rather than show an empty "not available" placeholder.
            const modeDatasets = await Promise.all(modes.map(async ({ label, suffix }) => {
                // The probe itself already answers plCData for the base (no-suffix/IndAS) format —
                // reuse it instead of re-fetching the identical sheet. Only when the probe actually
                // succeeded — if it failed outright, hasConsolidated fails open to true (fetch both,
                // same as before this change existed), but probePlCData is null, so it must NOT be
                // reused as a stand-in for a real fetch.
                const reusesProbe = probeSucceeded && hasConsolidated && suffix === "";
                const fetchC = (sheet) => hasConsolidated
                    ? fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                        method: "POST", headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ fincode, mode: "C", sector: sectorType, sheet })
                    })
                    : null;
                const fetchS = (sheet) => fetch("https://transcriptanalyser.com/goindiastock/annual_profitloss", {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ fincode, mode: "S", sector: sectorType, sheet })
                });

                const [cashCResp, cashSResp, qplCResp, qplSResp, bsCResp, bsSResp, plCResp, plSResp] =
                    await Promise.all([
                        fetchC(`CashFlow${suffix}`),
                        fetchS(`CashFlow${suffix}`),
                        fetchC(`QProfitLoss${suffix}`),
                        fetchS(`QProfitLoss${suffix}`),
                        fetchC(`BalanceSheet${suffix}`),
                        fetchS(`BalanceSheet${suffix}`),
                        reusesProbe ? null : fetchC(`ProfitLoss${suffix}`),
                        fetchS(`ProfitLoss${suffix}`),
                    ]);
                const [cashCData, cashSData, qplCData, qplSData, bsCData, bsSData, plCData, plSData] =
                    await Promise.all([
                        cashCResp ? cashCResp.json() : null, cashSResp.json(),
                        qplCResp ? qplCResp.json() : null, qplSResp.json(),
                        bsCResp ? bsCResp.json() : null, bsSResp.json(),
                        reusesProbe ? probePlCData : (plCResp ? plCResp.json() : null), plSResp.json(),
                    ]);
                return { label, cashCData, cashSData, qplCData, qplSData, bsCData, bsSData, plCData, plSData };
            }));

            // Fetch logo as base64 before entering Excel.run (fetch not allowed inside)
            let logoBase64 = null;
            try {
                const logoRes = await fetch("/assets/AppLogo_512x512.png");
                const blob = await logoRes.blob();
                logoBase64 = await new Promise((resolve) => {
                    const reader = new FileReader();
                    reader.onloadend = () => resolve(reader.result.split(",")[1]);
                    reader.readAsDataURL(blob);
                });
            } catch (e) {
                console.warn("Logo fetch failed, using text fallback.");
            }

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

                // --- Index Sheet ---
                const indexSheetName = "Index";
                let indexSheet;
                if (existingNames.includes(indexSheetName)) {
                    indexSheet = workbook.worksheets.getItem(indexSheetName);
                    indexSheet.getUsedRange()?.clear();
                    // Shapes are not cleared by getUsedRange().clear() — delete them explicitly
                    indexSheet.shapes.load("items");
                    await context.sync();
                    indexSheet.shapes.items.forEach(s => s.delete());
                } else {
                    indexSheet = workbook.worksheets.add(indexSheetName);
                }
                sheetsMap[indexSheetName] = indexSheet;

                // Move Index sheet to the first position
                indexSheet.position = 0;  

                // Setup appearance
                const topRow = indexSheet.getRange("A1:Z1");
                topRow.format.fill.color = "#173760";

                // Fix A1 height so we know exactly where A2 starts (shapes position from sheet top)
                const COL_W = 180; // pts — fixed column A width (wide enough for nav links)
                const A1_H  = 15;  // pts — thin blue header strip
                const A2_H  = 90;  // pts — tall enough to hold the logo with padding
                const LOGO   = 60; // pts — logo square size
                indexSheet.getRange("A1").format.rowHeight = A1_H;
                indexSheet.getRange("A2").format.rowHeight = A2_H;
                indexSheet.getRange("A:A").format.columnWidth = COL_W;

                // Insert logo — embed as base64 shape to bypass Excel's URL blocking
                if (logoBase64) {
                    const logoShape = indexSheet.shapes.addImage(logoBase64);
                    logoShape.width  = LOGO;
                    logoShape.height = LOGO;
                    // Center the logo within A2: vertically and horizontally
                    logoShape.top  = A1_H + (A2_H - LOGO) / 2;
                    logoShape.left = (COL_W - LOGO) / 2;
                } else {
                    const logoCell = indexSheet.getRange("A2");
                    logoCell.values = [["GO INDIA STOCKS"]];
                    logoCell.format.font.bold = true;
                    logoCell.format.font.size = 14;
                    logoCell.format.verticalAlignment = "Center";
                    logoCell.format.horizontalAlignment = "Center";
                }

                const titleCell = indexSheet.getRange("A3");
                titleCell.values = [["GO INDIA STOCKS"]];
                titleCell.format.font.bold = true;
                titleCell.format.font.size = 16;
                titleCell.format.horizontalAlignment = "CenterAcrossSelection";

                // Insert company name, updated date and model prepared text
                const nameCell = indexSheet.getRange("A4");
                nameCell.values = [[`Company: ${name}`]];
                nameCell.format.font.bold = true;
                nameCell.format.font.size = 12;
                nameCell.format.horizontalAlignment = "Left";

                const modelCell = indexSheet.getRange("A5");
                modelCell.values = [["Model prepared by Go India Stocks"]];
                modelCell.format.font.size = 11;
                modelCell.format.horizontalAlignment = "Left";

                const dateCell = indexSheet.getRange("A6");
                const today = new Date();
                const formattedDate = `${today.getDate().toString().padStart(2,"0")}-${(today.getMonth()+1).toString().padStart(2,"0")}-${today.getFullYear()}`;
                dateCell.values = [[`Updated as of ${formattedDate}`]];
                dateCell.format.font.size = 10;
                dateCell.format.horizontalAlignment = "Left";

                // Build index sections dynamically based on selected modes
                const sections = [
                    { name: "Key Financials", sheet: "Key Financials", marker: "Key Financials" },
                ];
                for (const { label } of modes) {
                    const ml = modes.length > 1 ? ` (${label})` : "";
                    sections.push({ name: `Quarterly Data - Quarterly P&L${ml}`, sheet: "Quarterly Data", marker: `Quarterly P&L${ml}` });
                }
                for (const { label } of modes) {
                    const ml = modes.length > 1 ? ` (${label})` : "";
                    sections.push(
                        { name: `Annual Data - Balance Sheet${ml}`, sheet: "Annual Data", marker: `Balance Sheet${ml}` },
                        { name: `Annual Data - Cash Flows${ml}`, sheet: "Annual Data", marker: `Cash Flows${ml}` },
                        { name: `Annual Data - Detailed P&L${ml}`, sheet: "Annual Data", marker: `Detailed P&L${ml}` }
                    );
                }

                // Add links to sheets created by other handlers (if they exist).
                // The "FM " sheets are the grouped, interlinked AI financial model.
                const extraSheets = [
                    "FM Summary", "FM Assumptions", "FM Operational", "FM P&L",
                    "FM Balance Sheet", "FM Capex & FCF", "FM Valuation & DCF",
                    "Operational Data", "Financial Model", "Company Comparison"
                ];
                for (const s of extraSheets) {
                    if (existingNames.includes(s)) {
                        sections.push({ name: s, sheet: s, marker: null });
                    }
                }

                // Insert navigation links starting from row 8
                let rowIdx = 8;
                for (const section of sections) {
                    const cell = indexSheet.getRange(`A${rowIdx}`);
                    cell.values = [[section.name]];

                    if (section.marker) {
                        cell.formulas = [[
                            `=HYPERLINK("#'${section.sheet}'!A" & MATCH("${section.marker}", '${section.sheet}'!A:A, 0), "${section.name}")`
                        ]];
                    } else {
                        cell.formulas = [[
                            `=HYPERLINK("#'${section.sheet}'!A1", "${section.name}")`
                        ]];
                    }

                    cell.format.font.color = "#1155CC";
                    cell.format.font.underline = "Single";
                    cell.format.horizontalAlignment = "Left";
                    rowIdx++;
                }

                // Align all content in A column from row 2 to the last one
                indexSheet.getRange(`A2:A${rowIdx - 1}`).format.horizontalAlignment = "CenterAcrossSelection";


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
                    keySheet.getRange("B:B").format.columnWidth = 90; 
                    keySheet.getRange(`B${startRow}`).values = [["(All values in Rs. Cr)"]];
                    keySheet.getRange(`B${startRow}`).format.font.bold = true;
                    keySheet.getRange(`B${startRow}`).format.horizontalAlignment = "CenterAcrossSelection";
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


                // --- Helper: Write one financial statement table ---
                const writeFinancialStatement = (sheet, startRow, data, sectionName, presentation) => {
                    // data === null means Consolidated was deliberately never fetched (the
                    // company has no "C" entry in basic_info, per the probe above) — write
                    // NOTHING for this section, not even a header or "not available" placeholder,
                    // so a standalone-only company's sheet shows Standalone data only.
                    if (presentation === "Consolidated" && data === null) return startRow;

                    const staticFields = ["Parameter"];

                    // --- Helper: get headers from data ---
                    // Accepts BOTH "Mar2024" (no separator, 4-digit year) and "Mar-24" (hyphenated,
                    // 2-digit year — a real format this API returns for at least some companies/
                    // tables) — missing the second form here would silently DROP those columns
                    // from the sheet entirely on a fresh download (worse than just a lookup miss
                    // downstream, which is what happens when buildIndex's own date-header regex,
                    // elsewhere in this file, has the same gap).
                    const MONYEAR_RE = /^[A-Z][a-z]{2}(\d{4}|-\d{2})$/;
                    const getHeaders = (data) => {
                        if (!data?.length) return [];
                        const headers = new Set();
                        data.forEach(row => {
                            Object.keys(row).forEach(k => {
                                if (!staticFields.includes(k) && MONYEAR_RE.test(k)) headers.add(k);
                            });
                        });
                        return Array.from(headers).sort((a, b) => {
                            const months = { Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5, Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11 };
                            const parseDate = s => {
                                const [_, mon, year] = s.match(/^([A-Za-z]+)-?(\d{2,4})$/) || [];
                                const y = year && year.length === 2 ? 2000 + parseInt(year) : parseInt(year);
                                return new Date(y, months[mon] ?? 0);
                            };
                            return parseDate(a) - parseDate(b);
                        });
                    };

                    // --- Flatten hierarchical data ---
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

                    // --- Section Title ---
                    sheet.getRange("B:B").format.columnWidth = 90; 
                    sheet.getRange(`B${startRow}`).values = [["(All values in Rs. Cr)"]];
                    
                    sheet.getRange(`B${startRow}`).format.font.bold = true;
                    sheet.getRange(`B${startRow}`).format.horizontalAlignment = "CenterAcrossSelection";
                    sheet.getRange(`A${startRow}`).values = [[sectionName]];
                    sheet.getRange(`A${startRow}`).format.font.bold = true;
                    sheet.getRange(`A${startRow}`).format.font.size = 14;
                    sheet.getRange(`A${startRow}`).format.fill.color = "#d9ead3";
                    sheet.getRange(`A${startRow}`).format.horizontalAlignment = "CenterAcrossSelection";
                    startRow++;

                    if (presentation === "Consolidated" && isConsolidatedEmpty(data)) {
                        sheet.getRange(`A${startRow}`).values = [["Consolidated not available for this company"]];
                        startRow += 2;
                        return startRow;
                    }

                    if (!data?.value?.length) {
                        sheet.getRange(`A${startRow}`).values = [[`${presentation} not available for this company`]];
                        startRow++;
                        return startRow;
                    }

                    const flatData = flattenHierarchicalData(data.value);
                    const headers = ["Parameter", ...getHeaders(data.value)];
                    const values = flatData.map(row => headers.map(h =>
                        h === "Parameter"
                            ? " ".repeat(row._level * 4) + (row.Parameter ?? "")
                            : row[h] ?? ""
                    ));

                    const dataStartRow = startRow;
                    startRow = formatTable(sheet, dataStartRow, presentation, headers, values);

                    // Smaller font for child rows
                    flatData.forEach((row, idx) => {
                        if (row._level > 0) {
                            const lastCol = getExcelColumnLetter(headers.length);
                            const r = sheet.getRange(`A${dataStartRow + 1 + idx}:${lastCol}${dataStartRow + 1 + idx}`);
                            r.format.font.size = 11;
                        }
                    });

                    return startRow;
                };


                // Write data for each selected mode
                let qRow = 3;
                let aRow = 3;
                const quarterlyStatements = [];
                const annualStatements = [];
                for (const dataset of modeDatasets) {
                    const ml = modes.length > 1 ? ` (${dataset.label})` : "";

                    quarterlyStatements.push({
                        sectionName: `Quarterly P&L${ml}`,
                        consolidatedData: dataset.qplCData,
                        standaloneData: dataset.qplSData
                    });

                    annualStatements.push(
                        { sectionName: `Balance Sheet${ml}`, consolidatedData: dataset.bsCData, standaloneData: dataset.bsSData },
                        { sectionName: `Cash Flows${ml}`, consolidatedData: dataset.cashCData, standaloneData: dataset.cashSData },
                        { sectionName: `Detailed P&L${ml}`, consolidatedData: dataset.plCData, standaloneData: dataset.plSData }
                    );
                }

                for (const statement of quarterlyStatements) {
                    qRow = writeFinancialStatement(sheetsMap["Quarterly Data"], qRow, statement.consolidatedData, statement.sectionName, "Consolidated");
                    qRow++;
                }

                for (const statement of quarterlyStatements) {
                    qRow = writeFinancialStatement(sheetsMap["Quarterly Data"], qRow, statement.standaloneData, statement.sectionName, "Standalone");
                    qRow++;
                }

                for (const statement of annualStatements) {
                    aRow = writeFinancialStatement(sheetsMap["Annual Data"], aRow, statement.consolidatedData, statement.sectionName, "Consolidated");
                    aRow++;
                }

                for (const statement of annualStatements) {
                    aRow = writeFinancialStatement(sheetsMap["Annual Data"], aRow, statement.standaloneData, statement.sectionName, "Standalone");
                    aRow++;
                }
                await context.sync();
            });

            console.log("✅ Data successfully written to Excel");

            fetchFinancialData();

        } catch (err) {
            console.error("❌ Error in refreshBtn:", err);
            showWarning("Failed to fetch company data. Check console for details.");
        }
    };
});


/* ═══════════════ DATAGPT CHAT (EXPERIMENTAL) ═══════════════
   One agent, DeepSeek V4 Pro (via OpenRouter), with two tool sources merged into a single
   tool-calling loop:
   1. GoIndia's own MCP server (goindia-mcp.fly.dev/sse) — connected to directly (the same
      server gia-chat2.js's own McpClient talks to; that file is a reference for the wire
      protocol only, its /api/chatai/gia-chat2 wrapper is a separate product with its own
      agent loop and is NOT called from here). Covers the full backend/routers/mcp.py data
      surface. Requires the user to hold an active MCP subscription (checked via dbcatalog's
      /get-mcp-key); a fresh MCP connection is opened per turn, mirroring gia-chat2.js's own
      "one client per chat turn" lifecycle.
   2. Local Excel tools (list/read/write the open workbook) — unchanged, still the only
      write path and still gated by the ask-before-edits confirmation flow.
   See the DataGPT plan doc for the original architecture proposal this has since diverged
   from (this file is the source of truth for the current design).
   To hide before deploy: set CHAT_ENABLED = false (the UI is removed and nothing
   is wired up), or comment out this whole block + the #chatFeature markup/style
   in taskpane.html. The financial-model feature does not depend on any of this. */
const CHAT_ENABLED = true; // re-enabled for local testing of the new Excel tools (get_workbook_overview/read_range/write_cells/fill_formula/recalculate_and_check/undo_last_change/trace_dependencies/search_workbook) — flip back to false before deploying if the wallet-feature update still isn't ready

(function setupAskClaudeChat() {
    if (!CHAT_ENABLED) {
        const f = document.getElementById("chatFeature");
        if (f) f.style.display = "none";
        return;
    }

    const OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions";
    const OPENROUTER_KEY = process.env.OPENROUTER_KEY; // injected at build time from .env — see webpack.config.js
    const AI_MODEL = "deepseek/deepseek-v4-pro"; // DataGPT — stronger multi-step tool-calling across the wider tool surface below

    const byId = (id) => document.getElementById(id);
    const fab = byId("chatFab"), panel = byId("chatPanel"), stream = byId("chatStream"),
        input = byId("chatInput"), sendBtn = byId("chatSend");
    const companyInput = byId("chatCompanyInput"), companyList = byId("chatCompanyList"),
        compareInput = byId("chatCompareInput"), compareList = byId("chatCompareList"), compareClear = byId("chatCompareClear");
    if (!fab || !panel) return;

    // Whether workbook edits are applied immediately or held for a per-write confirmation.
    // Defaults to OFF (ask first) — matches the "confirm before destructive edits" guardrail.
    // Presented as a compact mode-picker (button + popup with a tick by the active mode), the
    // same pattern Claude Code itself uses for its own mode switcher.
    const AUTO_APPLY_KEY = "goia_chatAutoApplyEdits";
    let autoApplyEdits = localStorage.getItem(AUTO_APPLY_KEY) === "1";
    const modeBtn = byId("chatModeBtn"), modeMenu = byId("chatModeMenu"),
        modeBtnLabel = byId("chatModeBtnLabel"), modeIcon = byId("chatModeIcon");
    const renderMode = () => {
        if (modeBtnLabel) modeBtnLabel.textContent = autoApplyEdits ? "Write automatically" : "Ask before writing";
        if (modeIcon) modeIcon.textContent = autoApplyEdits ? "⚡" : "✋";
        if (modeMenu) {
            modeMenu.querySelectorAll(".chat-mode-tick").forEach(t => {
                t.classList.toggle("active", t.dataset.tick === (autoApplyEdits ? "auto" : "ask"));
            });
        }
    };
    if (modeBtn && modeMenu) {
        renderMode();
        modeBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            modeMenu.classList.toggle("open");
        });
        modeMenu.querySelectorAll(".chat-mode-item").forEach(item => {
            item.addEventListener("click", () => {
                autoApplyEdits = item.dataset.mode === "auto";
                try { localStorage.setItem(AUTO_APPLY_KEY, autoApplyEdits ? "1" : "0"); } catch (e) { /* storage unavailable */ }
                renderMode();
                modeMenu.classList.remove("open");
            });
        });
        document.addEventListener("click", (e) => {
            if (modeMenu.classList.contains("open") && !modeMenu.contains(e.target) && e.target !== modeBtn) {
                modeMenu.classList.remove("open");
            }
        });
    }

    // Chat history survives taskpane reloads (window close, Excel restart) via localStorage.
    const CHAT_HISTORY_KEY = "goia_chatHistory";
    const MAX_HISTORY = 40;
    let history = [];                          // {role, content}
    try {
        const saved = JSON.parse(localStorage.getItem(CHAT_HISTORY_KEY) || "[]");
        if (Array.isArray(saved)) history = saved.slice(-MAX_HISTORY);
    } catch (e) { /* ignore corrupt data */ }
    const persistHistory = () => {
        try { localStorage.setItem(CHAT_HISTORY_KEY, JSON.stringify(history.slice(-MAX_HISTORY))); } catch (e) { /* storage full/unavailable */ }
    };
    let busy = false;

    // The chat keeps its own company selection so the user can switch companies (or add a
    // second one to compare) without leaving the panel. "primaryCo" mirrors the main
    // Company-tab dropdown until the user picks a company inside the chat itself.
    let primaryCo = null;      // {fincode, name, sector}
    let compareCo = null;      // {fincode, name, sector} | null
    let allCompanies = [];     // lazy-loaded once, reused by both pickers

    async function ensureCompanyList() {
        if (allCompanies.length) return allCompanies;
        try {
            const res = await fetch("https://transcriptanalyser.com/operational/companies_excel");
            allCompanies = await res.json();
        } catch (e) { allCompanies = []; }
        return allCompanies;
    }

    // Shared wiring for the "Acting on" / "Compare with" search boxes.
    function wireCompanyPicker(inputEl, listEl, onPick) {
        const render = (items) => {
            listEl.innerHTML = "";
            items.slice(0, 25).forEach(c => {
                const li = document.createElement("li");
                li.textContent = c.CompName;
                li.addEventListener("mousedown", (e) => {
                    e.preventDefault(); // fire before the input's blur tears the list down
                    onPick({ fincode: c.fincode, name: c.CompName, sector: c.sector_type });
                    inputEl.value = c.CompName;
                    listEl.classList.remove("open");
                });
                listEl.appendChild(li);
            });
            listEl.classList.toggle("open", items.length > 0);
        };
        inputEl.addEventListener("focus", async () => {
            inputEl.select();
            render(await ensureCompanyList());
        });
        inputEl.addEventListener("input", async () => {
            const q = inputEl.value.toLowerCase();
            const all = await ensureCompanyList();
            render(q ? all.filter(c => c.CompName.toLowerCase().includes(q)) : all);
        });
        inputEl.addEventListener("blur", () => setTimeout(() => listEl.classList.remove("open"), 150));
    }
    wireCompanyPicker(companyInput, companyList, (c) => { primaryCo = c; });
    wireCompanyPicker(compareInput, compareList, (c) => { compareCo = c; compareClear.style.display = "inline"; });
    compareClear.addEventListener("click", () => { compareCo = null; compareInput.value = ""; compareClear.style.display = "none"; });

    // Mirrors the main Company-tab dropdown until the user picks a company in-chat.
    const setCompany = () => {
        if (primaryCo) { companyInput.value = primaryCo.name; return; }
        const t = byId("dropdownToggle");
        if (t && t.value && t.dataset.value) {
            primaryCo = { fincode: t.dataset.value, name: t.value, sector: t.dataset.sector };
            companyInput.value = t.value;
        }
    };
    // Lightweight markdown renderer for DataGPT's own replies (headers, **bold**, *italics*,
    // `code`, bullet/numbered lists, --- rules, paragraph breaks) — the model's answers were
    // arriving as raw markdown text via textContent, showing literal "**" and "###" to the user.
    // User-typed messages are untouched (still plain textContent below) since they need no
    // rendering and shouldn't have their own literal * or # reinterpreted as formatting.
    const escapeHtml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const renderInline = (s) => {
        s = escapeHtml(s);
        s = s.replace(/`([^`]+)`/g, "<code>$1</code>");
        s = s.replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>");
        s = s.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, "$1<i>$2</i>");
        return s;
    };
    const renderMarkdown = (text) => {
        if (!text) return "";
        const lines = String(text).replace(/\r\n/g, "\n").split("\n");
        let html = "", listType = null, para = [];
        const closeList = () => { if (listType) { html += listType === "ul" ? "</ul>" : "</ol>"; listType = null; } };
        const flushPara = () => { if (para.length) { html += "<p>" + para.map(renderInline).join("<br>") + "</p>"; para = []; } };
        for (const raw of lines) {
            const line = raw.trim();
            let m;
            if (!line) { flushPara(); closeList(); }
            else if (/^(-{3,}|\*{3,})$/.test(line)) { flushPara(); closeList(); html += "<hr>"; }
            else if ((m = /^(#{1,4})\s+(.*)$/.exec(line))) { flushPara(); closeList(); html += `<div class="md-h md-h${m[1].length}">${renderInline(m[2])}</div>`; }
            else if ((m = /^[-*]\s+(.*)$/.exec(line))) { flushPara(); if (listType !== "ul") { closeList(); html += "<ul>"; listType = "ul"; } html += `<li>${renderInline(m[1])}</li>`; }
            else if ((m = /^\d+\.\s+(.*)$/.exec(line))) { flushPara(); if (listType !== "ol") { closeList(); html += "<ol>"; listType = "ol"; } html += `<li>${renderInline(m[1])}</li>`; }
            else { closeList(); para.push(line); }
        }
        flushPara(); closeList();
        return html;
    };
    const addMsg = (role, text) => {
        const m = document.createElement("div");
        m.className = "chat-msg " + role;
        const b = document.createElement("div");
        b.className = "chat-bubble";
        if (role === "bot") b.innerHTML = renderMarkdown(text);
        else b.textContent = text;
        m.appendChild(b);
        stream.appendChild(m);
        stream.scrollTop = stream.scrollHeight;
        return b;
    };
    const addCard = (text, cls) => {
        const c = document.createElement("div");
        c.className = "chat-card " + (cls || "run");
        c.dataset.createdAt = Date.now(); // read by settleCard below
        const t = document.createElement("span"); t.textContent = text;
        const s = document.createElement("span"); s.className = "chat-card-st";
        c.appendChild(t); c.appendChild(s);
        stream.appendChild(c);
        stream.scrollTop = stream.scrollHeight;
        return c;
    };
    // Flips a card from its pulsing "run" state to done/err. In-memory Excel API calls and cached
    // MCP lookups routinely finish inside a single microtask — fast enough that the browser never
    // actually paints the pulsing state before it's already flipped to done, so the card just
    // "pops in" already finished instead of visibly working. This tops up whatever's left of a
    // minimum visible run time first, so every card gets at least one real animated frame.
    const MIN_CARD_RUN_MS = 450;
    async function settleCard(card, cls, symbol) {
        if (!card) return; // some tools only show a card conditionally (e.g. toolNamedRanges)
        const elapsed = Date.now() - Number(card.dataset.createdAt || 0);
        if (elapsed < MIN_CARD_RUN_MS) await new Promise(r => setTimeout(r, MIN_CARD_RUN_MS - elapsed));
        card.className = "chat-card " + cls;
        const st = card.querySelector(".chat-card-st");
        if (st) st.textContent = symbol;
    }
    // Renders an inline approve/skip prompt and resolves once the user picks one —
    // lets the tool-calling loop below simply `await` a write decision like any other I/O.
    function addConfirmCard(text, opts) {
        const { approveLabel = "Write", rejectLabel = "Skip" } = opts || {};
        const c = document.createElement("div");
        c.className = "chat-confirm";
        const t = document.createElement("div"); t.className = "cf-text"; t.textContent = text;
        const actions = document.createElement("div"); actions.className = "cf-actions";
        const approve = document.createElement("button"); approve.className = "cf-approve"; approve.textContent = approveLabel;
        const reject = document.createElement("button"); reject.className = "cf-reject"; reject.textContent = rejectLabel;
        actions.appendChild(approve); actions.appendChild(reject);
        c.appendChild(t); c.appendChild(actions);
        stream.appendChild(c);
        stream.scrollTop = stream.scrollHeight;
        return new Promise((resolve) => {
            approve.addEventListener("click", () => {
                actions.remove(); c.classList.add("cf-done"); t.textContent = "✓ Approved — " + text;
                resolve(true);
            });
            reject.addEventListener("click", () => {
                actions.remove(); c.classList.add("cf-skipped"); t.textContent = "Cancelled — " + text;
                resolve(false);
            });
        });
    }

    const greet = () => addMsg("bot",
        "Hi — I'm connected to your GoIndia data and this workbook. Ask me about the selected company, or to write something into a sheet. (Experimental)");

    // Re-render whatever chat history survived from the last session.
    history.forEach(m => addMsg(m.role === "assistant" ? "bot" : "user", m.content));

    const openPanel = () => { panel.style.display = "flex"; fab.style.display = "none"; setCompany(); if (!stream.childElementCount) greet(); input.focus(); };
    const closePanel = () => { panel.style.display = "none"; fab.style.display = ""; if (modeMenu) modeMenu.classList.remove("open"); };
    fab.addEventListener("click", openPanel);

    // Ribbon handshake: the "Ask DataGPT" button's ShowTaskpane action points at
    // taskpane.html?view=datagpt (see manifest.xml and applyRibbonView near the top of this file,
    // which already hid every #addinUI section except #chatFeature for this exact view) —
    // auto-open the chat here too so the pane opens with DataGPT already in front and nothing
    // else, rather than an empty pane with just the fab to click.
    // Read ?view= directly rather than document.body.dataset.view: this whole IIFE runs
    // synchronously at script-parse time, before the Office.onReady callback that sets that
    // dataset attribute (applyRibbonView, above) has had a chance to fire — relying on it here
    // was a real race that left DataGPT's panel closed behind a fab the user had to click anyway.
    let initialView = "main";
    try { initialView = new URLSearchParams(location.search).get("view") || "main"; } catch (e) { /* malformed query string */ }
    if (initialView === "datagpt") openPanel();
    byId("chatClose").addEventListener("click", closePanel);
    const chatClear = byId("chatClear");
    if (chatClear) {
        chatClear.addEventListener("click", async () => {
            const ok = await addConfirmCard("Clear all chat history?", { approveLabel: "Clear", rejectLabel: "Cancel" });
            if (!ok) return;
            history = [];
            persistHistory();
            stream.innerHTML = "";
            greet();
        });
    }
    input.addEventListener("input", () => { input.style.height = "auto"; input.style.height = Math.min(96, input.scrollHeight) + "px"; });
    input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } });
    sendBtn.addEventListener("click", send);
    const ddToggle = byId("dropdownToggle");
    if (ddToggle) ddToggle.addEventListener("blur", setCompany);

    // ── GoIndia MCP access ──
    // TEMPORARY, local-testing-only: goindia-mcp.fly.dev has no CORS headers on /sse or
    // /messages/, so the browser blocks this add-in from calling it directly (curl/server-to-
    // server calls were never subject to that check — only cross-origin fetch() from a webpage
    // is). Pointed at the wallet backend's new /wallet/mcp/sse proxy (routers/wallet.py) instead,
    // which does the actual goindia-mcp hop server-to-server and relays the stream back — sidesteps
    // the problem without needing a fix on a server we don't control. Once this is confirmed
    // working against localhost:8000, point this at the equivalent proxy route once it's live on
    // transcriptanalyser.com, and revert this back to the direct goindia-mcp.fly.dev URL only if/
    // when that server actually adds its own CORS headers.
    const MCP_SERVER_URL = "https://transcriptanalyser.com/wallet/mcp/sse";
    const MCP_KEY_URL = "https://transcriptanalyser.com/dbcatalog/get-mcp-key";
    const TEST_MCP_KEY = process.env.TEST_MCP_KEY || ""; // injected at build time from .env — see webpack.config.js
    let mcpAccess = null; // { mcp_api_key, flag } | null — fetched once per panel session
    async function ensureMcpAccess() {
        if (mcpAccess) return mcpAccess;
        try {
            const user = JSON.parse(localStorage.getItem("user") || "{}");
            // "test123" is the Microsoft-certification bypass account — not a real DB user, so
            // /get-mcp-key can't resolve it — use the injected test key instead. Real accounts
            // (including UserId 6 / rakeshadmin@goindiaadvisors.com, now granted real MCP
            // access) go through the normal live lookup below, so their key is never stale.
            if (user.UserId === "test123") {
                return (mcpAccess = TEST_MCP_KEY ? { mcp_api_key: TEST_MCP_KEY, flag: 3 } : { mcp_api_key: null, flag: 0 });
            }
            if (!user.UserId) return (mcpAccess = { mcp_api_key: null, flag: 0 });
            const res = await fetch(`${MCP_KEY_URL}?user_id=${encodeURIComponent(user.UserId)}`);
            if (!res.ok) return (mcpAccess = { mcp_api_key: null, flag: 0 });
            const data = await res.json();
            mcpAccess = { mcp_api_key: data.mcp_api_key || null, flag: data.flag };
        } catch (e) { mcpAccess = { mcp_api_key: null, flag: 0 }; }
        return mcpAccess;
    }

    // ── MCP client (ported from gia-chat2.js's McpClient) ──
    // MCP's HTTP+SSE transport: 1) open an SSE GET, 2) the server sends an "endpoint" event
    // whose data is a relative URL to POST JSON-RPC requests to, 3) each request's response
    // arrives back over that same SSE stream, matched by request id. One client per chat turn
    // — opened at the start of send(), closed in its finally block.
    class McpClient {
        constructor(sseUrl, signal) {
            this.sseUrl = sseUrl;
            this.signal = signal;
            this.postUrl = null;
            this.id = 0;
            this.tools = [];
            this.pending = new Map();
            this._endpointResolve = null;
            this._endpointReject = null;
            this._endpointPromise = new Promise((resolve, reject) => {
                this._endpointResolve = resolve;
                this._endpointReject = reject;
            });
            this._endpointPromise.catch(() => {});
            this._closed = false;
        }
        _nextId() { return ++this.id; }
        async open() {
            const res = await fetch(this.sseUrl, { method: "GET", headers: { "Accept": "text/event-stream" }, signal: this.signal });
            if (!res.ok || !res.body) {
                const text = await res.text().catch(() => "");
                throw new Error(`MCP SSE open HTTP ${res.status}: ${text.slice(0, 300)}`);
            }
            this._readSse(res.body).catch((err) => {
                if (!this._closed) {
                    this._endpointReject?.(err);
                    for (const { reject } of this.pending.values()) reject(new Error(`MCP SSE stream closed: ${err.message ?? err}`));
                    this.pending.clear();
                }
            });
            this.postUrl = await this._endpointPromise;
        }
        async _readSse(body) {
            const reader = body.getReader();
            const decoder = new TextDecoder("utf-8");
            let buf = "", evName = "", evData = "";
            const flushEvent = () => {
                if (!evName && !evData) return;
                this._handleSseEvent(evName || "message", evData);
                evName = ""; evData = "";
            };
            while (true) {
                const { value, done } = await reader.read();
                if (done) break;
                buf += decoder.decode(value, { stream: true });
                let nl;
                while ((nl = buf.indexOf("\n")) !== -1) {
                    const line = buf.slice(0, nl).replace(/\r$/, "");
                    buf = buf.slice(nl + 1);
                    if (line === "") { flushEvent(); continue; }
                    if (line.startsWith(":")) continue;
                    if (line.startsWith("event:")) evName = line.slice(6).trim();
                    else if (line.startsWith("data:")) { const chunk = line.slice(5).replace(/^ /, ""); evData = evData ? `${evData}\n${chunk}` : chunk; }
                }
            }
            flushEvent();
        }
        _handleSseEvent(name, data) {
            if (name === "endpoint") {
                let resolved;
                try { resolved = new URL(data, this.sseUrl).toString(); } catch { resolved = data; }
                this._endpointResolve?.(resolved);
                return;
            }
            if (name === "message" || name === "") {
                if (!data) return;
                let obj;
                try { obj = JSON.parse(data); } catch { return; }
                if (obj && Object.prototype.hasOwnProperty.call(obj, "id") && this.pending.has(obj.id)) {
                    const { resolve } = this.pending.get(obj.id);
                    this.pending.delete(obj.id);
                    resolve(obj);
                }
            }
        }
        async _rpc(method, params) {
            if (!this.postUrl) throw new Error("MCP client not opened");
            const id = this._nextId();
            const body = { jsonrpc: "2.0", id, method, params: params ?? {} };
            const responsePromise = new Promise((resolve, reject) => { this.pending.set(id, { resolve, reject }); });
            responsePromise.catch(() => {});
            let res;
            try {
                res = await fetch(this.postUrl, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body), signal: this.signal });
            } catch (err) { this.pending.delete(id); throw err; }
            if (!res.ok && res.status !== 202) {
                this.pending.delete(id);
                const text = await res.text().catch(() => "");
                throw new Error(`MCP ${method} POST HTTP ${res.status}: ${text.slice(0, 300)}`);
            }
            try {
                const ctype = res.headers.get("content-type") || "";
                if (ctype.includes("application/json")) {
                    const inline = await res.json();
                    if (inline && inline.id === id) {
                        this.pending.delete(id);
                        if (inline.error) throw new Error(`MCP ${method} error: ${inline.error.message ?? JSON.stringify(inline.error)}`);
                        return inline.result;
                    }
                }
            } catch { /* fall through to SSE-delivered response */ }
            const payload = await responsePromise;
            if (payload.error) throw new Error(`MCP ${method} error: ${payload.error.message ?? JSON.stringify(payload.error)}`);
            return payload.result;
        }
        async _notify(method, params) {
            if (!this.postUrl) return;
            try { await fetch(this.postUrl, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ jsonrpc: "2.0", method, params: params ?? {} }), signal: this.signal }); }
            catch { /* best-effort */ }
        }
        async initialize() {
            const result = await this._rpc("initialize", { protocolVersion: "2025-03-26", capabilities: {}, clientInfo: { name: "goindia-excel-addin", version: "1.0.0" } });
            await this._notify("notifications/initialized");
            return result;
        }
        async listTools() {
            const result = await this._rpc("tools/list", {});
            this.tools = (result?.tools ?? []).map((t) => ({ name: t.name, description: t.description ?? "", input_schema: t.inputSchema ?? t.input_schema ?? { type: "object", properties: {} } }));
            return this.tools;
        }
        async callTool(name, args) {
            const result = await this._rpc("tools/call", { name, arguments: args ?? {} });
            const raw = result?.content ?? "";
            const text = typeof raw === "string" ? raw : Array.isArray(raw) ? raw.map((c) => (typeof c === "string" ? c : (c?.text ?? ""))).join("") : JSON.stringify(raw);
            const isError = !!result?.isError || text.includes("Query execution failed") || text.includes("DatabaseError") || text.includes("Error:") || text.startsWith("❌");
            return { text, isError };
        }
        async close() { this._closed = true; this.pending.clear(); }
    }
    const toOpenAITools = (mcpTools) => mcpTools.map((t) => ({ type: "function", function: { name: t.name, description: t.description, parameters: t.input_schema } }));

    // ── Excel-side tools ──
    // Ported from sbroenne/pi-for-excel (src/tools/*.ts, src/excel/helpers.ts) — a real Office.js
    // add-in that implements this same tool set. Adapted to plain JS (this file has no TypeScript/
    // TypeBox) and this file's existing style: OpenAI function-calling schemas in EXCEL_TOOLS,
    // dispatched by callTool(name, args), with mutating tools managing their own confirm/running
    // card lifecycle exactly like the old write_excel_sheet did (see the dispatch split in send()
    // below) rather than the generic per-tool card the read-only tools get there.
    const TEXT_CAP = 3500;
    const capText = (s, n = TEXT_CAP) => {
        s = (s || "").trim();
        return s.length > n ? s.slice(0, n) + "\n…[truncated]" : s;
    };

    // Address/range math — 0-indexed column, 1-indexed row, mirrors pi-for-excel's helpers.ts.
    const chatColToLetter = (col) => { let letter = "", c = col; while (c >= 0) { letter = String.fromCharCode((c % 26) + 65) + letter; c = Math.floor(c / 26) - 1; } return letter; };
    const chatLetterToCol = (letters) => { let col = 0; for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64); return col - 1; };
    const chatParseCell = (cell) => {
        const clean = cell.includes("!") ? cell.slice(cell.indexOf("!") + 1) : cell;
        const m = clean.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
        if (!m) throw new Error(`Invalid cell address: ${cell}`);
        return { col: chatLetterToCol(m[1].toUpperCase()), row: parseInt(m[2], 10) };
    };
    const chatCellAddress = (col, row) => `${chatColToLetter(col)}${row}`;
    const chatComputeRangeAddress = (startCell, rows, cols) => {
        const { col, row } = chatParseCell(startCell);
        return `${startCell}:${chatCellAddress(col + cols - 1, row + rows - 1)}`;
    };
    const chatParseRangeRef = (ref) => {
        if (ref.includes("!")) {
            const idx = ref.indexOf("!");
            const sheet = ref.slice(0, idx).replace(/^'|'$/g, "").replace(/''/g, "'");
            return { sheet, address: ref.slice(idx + 1) };
        }
        return { sheet: null, address: ref };
    };
    const chatQualifiedAddress = (sheetName, address) => {
        const clean = address.includes("!") ? address.slice(address.indexOf("!") + 1) : address;
        const needsQuote = /[\s']/.test(sheetName);
        return `${needsQuote ? `'${sheetName.replace(/'/g, "''")}'` : sheetName}!${clean}`;
    };
    const chatGetRangeAndSheet = (ctx, ref) => {
        const parsed = chatParseRangeRef(ref);
        const sheet = parsed.sheet ? ctx.workbook.worksheets.getItem(parsed.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
        return { sheet, range: sheet.getRange(parsed.address) };
    };
    // A cell VALUE starting with "#" (e.g. "#REF!", "#DIV/0!") is Excel's own error rendering —
    // simplest reliable way to detect a formula error without a special-cells API call.
    const chatFindErrors = (values, startAddress) => {
        const errors = [];
        const start = chatParseCell(startAddress);
        for (let r = 0; r < (values || []).length; r++) {
            const row = values[r] || [];
            for (let c = 0; c < row.length; c++) {
                const v = row[c];
                if (typeof v === "string" && v.startsWith("#")) errors.push({ address: chatCellAddress(start.col + c, start.row + r), error: v });
            }
        }
        return errors;
    };
    const chatCountOccupied = (values, formulas) => {
        let n = 0;
        for (let r = 0; r < (values || []).length; r++) {
            const vr = values[r] || [], fr = (formulas && formulas[r]) || [];
            for (let c = 0; c < vr.length; c++) {
                const hasValue = vr[c] !== null && vr[c] !== undefined && vr[c] !== "";
                const hasFormula = typeof fr[c] === "string" && fr[c].startsWith("=");
                if (hasValue || hasFormula) n++;
            }
        }
        return n;
    };
    const chatPadValues = (values) => {
        const cols = Math.max(1, ...values.map(r => r.length));
        return { padded: values.map(row => { const r = row.slice(); while (r.length < cols) r.push(""); return r; }), rows: values.length, cols };
    };
    // Same syntax checks pi-for-excel's write-cells.ts runs before ever touching Excel — catches
    // an unbalanced paren/quote or a formula that trails an operator client-side, instead of
    // finding out only after the cell renders #NAME?/#VALUE!.
    const chatValidateFormula = (formula) => {
        if (!formula.startsWith("=")) return null;
        const body = formula.slice(1);
        if (body.trim().length === 0) return "Empty formula";
        if (((body.match(/"/g) || []).length) % 2 !== 0) return "Unbalanced quotes";
        let depth = 0, inString = false;
        for (const ch of body) {
            if (ch === "\"") { inString = !inString; continue; }
            if (inString) continue;
            if (ch === "(") depth++;
            if (ch === ")") { depth--; if (depth < 0) return "Unbalanced parentheses"; }
        }
        if (depth !== 0) return "Unbalanced parentheses";
        if (/[+\-*/^&,]$/.test(body.trim())) return "Formula ends with an operator";
        return null;
    };
    const formatMiniTable = (rows) => {
        if (!rows || !rows.length) return "(empty)";
        return rows.map(row => "| " + row.map(v => (v === null || v === undefined || v === "") ? "" : String(v)).join(" | ") + " |").join("\n");
    };

    // Undo support: a small in-memory stack of "before" snapshots, one per successful write_cells/
    // fill_formula call, so undo_last_change can restore exactly what was there — Ctrl+Z doesn't
    // reliably undo programmatic Office.js writes, so this is the only real safety net for a bad
    // AI edit. Capped at 10 entries; oldest falls off first. Cleared on panel reload (session-only,
    // same lifetime as chat history's in-memory state before persistHistory).
    const undoStack = [];
    const pushUndoSnapshot = (snap) => { undoStack.push(snap); if (undoStack.length > 10) undoStack.shift(); };

    const EXCEL_TOOLS = [
        { type: "function", function: { name: "get_workbook_overview", description: "Get a structural overview of the workbook: sheet names, dimensions, header rows, named ranges, tables, and chart/pivot counts. Call this at the start of a task, or when you need to understand the workbook's structure before reading specific ranges. Pass sheet_name for a detailed view of one sheet (headers, tables, named ranges, a data preview).", parameters: { type: "object", properties: { sheet_name: { type: "string", description: "Optional — return detailed info for this one sheet instead of the whole-workbook summary." } } } } },
        { type: "function", function: { name: "read_range", description: "Read cell values, formulas, and number formats from a range in the user's open workbook (e.g. \"A1:D10\" or \"Sheet2!B3:B20\"). Always call this before writing a formula that references existing cells — never guess row/column positions.", parameters: { type: "object", properties: { range: { type: "string", description: "A1-notation range, optionally sheet-qualified, e.g. \"A1:D10\" or \"Sheet2!A1:B5\". If no sheet is given, uses the active sheet." } }, required: ["range"] } } },
        { type: "function", function: { name: "create_sheet", description: "Create a new, blank worksheet. write_cells/fill_formula only work on sheets that already exist — call this first if the sheet the user wants doesn't exist yet.", parameters: { type: "object", properties: { name: { type: "string", description: "Name for the new sheet (max 31 characters, can't contain \\ / ? * [ ] or :)." }, activate: { type: "boolean", description: "Switch the workbook view to the new sheet after creating it. Default true." } }, required: ["name"] } } },
        { type: "function", function: { name: "write_cells", description: "Write values and/or formulas (strings starting with \"=\") to a range of cells, starting at start_cell. Only touches the cells you specify — everything else on the sheet is left alone. Blocks by default if the target already has data; only pass allow_overwrite:true after the user has actually confirmed they want that data overwritten.", parameters: { type: "object", properties: { start_cell: { type: "string", description: "Top-left cell to write from, e.g. \"A1\" or \"Sheet2!B3\"." }, values: { type: "array", items: { type: "array" }, description: "2D array of values/formulas. Each inner array is a row." }, allow_overwrite: { type: "boolean", description: "Set true to overwrite existing data in the target range. Default false." } }, required: ["start_cell", "values"] } } },
        { type: "function", function: { name: "fill_formula", description: "Fill one formula across a range using Excel's own AutoFill — relative references adjust per cell automatically, the same as dragging the fill handle. Use this instead of building a large 2D formula array by hand.", parameters: { type: "object", properties: { range: { type: "string", description: "Target range to fill, e.g. \"B2:B20\" or \"Sheet1!C3:F20\". Single contiguous range only." }, formula: { type: "string", description: "Formula to fill, starting with \"=\", e.g. \"=SUM(B2:B10)\"." }, allow_overwrite: { type: "boolean", description: "Set true to overwrite existing data in the target range. Default false." } }, required: ["range", "formula"] } } },
        { type: "function", function: { name: "recalculate_and_check", description: "Force Excel to recalculate every formula, then scan for formula errors (#REF!, #DIV/0!, #VALUE!, etc.). Call this after write_cells/fill_formula to confirm an edit didn't break anything downstream, instead of just assuming it worked.", parameters: { type: "object", properties: { sheet_name: { type: "string", description: "Optional — scope the error scan to just this sheet. Omit to scan every visible sheet." } } } } },
        { type: "function", function: { name: "audit_workbook", description: "Proactively check whether the workbook/model is actually healthy — not just after a specific edit. Recalculates, then scans for formula errors AND for any \"check\"/reconciliation row (e.g. a balance check) whose value isn't close to zero the way it should be. Use this for a request like \"is my model healthy?\" or \"does this balance?\", not just reactively after write_cells.", parameters: { type: "object", properties: { sheet_name: { type: "string", description: "Optional — scope the audit to just this sheet. Omit to audit every visible sheet." }, tolerance: { type: "number", description: "How far from zero a check-row value can be before it's flagged. Default 0.01 (covers ordinary floating-point rounding)." } } } } },
        { type: "function", function: { name: "undo_last_change", description: "Revert the most recent write_cells or fill_formula change made this session, restoring exactly what was there before. Regular Ctrl+Z does not reliably undo programmatic add-in writes, so this is the only dependable way to back out a bad edit.", parameters: { type: "object", properties: {} } } },
        { type: "function", function: { name: "trace_dependencies", description: "Trace formula lineage for a single cell — precedents (what feeds it, upstream) or dependents (what it feeds, downstream). Use this to explain WHY a value is what it is by tracing the real chain of cells behind it, instead of reading one formula at a time and guessing.", parameters: { type: "object", properties: { cell: { type: "string", description: "Cell to trace, e.g. \"D10\" or \"Sheet2!F5\". Must be a single cell, not a range." }, mode: { type: "string", enum: ["precedents", "dependents"], description: "Trace direction. Default precedents." }, depth: { type: "number", description: "How many levels to trace. Default 2, max 5." } }, required: ["cell"] } } },
        { type: "function", function: { name: "search_workbook", description: "Search for text/values (or formula text) across every visible sheet in the workbook. Use this to find where something is calculated or referenced — e.g. \"where is EBITDA calculated?\" — without reading every sheet one by one.", parameters: { type: "object", properties: { query: { type: "string", description: "Search term." }, search_formulas: { type: "boolean", description: "Search formula text instead of displayed values. Useful for finding cross-sheet references." }, use_regex: { type: "boolean", description: "Treat query as a case-insensitive regular expression." }, sheet_name: { type: "string", description: "Restrict the search to this sheet. Omit to search every visible sheet." }, max_results: { type: "number", description: "Maximum matches to return. Default 20." } }, required: ["query"] } } },
        { type: "function", function: { name: "format_cells", description: "Apply VISUAL formatting to a range — font (bold/italic/underline/color/size/name), fill color, number format, alignment, borders, column width/row height, or merge. Does NOT change cell values or formulas — use write_cells for that.", parameters: { type: "object", properties: { range: { type: "string", description: "Range to format, e.g. \"A1:D1\" or \"Sheet2!B3:B20\"." }, bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, font_color: { type: "string", description: "Hex color, e.g. \"#0000FF\"." }, font_size: { type: "number" }, font_name: { type: "string" }, fill_color: { type: "string", description: "Hex background color, e.g. \"#FFFF00\"." }, number_format: { type: "string", description: "A preset (\"currency\", \"percent\", \"number\", \"integer\", \"text\") or a raw Excel number-format string, e.g. \"#,##0.00\"." }, horizontal_alignment: { type: "string", enum: ["Left", "Center", "Right", "General"] }, vertical_alignment: { type: "string", enum: ["Top", "Center", "Bottom"] }, wrap_text: { type: "boolean" }, borders: { type: "string", enum: ["thin", "medium", "thick", "none"], description: "Border weight applied to all four edges plus inside gridlines." }, border_color: { type: "string", description: "Hex color for borders set via `borders`. Default black." }, column_width: { type: "number", description: "Column width in points." }, row_height: { type: "number", description: "Row height in points." }, merge: { type: "boolean", description: "Merge (true) or unmerge (false) the range." } }, required: ["range"] } } },
        { type: "function", function: { name: "create_chart", description: "Create a chart from a data range — the core way to turn a comparison table into a visual. Use this whenever the user asks to \"chart\", \"graph\", \"visualize\", or \"plot\" something.", parameters: { type: "object", properties: { source_range: { type: "string", description: "Data range for the chart, e.g. \"A1:D12\" or \"Sheet1!A1:D12\" (include header row/column for series names)." }, chart_type: { type: "string", enum: ["column", "column_stacked", "bar", "bar_stacked", "line", "line_markers", "area", "area_stacked", "pie", "doughnut", "scatter", "radar"], description: "Chart type. Default column." }, title: { type: "string", description: "Chart title." }, x_axis_title: { type: "string" }, y_axis_title: { type: "string" }, legend_position: { type: "string", enum: ["right", "left", "top", "bottom", "none"] } }, required: ["source_range"] } } },
        { type: "function", function: { name: "conditional_format", description: "Add or clear conditional formatting rules — highlight cells that meet a condition (e.g. \"highlight negative numbers in red\"), or apply a color scale/gradient across a range (e.g. \"color-scale this range\").", parameters: { type: "object", properties: { action: { type: "string", enum: ["add", "clear"], description: "\"add\" a rule, or \"clear\" every rule in the range." }, range: { type: "string", description: "Target range, e.g. \"A1:D10\" or \"Sheet2!B2:B50\"." }, type: { type: "string", enum: ["cell_value", "formula", "color_scale"], description: "Rule type for action=\"add\"." }, operator: { type: "string", enum: ["Between", "NotBetween", "EqualTo", "NotEqualTo", "GreaterThan", "LessThan", "GreaterThanOrEqual", "LessThanOrEqual"], description: "Required for type=\"cell_value\"." }, value: { description: "Comparison value for type=\"cell_value\", e.g. 0 or \"=$B$2\"." }, value2: { description: "Second value, for Between/NotBetween." }, formula: { type: "string", description: "Custom formula for type=\"formula\", e.g. \"=A1<0\"." }, fill_color: { type: "string", description: "Hex fill color applied when the rule matches (cell_value/formula rules)." }, font_color: { type: "string" }, bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, min_color: { type: "string", description: "Color-scale low-end hex color. Default red." }, mid_color: { type: "string", description: "Color-scale midpoint hex color. Default yellow." }, max_color: { type: "string", description: "Color-scale high-end hex color. Default green." } }, required: ["action", "range"] } } },
        { type: "function", function: { name: "modify_structure", description: "Insert/delete rows or columns, or rename/delete/hide/unhide a sheet. Deleting anything asks for confirmation first (unless auto-apply is on) since it can destroy data — inserting, renaming, hiding, and unhiding do not, since nothing is lost.", parameters: { type: "object", properties: { action: { type: "string", enum: ["insert_row", "delete_row", "insert_column", "delete_column", "rename_sheet", "delete_sheet", "hide_sheet", "unhide_sheet"] }, sheet: { type: "string", description: "Target sheet name. Required for sheet actions; for row/column actions, omit to use the active sheet." }, position: { type: "number", description: "1-indexed row number (insert_row/delete_row) or 1-indexed column number (insert_column/delete_column)." }, count: { type: "number", description: "Number of rows/columns to insert or delete. Default 1." }, new_name: { type: "string", description: "New name, for rename_sheet." } }, required: ["action"] } } },
        { type: "function", function: { name: "named_ranges", description: "Manage named ranges — lets you (and the AI) refer to \"WACC\" instead of a guessed cell address. Deleting a name asks for confirmation first since any formula referencing it by name would break.", parameters: { type: "object", properties: { action: { type: "string", enum: ["add", "update", "delete", "list"] }, name: { type: "string", description: "The named range's name. Required for add/update/delete." }, reference: { type: "string", description: "Cell/range reference for add/update, e.g. \"Sheet1!$B$8\" or \"='Assumptions'!$B$8\"." }, scope: { type: "string", description: "Optional sheet name to scope the name to that sheet instead of the whole workbook." } }, required: ["action"] } } },
        { type: "function", function: { name: "autofit", description: "Auto-fit column widths and/or row heights to their content — use after writing a new sheet/table where columns are too narrow to read.", parameters: { type: "object", properties: { range: { type: "string", description: "Range/sheet to auto-fit, e.g. \"A1:F1\" or \"Sheet2!A:F\". Omit to auto-fit the active sheet's whole used range." }, columns: { type: "boolean", description: "Auto-fit column widths. Default true." }, rows: { type: "boolean", description: "Auto-fit row heights. Default true." } } } } },
        { type: "function", function: { name: "tables", description: "Convert a range into a proper Excel Table (sortable/filterable, with structured references) — use when the user asks to make a range \"a proper table,\" or wants to sort a range that already is one. \"delete\" only removes the Table wrapper (filters/styling) — the underlying data is left exactly as-is, same as Excel's own \"Convert to Range.\"", parameters: { type: "object", properties: { action: { type: "string", enum: ["create", "sort", "rename", "delete"] }, range: { type: "string", description: "Source range for action=\"create\", e.g. \"A1:E20\" or \"Sheet1!A1:E20\"." }, has_headers: { type: "boolean", description: "Whether the first row is a header row. Default true." }, name: { type: "string", description: "Table name — assigned name for create, or the target table's current name for sort/rename/delete." }, new_name: { type: "string", description: "New name, for action=\"rename\"." }, sort_column: { type: "number", description: "0-indexed column within the table to sort by, for action=\"sort\". Default 0." }, sort_ascending: { type: "boolean", description: "Sort ascending. Default true." } }, required: ["action"] } } },
        { type: "function", function: { name: "comments", description: "Add, reply to, resolve/reopen, delete, or list cell comments — good for leaving a note explaining WHY a change was made, pairing with undo_last_change's audit trail.", parameters: { type: "object", properties: { action: { type: "string", enum: ["add", "reply", "resolve", "reopen", "delete", "list"] }, cell: { type: "string", description: "Single cell the comment is attached to, e.g. \"B2\" or \"Sheet1!B2\". Required for add/reply/resolve/reopen/delete." }, content: { type: "string", description: "Comment/reply text. Required for add/reply." }, sheet: { type: "string", description: "Sheet to list comments from, for action=\"list\". Omit to use the active sheet." } }, required: ["action"] } } },
    ];
    const EXCEL_TOOL_NAMES = new Set(EXCEL_TOOLS.map(t => t.function.name));
    // These three manage their own confirm-card/running-card lifecycle (see each function below),
    // the same way the old write_excel_sheet did via maybeConfirmAndWrite — everything else in
    // EXCEL_TOOL_NAMES gets the generic "Running X…" wrapper send() applies itself.
    const EXCEL_SELF_MANAGED_UI_TOOLS = new Set(["create_sheet", "write_cells", "fill_formula", "undo_last_change", "modify_structure", "named_ranges"]);

    async function toolWorkbookOverview({ sheet: sheetName } = {}) {
        let text;
        await Excel.run(async (ctx) => {
            if (sheetName) {
                const sheet = ctx.workbook.worksheets.getItem(sheetName);
                sheet.load("name,visibility");
                const used = sheet.getUsedRangeOrNullObject();
                used.load("rowCount,columnCount,address,values");
                const headerRange = sheet.getRange("1:1").getUsedRangeOrNullObject();
                headerRange.load("values");
                const tables = sheet.tables;
                tables.load("items/name,items/rows/count,items/columns/count");
                const names = ctx.workbook.names;
                names.load("items/name,items/value,items/visible");
                const charts = sheet.charts;
                charts.load("count");
                const pivotTables = sheet.pivotTables;
                pivotTables.load("items/name");
                await ctx.sync();

                const lines = [];
                const vis = sheet.visibility === "Visible" ? "" : ` (${sheet.visibility})`;
                const dims = used.isNullObject ? "empty" : `${used.rowCount} rows × ${used.columnCount} cols`;
                lines.push(`## Sheet: ${sheet.name}${vis}`, `Dimensions: ${dims}`);
                const headers = used.isNullObject || headerRange.isNullObject ? [] : (headerRange.values[0] || []).filter(v => v !== null && v !== undefined && v !== "");
                if (headers.length) lines.push(`Headers: ${headers.join(", ")}`);
                if (tables.items.length) {
                    lines.push("", `### Tables (${tables.items.length})`);
                    for (const t of tables.items) lines.push(`- **${t.name}** (${t.rows.count} rows × ${t.columns.count} cols)`);
                }
                const sheetPrefix = `${sheet.name}!`.toLowerCase(), sheetQPrefix = `'${sheet.name}'!`.toLowerCase();
                const relevantNames = names.items.filter(n => n.visible && typeof n.value === "string" && (n.value.toLowerCase().startsWith(sheetPrefix) || n.value.toLowerCase().startsWith(sheetQPrefix)));
                if (relevantNames.length) {
                    lines.push("", `### Named Ranges (${relevantNames.length})`);
                    for (const n of relevantNames) lines.push(`- **${n.name}** = ${n.value}`);
                }
                const pivotTotal = pivotTables.items.length;
                if ((charts.count || 0) + pivotTotal > 0) {
                    lines.push("", "### Objects", `- Charts: ${charts.count || 0}`, `- Pivot tables (${pivotTotal}): ${pivotTables.items.map(p => p.name).join(", ") || "(none)"}`);
                }
                if (!used.isNullObject && used.rowCount > 0) {
                    const n = Math.min(5, used.rowCount);
                    lines.push("", `### Preview (first ${n} rows)`, formatMiniTable(used.values.slice(0, n)));
                }
                text = lines.join("\n");
            } else {
                const wb = ctx.workbook;
                wb.load("name");
                const sheets = wb.worksheets;
                sheets.load("items/name,items/position,items/visibility");
                const names = wb.names;
                names.load("items/name,items/value,items/visible");
                await ctx.sync();

                const lines = [`## Workbook: ${wb.name}`, "", `### Sheets (${sheets.items.length})`];
                const perSheet = sheets.items.map(s => {
                    const used = s.getUsedRangeOrNullObject();
                    used.load("rowCount,columnCount");
                    const headerRange = s.getRange("1:1").getUsedRangeOrNullObject();
                    headerRange.load("values");
                    const tables = s.tables;
                    tables.load("items/name");
                    const charts = s.charts;
                    charts.load("count");
                    return { s, used, headerRange, tables, charts };
                });
                await ctx.sync();
                for (const { s, used, headerRange, tables, charts } of perSheet) {
                    const vis = s.visibility === "Visible" ? "" : ` (${s.visibility})`;
                    const dims = used.isNullObject ? "empty" : `${used.rowCount} rows × ${used.columnCount} cols`;
                    lines.push(`${s.position + 1}. **${s.name}**${vis} — ${dims}`);
                    const headers = used.isNullObject || headerRange.isNullObject ? [] : (headerRange.values[0] || []).filter(v => v !== null && v !== undefined && v !== "");
                    if (headers.length) {
                        const display = headers.length > 8 ? headers.slice(0, 8).join(", ") + `, … (+${headers.length - 8} more)` : headers.join(", ");
                        lines.push(`   Headers: ${display}`);
                    }
                    if (tables.items.length) for (const t of tables.items) lines.push(`   Table: "${t.name}"`);
                    if (charts.count) lines.push(`   Charts: ${charts.count}`);
                }
                const visibleNames = names.items.filter(n => n.visible);
                if (visibleNames.length) {
                    lines.push("", `### Named Ranges (${visibleNames.length})`);
                    for (const n of visibleNames) lines.push(`- **${n.name}** = ${n.value}`);
                }
                text = lines.join("\n");
            }
        });
        return { overview: capText(text, 5000) };
    }

    async function toolReadRange({ range: rangeRef }) {
        if (!rangeRef) return { error: "range is required" };
        let result;
        await Excel.run(async (ctx) => {
            const { sheet, range } = chatGetRangeAndSheet(ctx, rangeRef);
            range.load("values,formulas,numberFormat,address,rowCount,columnCount");
            sheet.load("name");
            await ctx.sync();
            const cellPart = range.address.includes("!") ? range.address.slice(range.address.indexOf("!") + 1) : range.address;
            const errors = chatFindErrors(range.values, cellPart.split(":")[0]);
            result = {
                sheet: sheet.name,
                address: chatQualifiedAddress(sheet.name, range.address),
                rows: range.rowCount,
                cols: range.columnCount,
                values: range.values,
                formulas: range.formulas,
                numberFormats: range.numberFormat,
                ...(errors.length ? { errors } : {}),
            };
        });
        return result;
    }

    // Blocks on existing data (unless allow_overwrite) BEFORE ever showing a confirm card — a
    // blocked write never touched Excel at all, so there's nothing to confirm; the model sees the
    // existing data and can ask the user whether to proceed with allow_overwrite:true. Only an
    // ACTUAL pending write (empty target, or allow_overwrite already true) reaches the confirm gate.
    async function toolWriteCells({ start_cell, values, allow_overwrite }) {
        if (!start_cell || !Array.isArray(values) || !values.length) return { error: "start_cell and a non-empty values array are required" };
        const startCellRef = start_cell.includes("!") ? start_cell.slice(start_cell.indexOf("!") + 1) : start_cell;
        if (startCellRef.includes(":")) return { error: "start_cell must be a single cell (e.g. \"A1\"), not a range" };
        let startParsed;
        try { startParsed = chatParseCell(startCellRef); } catch (e) { return { error: String(e.message || e) }; }
        const { padded, rows, cols } = chatPadValues(values);

        for (let r = 0; r < padded.length; r++) {
            for (let c = 0; c < padded[r].length; c++) {
                const v = padded[r][c];
                if (typeof v === "string" && v.startsWith("=")) {
                    const bad = chatValidateFormula(v);
                    if (bad) return { error: `Invalid formula at ${chatCellAddress(startParsed.col + c, startParsed.row + r)}: ${v} (${bad})` };
                }
            }
        }

        let check;
        try {
            await Excel.run(async (ctx) => {
                const { sheet } = chatGetRangeAndSheet(ctx, start_cell);
                sheet.load("name");
                const rangeAddr = chatComputeRangeAddress(startCellRef, rows, cols);
                const target = sheet.getRange(rangeAddr);
                target.load("values,formulas");
                await ctx.sync();
                check = { sheetName: sheet.name, address: rangeAddr, beforeValues: target.values, beforeFormulas: target.formulas };
            });
        } catch (e) { return { error: String(e.message || e) }; }

        const occupied = chatCountOccupied(check.beforeValues, check.beforeFormulas);
        const fullAddr = chatQualifiedAddress(check.sheetName, check.address);
        if (occupied > 0 && !allow_overwrite) {
            return { blocked: true, address: fullAddr, existingCount: occupied, existingValues: check.beforeValues };
        }

        if (!autoApplyEdits) {
            const verb = occupied > 0 ? `overwrite ${occupied} existing cell(s) in` : "write to";
            const approved = await addConfirmCard(`DataGPT wants to ${verb} ${fullAddr} (${rows}×${cols}).`);
            if (!approved) return { ok: false, rejected: true, address: fullAddr };
        }
        const card = addCard(`Writing ${fullAddr}…`, "run");
        try {
            let written;
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getItem(check.sheetName);
                sheet.getRange(check.address).values = padded;
                await ctx.sync();
                const verify = sheet.getRange(check.address);
                verify.load("values,formulas");
                await ctx.sync();
                written = { readBackValues: verify.values, readBackFormulas: verify.formulas };
            });
            pushUndoSnapshot({ kind: "cells", sheetName: check.sheetName, address: check.address, beforeValues: check.beforeValues, beforeFormulas: check.beforeFormulas });
            const errors = chatFindErrors(written.readBackValues, startCellRef);
            await settleCard(card, errors.length ? "err" : "done", errors.length ? "✕" : "✓");
            return { ok: true, address: fullAddr, rows, cols, ...(errors.length ? { errors } : {}) };
        } catch (e) {
            await settleCard(card, "err", "✕");
            return { error: String(e.message || e) };
        }
    }

    async function toolFillFormula({ range: rangeRef, formula, allow_overwrite }) {
        if (!rangeRef || !formula) return { error: "range and formula are required" };
        if (!formula.startsWith("=")) return { error: "formula must start with \"=\"" };
        const bad = chatValidateFormula(formula);
        if (bad) return { error: `Invalid formula (${bad})` };
        if (/[;,]/.test(rangeRef)) return { error: "fill_formula only supports a single contiguous range" };

        let check;
        try {
            await Excel.run(async (ctx) => {
                const { sheet, range } = chatGetRangeAndSheet(ctx, rangeRef);
                sheet.load("name");
                range.load("address,rowCount,columnCount,values,formulas");
                await ctx.sync();
                check = { sheetName: sheet.name, address: range.address, rowCount: range.rowCount, columnCount: range.columnCount, beforeValues: range.values, beforeFormulas: range.formulas };
            });
        } catch (e) { return { error: String(e.message || e) }; }

        const occupied = chatCountOccupied(check.beforeValues, check.beforeFormulas);
        const fullAddr = chatQualifiedAddress(check.sheetName, check.address);
        if (occupied > 0 && !allow_overwrite) {
            return { blocked: true, address: fullAddr, existingCount: occupied, existingValues: check.beforeValues };
        }

        if (!autoApplyEdits) {
            const verb = occupied > 0 ? `overwrite ${occupied} existing cell(s) in` : "fill";
            const approved = await addConfirmCard(`DataGPT wants to ${verb} ${fullAddr} (${check.rowCount}×${check.columnCount}) with "${formula}".`);
            if (!approved) return { ok: false, rejected: true, address: fullAddr };
        }
        const card = addCard(`Filling formula across ${fullAddr}…`, "run");
        try {
            let filled;
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getItem(check.sheetName);
                const range = sheet.getRange(check.address);
                const topLeft = range.getCell(0, 0);
                topLeft.formulas = [[formula]];
                topLeft.autoFill(range, "FillDefault");
                range.load("values,formulas");
                await ctx.sync();
                filled = { readBackValues: range.values, readBackFormulas: range.formulas };
            });
            pushUndoSnapshot({ kind: "cells", sheetName: check.sheetName, address: check.address, beforeValues: check.beforeValues, beforeFormulas: check.beforeFormulas });
            const cellPart = check.address.includes("!") ? check.address.slice(check.address.indexOf("!") + 1) : check.address;
            const errors = chatFindErrors(filled.readBackValues, cellPart.split(":")[0]);
            await settleCard(card, errors.length ? "err" : "done", errors.length ? "✕" : "✓");
            return { ok: true, address: fullAddr, rows: check.rowCount, cols: check.columnCount, formula, ...(errors.length ? { errors } : {}) };
        } catch (e) {
            await settleCard(card, "err", "✕");
            return { error: String(e.message || e) };
        }
    }

    // Not gated by a confirm card — recalculating is always-safe, expected Excel behavior (not an
    // AI-authored edit), same reasoning as why read-only tools don't need approval either.
    async function toolRecalculate({ sheet: sheetName } = {}) {
        let result;
        await Excel.run(async (ctx) => {
            ctx.workbook.application.calculate("Recalculate");
            await ctx.sync();

            let sheets;
            if (sheetName) {
                const s = ctx.workbook.worksheets.getItem(sheetName);
                s.load("name");
                sheets = [s];
            } else {
                ctx.workbook.worksheets.load("items/name,items/visibility");
                await ctx.sync();
                sheets = ctx.workbook.worksheets.items.filter(s => s.visibility === "Visible");
            }
            // Plain values scan per sheet, same technique as chatFindErrors elsewhere in this file
            // — could be slow on a handful of very large sheets, but every sheet this add-in itself
            // writes is bounded to a few hundred rows at most, and that's the common case here.
            const usedRanges = sheets.map(s => { const u = s.getUsedRangeOrNullObject(); u.load("values,address,isNullObject"); return { sheet: s, used: u }; });
            await ctx.sync();

            const errors = [];
            for (const { sheet, used } of usedRanges) {
                if (used.isNullObject) continue;
                const cellPart = used.address.includes("!") ? used.address.slice(used.address.indexOf("!") + 1) : used.address;
                for (const e of chatFindErrors(used.values, cellPart.split(":")[0])) errors.push({ sheet: sheet.name, ...e });
            }
            result = { recalculated: true, errorCount: errors.length, errors: errors.slice(0, 50) };
        });
        return result;
    }

    // Purpose-built for this product specifically, not a generic Excel feature — a direct
    // application of the same "does this model actually tie out" checks the AI Financial Model
    // builder already runs on itself (balance_check ≈ 0, etc.), but on demand for ANY workbook,
    // not just one this add-in built. Recalculates first (unlike recalculate_and_check, which
    // assumes the caller already knows what changed — this is the "give me the full picture,
    // unprompted" tool, so it can't rely on a fresh recalc having already happened), then reports
    // two things: formula errors everywhere, and any "check" row whose value isn't close to zero
    // the way a reconciliation/sanity-check row is supposed to be. A "check" row is found by
    // label, not by key — this works on someone else's workbook too, not just ones this add-in
    // built with a known balance_check key.
    async function toolAuditWorkbook({ sheet: sheetName, tolerance } = {}) {
        const tol = typeof tolerance === "number" && isFinite(tolerance) ? Math.abs(tolerance) : 0.01;
        let result;
        await Excel.run(async (ctx) => {
            ctx.workbook.application.calculate("Recalculate");
            await ctx.sync();

            let sheets;
            if (sheetName) {
                const s = ctx.workbook.worksheets.getItem(sheetName);
                s.load("name");
                sheets = [s];
            } else {
                ctx.workbook.worksheets.load("items/name,items/visibility");
                await ctx.sync();
                sheets = ctx.workbook.worksheets.items.filter(s => s.visibility === "Visible");
            }
            const usedRanges = sheets.map(s => { const u = s.getUsedRangeOrNullObject(); u.load("values,address,isNullObject"); return { sheet: s, used: u }; });
            await ctx.sync();

            const formulaErrors = [];
            const checkRowIssues = [];
            for (const { sheet, used } of usedRanges) {
                if (used.isNullObject) continue;
                const cellPart = used.address.includes("!") ? used.address.slice(used.address.indexOf("!") + 1) : used.address;
                const startCell = cellPart.split(":")[0];
                const start = chatParseCell(startCell);

                for (const e of chatFindErrors(used.values, startCell)) formulaErrors.push({ sheet: sheet.name, ...e });

                for (let r = 0; r < used.values.length; r++) {
                    const row = used.values[r] || [];
                    const label = row.find(v => typeof v === "string" && v.trim() !== "");
                    if (typeof label !== "string" || !/check/i.test(label)) continue;
                    const rowNum = start.row + r;
                    const offenders = [];
                    for (let c = 0; c < row.length; c++) {
                        const v = row[c];
                        if (typeof v === "number" && Math.abs(v) > tol) offenders.push({ address: chatCellAddress(start.col + c, rowNum), value: v });
                    }
                    if (offenders.length) checkRowIssues.push({ sheet: sheet.name, row: rowNum, label: label.trim(), offenders: offenders.slice(0, 10) });
                }
            }
            result = {
                healthy: formulaErrors.length === 0 && checkRowIssues.length === 0,
                formulaErrors: formulaErrors.slice(0, 50),
                checkRowIssues: checkRowIssues.slice(0, 20),
            };
        });
        return result;
    }

    // Only checks for a name collision (Excel can't silently overwrite a sheet by creating a
    // duplicate-named one the way write_cells can silently overwrite occupied cells) — so unlike
    // write_cells/fill_formula there's no allow_overwrite escape hatch, just a clear error telling
    // the model to target the existing sheet instead of trying to create a duplicate.
    async function toolCreateSheet({ name, activate } = {}) {
        const sheetName = String(name || "").trim().slice(0, 31);
        if (!sheetName) return { error: "name is required" };
        if (/[\\/?*[\]:]/.test(sheetName)) return { error: "Sheet names can't contain \\ / ? * [ ] or :" };

        let exists = false;
        try {
            await Excel.run(async (ctx) => {
                ctx.workbook.worksheets.load("items/name");
                await ctx.sync();
                exists = ctx.workbook.worksheets.items.some(s => s.name.toLowerCase() === sheetName.toLowerCase());
            });
        } catch (e) { return { error: String(e.message || e) }; }
        if (exists) return { error: `A sheet named "${sheetName}" already exists — use write_cells/fill_formula/read_range on it directly instead of creating a duplicate.` };

        if (!autoApplyEdits) {
            const approved = await addConfirmCard(`DataGPT wants to create a new sheet named "${sheetName}".`);
            if (!approved) return { ok: false, rejected: true, name: sheetName };
        }
        const card = addCard(`Creating sheet "${sheetName}"…`, "run");
        try {
            await Excel.run(async (ctx) => {
                const sh = ctx.workbook.worksheets.add(sheetName);
                if (activate !== false) sh.activate();
                await ctx.sync();
            });
            pushUndoSnapshot({ kind: "createSheet", sheetName });
            await settleCard(card, "done", "✓");
            return { ok: true, name: sheetName };
        } catch (e) {
            await settleCard(card, "err", "✕");
            return { error: String(e.message || e) };
        }
    }

    async function toolUndoLastChange() {
        const snap = undoStack.pop();
        if (!snap) return { ok: false, error: "Nothing to undo — no tracked change this session." };

        if (snap.kind === "createSheet") {
            if (!autoApplyEdits) {
                const approved = await addConfirmCard(`DataGPT wants to undo creating the sheet "${snap.sheetName}" (deletes it).`);
                if (!approved) { undoStack.push(snap); return { ok: false, rejected: true, name: snap.sheetName }; }
            }
            const card = addCard(`Deleting sheet "${snap.sheetName}"…`, "run");
            try {
                await Excel.run(async (ctx) => {
                    ctx.workbook.worksheets.getItem(snap.sheetName).delete();
                    await ctx.sync();
                });
                await settleCard(card, "done", "✓");
                return { ok: true, deletedSheet: snap.sheetName };
            } catch (e) {
                await settleCard(card, "err", "✕");
                undoStack.push(snap);
                return { error: String(e.message || e) };
            }
        }

        const fullAddr = chatQualifiedAddress(snap.sheetName, snap.address);
        if (!autoApplyEdits) {
            const approved = await addConfirmCard(`DataGPT wants to undo the last change to ${fullAddr}.`);
            if (!approved) { undoStack.push(snap); return { ok: false, rejected: true, address: fullAddr }; }
        }
        const card = addCard(`Undoing change to ${fullAddr}…`, "run");
        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getItem(snap.sheetName);
                // Restore the formula where the original cell had one, otherwise the plain value —
                // a cell that was blank before must go back to blank, not to whatever the write
                // computed. Range.formulas accepts a mix of formula strings and plain values in the
                // same 2D array, same as .values would for the non-formula entries.
                const restore = snap.beforeValues.map((row, r) => row.map((v, c) => {
                    const f = snap.beforeFormulas[r] && snap.beforeFormulas[r][c];
                    return (typeof f === "string" && f.startsWith("=")) ? f : v;
                }));
                sheet.getRange(snap.address).formulas = restore;
                await ctx.sync();
            });
            await settleCard(card, "done", "✓");
            return { ok: true, address: fullAddr };
        } catch (e) {
            await settleCard(card, "err", "✕");
            undoStack.push(snap);
            return { error: String(e.message || e) };
        }
    }

    // Uses Excel's own getDirectPrecedents()/getDirectDependents() when the host supports them
    // (falls back gracefully otherwise). For precedents, an unavailable API falls back to scanning
    // the cell's own formula text for references. There's deliberately NO workbook-wide scanning
    // fallback for dependents — finding every cell that references this one with no native API
    // means scanning every sheet's formulas by hand, which is expensive and out of scope here; if
    // the API isn't available, dependents tracing just reports that plainly instead.
    async function toolTraceDependencies({ cell, mode, depth }) {
        if (!cell) return { error: "cell is required" };
        const traceMode = mode === "dependents" ? "dependents" : "precedents";
        const maxDepth = Math.min(Math.max(parseInt(depth, 10) || 2, 1), 5);
        const MAX_NODES = 40;
        let apiUnavailable = false;
        const nodes = {};
        let rootAddr;
        try {
            await Excel.run(async (ctx) => {
                const queue = [{ ref: cell, depth: 0 }];
                const visited = new Set();
                while (queue.length && Object.keys(nodes).length < MAX_NODES) {
                    const { ref, depth: d } = queue.shift();
                    let sheet, range;
                    try { ({ sheet, range } = chatGetRangeAndSheet(ctx, ref)); } catch (e) { continue; }
                    range.load("values,formulas,address");
                    sheet.load("name");
                    await ctx.sync();
                    const cellPart = range.address.includes("!") ? range.address.slice(range.address.indexOf("!") + 1) : range.address;
                    const qualified = chatQualifiedAddress(sheet.name, cellPart.split(":")[0]);
                    if (visited.has(qualified)) continue;
                    visited.add(qualified);
                    if (d === 0) rootAddr = qualified;
                    const formula = range.formulas[0] && range.formulas[0][0];
                    const formulaStr = (typeof formula === "string" && formula.startsWith("=")) ? formula : undefined;
                    const node = { address: qualified, value: range.values[0] && range.values[0][0], ...(formulaStr ? { formula: formulaStr } : {}), links: [] };
                    nodes[qualified] = node;
                    if (d >= maxDepth) continue;

                    let related = [];
                    try {
                        const rel = traceMode === "precedents" ? range.getDirectPrecedents() : range.getDirectDependents();
                        rel.load("addresses");
                        await ctx.sync();
                        related = (rel.addresses || []).flatMap(s => s.split(",").map(x => x.trim()).filter(Boolean));
                    } catch (e) {
                        apiUnavailable = true;
                        if (traceMode === "precedents" && formulaStr) {
                            const re = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)?!?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g;
                            const seen = new Set();
                            for (const m of formulaStr.matchAll(re)) {
                                const token = m[0];
                                if (!/\d/.test(token)) continue;
                                const ref2 = (token.includes("!") ? token : `${sheet.name}!${token}`).replace(/\$/g, "");
                                if (!seen.has(ref2)) { seen.add(ref2); related.push(ref2); }
                            }
                        }
                    }
                    related = related.slice(0, 15); // cap fan-out per node so a wide dependency web can't runaway
                    for (const r of related) {
                        // fullRef always has a sheet prefix by construction — Excel sheet names can
                        // never contain "!", so splitting on the first one is always unambiguous.
                        const fullRef = r.includes("!") ? r : `${sheet.name}!${r}`;
                        const bangIdx = fullRef.indexOf("!");
                        node.links.push(chatQualifiedAddress(fullRef.slice(0, bangIdx), fullRef.slice(bangIdx + 1)));
                        queue.push({ ref: fullRef, depth: d + 1 });
                    }
                }
            });
        } catch (e) {
            return { error: String(e.message || e) };
        }
        return {
            mode: traceMode, root: rootAddr, depth: maxDepth, nodes: Object.values(nodes),
            ...(apiUnavailable ? { note: traceMode === "dependents" ? "Direct-dependents API unavailable in this Excel version/host — dependents tracing could not run." : "Direct-precedents API unavailable — fell back to scanning each cell's own formula text." } : {}),
        };
    }

    async function toolSearchWorkbook({ query, search_formulas, use_regex, sheet, max_results }) {
        if (!query) return { error: "query is required" };
        const maxResults = Math.max(parseInt(max_results, 10) || 20, 1);
        const searchFormulas = !!search_formulas;
        let regex;
        if (use_regex) {
            try { regex = new RegExp(query, "i"); } catch (e) { return { error: `Invalid regex: ${e.message}` }; }
        }
        const queryLower = query.toLowerCase();
        let result;
        await Excel.run(async (ctx) => {
            const sheets = ctx.workbook.worksheets;
            sheets.load("items/name,items/visibility");
            await ctx.sync();
            const targets = sheet ? sheets.items.filter(s => s.name === sheet) : sheets.items.filter(s => s.visibility === "Visible");

            const loaded = targets.map(s => { const u = s.getUsedRangeOrNullObject(); u.load("values,formulas,address,isNullObject"); return { sheet: s, used: u }; });
            await ctx.sync();

            const matches = [];
            let hasMore = false;
            outer: for (const { sheet: sh, used } of loaded) {
                if (used.isNullObject) continue;
                const cellPart = used.address.includes("!") ? used.address.slice(used.address.indexOf("!") + 1) : used.address;
                let start;
                try { start = chatParseCell(cellPart.split(":")[0]); } catch (e) { continue; }
                for (let r = 0; r < used.values.length; r++) {
                    const valueRow = used.values[r] || [], formulaRow = used.formulas[r] || [];
                    for (let c = 0; c < valueRow.length; c++) {
                        const value = valueRow[c], formula = formulaRow[c];
                        let target = "";
                        if (searchFormulas) {
                            if (typeof formula !== "string" || !formula) continue;
                            target = formula;
                        } else {
                            if (value === null || value === undefined || value === "") continue;
                            target = String(value);
                        }
                        const isMatch = regex ? regex.test(target) : target.toLowerCase().includes(queryLower);
                        if (!isMatch) continue;
                        matches.push({
                            sheet: sh.name,
                            address: chatCellAddress(start.col + c, start.row + r),
                            value,
                            ...(typeof formula === "string" && formula.startsWith("=") ? { formula } : {}),
                        });
                        if (matches.length >= maxResults) { hasMore = true; break outer; }
                    }
                }
            }
            result = { matches, hasMore };
        });
        return result;
    }

    // ── Second tool batch: chart/formatting-focused, ported from the same reference
    // (pi-for-excel-main's src/tools/format-cells.ts, charts.ts, conditional-format.ts,
    // modify-structure.ts, comments.ts) where it has an equivalent; named_ranges/autofit/tables
    // have no dedicated file there, so those are written directly from the plain Excel JS API.
    // Risk classification for the confirm-gate (see EXCEL_SELF_MANAGED_UI_TOOLS below): anything
    // that only changes PRESENTATION (formatting, conditional formatting, charts, autofit, table
    // wrapping, comments) can't destroy data and is trivially reversible, so none of those prompt
    // for approval — same reasoning as why read-only tools don't. modify_structure's delete_*
    // actions and named_ranges' delete action CAN cause real damage (deleting a row/column/sheet
    // destroys data; deleting a name breaks every formula that referenced it by name), so those two
    // tools manage their own confirm logic internally, gating only their destructive actions.

    const CHAT_NUMBER_FORMAT_PRESETS = { currency: "#,##0.00", integer: "#,##0", number: "#,##0.00", percent: "0.0%", text: "@" };
    const CHAT_BORDER_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
    const CHAT_BORDER_WEIGHTS = { thin: "Thin", medium: "Medium", thick: "Thick" };

    async function toolFormatCells(p) {
        if (!p || !p.range) return { error: "range is required" };
        const applied = [];
        try {
            let fullAddr;
            await Excel.run(async (ctx) => {
                const { sheet, range } = chatGetRangeAndSheet(ctx, p.range);
                sheet.load("name");
                range.load("address,rowCount,columnCount");
                await ctx.sync();
                const fmt = range.format;
                if (typeof p.bold === "boolean") { fmt.font.bold = p.bold; applied.push(p.bold ? "bold" : "not bold"); }
                if (typeof p.italic === "boolean") { fmt.font.italic = p.italic; applied.push(p.italic ? "italic" : "not italic"); }
                if (typeof p.underline === "boolean") { fmt.font.underline = p.underline ? "Single" : "None"; applied.push(p.underline ? "underline" : "no underline"); }
                if (p.font_color) { fmt.font.color = p.font_color; applied.push(`font color ${p.font_color}`); }
                if (typeof p.font_size === "number") { fmt.font.size = p.font_size; applied.push(`${p.font_size}pt`); }
                if (p.font_name) { fmt.font.name = p.font_name; applied.push(`font ${p.font_name}`); }
                if (p.fill_color) { fmt.fill.color = p.fill_color; applied.push(`fill ${p.fill_color}`); }
                if (p.number_format) {
                    const nf = CHAT_NUMBER_FORMAT_PRESETS[p.number_format] || p.number_format;
                    range.numberFormat = Array.from({ length: range.rowCount }, () => Array.from({ length: range.columnCount }, () => nf));
                    applied.push(`format "${nf}"`);
                }
                if (p.horizontal_alignment) { fmt.horizontalAlignment = p.horizontal_alignment; applied.push(`align ${p.horizontal_alignment}`); }
                if (p.vertical_alignment) { fmt.verticalAlignment = p.vertical_alignment; applied.push(`v-align ${p.vertical_alignment}`); }
                if (typeof p.wrap_text === "boolean") { fmt.wrapText = p.wrap_text; applied.push(p.wrap_text ? "wrap" : "no wrap"); }
                if (p.borders === "none") {
                    for (const edge of CHAT_BORDER_EDGES) fmt.borders.getItem(edge).style = "None";
                    applied.push("borders removed");
                } else if (p.borders) {
                    const weight = CHAT_BORDER_WEIGHTS[p.borders];
                    for (const edge of CHAT_BORDER_EDGES) {
                        const b = fmt.borders.getItem(edge);
                        b.style = "Continuous"; b.weight = weight;
                        if (p.border_color) b.color = p.border_color;
                    }
                    applied.push(`${p.borders} borders`);
                }
                if (typeof p.column_width === "number") { fmt.columnWidth = p.column_width; applied.push(`col width ${p.column_width}pt`); }
                if (typeof p.row_height === "number") { fmt.rowHeight = p.row_height; applied.push(`row height ${p.row_height}pt`); }
                if (typeof p.merge === "boolean") { if (p.merge) range.merge(); else range.unmerge(); applied.push(p.merge ? "merged" : "unmerged"); }
                await ctx.sync();
                fullAddr = chatQualifiedAddress(sheet.name, range.address);
            });
            return { ok: true, address: fullAddr, applied };
        } catch (e) {
            return { error: String(e.message || e) };
        }
    }

    const CHAT_CHART_TYPE_MAP = {
        column: "ColumnClustered", column_stacked: "ColumnStacked", bar: "BarClustered", bar_stacked: "BarStacked",
        line: "Line", line_markers: "LineMarkers", area: "Area", area_stacked: "AreaStacked",
        pie: "Pie", doughnut: "Doughnut", scatter: "XYScatter", radar: "Radar",
    };
    async function toolCreateChart(p) {
        if (!p || !p.source_range) return { error: "source_range is required" };
        const excelType = CHAT_CHART_TYPE_MAP[p.chart_type] || "ColumnClustered";
        try {
            let result;
            await Excel.run(async (ctx) => {
                const { sheet, range } = chatGetRangeAndSheet(ctx, p.source_range);
                sheet.load("name");
                const chart = sheet.charts.add(excelType, range, "Auto");
                if (p.title) chart.title.text = p.title;
                if (p.legend_position === "none") chart.legend.visible = false;
                else if (p.legend_position) { chart.legend.visible = true; chart.legend.position = p.legend_position; }
                if (p.x_axis_title) chart.axes.categoryAxis.title.text = p.x_axis_title;
                if (p.y_axis_title) chart.axes.valueAxis.title.text = p.y_axis_title;
                chart.name = `${p.title || "Chart"}_${Date.now() % 100000}`;
                chart.load("name");
                await ctx.sync();
                result = { name: chart.name, sheet: sheet.name };
            });
            return { ok: true, ...result };
        } catch (e) {
            return { error: String(e.message || e) };
        }
    }

    async function toolConditionalFormat(p) {
        if (!p || !p.range) return { error: "range is required" };
        try {
            let fullAddr;
            await Excel.run(async (ctx) => {
                const { sheet, range } = chatGetRangeAndSheet(ctx, p.range);
                sheet.load("name");
                range.load("address");
                await ctx.sync();
                fullAddr = chatQualifiedAddress(sheet.name, range.address);

                if (p.action === "clear") {
                    range.conditionalFormats.clearAll();
                    await ctx.sync();
                    return;
                }

                const applyFmt = (fmt) => {
                    if (p.fill_color) fmt.fill.color = p.fill_color;
                    if (p.font_color) fmt.font.color = p.font_color;
                    if (typeof p.bold === "boolean") fmt.font.bold = p.bold;
                    if (typeof p.italic === "boolean") fmt.font.italic = p.italic;
                    if (typeof p.underline === "boolean") fmt.font.underline = p.underline ? "Single" : "None";
                };

                if (p.type === "color_scale") {
                    const cf = range.conditionalFormats.add("ColorScale");
                    cf.colorScale.criteria = {
                        minimum: { formula: null, type: "LowestValue", color: p.min_color || "#F8696B" },
                        midpoint: { formula: "50", type: "Percentile", color: p.mid_color || "#FFEB84" },
                        maximum: { formula: null, type: "HighestValue", color: p.max_color || "#63BE7B" },
                    };
                } else if (p.type === "formula") {
                    if (!p.formula) throw new Error("formula is required for type=\"formula\"");
                    const cf = range.conditionalFormats.add("Custom");
                    cf.custom.rule.formula = p.formula;
                    applyFmt(cf.custom.format);
                } else {
                    if (!p.operator) throw new Error("operator is required for type=\"cell_value\"");
                    const cf = range.conditionalFormats.add("CellValue");
                    const rule = { formula1: String(p.value ?? ""), operator: p.operator };
                    if (p.value2 !== undefined) rule.formula2 = String(p.value2);
                    cf.cellValue.rule = rule;
                    applyFmt(cf.cellValue.format);
                }
                await ctx.sync();
            });
            return { ok: true, address: fullAddr, action: p.action };
        } catch (e) {
            return { error: String(e.message || e) };
        }
    }

    const CHAT_STRUCTURE_DESTRUCTIVE = new Set(["delete_row", "delete_column", "delete_sheet"]);
    async function toolModifyStructure(p) {
        if (!p || !p.action) return { error: "action is required" };
        if (CHAT_STRUCTURE_DESTRUCTIVE.has(p.action) && !autoApplyEdits) {
            const desc = p.action === "delete_sheet" ? `delete the sheet "${p.sheet}"`
                : p.action === "delete_row" ? `delete row(s) ${p.position}-${(p.position || 0) + (p.count || 1) - 1} on "${p.sheet || "the active sheet"}"`
                    : `delete column(s) starting at ${chatColToLetter((p.position || 1) - 1)} on "${p.sheet || "the active sheet"}"`;
            const approved = await addConfirmCard(`DataGPT wants to ${desc}. This can destroy data.`);
            if (!approved) return { ok: false, rejected: true };
        }
        const card = addCard(`${p.action.replace(/_/g, " ")}…`, "run");
        try {
            let result;
            await Excel.run(async (ctx) => {
                const getSheet = () => p.sheet ? ctx.workbook.worksheets.getItem(p.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
                const count = Math.max(1, parseInt(p.count, 10) || 1);
                switch (p.action) {
                    case "insert_row": {
                        if (!p.position) throw new Error("position is required for insert_row");
                        const sheet = getSheet(); sheet.load("name");
                        sheet.getRange(`${p.position}:${p.position + count - 1}`).insert("Down");
                        await ctx.sync();
                        result = { message: `Inserted ${count} row(s) at row ${p.position} on "${sheet.name}"` };
                        break;
                    }
                    case "delete_row": {
                        if (!p.position) throw new Error("position is required for delete_row");
                        const sheet = getSheet(); sheet.load("name");
                        sheet.getRange(`${p.position}:${p.position + count - 1}`).delete("Up");
                        await ctx.sync();
                        result = { message: `Deleted ${count} row(s) starting at row ${p.position} on "${sheet.name}"` };
                        break;
                    }
                    case "insert_column": {
                        if (!p.position) throw new Error("position is required for insert_column");
                        const startLetter = chatColToLetter(p.position - 1), endLetter = chatColToLetter(p.position + count - 2);
                        const sheet = getSheet(); sheet.load("name");
                        sheet.getRange(`${startLetter}:${endLetter}`).insert("Right");
                        await ctx.sync();
                        result = { message: `Inserted ${count} column(s) at column ${startLetter} on "${sheet.name}"` };
                        break;
                    }
                    case "delete_column": {
                        if (!p.position) throw new Error("position is required for delete_column");
                        const startLetter = chatColToLetter(p.position - 1), endLetter = chatColToLetter(p.position + count - 2);
                        const sheet = getSheet(); sheet.load("name");
                        sheet.getRange(`${startLetter}:${endLetter}`).delete("Left");
                        await ctx.sync();
                        result = { message: `Deleted ${count} column(s) starting at column ${startLetter} on "${sheet.name}"` };
                        break;
                    }
                    case "rename_sheet": {
                        if (!p.sheet || !p.new_name) throw new Error("sheet and new_name are required for rename_sheet");
                        getSheet().name = p.new_name;
                        await ctx.sync();
                        result = { message: `Renamed "${p.sheet}" to "${p.new_name}"` };
                        break;
                    }
                    case "delete_sheet": {
                        if (!p.sheet) throw new Error("sheet is required for delete_sheet");
                        getSheet().delete();
                        await ctx.sync();
                        result = { message: `Deleted sheet "${p.sheet}"` };
                        break;
                    }
                    case "hide_sheet": {
                        if (!p.sheet) throw new Error("sheet is required for hide_sheet");
                        getSheet().visibility = "Hidden";
                        await ctx.sync();
                        result = { message: `Hid sheet "${p.sheet}"` };
                        break;
                    }
                    case "unhide_sheet": {
                        if (!p.sheet) throw new Error("sheet is required for unhide_sheet");
                        getSheet().visibility = "Visible";
                        await ctx.sync();
                        result = { message: `Unhid sheet "${p.sheet}"` };
                        break;
                    }
                    default: throw new Error(`Unknown action "${p.action}"`);
                }
            });
            await settleCard(card, "done", "✓");
            return { ok: true, ...result };
        } catch (e) {
            await settleCard(card, "err", "✕");
            return { error: String(e.message || e) };
        }
    }

    async function toolNamedRanges(p) {
        if (!p || !p.action) return { error: "action is required" };
        if (p.action === "delete" && !autoApplyEdits) {
            const approved = await addConfirmCard(`DataGPT wants to delete the named range "${p.name}". Any formula referencing it by name will break.`);
            if (!approved) return { ok: false, rejected: true };
        }
        const selfManaged = p.action === "delete";
        const card = selfManaged ? addCard(`Deleting named range "${p.name}"…`, "run") : null;
        try {
            let result;
            await Excel.run(async (ctx) => {
                const names = p.scope ? ctx.workbook.worksheets.getItem(p.scope).names : ctx.workbook.names;
                switch (p.action) {
                    case "list": {
                        names.load("items/name,items/value,items/type,items/visible");
                        await ctx.sync();
                        result = { names: names.items.filter(n => n.visible).map(n => ({ name: n.name, value: n.value, type: n.type })) };
                        break;
                    }
                    case "add": {
                        if (!p.name || !p.reference) throw new Error("name and reference are required for add");
                        names.add(p.name, p.reference.startsWith("=") ? p.reference : `=${p.reference}`);
                        await ctx.sync();
                        result = { message: `Created named range "${p.name}" = ${p.reference}` };
                        break;
                    }
                    case "update": {
                        if (!p.name || !p.reference) throw new Error("name and reference are required for update");
                        const item = names.getItem(p.name);
                        item.formula = p.reference.startsWith("=") ? p.reference : `=${p.reference}`;
                        await ctx.sync();
                        result = { message: `Updated named range "${p.name}" to ${p.reference}` };
                        break;
                    }
                    case "delete": {
                        if (!p.name) throw new Error("name is required for delete");
                        names.getItem(p.name).delete();
                        await ctx.sync();
                        result = { message: `Deleted named range "${p.name}"` };
                        break;
                    }
                    default: throw new Error(`Unknown action "${p.action}"`);
                }
            });
            await settleCard(card, "done", "✓");
            return { ok: true, ...result };
        } catch (e) {
            await settleCard(card, "err", "✕");
            return { error: String(e.message || e) };
        }
    }

    async function toolAutofit({ range: rangeRef, columns, rows } = {}) {
        try {
            let fullAddr;
            await Excel.run(async (ctx) => {
                let sheet, range;
                if (rangeRef) { ({ sheet, range } = chatGetRangeAndSheet(ctx, rangeRef)); }
                else { sheet = ctx.workbook.worksheets.getActiveWorksheet(); range = sheet.getUsedRange(); }
                sheet.load("name");
                range.load("address");
                await ctx.sync();
                if (columns !== false) range.format.autofitColumns();
                if (rows !== false) range.format.autofitRows();
                await ctx.sync();
                fullAddr = chatQualifiedAddress(sheet.name, range.address);
            });
            return { ok: true, address: fullAddr };
        } catch (e) {
            return { error: String(e.message || e) };
        }
    }

    async function toolTables(p) {
        if (!p || !p.action) return { error: "action is required" };
        try {
            let result;
            await Excel.run(async (ctx) => {
                switch (p.action) {
                    case "create": {
                        if (!p.range) throw new Error("range is required for create");
                        const { sheet, range } = chatGetRangeAndSheet(ctx, p.range);
                        sheet.load("name");
                        const table = sheet.tables.add(range, p.has_headers !== false);
                        if (p.name) table.name = p.name;
                        table.load("name,id");
                        await ctx.sync();
                        result = { message: `Created table "${table.name}" from ${p.range}`, name: table.name };
                        break;
                    }
                    case "sort": {
                        if (!p.name) throw new Error("name is required for sort");
                        const table = ctx.workbook.tables.getItem(p.name);
                        table.sort.apply([{ key: p.sort_column || 0, ascending: p.sort_ascending !== false }]);
                        await ctx.sync();
                        result = { message: `Sorted table "${p.name}" by column ${p.sort_column || 0}` };
                        break;
                    }
                    case "rename": {
                        if (!p.name || !p.new_name) throw new Error("name and new_name are required for rename");
                        ctx.workbook.tables.getItem(p.name).name = p.new_name;
                        await ctx.sync();
                        result = { message: `Renamed table "${p.name}" to "${p.new_name}"` };
                        break;
                    }
                    case "delete": {
                        // Removes the Table wrapper only (filters/styling/structured refs) — the
                        // underlying cell data is left exactly as-is, same as Excel's own "Convert
                        // to Range." Nothing gets deleted, so this isn't confirm-gated.
                        if (!p.name) throw new Error("name is required for delete");
                        ctx.workbook.tables.getItem(p.name).convertToRange();
                        await ctx.sync();
                        result = { message: `Converted table "${p.name}" back to a plain range` };
                        break;
                    }
                    default: throw new Error(`Unknown action "${p.action}"`);
                }
            });
            return { ok: true, ...result };
        } catch (e) {
            return { error: String(e.message || e) };
        }
    }

    async function toolComments(p) {
        if (!p || !p.action) return { error: "action is required" };
        try {
            let result;
            await Excel.run(async (ctx) => {
                switch (p.action) {
                    case "list": {
                        const sheet = p.sheet ? ctx.workbook.worksheets.getItem(p.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
                        sheet.load("name");
                        const comments = sheet.comments;
                        comments.load("items/content,items/authorName,items/resolved");
                        await ctx.sync();
                        const withLoc = comments.items.map(c => { const loc = c.getLocation(); loc.load("address"); return { c, loc }; });
                        await ctx.sync();
                        result = { comments: withLoc.map(({ c, loc }) => ({ address: chatQualifiedAddress(sheet.name, loc.address), content: c.content, author: c.authorName, resolved: c.resolved })) };
                        break;
                    }
                    case "add": {
                        if (!p.cell || !p.content) throw new Error("cell and content are required for add");
                        const { sheet, range } = chatGetRangeAndSheet(ctx, p.cell);
                        sheet.load("name");
                        sheet.comments.add(range, p.content);
                        await ctx.sync();
                        result = { message: `Added comment to ${chatQualifiedAddress(sheet.name, range.address || p.cell)}` };
                        break;
                    }
                    case "reply": {
                        if (!p.cell || !p.content) throw new Error("cell and content are required for reply");
                        const comment = await chatFindCommentAtCell(ctx, p.cell);
                        if (!comment) throw new Error(`No comment found at ${p.cell}`);
                        comment.replies.add(p.content);
                        await ctx.sync();
                        result = { message: `Replied to comment at ${p.cell}` };
                        break;
                    }
                    case "resolve":
                    case "reopen": {
                        if (!p.cell) throw new Error("cell is required for resolve/reopen");
                        const comment = await chatFindCommentAtCell(ctx, p.cell);
                        if (!comment) throw new Error(`No comment found at ${p.cell}`);
                        comment.resolved = p.action === "resolve";
                        await ctx.sync();
                        result = { message: `${p.action === "resolve" ? "Resolved" : "Reopened"} comment at ${p.cell}` };
                        break;
                    }
                    case "delete": {
                        if (!p.cell) throw new Error("cell is required for delete");
                        const comment = await chatFindCommentAtCell(ctx, p.cell);
                        if (!comment) throw new Error(`No comment found at ${p.cell}`);
                        comment.delete();
                        await ctx.sync();
                        result = { message: `Deleted comment at ${p.cell}` };
                        break;
                    }
                    default: throw new Error(`Unknown action "${p.action}"`);
                }
            });
            return { ok: true, ...result };
        } catch (e) {
            return { error: String(e.message || e) };
        }
    }
    // Shared by reply/resolve/reopen/delete — Office.js has no "get comment at this cell" lookup,
    // only "get this comment's cell", so finding the right comment means loading every comment on
    // the sheet and matching addresses ourselves. Must run inside the SAME Excel.run as the caller
    // (shares the caller's context/sync cadence), so this takes `ctx` rather than opening its own.
    async function chatFindCommentAtCell(ctx, cellRef) {
        const { sheet, range } = chatGetRangeAndSheet(ctx, cellRef);
        sheet.load("name");
        range.load("address");
        const comments = sheet.comments;
        comments.load("items");
        await ctx.sync();
        const targetAddr = range.address.includes("!") ? range.address.slice(range.address.indexOf("!") + 1) : range.address;
        const withLoc = comments.items.map(c => { const loc = c.getLocation(); loc.load("address"); return { c, loc }; });
        await ctx.sync();
        const match = withLoc.find(({ loc }) => {
            const a = loc.address.includes("!") ? loc.address.slice(loc.address.indexOf("!") + 1) : loc.address;
            return a === targetAddr;
        });
        return match ? match.c : null;
    }

    function toolLabel(name, args) {
        switch (name) {
            case "get_workbook_overview": return args.sheet_name ? `Inspecting sheet "${args.sheet_name}"` : "Getting workbook overview";
            case "read_range": return `Reading ${args.range}`;
            case "recalculate_and_check": return "Recalculating and checking for errors";
            case "audit_workbook": return args.sheet_name ? `Auditing sheet "${args.sheet_name}"` : "Auditing workbook health";
            case "trace_dependencies": return `Tracing ${args.mode === "dependents" ? "dependents" : "precedents"} of ${args.cell}`;
            case "search_workbook": return `Searching for "${args.query}"`;
            case "format_cells": return `Formatting ${args.range}`;
            case "create_chart": return `Creating chart from ${args.source_range}`;
            case "conditional_format": return args.action === "clear" ? `Clearing conditional formatting on ${args.range}` : `Adding conditional formatting to ${args.range}`;
            case "autofit": return args.range ? `Auto-fitting ${args.range}` : "Auto-fitting the active sheet";
            case "tables": return `${(args.action || "").replace(/^./, c => c.toUpperCase())} table${args.name ? ` "${args.name}"` : ""}`;
            case "comments": return `${(args.action || "").replace(/^./, c => c.toUpperCase())} comment${args.cell ? ` at ${args.cell}` : ""}`;
            default: return name;
        }
    }

    async function callTool(name, args) {
        switch (name) {
            case "get_workbook_overview": return toolWorkbookOverview({ sheet: args.sheet_name });
            case "read_range": return toolReadRange(args);
            case "create_sheet": return toolCreateSheet(args);
            case "write_cells": return toolWriteCells(args);
            case "fill_formula": return toolFillFormula(args);
            case "recalculate_and_check": return toolRecalculate({ sheet: args.sheet_name });
            case "audit_workbook": return toolAuditWorkbook({ sheet: args.sheet_name, tolerance: args.tolerance });
            case "undo_last_change": return toolUndoLastChange();
            case "trace_dependencies": return toolTraceDependencies(args);
            case "search_workbook": return toolSearchWorkbook({ ...args, sheet: args.sheet_name });
            case "format_cells": return toolFormatCells(args);
            case "create_chart": return toolCreateChart(args);
            case "conditional_format": return toolConditionalFormat(args);
            case "modify_structure": return toolModifyStructure(args);
            case "named_ranges": return toolNamedRanges(args);
            case "autofit": return toolAutofit(args);
            case "tables": return toolTables(args);
            case "comments": return toolComments(args);
            default: return { error: `Unknown tool "${name}"` };
        }
    }

    async function send() {
        if (busy) return;
        if (!chatBalanceSufficient) return; // button is already disabled; this covers Enter-to-send
        const q = input.value.trim();
        if (!q) return;

        // Fall back to the main Company-tab dropdown if the user never touched the chat's own picker.
        if (!primaryCo) {
            const t = byId("dropdownToggle");
            if (t && t.dataset.value) primaryCo = { fincode: t.dataset.value, name: t.value, sector: t.dataset.sector };
        }

        addMsg("user", q);
        input.value = ""; input.style.height = "auto";
        history.push({ role: "user", content: q });
        persistHistory();
        const thinking = addMsg("bot", "…");
        busy = true; sendBtn.disabled = true;
        const abortController = new AbortController();
        let mcpClient = null;
        try {
            const contextNote = primaryCo
                ? `Selected in the add-in — primary: ${primaryCo.name} (fincode ${primaryCo.fincode}, sector ${primaryCo.sector || "unknown"})` +
                  (compareCo ? `; compare: ${compareCo.name} (fincode ${compareCo.fincode}, sector ${compareCo.sector || "unknown"}).` : ".")
                : "No company is currently selected in the add-in.";

            // Connect to GoIndia's MCP server directly and pull its tool list in for this turn.
            let mcpOpenAITools = [], mcpUnavailableReason = "";
            const access = await ensureMcpAccess();
            if (access.flag !== 3) {
                mcpUnavailableReason = access.flag === 2 ? "MCP access has been revoked for this account."
                    : access.flag === 1 ? "this account doesn't have an active MCP research subscription."
                    : "couldn't verify MCP research access for this account.";
            } else {
                const mcpCard = addCard("Connecting to GoIndia MCP…", "run");
                const mcpSt = mcpCard.querySelector(".chat-card-st");
                try {
                    const mcpUrl = `${MCP_SERVER_URL}?token=${encodeURIComponent(access.mcp_api_key)}&skills=false`;
                    mcpClient = new McpClient(mcpUrl, abortController.signal);
                    await mcpClient.open();
                    await mcpClient.initialize();
                    mcpOpenAITools = toOpenAITools(await mcpClient.listTools());
                    mcpCard.className = "chat-card done"; mcpSt.textContent = "✓";
                } catch (e) {
                    mcpUnavailableReason = String(e.message || e);
                    mcpCard.className = "chat-card err"; mcpSt.textContent = "✕";
                }
            }

            const system =
                `You are GoIndia Stocks' in-Excel research assistant. ${contextNote}\n` +
                (mcpOpenAITools.length
                    ? `You have direct tool access to GoIndia's research data platform (financials, earnings calls, sentiment, broker reports, sector/market data, and more) via the tools below, plus tools to read/write the user's open workbook. Never invent or guess figures — fetch them.\n`
                    : `GoIndia's research data tools are unavailable this turn (${mcpUnavailableReason}) — you only have the Excel read/write tools below. If the user's request needs data you don't already have from this conversation, say so rather than inventing numbers.\n`) +
                `Rules:\n` +
                `- If a data tool needs a company identifier you don't have, look it up via the available tools first — never guess.\n` +
                `- Call get_workbook_overview at the start of a task (or read_range on the specific area) before writing any formula that references existing cells — never guess row/column positions. Prefer a named range from the overview over a guessed cell address when one exists.\n` +
                `- write_cells/fill_formula only work on a sheet that already exists — if the user names a sheet that isn't in the workbook (check get_workbook_overview first), call create_sheet before trying to write to it, rather than reporting that you "can't" create it.\n` +
                `- Use write_cells for values/formulas at specific cells, and fill_formula (never a hand-built 2D formula array) when the same formula repeats down or across a range — relative references adjust automatically. Both only touch the cells you specify; nothing else on the sheet is affected.\n` +
                `- Every write is shown to the user for approval before it happens unless they've turned on auto-apply. If a write_cells/fill_formula/undo_last_change result comes back with rejected: true or blocked: true, tell the user rather than assuming it happened — only retry with allow_overwrite:true if the user actually confirms they want that data overwritten.\n` +
                `- After a write that could affect other formulas, call recalculate_and_check and mention any errors it finds rather than assuming success. If a change turns out wrong, use undo_last_change rather than trying to manually reconstruct the previous state.\n` +
                `- Use trace_dependencies to explain WHY a value is what it is (trace its precedents) instead of re-reading formulas one cell at a time. Use search_workbook to locate where something is calculated or referenced instead of reading every sheet.\n` +
                `- Visual/structural requests have their own tools — format_cells (bold, currency, colors, borders), conditional_format (highlight/color-scale by value), create_chart (visualize a comparison), autofit (fix narrow columns after writing new data), tables (make a range sortable/filterable), named_ranges (so a cell can be called "WACC" instead of guessed by address), modify_structure (insert/delete rows/columns, rename/hide a sheet), and comments (leave a note explaining a change). Prefer these over write_cells for anything that isn't a value or formula.\n` +
                `- Compute derived figures yourself (YoY growth, CAGR, margins, ratios, averages) from the numbers you fetch.\n` +
                `- Be concise in your final answer.`;

            const combinedTools = EXCEL_TOOLS.concat(mcpOpenAITools);
            const convo = [{ role: "system", content: system }, ...history.slice(-8)];
            let full = "";
            // Summed across every OpenRouter call this question takes (tool-calling round trips
            // included, plus the forced-final-answer fallback below if it fires) — deducted from
            // the wallet and logged once the full answer is in, see deductChatCost near the bottom.
            let totalCostUsd = 0;
            // A real MCP research/model-editing question routinely needs many round trips —
            // search_company, get_catalog, get_schema, one or more run_query attempts (the first
            // query often isn't quite right and needs a retry), then possibly audit_workbook/
            // trace_dependencies/write_cells on top of that. 6 was too tight and left "full" empty
            // (see the fallback below) even though the model was mid-investigation, not stuck or
            // erroring; raised again since even 10 was cutting off legitimately multi-step work,
            // then again from 20 to 30 for the same reason.
            const MAX_TOOL_ITERS = 30;
            for (let iter = 0; iter < MAX_TOOL_ITERS; iter++) {
                const res = await fetch(OPENROUTER_URL, {
                    method: "POST",
                    headers: { "Content-Type": "application/json", "Authorization": `Bearer ${OPENROUTER_KEY}` },
                    body: JSON.stringify({ model: AI_MODEL, messages: convo, tools: combinedTools, temperature: 0.2, max_tokens: 3000, usage: { include: true } })
                });
                if (!res.ok) throw new Error(`OpenRouter ${res.status}: ${(await res.text()).slice(0, 160)}`);
                const data = await res.json();
                const cost = data.usage?.cost;
                if (cost != null) totalCostUsd += Number(cost) || 0;
                console.log(`[Chat] 💰 turn ${iter + 1} cost: ${cost != null ? "$" + Number(cost).toFixed(5) : "n/a"} | tokens: ${data.usage?.total_tokens ?? "?"}`);

                const msg = data.choices?.[0]?.message;
                if (!msg) break;
                convo.push(msg);

                const calls = msg.tool_calls || [];
                if (!calls.length) { full = msg.content || ""; break; }

                for (const call of calls) {
                    let args = {};
                    try { args = JSON.parse(call.function.arguments || "{}"); } catch (e) { /* leave empty */ }
                    let result;
                    if (EXCEL_SELF_MANAGED_UI_TOOLS.has(call.function.name)) {
                        try { result = await callTool(call.function.name, args); } // renders its own confirm/running card
                        catch (e) { result = { error: String(e.message || e) }; }
                    } else if (EXCEL_TOOL_NAMES.has(call.function.name)) {
                        const card = addCard(`${toolLabel(call.function.name, args)}…`, "run");
                        try { result = await callTool(call.function.name, args); await settleCard(card, "done", "✓"); }
                        catch (e) { result = { error: String(e.message || e) }; await settleCard(card, "err", "✕"); }
                    } else {
                        // MCP-provided tool
                        const card = addCard(`Running "${call.function.name}"…`, "run");
                        try {
                            const r = await mcpClient.callTool(call.function.name, args);
                            result = r.isError ? { error: r.text } : { text: capText(r.text) };
                            await settleCard(card, r.isError ? "err" : "done", r.isError ? "✕" : "✓");
                        } catch (e) { result = { error: String(e.message || e) }; await settleCard(card, "err", "✕"); }
                    }
                    convo.push({ role: "tool", tool_call_id: call.id, content: JSON.stringify(result).slice(0, 6000) });
                }
            }

            // Ran out of iterations while the model was still calling tools (never hit the
            // !calls.length break above) — force one last text-ONLY turn (no "tools" param at all,
            // so it physically cannot call another one) so whatever was actually discovered along
            // the way still reaches the user, instead of silently showing "(no response)" for a
            // question that was genuinely being worked on, not stuck or erroring.
            if (!full) {
                try {
                    const res = await fetch(OPENROUTER_URL, {
                        method: "POST",
                        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${OPENROUTER_KEY}` },
                        body: JSON.stringify({
                            model: AI_MODEL,
                            messages: [...convo, { role: "user", content: "Give your best answer now based on everything gathered above — no more tool calls." }],
                            temperature: 0.2, max_tokens: 3000, usage: { include: true },
                        }),
                    });
                    if (res.ok) {
                        const data = await res.json();
                        full = data.choices?.[0]?.message?.content || "";
                        const cost = data.usage?.cost;
                        if (cost != null) totalCostUsd += Number(cost) || 0;
                    }
                } catch (e) { /* fall through to "(no response)" below if even this fails */ }
            }

            // (No fenced-```excel``` fallback anymore — that existed only to catch write_excel_sheet
            // being emitted as text instead of a real tool call. write_cells/fill_formula are plain
            // function-calling tools like every other tool here, with no bespoke action-JSON shape
            // to fall back to; if the model doesn't call a tool properly, it simply doesn't act —
            // the same failure mode as any other tool-calling miss.)
            //
            // "thinking" was appended to the stream FIRST, before any of this turn's tool-progress
            // cards — left in place, the answer would fill in at the TOP of a growing list of cards
            // instead of at the bottom, so as the window grew taller the user had to scroll back UP
            // to notice it had finished. Move it to the end now that every card for this turn has
            // already been added, so the answer always lands as the last, most-visible thing.
            stream.appendChild(thinking.parentElement); // move the .chat-msg wrapper, not just the inner .chat-bubble — moving the bubble alone detaches it from the wrapper whose .chat-msg.bot .chat-bubble rule supplies its background/border
            stream.scrollTop = stream.scrollHeight;
            thinking.innerHTML = renderMarkdown(full || "(no response)");
            history.push({ role: "assistant", content: full });
            persistHistory();
            // Query is fully complete (tool calls done, final answer rendered) — deduct+log the
            // summed cost now. Fire-and-forget, same as the financial-model deduction elsewhere.
            if (totalCostUsd > 0) deductChatCost({ question: q, costUsd: totalCostUsd });
        } catch (e) {
            stream.appendChild(thinking.parentElement); // move the .chat-msg wrapper, not just the inner .chat-bubble — moving the bubble alone detaches it from the wrapper whose .chat-msg.bot .chat-bubble rule supplies its background/border
            stream.scrollTop = stream.scrollHeight;
            thinking.innerHTML = renderMarkdown("⚠ " + e.message);
            console.error("[Chat] error", e);
        } finally {
            if (mcpClient) { try { await mcpClient.close(); } catch (e) { /* already torn down */ } }
            try { abortController.abort(); } catch (e) { /* no-op if already settled */ }
            busy = false; sendBtn.disabled = !chatBalanceSufficient;
        }
    }
})();
