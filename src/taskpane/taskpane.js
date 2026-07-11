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
        const res = await fetch(`http://localhost:8000/addin/access_status?user_id=${userId}`);
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
                const resStatus = await fetch(`http://localhost:8000/addin/access_status?user_id=${user.UserId}`);
                const data = await resStatus.json();

                if (data.status === "Approved") {
                    // Treat this user as an enterprise member
                    localStorage.setItem("user", JSON.stringify(user));
                    localStorage.setItem("membership", JSON.stringify(membership));

                    document.getElementById("memberContent").style.display = "block";
                    document.getElementById("addinUI").style.display = "block";
                    document.getElementById("logoutBtn").style.display = "block";

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
                    const res = await fetch("http://localhost:8000/addin/request_access", {
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
                    const resStatus = await fetch(`http://localhost:8000/addin/access_status?user_id=${user.UserId}`);
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
            document.getElementById("addinUI").style.display = "block";
            document.getElementById("logoutBtn").style.display = "block";

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
            document.getElementById("addinUI").style.display = "block";
            document.getElementById("logoutBtn").style.display = "block";

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
            document.getElementById("addinUI").style.display = "block";
            document.getElementById("logoutBtn").style.display = "block";

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

        // Show blocked state on error
        document.getElementById("loginScreen").style.display = "none";
        document.getElementById("memberContent").style.display = "none";
        document.getElementById("blockedContainer").style.display = "flex";
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
                    const resStatus = await fetch(`http://localhost:8000/addin/access_status?user_id=${user.UserId}`);
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
                document.getElementById("addinUI").style.display = "block";
                document.getElementById("logoutBtn").style.display = "block";

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

const MAX_MODEL_COST_INR = 25; // matches the AI Financial Model card's own displayed cost ceiling
const TEST_BYPASS_USER_ID = "test123"; // ut@gmail.com certification account — see refreshWalletBalance

// Disables the AI Financial Model toggle (and unchecks it if it was on); re-enables it otherwise.
function setAiModelAvailability(sufficientBalance) {
    const check = document.getElementById("aiModelCheck");
    if (!check) return;
    check.disabled = !sufficientBalance;
    if (!sufficientBalance) check.checked = false;
}

// Fetches the current user's wallet balance (sum of all three credit buckets, per backend
// convention), updates the pill in the header, and gates the AI Financial Model toggle on it.
// Fails CLOSED: a real account with an unverifiable balance (network error, non-200, etc.) gets
// the toggle disabled, same as a confirmed-insufficient balance — we should never let a real user
// generate a model we don't know they can afford. The one exemption is the ut@gmail.com
// certification bypass (UserId "test123"), which always fails this fetch since it isn't a real
// wallet user — it stays enabled unconditionally so Microsoft's reviewers can test the feature.
async function refreshWalletBalance() {
    const walletBalanceEl = document.getElementById("walletBalance");
    if (!walletBalanceEl) return;
    let userId = null;
    try {
        const user = JSON.parse(localStorage.getItem("user") || "{}");
        userId = user.UserId || null;
        if (!userId) { walletBalanceEl.textContent = "—"; setAiModelAvailability(false); return; }
        const isTestAccount = userId === TEST_BYPASS_USER_ID;
        if (isTestAccount) setAiModelAvailability(true);

        const res = await fetch("https://transcriptanalyser.com/wallet/user_credit_balance", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ user_id: String(userId) })
        });
        if (!res.ok) {
            walletBalanceEl.textContent = "—";
            if (!isTestAccount) setAiModelAvailability(false);
            return;
        }
        const data = await res.json();
        const c = data.credits || {};
        const total = (c.whatsapp_credit || 0) + (c.stockgpt_credit || 0) + (c.balance_credit || 0);
        walletBalanceEl.textContent = `₹${total.toFixed(2)}`;
        setAiModelAvailability(isTestAccount || total >= MAX_MODEL_COST_INR);
    } catch (e) {
        console.warn("Wallet balance fetch failed:", e);
        walletBalanceEl.textContent = "—";
        if (userId !== TEST_BYPASS_USER_ID) setAiModelAvailability(false);
    }
}

// wallet.py's two excel_-prefixed endpoints, deployed alongside the pre-existing
// /wallet/user_credit_balance on the same backend.
const WALLET_BASE_URL = "https://transcriptanalyser.com/wallet";
const WALLET_DEDUCT_URL = `${WALLET_BASE_URL}/excel_deduct_model_cost`;

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

// ── Wallet history popup ──
// Reads /wallet/excel_cost_history (fincode + cost only), then resolves each unique fincode to a
// company name via the existing company_search endpoint (passing a fincode as searchtxt returns
// that company as a match, same as passing a name) — cached per-open so repeat fincodes in the
// list don't refetch.
async function fetchWalletHistory() {
    const listEl = document.getElementById("walletHistoryList");
    if (!listEl) return;
    listEl.textContent = "Loading…";
    try {
        const { userId, isEnterprise, enterpriseId } = getWalletIdentity();
        if (!userId) { listEl.textContent = "Not logged in."; return; }
        const qs = isEnterprise ? `enterprise_id=${enterpriseId}` : `user_id=${encodeURIComponent(userId)}`;
        const res = await fetch(`${WALLET_BASE_URL}/excel_cost_history?${qs}&limit=20`);
        if (!res.ok) { listEl.textContent = "Couldn't load history."; return; }
        const data = await res.json();
        const entries = data.entries || [];
        if (!entries.length) { listEl.textContent = "No AI model charges yet."; return; }

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

        listEl.innerHTML = "";
        for (const entry of entries) {
            const name = await resolveCompanyName(entry.fincode);
            const row = document.createElement("div");
            row.style.cssText = "padding:8px 0; border-bottom:1px solid #eee;";
            const dt = entry.created_utc ? new Date(entry.created_utc) : null;
            row.innerHTML = `
                <div style="display:flex; justify-content:space-between; gap:8px;">
                    <span>${name}</span>
                    <span style="font-weight:700; color:#173760; white-space:nowrap;">₹${Number(entry.cost_inr).toFixed(2)}</span>
                </div>
                <div style="font-size:10.5px; color:#999; margin-top:2px;">${dt ? dt.toLocaleString() : ""}${entry.used_fallback ? " · fallback model" : ""}</div>
            `;
            listEl.appendChild(row);
        }
    } catch (e) {
        console.warn("[Wallet] excel_cost_history fetch failed:", e);
        listEl.textContent = "Couldn't load history.";
    }
}

// webpack scopes top-level functions inside its own module wrapper, not onto window — so
// taskpane.html's separate inline <script> can't see these by bare name even though a
// `typeof x === "function"` guard makes that failure silent instead of throwing. Explicitly
// exposing the two functions taskpane.html actually calls across that boundary.
window.refreshWalletBalance = refreshWalletBalance;
window.fetchWalletHistory = fetchWalletHistory;

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
                document.getElementById("addinUI").style.display = "block";
                document.getElementById("logoutBtn").style.display = "block";

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



Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) return;

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
                if (ranges["Annual Data"])    annualText    = readSheet(ranges["Annual Data"].text, 120, 30);
                if (ranges["Quarterly Data"]) quarterlyText = readSheet(ranges["Quarterly Data"].text, 120, 30);
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
            const buildIndex = (grid, sheetName) => {
                const map = new Map();
                if (!Array.isArray(grid) || !grid.length) return map;
                const headers = []; // each "Parameter" header row + its FY->column map
                for (let r = 0; r < grid.length; r++) {
                    if (normLabel(grid[r] && grid[r][0]) === "parameter") {
                        const colByFY = {}, colIdxByFY = {}, rowArr = grid[r] || [];
                        for (let c = 1; c < rowArr.length; c++) {
                            const h = String(rowArr[c] || "").trim();
                            const m = h.match(/^FY(\d{4})E?$/) || h.match(/^[A-Za-z]{3}(\d{4})$/); // FY2024 or Mar2024
                            if (m) { const fy = parseInt(m[1]); colByFY[fy] = colLetter(c + 1); colIdxByFY[fy] = c; }
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
                        // keep the LAST occurrence (Consolidated tables are written after Standalone)
                        map.set(normLabel(label), { sheet: sheetName, label: exact, colByFY: h.colByFY, r1, r2, row: r + 1, valByFY });
                    }
                }
                return map;
            };
            const idxKF = buildIndex(grids["Key Financials"], "Key Financials");
            const idxAnnual = buildIndex(grids["Annual Data"], "Annual Data");
            const resolveLink = (label) => idxKF.get(normLabel(label)) || idxAnnual.get(normLabel(label)) || null;

            // Resolve a model row to a data-sheet cell. An explicit "link" is trusted; otherwise the
            // row's label/key is tried but ACCEPTED only when a data value matches one of the row's
            // own historicals — so we link far more rows without risking a wrong auto-match.
            const stripUnit = (s) => String(s || "").replace(/\([^)]*\)/g, " ").trim();
            const valuesAgree = (a, b) => { if (a == null || b == null) return false; const m = Math.max(Math.abs(a), Math.abs(b), 1); return Math.abs(a - b) / m < 0.02; };
            const hit = (a, b) => a != null && b != null && Math.abs(a) > 0.5 && valuesAgree(a, b); // non-trivial agreement
            const autoLink = (item) => {
                if (!item) return null;
                if (item.link) { const e = resolveLink(item.link); if (e) return e; } // explicit link — trusted
                const hist = (Array.isArray(item.historical) ? item.historical : []).map(parseNum);
                // fast path: a label/key match confirmed by an agreeing value
                for (const cand of [item.label && resolveLink(stripUnit(item.label)), item.key && resolveLink(String(item.key).replace(/_/g, " "))]) {
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

            // Step 1a: Fetch operational data directly from the dashboard API (no truncation).
            // Honour the user's selection: skip entirely if unchecked, and keep only the
            // chosen sections when a subset was picked in the modal.
            let operationalText = "";
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

            // ── Step 2: Period axis — 5 historical actuals + 5 forecast years (UltraTech-style) ──
            // Latest actual fiscal year = max non-forecast FY found in the Key Financials header.
            const fyMatches = (financialText.match(/FY(\d{4})(?!E)/g) || [])
                .map(s => parseInt(s.slice(2))).filter(n => n >= 2000 && n <= 2100);
            const latestActualFY = fyMatches.length ? Math.max(...fyMatches) : new Date().getFullYear();
            const HIST = [], FCST = [];
            for (let i = 4; i >= 0; i--) HIST.push(latestActualFY - i);   // 5 historical years
            for (let i = 1; i <= 5; i++) FCST.push(latestActualFY + i);   // 5 forecast years
            const histLabels = HIST.map(y => `FY${y}A`);
            const fcLabels = FCST.map(y => `FY${y}E`);
            const periodLabels = [...histLabels, ...fcLabels];            // 10 column labels (B..K)
            const HIST_COLS = ["B", "C", "D", "E", "F"];
            const FC_COLS = ["G", "H", "I", "J", "K"];
            const ALL_COLS = [...HIST_COLS, ...FC_COLS];
            const CAGR_COL = "L";
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

            // Data bundles fed to the calls (statement sheets get the numeric block;
            // the assumptions call also gets the qualitative sources).
            const DATA_BASIS_NOTE = `DATA BASIS — the figures below may include more than one presentation (e.g. an "IndAS" series and a more granular "Detailed" series, and standalone vs consolidated). YOU decide which to use: prefer CONSOLIDATED figures on a SINGLE consistent reporting basis across every year and every sheet; use the Detailed breakdown only for extra line-item granularity. Do not mix bases within a series. When more than one row could plausibly be "the" figure for a line item (e.g. two similarly-labelled EBIT/EBITDA rows from different presentations), deterministically prefer whichever appears FIRST in the Key Financials sheet — do not switch between presentations across a rerun of this same prompt, since that produces a different "historical actual" figure for the same real year each time, which should never happen for a reported number.\nLINKING — whenever a row's historicals come from a line item present in the data, set "link" to that item's EXACT label (from Key Financials OR the Detailed Annual Financials below) so its historical cells become LIVE references; ALSO include "historical" numbers as a fallback. If an item is not in the data, just give "historical" numbers. Do NOT link percentage / margin / ratio rows — provide those as DECIMAL fractions (0.069 for 6.9%, never 6.9) or compute them; only link absolute figures (currency amounts, volumes, counts, share counts).\nFLOW vs STOCK — do not link a row to a similarly-worded label of the WRONG kind (e.g. a per-year "Depreciation" flow row must link a P&L-style annual depreciation label, never an "Accumulated/Cumulative Depreciation" balance label; a "Capex" flow row must never link a "Gross Block" balance label). If you are not sure a label matches the concept, prefer "historical" numbers you derive yourself over a mismatched link.\nROLL-FORWARD CHECK — for balance-sheet stock items you carry across years (gross block, accumulated depreciation, equity, debt), the LATEST historical year's linked/entered value should be consistent with prior-year value + this year's flow (e.g. Gross Block ≈ prior Gross Block + this year's Capex; Accumulated Depreciation ≈ prior Accumulated Depreciation + this year's Depreciation; Equity ≈ prior Equity + this year's PAT − Dividends). If the source data is stale for the latest year (identical to the prior year despite a non-zero flow that year — a common data-lag artifact), use the roll-forward-implied value for that one year instead of silently repeating the prior year's balance.`;
            const finBlock = `${DATA_BASIS_NOTE}\n\nKEY FINANCIALS (downloaded sheet):\n${financialText}\n`
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
 "historical":[v1,v2,v3,v4,v5]            (the 5 historical values ${histLabels.join("/")} read straight from the data; ALWAYS provide these — EXCEPT for a row whose own instructions below explicitly say to give ONLY a formula, e.g. "do NOT give this row historical or link". For those specific rows, omit BOTH "historical" AND "link" entirely — do not guess numbers just because this schema default says to always provide them; a guessed number there silently overrides the required formula and has caused real bugs (a working-capital-change row showing round guessed numbers, an operating-cash-flow row showing a plain PAT copy) instead of the correct cross-sheet figure),
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
 "historical":[v1,v2,v3,v4,v5]            (5 historical values ${histLabels.join("/")} read straight from the data; ALWAYS provide these),
 "forecast":[v1,v2,v3,v4,v5]              (5 FORECAST input values ${fcLabels.join("/")}),
 "formula":"<recurrence per the FORMULA GRAMMAR>" (OPTIONAL alternative to "forecast" for computed driver rows — normally same-sheet {R:}/{P:} for driver-to-driver relationships; {S:sheet.key}/{PS:sheet.key} IS also valid here, e.g. a repayment schedule referencing {S:pnl.pat}. Do NOT put an ABSOLUTE RUPEE line item on THIS sheet if computing it correctly needs a figure that only exists on another sheet, like revenue — this sheet has no revenue row of its own, so "{R:revenue}"-style same-sheet references for it will silently fail and resolve to 0. Keep such figures as a RATIO driver here (e.g. capex as % of revenue) and let the sheet that actually has revenue compute the rupee amount),
 "source":"<short citation>"              (REQUIRED whenever "input":true AND the number is a market/judgment figure not directly implied by the downloaded historicals — e.g. beta, risk-free rate, equity risk premium, cost of debt, target multiples, terminal growth, or a forward margin/growth call. Keep it short and concrete, e.g. "NSE 3Y adjusted beta", "10Y G-Sec yield, Jul-2026", "Sector avg EV/EBITDA (peer comps)", "Mgmt guidance, Q1FY27 earnings call". Omit for rows that are plainly derived/linked from the data.)}
The sheet is laid out ACROSS YEARS (columns ${periodLabels.join(", ")}). Other sheets read a driver via {A:key} from the SAME year column, so for CONSTANT global inputs repeat the same value in ALL 5 historical AND 5 forecast cells. None of the constant global inputs may be 0 or blank. wacc MUST be at least 1.5 percentage points greater than terminal_growth (typical: wacc approx 0.11-0.14, terminal_growth approx 0.03-0.05) so the DCF terminal-value division never approaches zero.`;
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
            const [structJson, judgJson] = await Promise.all([
                callLLM("Assumptions (Structural)",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${assumSchema}\n\n${finBlock}\nBuild PART of a DETAILED ASSUMPTIONS / DRIVER SCHEDULE for ${companyName} — ONLY these sections (a separate call handles WACC/growth/margin/capex/working-capital-day assumptions, so do NOT include those here): SHARE CAPITAL (shares outstanding, face value — key=shares_out), RESERVES & DEBT (reserves, secured/unsecured debt balances), FIXED ASSETS ROLL-FORWARD (gross block beginning/additions/ending, accumulated depreciation, net block, CWIP, depreciation as % of avg gross block — key=depreciation_pct_gross_block: give this as a PLAIN POSITIVE DECIMAL VALUE (NOT a formula), computed EXACTLY as this company's OWN annual depreciation charge / average reported gross block, using ITS actual reported figures — do NOT force it toward a generic "typical" range; the correct ratio varies enormously by industry (asset-heavy, long-asset-life businesses like refining, oil & gas, power, telecom towers legitimately run well under 3%, while asset-light services businesses can run well over 8% — trust this company's own historical arithmetic, not a rule of thumb), and repeat that SAME value in all 10 year cells. Do NOT compute it with a same-sheet formula that divides by a "gross block ending" roll-forward row — that roll-forward starts unanchored and can collapse to ~0, which makes this ratio 0 and SILENTLY ZEROES the entire forecast depreciation → EBIT → PAT chain on the P&L. It MUST be a non-zero positive decimal in every cell. Also make the "gross block ending" display row anchor correctly as beginning + additions using same-year {R:} references (e.g. "{R:gross_block_beginning}+{R:capex}"), never rolled off only the prior-year ending. CONTINUITY SELF-CHECK — before finalizing, multiply your depreciation_pct_gross_block by the projected average gross block for the FIRST forecast year and confirm the result is reasonably close to (not a multi-x jump from) the LAST historical year's actual depreciation charge, since gross block does not change overnight; if your ratio produces a first-forecast-year figure several times larger or smaller than the last actual, you have the wrong ratio — recompute it directly from the historical annual-charge ÷ average-gross-block arithmetic instead. The P&L sheet references this EXACT key to forecast the annual depreciation charge), INVESTMENTS, WORKING CAPITAL BALANCES (inventories, debtors, cash, loans & advances, payables, provisions — the ACTUAL reported rupee amounts; do NOT include the days/ratio assumptions for these, that's the other call). This is pure transcription/linking from the data, no forward-looking judgment needed here.\nHISTORICAL-ONLY ROWS — every row you produce in THIS ENTIRE CALL, in EVERY section (SHARE CAPITAL, RESERVES & DEBT, FIXED ASSETS ROLL-FORWARD, INVESTMENTS, and WORKING CAPITAL BALANCES alike — this explicitly includes gross_block_beginning, the additions/capex row, and accumulated_depreciation_beginning under FIXED ASSETS ROLL-FORWARD, not only the 3 sections that happen to be named in this sentence), EXCEPT shares_out and depreciation_pct_gross_block, is a REFERENCE/DISPLAY snapshot only — no other sheet reads it. Give these rows ONLY "link"/"historical" (numbers straight from the data). Do NOT give them a "forecast" array or a "formula", and above all do NOT invent a growth-rate key like "{A:long_term_debt_growth}" or a cross-sheet pull like "{S:pnl.dividends}"/"{S:pnl.revenue}"/"{S:pnl.cogs}"/"{S:capex.capex_total}" for them — none of those keys exist anywhere (the real ones are net_revenue, not revenue; capex, not capex_total; there is no separate "cf" sheet at all, cash flow lives on capex; and there is no "cogs" or "dividends" row exposed by any sheet), and referencing an invented one resolves to a visible #MISSING error. The Balance Sheet sheet is the ACTUAL, correctly-forecast source for every one of these line items (receivables, inventory, payables, provisions, borrowings, investments) — leaving these rows historical-only here causes them to display a sensible flat continuation automatically, which is the correct behavior for a reference snapshot.\nSHARES OUTSTANDING UNIT — shares_out MUST be expressed in CRORES OF SHARES, consistent with every other figure in this model being in Rs Crores (e.g. a company with 676 crore shares outstanding — 6.76 billion / 6,760 million shares — must show shares_out = 676, NOT 6760000000, NOT 6760, NOT 6.76). If the source data reports the share count in millions or as a raw absolute count, CONVERT it yourself (crore = raw_count / 1e7 = million_count / 10) before writing shares_out. Only ONE row may describe share count (key="shares_out") — do not also add a separate informational "shares outstanding" row.\nYou MUST include rows keyed EXACTLY, with a value in every one of the 10 year cells: shares_out (source citation REQUIRED), depreciation_pct_gross_block (no source citation needed, this comes from the fixed-asset roll-forward, not market data). Be thorough (20-30 rows including section headers).`,
                    STATIC_MODEL),
                callLLM("Assumptions (Judgment)",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${assumSchema}\n\n${finBlockJudgment}${qualBlock}\nBuild PART of a DETAILED ASSUMPTIONS / DRIVER SCHEDULE for ${companyName} — ONLY these sections (a separate call handles share capital/debt/fixed-asset/investment/working-capital-balance transcription, so do NOT include those here): CAPEX (CAPEX-AS-%-OF-REVENUE as a RATIO driver ONLY — key=capex_pct_revenue, informed by any capex guidance in the qualitative sources; do NOT put an absolute-rupee "capex" amount on this sheet: this sheet has no revenue row to multiply against, so a same-sheet reference for it would silently resolve to 0 — the Capex & FCF sheet computes the actual rupee capex from this ratio via {S:pnl.net_revenue}), WORKING CAPITAL DRIVERS (debtor_days, inventory_days, payable_days, min_cash_pct_revenue — a minimum operating-cash buffer as a decimal fraction of revenue, informed by the company's own average historical cash-to-revenue level; this floor is what keeps the Balance Sheet's cash plug from ever going negative, so it MUST be set from the company's real historical cash/revenue ratio, not left at a token value), OTHER INCOME, P&L DRIVERS (revenue growth %, EBITDA margin %, tax rate, dividend payout — informed by earnings-call guidance, broker views, and management commentary where available), and VALUATION & WACC INPUTS. These are the rows that genuinely require interpreting the qualitative sources or making a forward-looking call — use them.\nSELF-CHECK before finalizing any per-share-driven assumption: values should stay in a plausible range given the company's scale — if something looks off by orders of magnitude, you likely have a unit mismatch; fix it, do not leave it. Also keep capex intensity (capex_pct_revenue) and dividend payout consistent with management's stated plans and the company's recent history — do NOT let a genuinely capex-heavy, debt-carrying company deleverage all the way to NET CASH within the forecast window unless the data clearly supports it (that unrealistically collapses forecast interest expense toward zero); if the company is guiding continued heavy investment, keep capex_pct_revenue and payout high enough that net debt stays realistic.\nYou MUST include rows keyed EXACTLY (constant global inputs — repeat the same value in every one of the 10 year cells), every one with a "source" citation: current_price, tax_rate, pat_retention, wacc, terminal_growth, risk_free_rate, equity_risk_premium, beta, cost_of_debt, target_pe, target_ev_ebitda. wacc MUST be at least 1.5 percentage points greater than terminal_growth (typical: wacc approx 0.11-0.14, terminal_growth approx 0.03-0.05) so the DCF terminal-value division never approaches zero. You MUST ALSO include these EXACT keys (no source citation needed, these come from working-capital/capex analysis not market data): capex_pct_revenue (a decimal fraction of revenue, e.g. 0.14 for 14%), debtor_days, inventory_days, payable_days (all in DAYS, e.g. 45, not decimals), min_cash_pct_revenue (a decimal fraction of revenue, e.g. 0.03 for 3% — set from the company's own historical average cash/revenue ratio; must be a positive decimal in every one of the 10 year cells, never 0 or blank, since the Balance Sheet's revolver formula divides the whole forecast's cash-balancing logic on this key), revenue_growth, ebitda_margin (the TOTAL-COMPANY forward growth/margin call, as decimals — the Operational Model sheet references these EXACT keys for every segment's forecast, since it has no way to know segment-specific splits and this call has no way to know segment names in advance; one blended company-wide rate is the deliberate, honest simplification here, not a bug). REALISTIC DRIVER PATHS — do NOT repeat a single flat value across all 5 forecast years for the OPERATING drivers (revenue_growth, ebitda_margin, capex_pct_revenue, debtor_days, inventory_days, payable_days): give each a realistic YEAR-BY-YEAR path in its 5-value forecast array (e.g. revenue growth tapering from a higher near-term rate toward a sustainable long-run rate; margins drifting toward a normalized level; capex intensity easing as a build-out completes) — informed by the qualitative guidance where available. Only the genuine market/valuation inputs (WACC, beta, risk_free_rate, equity_risk_premium, cost_of_debt, terminal_growth, tax_rate, target_pe, target_ev_ebitda) stay constant across the years.\nWORKING-CAPITAL DAYS MUST ANCHOR TO THE LAST ACTUAL — debtor_days/inventory_days/payable_days each drive a Balance Sheet formula of the form days/365*revenue, applied starting in the FIRST forecast year. The Balance Sheet's LAST HISTORICAL year, by contrast, uses the REAL reported receivables/inventory/payables figure straight from the data — a different basis. If your FIRST forecast value for any of these three days drivers is set independently of what the company's own last reported balance sheet actually implies, the first forecast year's receivables/inventory/payables will JUMP discontinuously away from the real prior-year actual, producing a large, artificial one-off working-capital swing that distorts free cash flow and the DCF (this has actually happened — a bogus ~₹69,000 Cr working-capital "change" in the very first forecast year, large enough to flip FCFF and the DCF value per share negative). To prevent this: before setting each days driver's forecast values, compute its implied historical ratio directly from the LAST historical year's own reported figures — debtor_days ≈ last-year receivables / last-year revenue × 365 (same logic for inventory_days off inventory, payable_days off payables) — and set that driver's FIRST forecast-year value to (very close to) this computed number; only THEN taper it gradually across the remaining forecast years per the guidance above. Do not start any of the three days ratios at a generic/textbook value that ignores what the company's own last reported balance sheet implies. Express percentages as decimals. Be thorough (15-20 rows including section headers).`,
                    DYNAMIC_MODEL, 40000)
            ]);
            const assumRows = [
                ...(Array.isArray(structJson.rows) ? structJson.rows : []),
                ...(Array.isArray(judgJson.rows) ? judgJson.rows : [])
            ];
            planRows("assum", assumRows);
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
- {LAST:key} or {LAST:sheet.key} -> that row's LAST (5th, terminal) forecast-year value (e.g. ending net debt for the DCF bridge).
- {FWD1:key} or {FWD1:sheet.key} -> that row's FIRST (near-term/NTM) forecast-year value — use this, NOT {LAST:}, for relative-valuation target prices: {LAST:} is 5 years out and undiscounted, so multiplying it by a multiple wildly overstates fair value versus the DCF.
- SCALAR row {"scalar":true,...}: a single output value (NOT a per-year series); give a "formula" (using {A:},{R:other_scalar},{SUM:},{LAST:},{FWD1:}) or a plain "value". All scalar values live in one column, so {R:another_scalar_key} works between scalars.
- Per-year DCF flow rows: set "forecast_only":true (don't compute historicals) AND still give a "formula" — "forecast_only" ONLY means "skip the historical columns"; it does NOT mean "give plain forecast VALUES instead of a formula." Every row in this section (FCFF, discount factor, PV of FCFF, etc.) is a MECHANICAL calculation with exactly one correct formula, never a judgment call — so it must never be a bare "forecast" array of guessed or placeholder numbers. In particular, the discount factor MUST use "1/(1+{A:wacc})^{N}" verbatim as its "formula" — writing it as a flat forecast array (e.g. [1,1,1,1,1]) silently disables ALL discounting and has actually happened, producing a DCF where every "present value" simply equals the raw undiscounted cash flow.`;

            // 3b–3g — generate the remaining sheets IN PARALLEL.
            aiStatus("AI (step 2/2): building all model sheets in parallel...");
            const [opJson, pnlJson, bsJson, capexJson, valJson, sumJson] = await Promise.all([
                callLLM("Operational Model",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild the OPERATIONAL MODEL for ${companyName}, including a per-segment/product-line breakdown where the company reports one (e.g. by business segment). Capacity, utilisation, volumes/throughput, realisations/ASP per segment, and cost-per-unit are DISCLOSURE rows — informational KPIs that do NOT feed revenue or EBITDA. Each disclosure row's forecast MUST be a FLAT carry-forward of its last actual, written EXACTLY as "{P:<this row's own key>}". Do NOT attach a growth rate, and above all do NOT reference an invented assumption key such as "{A:jio_arpu_growth}" or "{A:o2c_production_growth}" — those keys do not exist (the ONLY {A:...} keys you may use are those in the ASSUMPTION KEYS list), so they render as a #MISSING error. If you have no real driver for a KPI, a flat carry-forward IS the correct, honest forecast.\nFORECASTING SEGMENT/TOTAL REVENUE AND EBITDA — do NOT compute forecast revenue by multiplying a forecast "volume" row by a forecast "realisation/ASP" row: those two drivers' units have repeatedly ended up mismatched (observed real failures: a 10x, a 100x, and a 1000x understatement on three different segments of the same company in one run, because volume and realisation were each forecast independently and their product no longer matched revenue's own actual unit scale). Instead, forecast each segment's (and the total's) revenue as a GROWTH-RATE recurrence anchored to its OWN historical actual level, which is unit-safe by construction since it only ever multiplies a real linked/historical figure by a dimensionless growth rate — e.g. segment revenue: "{P:segment_key}*(1+{A:revenue_growth})". Forecast segment EBITDA the same way, off a margin assumption applied to that segment's OWN revenue row: "{R:segment_key_revenue}*{A:ebitda_margin}". The canonical TOTAL rows MUST use these EXACT dictated formulas (do NOT invent your own): net_revenue forecast = "{P:net_revenue}*(1+{A:revenue_growth})", and ebitda forecast = "{R:net_revenue}*{A:ebitda_margin}" (revenue times the margin assumption — this equals the sum of the segments, since the SAME company-wide margin is applied to every segment). NEVER write a growth-style recurrence for the total EBITDA row such as "{P:ebitda}*(1+{A:ebitda_growth})" — there is NO "ebitda_growth" assumption key, so it resolves to a #MISSING error that poisons pnl.ebitda, the DCF, and every sheet pulling op.ebitda.\nUSE THE EXACT ASSUMPTION KEYS "revenue_growth" AND "ebitda_margin" FROM THE ASSUMPTION KEYS LIST ABOVE for every segment revenue/EBITDA row. The ONLY growth/margin keys that exist are revenue_growth and ebitda_margin; do NOT invent "segment_x_growth", "ebitda_growth", or any other {A:...} key — an invented key exists nowhere and resolves to a #MISSING error. Applying the SAME company-wide growth/margin to every segment is the correct, deliberate simplification here.\nNO BLANK FORECASTS — every assumption you reference via {A:...} (growth rates, EBITDA margins, etc.) must itself have a value in ALL 10 year columns, never left blank — a blank/missing forecast cell is read as ZERO by Excel and silently zeroes out everything downstream that multiplies by it (e.g. an EBITDA margin left blank makes segment EBITDA compute to 0 even though revenue is fine).\nYou MUST include rows keyed exactly: net_revenue, ebitda (and volume if applicable, informational only — not used to derive net_revenue).`,
                    STATIC_MODEL),
                callLLM("P&L Model",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild a DETAILED P&L MODEL for ${companyName}.\nCANONICAL EBITDA — key=ebitda, formula = "{S:op.ebitda}" (pulled DIRECTLY from Operational; this is the single source of truth, per CANON above). Do NOT compute EBITDA as net_revenue minus the expense lines below — that makes EBITDA depend on four fragile disclosure rows instead of the one already-correct canonical figure, and if any one of those breaks, EBITDA silently breaks with it.\nDISCLOSURE expense build (informational detail only, does NOT feed EBITDA): cost of materials/COGS, power & fuel, employee cost, other expenses (for EACH of these four: FORECAST formula = "{P:<this row's own key>}/{P:net_revenue}*{R:net_revenue}" — holds that expense's ratio-to-revenue constant at its last historical/prior-year level. Use {P:net_revenue}, NOT {PS:pnl.net_revenue} — net_revenue is a row on THIS SAME sheet, so it's addressed the same-sheet way ({P:}/{R:}), never with a sheet-qualified {PS:sheet.key}/{S:sheet.key} reference to your own sheet. Do NOT invent a new assumption key like "cogs_pct_revenue" for these — none exists in the ASSUMPTION KEYS list above, and referencing one that doesn't exist resolves to a visible #MISSING error, not a usable number. This ratio-preserving formula needs no assumption key at all), total expenditure (key=total_expenditure, formula = "{R:net_revenue}-{R:ebitda}" — a DERIVED memo line computed from the two authoritative figures, NOT a sum of the four disclosure lines above, so it can never disagree with EBITDA even if a disclosure line is off), EBITDA margin (0.0%), depreciation (key=depreciation — the ANNUAL depreciation CHARGE for the year; if the data has both a P&L-style annual figure and a Balance-Sheet-style "Accumulated/Cumulative Depreciation" balance, link the ANNUAL one — never the cumulative one. PLAUSIBILITY CHECK: annual depreciation is typically 3-10% of average gross block per year, and should NOT be a monotonically-increasing-by-similar-amounts series across historical years (that pattern is a signature of an accumulated/cumulative balance, not an annual charge) — if your linked figure fails this check, you linked the WRONG (cumulative) row; either find the correct annual-charge label, or — if only the cumulative balance is available in the data — compute the annual figure yourself as this year's cumulative balance MINUS last year's (do this for every historical year using the raw source numbers, then put the resulting 5 numbers directly in "historical", not a link). This row's FORECAST formula is REQUIRED and must NOT be left blank — a blank forecast is read as 0 by Excel, which silently makes EBIT equal EBITDA and inflates every line below it. Tie it to the gross block build using the EXACT key "depreciation_pct_gross_block" from the ASSUMPTION KEYS list: "{A:depreciation_pct_gross_block}*({S:capex.gross_block}+{PS:capex.gross_block})/2"), EBIT, EBIT margin, other income, finance cost (key=interest — charge it on OPENING (prior-year) net debt ONLY, e.g. "MAX(0,{A:cost_of_debt}*{PS:bs.net_debt})" — do NOT average current-year and prior-year net debt: current-year net debt depends on the cash plug, which depends on retained earnings, which depends on PAT, which depends on this interest figure, so including current-year net debt here creates a genuine circular reference that silently evaluates to 0 across the whole PBT/tax/PAT/EPS chain once it closes the loop. Opening-balance-only breaks the loop by construction, since the prior year's net debt is already fully resolved before this year is computed. Also floor it at zero so it can never go negative even if net debt turns negative/net-cash in later forecast years), PBT, exceptional items, tax, effective tax rate (0.0%), PAT, PAT margin, minority interest where relevant (key it minority_interest and do NOT leave it a flat carry-forward — give it a forecast formula that grows it with earnings, holding it at a constant share of PAT: "{P:minority_interest}/{P:pat}*{R:pat}"), EPS, DPS (link/enter the ACTUAL reported dividend per share for EACH historical year — do NOT leave historical DPS at 0 when the company actually paid a dividend; forecast DPS from the dividend-payout assumption so it stays continuous with the reported history). Pull revenue & EBITDA from operational via {S:op.net_revenue} / {S:op.ebitda}. You MUST include rows keyed exactly: net_revenue, ebitda, depreciation, ebit, interest, tax, pat, eps — and NONE of these forecast formulas may be left blank/omitted.`,
                    STATIC_MODEL),
                callLLM("Balance Sheet",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild a DETAILED BALANCE SHEET & RETURNS sheet for ${companyName}.\nASSETS (use these EXACT keys so the cash plug below can reference them): receivables (key=receivables — HISTORICAL: link/enter from the data; FORECAST formula = "{A:debtor_days}/365*{S:pnl.net_revenue}" — do NOT leave this as historical-only, it must scale with revenue), inventory (key=inventory — HISTORICAL: link/enter from the data; FORECAST formula = "{A:inventory_days}/365*{S:pnl.net_revenue}"), other_current_assets (HISTORICAL: link/enter; FORECAST formula = "{P:other_current_assets}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}" — scales flat with revenue since there's no dedicated days assumption for it), net_fixed_assets (key=net_fixed_assets — HISTORICAL: link/enter the REAL reported Net Fixed Assets figure from the data, exactly like every other asset row above — do NOT leave this historical-derived-from-formula; FORECAST formula = "{S:capex.gross_block}-{R:acc_depreciation}" — do NOT track your own separate gross block, capex.gross_block is the single source of truth, and this formula applies ONLY to the forecast columns, never the historical ones), acc_depreciation (key=acc_depreciation — a STOCK/balance, NOT the same thing as pnl.depreciation which is the annual CHARGE. HISTORICAL: do NOT try to independently look up or link a separate "accumulated depreciation" line from the data — unlike gross block and net block, it is often not cleanly reported as one unambiguous figure, and guessing at it has actually produced a large understatement in a real run (accumulated depreciation understated by over ₹220,000 Cr for a company with a ~₹18 lakh Cr gross block), which overstated the very next forecast year's Net Fixed Assets by the same amount — a phantom multi-lakh-crore asset increase that then forces the short_term_borrowings revolver into an enormous, implausible draw to fund it, even though the revolver formula itself, the capex driver, and the working-capital formulas were all completely correct. Instead, DERIVE each historical year's value directly as (that year's capex.gross_block − that year's net_fixed_assets, both already reliably known from the two rows above) and put these 5 COMPUTED numbers directly in "historical" — do not attempt a "link" for this row. FORECAST formula = "{P:acc_depreciation}+{S:pnl.depreciation}" (this only rolls forward correctly because the historical base is properly derived, not guessed). SELF-CHECK before finalizing: gross_block(last historical year) − acc_depreciation(last historical year) MUST equal net_fixed_assets(last historical year) EXACTLY — if it does not, you have not derived acc_depreciation correctly and every forecast year's Net Fixed Assets will be wrong), cwip, investments, other_non_current_assets (do NOT freeze these flat across the forecast — give EACH a revenue-scaling forecast formula that holds its ratio-to-revenue constant at the last actual level: "{P:<this row's own key>}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}", so the balance sheet grows coherently instead of showing an identical number in every forecast year), and cash (key=cash — historicals from the data as usual, but see BALANCING below for the forecast).\nTotal assets: key=total_assets, formula/sum = "{R:cash}+{R:receivables}+{R:inventory}+{R:other_current_assets}+{R:net_fixed_assets}+{R:cwip}+{R:investments}+{R:other_non_current_assets}".\nLIABILITIES & EQUITY: share_capital (key=share_capital, flat/no formula is acceptable — no fresh issuance is being modeled), reserves (roll forward "{P:equity}+{S:pnl.pat}*{A:pat_retention}"), shareholders' equity (key=equity), long_term_borrowings (key=long_term_borrowings, flat/no formula is acceptable — no repayment/drawdown schedule is being modeled), short_term_borrowings (key=short_term_borrowings — THIS IS THE REVOLVER/FUNDING PLUG, not a flat carry: HISTORICAL — link/enter the real reported figure as usual. FORECAST — use this EXACT dictated formula, copied verbatim (do not shorten or redesign it): "{P:short_term_borrowings}+MAX(0,{A:min_cash_pct_revenue}*{S:pnl.net_revenue}-({R:share_capital}+{R:reserves}+{R:long_term_borrowings}+{P:short_term_borrowings}+{R:payables}+{R:provisions}+{R:deferred_tax_liabilities}+{R:other_non_current_liabilities}+{R:other_current_liabilities}-{R:receivables}-{R:inventory}-{R:other_current_assets}-{R:net_fixed_assets}-{R:cwip}-{R:investments}-{R:other_non_current_assets}))" — this holds short_term_borrowings at last year's level UNLESS funding the rest of the balance sheet at last year's borrowing level would push cash below the minimum buffer, in which case it draws exactly enough extra short-term debt to keep cash at that minimum. This is what prevents the cash plug below from ever going negative — do NOT simplify this formula away or replace it with a flat carry-forward, and do NOT change the cash formula to "fix" a negative value directly; the fix belongs on THIS row, not on cash. If this draw comes out implausibly large (e.g. short-term borrowings jumping to several times their historical level in a single year), that is almost always a symptom of an unrealistic upstream working-capital or capex assumption — most commonly a debtor/inventory/payable-days driver whose FIRST forecast value was not anchored to the company's own last reported balance sheet (see the working-capital-days anchoring requirement in the Assumptions prompt), which manufactures a large artificial one-off funding gap. Do not "fix" a large draw by capping or overriding this formula — the correct fix is upstream, in the days/capex assumptions), payables (key=payables — HISTORICAL: link/enter from the data; FORECAST formula = "{A:payable_days}/365*{S:pnl.net_revenue}" — do NOT leave this as historical-only, it must scale with revenue), provisions (HISTORICAL: link/enter; FORECAST formula = "{P:provisions}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}"), deferred_tax_liabilities, other_non_current_liabilities, other_current_liabilities (do NOT freeze these flat — give EACH the same revenue-scaling forecast formula "{P:<this row's own key>}/{PS:pnl.net_revenue}*{S:pnl.net_revenue}" so they scale with the business rather than repeating an identical number every forecast year).\nNONE of receivables/inventory/other_current_assets/payables/provisions may be left without a forecast formula — a flat/frozen working-capital line while revenue grows is a modeling error, not a simplification.\nTotal liabilities & equity: key=total_liabilities_equity, formula/sum of every liabilities & equity row above (via {R:} references, mirroring total_assets).\nPlus: net working capital (key=net_working_capital = "{R:receivables}+{R:inventory}+{R:other_current_assets}-{R:payables}-{R:provisions}" — the Capex & FCF sheet's change-in-WC pulls this), net debt (key=net_debt, formula = "{R:long_term_borrowings}+{R:short_term_borrowings}-{R:cash}" — do NOT leave this without a formula, it must move with the computed cash plug, not stay flat), capital employed (key=capital_employed), working-capital days, and returns: ROE (key=roe, formula = "{S:pnl.pat}/{R:equity}", fmt "0.0%" — this MUST be a "formula", never a plain "forecast" array of guessed numbers; a bare zero here has actually happened when the formula was omitted), ROCE (key=roce, formula = "{S:pnl.ebit}/{R:capital_employed}", fmt "0.0%" — same requirement, must be a formula), net debt/EBITDA (key=net_debt_ebitda, formula = "{R:net_debt}/{S:pnl.ebitda}", fmt "0.00").\nHISTORICAL BALANCE (the 5 ACTUAL-year columns) — a REPORTED balance sheet ALWAYS balances, so your historical columns MUST balance too, not only the forecast columns: for EVERY historical year the summed asset rows must equal the summed liabilities-&-equity rows (balance_check ≈ 0 in the actual years as well). Historical cash is NOT a plug — take it from the data. To make the actuals tie out: (a) transcribe EVERY reported line from the downloaded Balance Sheet; (b) NEVER double-count — cash, current investments and non-current investments are DISTINCT lines (include each exactly once) and reserves are counted once inside equity, not again as a separate asset/liability; (c) cross-check your historical total_assets against the reported "Total Assets" figure in the downloaded data — if it differs by more than ~1% you have mislinked or double-counted a line, so find and fix it before returning (a real prior run overstated total assets by ~18%); (d) reconcile any small residual through the low-materiality catch-all rows (other_current_assets / other_non_current_assets on the asset side, other_current_liabilities / other_non_current_liabilities on the liabilities side) so the historical balance_check lands at ~0; (e) EVERY historical year must reflect that SPECIFIC year's own reported figures — do NOT copy or repeat the prior year's balance-sheet numbers into the next year because a distinct source figure was hard to find (this has actually happened: the latest two historical years showed IDENTICAL balance-sheet values even though the P&L for those same years was correctly distinct). Before returning, spot-check total_assets, cash, and equity across adjacent historical years — if any two adjacent years show the SAME value for several rows at once, you have accidentally duplicated one year's column into another; go back to the source data and locate that specific year's real figures instead. Historical net_debt (borrowings − cash) must also match the company's REPORTED net debt — if it comes out roughly double the reported figure you have linked the wrong borrowings or cash line; fix it.\nBALANCING — cash's FORECAST formula (columns G-K only; historicals still come from the data) MUST be the plug: "{R:total_liabilities_equity}-{R:receivables}-{R:inventory}-{R:other_current_assets}-{R:net_fixed_assets}-{R:cwip}-{R:investments}-{R:other_non_current_assets}" (total liabilities & equity minus every OTHER asset row) — this makes total assets equal total liabilities & equity by construction in every forecast year. Do NOT drive forecast cash off a %-of-sales ratio; that breaks the balance. Because short_term_borrowings above already includes the revolver top-up whenever needed, this cash plug should NEVER compute to a negative number in any forecast year — a negative result here means the short_term_borrowings revolver formula was not applied exactly as dictated; re-check that row before returning, since a company literally cannot hold negative cash. Add a final check row "Balance Check (Assets - Liab.&Equity, should be ~0)" key=balance_check, formula "{R:total_assets}-{R:total_liabilities_equity}".\nYou MUST include rows keyed exactly: equity, net_debt, net_debt_ebitda, capital_employed, net_working_capital, total_assets, total_liabilities_equity, cash, receivables, inventory, other_current_assets, net_fixed_assets, acc_depreciation, cwip, investments, other_non_current_assets, payables, provisions, share_capital, long_term_borrowings, short_term_borrowings, roe, roce.`,
                    STATIC_MODEL),
                callLLM("Capex & FCF",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${CANON}\n\n${rowSchemaDoc}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild the CAPEX & FCF sheet for ${companyName}: total capex (key=capex — HISTORICAL: link/enter the real reported figure from the data as usual; FORECAST formula = "{A:capex_pct_revenue}*{S:pnl.net_revenue}" — use this EXACT formula, referencing the capex_pct_revenue assumption; do NOT reference a "capex" row on the Assumptions sheet, there isn't one), capex/revenue, capex per unit, operating cash flow (key=ocf — do NOT give this row "historical" or "link"; give ONLY formula="{S:pnl.ebitda}-{S:pnl.interest}-{S:pnl.tax}" for EVERY column, historical and forecast alike — this is a LEVERED cash flow measure since it nets out interest; do NOT substitute PAT or any other guessed historical proxy), change in net working capital (key=wc_change — do NOT give this row "historical" or "link"; give ONLY formula="{S:bs.net_working_capital}-{PS:bs.net_working_capital}" for EVERY column — pull this from the Balance Sheet rather than re-deriving or guessing it, so it can never disagree with the DCF sheet which pulls the same figure), free cash flow (key=fcf, formula "{R:ocf}-{R:capex}" — label it "Free Cash Flow (levered)" and note it is NOT expected to equal the DCF sheet's FCFF, which is an unlevered measure computed before interest), cumulative FCF, gross block (key=gross_block — HISTORICAL: link/enter from the data, but sanity-check it rolls forward (Gross Block(t) ≈ Gross Block(t-1)+Capex(t)) for the latest historical year per the ROLL-FORWARD CHECK above; FORECAST formula = "{P:gross_block}+{R:capex}" — this is the ONLY place gross block is computed; the Balance Sheet references it rather than tracking its own copy), depreciation (key=depreciation — do NOT give this row "historical" or "link"; give ONLY formula="{S:pnl.depreciation}" so every column, historical and forecast alike, pulls the exact same figure pnl.depreciation already resolved — never link or derive this independently on this sheet). You MUST include rows keyed exactly: capex, ocf, wc_change, fcf, gross_block, depreciation.`,
                    STATIC_MODEL),
                callLLM("Valuation & DCF",
                    `${periodsNote}\n\n${GRAMMAR}\n\n${VAL_GRAMMAR}\n\n${CANON}\n\nASSUMPTION KEYS:\n${assumKeyList}\n\n${finBlock}\nBuild a detailed VALUATION & DCF sheet for ${companyName}. Return JSON {"rows":[...]} with these sections (use {"section":"NAME"} dividers):\n(1) "RELATIVE VALUATION" — per-year live ratios: EPS key=eps ("{S:pnl.eps}"), BVPS key=bvps ("{S:bs.equity}/{A:shares_out}"), PER ("{A:current_price}/{R:eps}", fmt 0.0), P/BV ("{A:current_price}/{R:bvps}", 0.0), EV/EBITDA ("(({A:current_price}*{A:shares_out})+{S:bs.net_debt})/{S:pnl.ebitda}", 0.0), RoE ("{S:pnl.pat}/{S:bs.equity}", 0.0%), RoCE ("{S:pnl.ebit}/{S:bs.capital_employed}", 0.0%), Net debt/EBITDA ("{S:bs.net_debt}/{S:pnl.ebitda}", 0.00).\n(2) "DCF — FREE CASH FLOW TO FIRM" (every row "forecast_only":true): EBIT key=ebit ("{S:pnl.ebit}"), NOPAT key=nopat ("{R:ebit}*(1-{A:tax_rate})"), Add depreciation key=dep ("{S:pnl.depreciation}"), Less change in working capital key=wc_change ("{S:capex.wc_change}" — pull the SAME figure the Capex & FCF sheet computes; do NOT leave this without a formula or default it to zero), FCFF key=fcff ("{R:nopat}+{R:dep}-{R:wc_change}-{S:capex.capex}"), Discount factor key=discount_factor ("1/(1+{A:wacc})^{N}", fmt 0.000), PV of FCFF key=pv_fcff ("{R:fcff}*{R:discount_factor}").\n(3) "DCF VALUATION" — SCALAR rows: Sum of PV key=sum_pv ("{SUM:pv_fcff}"), Terminal value key=tv — use this EXACT dictated formula, copied verbatim (do not redesign it): "MAX(0,({LAST:nopat}-{A:terminal_growth}*{LAST:bs.net_working_capital})*(1+{A:terminal_growth})/({A:wacc}-{A:terminal_growth}))". Do NOT build this off {LAST:fcff} (the raw final explicit-year FCFF) — for a still-investing, capex-heavy company the LAST explicit forecast year's FCFF can itself be depressed or even negative (heavy capex, a working-capital step-up, etc.), and extrapolating THAT into a perpetuity produces a nonsensical negative terminal value (this has actually happened — a real run's terminal FCFF was negative, making equity value collapse to a large negative per-share figure). The dictated formula instead uses a NORMALIZED steady-state terminal cash flow: terminal NOPAT (which excludes capex/working-capital entirely, so it stays sensible even in a buildout year) less a working-capital investment that grows at the terminal growth rate (g × the terminal year's own net working capital LEVEL, not the raw year-over-year change row, which is more robust to any residual basis discontinuity in an early forecast year) — implicitly assuming steady-state maintenance capex equals depreciation, the standard textbook simplification — wrapped in MAX(0,...) so a structurally unprofitable terminal NOPAT can never produce a negative terminal value. PV of terminal value key=pv_tv ("{R:tv}/(1+{A:wacc})^5"), Enterprise value key=ev ("{R:sum_pv}+{R:pv_tv}"), Less net debt key=net_debt_last ("{LAST:bs.net_debt}"), Equity value key=equity_value ("{R:ev}-{R:net_debt_last}"), Value per share key=value_per_share ("{R:equity_value}/{A:shares_out}", emphasis:"highlight"), Current price ("{A:current_price}"), DCF upside key=upside ("{R:value_per_share}/{A:current_price}-1", fmt 0.0%).\n(4) "TARGET PRICE (relative, NEAR-TERM FY+1 basis)" — SCALAR rows, built off the FIRST forecast year via {FWD1:} (NOT {LAST:}, which is 5 years out and undiscounted — using it overstates the target price versus the DCF): P/E-based ("{FWD1:pnl.eps}*{A:target_pe}"), EV/EBITDA-based ("(({FWD1:pnl.ebitda}*{A:target_ev_ebitda})-{FWD1:bs.net_debt})/{A:shares_out}"), Blended target price key=target_price (average of the two, emphasis:"highlight"), Upside to target ("{R:target_price}/{A:current_price}-1", fmt 0.0%).\n(5) "WACC BUILD-UP" — SCALAR rows: risk free ("{A:risk_free_rate}", 0.0%), equity risk premium ("{A:equity_risk_premium}", 0.0%), beta ("{A:beta}", 0.00), cost of equity ("{A:risk_free_rate}+{A:beta}*{A:equity_risk_premium}", 0.0%), cost of debt ("{A:cost_of_debt}", 0.0%), tax rate ("{A:tax_rate}", 0.0%), WACC ("{A:wacc}", 0.0%).\nYou MUST emit scalar keys: value_per_share, target_price, upside (plus helpers sum_pv, pv_tv, tv, ev, net_debt_last, equity_value). Use fmt "#,##0" for currency rows. Be thorough.`,
                    STATIC_MODEL),
                callLLM("Summary Dashboard",
                    `Build a SUMMARY DASHBOARD for ${companyName} that links the most decision-relevant lines from the other sheets, referenced ONLY by these canonical keys:\nop.net_revenue, op.ebitda, op.volume\npnl.net_revenue, pnl.ebitda, pnl.depreciation, pnl.ebit, pnl.interest, pnl.tax, pnl.pat, pnl.eps\nbs.equity, bs.net_debt, bs.capital_employed\ncapex.capex, capex.ocf, capex.fcf, capex.gross_block\nReturn JSON: {"rows":[ {"section":"SECTION NAME"} OR {"label":"Display label","ref":"sheet.canonical_key","fmt":"#,##0"|"0.00","cagr":true|false} ]}. Group into OPERATIONAL, P&L, BALANCE SHEET, CAPEX & FCF. Pick ~15-22 lines. EVERY ref points to an ABSOLUTE canonical value (rupee levels, EPS, or share/volume counts) — there are NO margin/ratio/percentage canonical keys available here, so do NOT add EBITDA-margin, ROE, ROCE, or any other "%" row: applying a percentage format to an absolute-rupee figure renders a nonsensical multi-million-percent number (a real failure produced "11046000%"). Use ONLY fmt "#,##0" for rupee amounts or "0.00" for EPS / x-multiples — NEVER a "0.0%" percentage format on this sheet.`,
                    STATIC_MODEL)
            ]);

            const opRows = Array.isArray(opJson.rows) ? opJson.rows : [];
            const pnlRows = Array.isArray(pnlJson.rows) ? pnlJson.rows : [];
            const bsRows = Array.isArray(bsJson.rows) ? bsJson.rows : [];
            const capexRows = Array.isArray(capexJson.rows) ? capexJson.rows : [];
            const valRows = Array.isArray(valJson.rows) ? valJson.rows : [];
            const summaryRows = Array.isArray(sumJson.rows) ? sumJson.rows : [];
            planRows("op", opRows);
            planRows("pnl", pnlRows);
            planRows("bs", bsRows);
            planRows("capex", capexRows);
            planRows("val", valRows);

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
            };
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
                const derived = [0, 1, 2, 3, 4].map(i => {
                    const gb = Number(gbItem.historical[i]), nfa = Number(nfaItem.historical[i]);
                    return (isFinite(gb) && isFinite(nfa)) ? gb - nfa : null;
                });
                if (derived.some(v => v === null)) {
                    console.warn("[AI Model] bs.net_fixed_assets or capex.gross_block historicals were incomplete — could not derive all 5 years of bs.acc_depreciation; verify Net Fixed Assets ties out at the FY forecast boundary manually.");
                    return;
                }
                adItem.historical = derived;
                delete adItem.link; // a derived figure, not a linkable data-sheet cell
                console.log(`[AI Model] Derived bs.acc_depreciation historicals as gross_block − net_fixed_assets (guarantees Net Fixed Assets ties out at the forecast boundary):`, derived.map(v => Math.round(v)));
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
                { re: /capex|capital_work_in_progress|^cwip$/i, need: "capex_pct_revenue", formula: "{A:capex_pct_revenue}*{S:pnl.net_revenue}" },
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
            const REF_RE = /\{(SUM|LAST|FWD1|PS|S|A|R|P):([^}]+)\}/g;
            const refResolves = (sheetKey, kind, body) => {
                if (kind === "A") return !!symbols.assum[body];
                if (kind === "R" || kind === "P") return !!(symbols[sheetKey] && symbols[sheetKey][body]);
                if (kind === "SUM" || kind === "LAST" || kind === "FWD1") {
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

            aiStatus("Writing model to Excel...");

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

                // ── Formula placeholder resolver ({A}/{R}/{P}/{S}/{PS}) -> real Excel addresses ──
                const colAt = (i) => ALL_COLS[i];
                const prevColAt = (i) => ALL_COLS[i - 1];
                const refMissing = [];
                const missingForecast = [];
                let curSheetKey = null;
                let curRowKey = null; // set per-row in writeDataSheet, so a bad reference logs WHERE it was written, not just the raw template
                const resolve = (template, ci) => {
                    const col = colAt(ci), pcol = prevColAt(ci) || col;
                    const nIdx = ci >= 5 ? (ci - 4) : (ci + 1); // {N} = forecast year number (1..5)
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
                        .replace(/\{(SUM|LAST|FWD1|PS|S|A|R|P):([^}]+)\}/g, (m, kind, body) => {
                            if (kind === "A") {
                                const row = symbols.assum[body];
                                if (!row) return markBad(`A:${body}`);
                                // Assumptions is a multi-year schedule — read the SAME year column.
                                return `'${SHEETS.assum}'!${col}${row}`;
                            }
                            if (kind === "R" || kind === "P") {
                                const row = symbols[curSheetKey] && symbols[curSheetKey][body];
                                if (!row) return markBad(`${kind}:${body}`);
                                return `${kind === "P" ? pcol : col}${row}`;
                            }
                            if (kind === "SUM" || kind === "LAST" || kind === "FWD1") {
                                const a = aggRowOf(body);
                                if (!a) return markBad(`${kind}:${body}`);
                                if (kind === "SUM") return `SUM(${a.pfx}${FC_COLS[0]}${a.row}:${FC_COLS[4]}${a.row})`;
                                // LAST = final forecast year; FWD1 = first (near-term/NTM) forecast year.
                                return `${a.pfx}${kind === "FWD1" ? FC_COLS[0] : FC_COLS[4]}${a.row}`;
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
                    const rng = sh.getRange("A1:L1");
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
                        sh.getRange(`${ALL_COLS[i]}3`).format.fill.color = i < 5 ? "#e8edf3" : "#d0d9f0";
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
                        sh.getRange(`A${rr}`).values = [[item.label || item.key || ""]];
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
                        // Use the AI's decimals / a derived formula for those instead.
                        const linkEntry = /%/.test(fmt) ? null : autoLink(item);
                        if (linkEntry) {
                            const histNums = Array.isArray(item.historical) ? item.historical : [];
                            for (let i = 0; i < 5; i++) {
                                const col = linkEntry.colByFY[HIST[i]];
                                const cell = sh.getRange(`${HIST_COLS[i]}${rr}`);
                                if (col) {
                                    cell.formulas = [[`='${linkEntry.sheet}'!${col}${linkEntry.row}`]];
                                } else if (histNums[i] !== null && histNums[i] !== undefined && histNums[i] !== "") {
                                    cell.values = [[histNums[i]]];
                                }
                            }
                        } else if (Array.isArray(item.historical)) {
                            const vals = item.historical.slice(0, 5);
                            while (vals.length < 5) vals.push("");
                            sh.getRange(`B${rr}:F${rr}`).values = [vals.map(v => (v === null || v === undefined) ? "" : v)];
                            if (opts.blue) sh.getRange(`B${rr}:F${rr}`).format.font.color = "#1155CC";
                        } else if (item.formula && !item.forecast_only) {
                            // derived row (e.g. a margin) — compute historicals too
                            for (let i = 0; i < 5; i++) sh.getRange(`${HIST_COLS[i]}${rr}`).formulas = [[resolve(item.formula, i)]];
                        }

                        // Forecast columns (G..K): recurrence formula.
                        if (item.formula) {
                            for (let i = 0; i < 5; i++) sh.getRange(`${FC_COLS[i]}${rr}`).formulas = [[resolve(item.formula, 5 + i)]];
                        } else if (Array.isArray(item.forecast)) {
                            const vals = item.forecast.slice(0, 5);
                            while (vals.length < 5) vals.push("");
                            sh.getRange(`G${rr}:K${rr}`).values = [vals.map(v => (v === null || v === undefined) ? "" : v)];
                            if (opts.blue) sh.getRange(`G${rr}:K${rr}`).format.font.color = "#1155CC";
                        } else if (!item.forecast_only && (linkEntry || Array.isArray(item.historical))) {
                            // The model gave this row real historicals but NEITHER a forecast "formula"
                            // NOR "forecast" values — left as-is, Excel reads the blank forecast cells as
                            // 0, which silently zeroes out everything downstream (e.g. an EBITDA margin
                            // row defaulting to 0% makes segment EBITDA compute to 0 even though revenue
                            // is fine, or depreciation defaulting to 0 makes EBIT = EBITDA). Flat-carry-
                            // forward the last historical value instead of a silent zero, and flag it —
                            // this is a real gap that should get a real forecast, not just a duplicate.
                            for (let i = 0; i < 5; i++) sh.getRange(`${FC_COLS[i]}${rr}`).formulas = [[`=IFERROR(${i === 0 ? "F" : FC_COLS[i - 1]}${rr},0)`]];
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

                        sh.getRange(`B${rr}:K${rr}`).numberFormat = [new Array(10).fill(normalizeFmt(fmt))];
                        if (item.cagr) {
                            sh.getRange(`${CAGR_COL}${rr}`).formulas = [[`=IFERROR((K${rr}/F${rr})^(1/5)-1,"")`]];
                            sh.getRange(`${CAGR_COL}${rr}`).numberFormat = [[normalizeFmt("0.0%")]];
                        }
                    }
                };

                // ── Assumptions: multi-year driver schedule (blue = input you can flex) ──
                writeDataSheet("assum", assumRows, { blue: true, unitLabel: "in row's stated unit" });
                made.assum.getRange("A2").values = [["Blue = hardcoded input you can flex · Black = formula / linked actual · each driver feeds the other sheets via the same year column"]];
                made.assum.getRange("A2").format.font.size = 9;
                made.assum.getRange("A2").format.font.color = "#777777";

                // Source citations on hardcoded assumption inputs (beta, WACC components, multiples,
                // etc.), as Excel cell comments on the forecast input cell. Best-effort: the Comments
                // API needs a reasonably current Excel client, and a source is only as good as the
                // model's own citation — failures here must never abort the rest of the model build.
                try {
                    let sourced = 0;
                    for (const item of assumRows) {
                        if (item && item.key && item.source && item._row) {
                            wb.comments.add(made.assum.getRange(`G${item._row}`), String(item.source).slice(0, 300));
                            sourced++;
                        }
                    }
                    await context.sync();
                    if (sourced) console.log(`[AI Model] Added ${sourced} source citation(s) to the Assumptions sheet.`);
                } catch (e) {
                    console.warn("[AI Model] Could not attach source comments (Comments API unavailable on this Excel client?):", e.message);
                }

                writeDataSheet("op", opRows);
                writeDataSheet("pnl", pnlRows);
                writeDataSheet("bs", bsRows);
                writeDataSheet("capex", capexRows);
                writeDataSheet("val", valRows);

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
                            sh.getRange(`B${rr}:K${rr}`).numberFormat = [new Array(10).fill(normalizeFmt(item.fmt || "#,##0"))];
                            // Same canonical-key override as the source sheets: force CAGR for
                            // absolute-value metrics regardless of the model's own "cagr" flag.
                            if (item.cagr || (CAGR_KEYS[sk] || []).includes(rk)) {
                                sh.getRange(`${CAGR_COL}${rr}`).formulas = [[`=IFERROR((K${rr}/F${rr})^(1/5)-1,"")`]];
                                sh.getRange(`${CAGR_COL}${rr}`).numberFormat = [[normalizeFmt("0.0%")]];
                            }
                        }
                        if (rr % 2 === 0) sh.getRange(`A${rr}:${CAGR_COL}${rr}`).format.fill.color = "#f9f9f9";
                        rr++;
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

                // Verify the balance-sheet revolver actually held: forecast cash should never compute
                // negative, and the balance check should stay near 0 (including historicals). Read back
                // the ACTUAL Excel-computed post-recalc values rather than trusting the formula text.
                const cashRow = symbols.bs && symbols.bs.cash;
                const balRow = symbols.bs && symbols.bs.balance_check;
                const taRow = symbols.bs && symbols.bs.total_assets;
                let cashRange, balRange, taRange;
                if (cashRow) { cashRange = made.bs.getRange(`B${cashRow}:K${cashRow}`); cashRange.load("values"); }
                if (balRow) { balRange = made.bs.getRange(`B${balRow}:K${balRow}`); balRange.load("values"); }
                if (taRow) { taRange = made.bs.getRange(`B${taRow}:K${taRow}`); taRange.load("values"); }
                if (cashRow || balRow || taRow) await context.sync();

                if (cashRange) {
                    const cashVals = cashRange.values[0];
                    const negatives = cashVals.map((v, i) => ({ v, i })).filter(x => typeof x.v === "number" && x.v < 0);
                    if (negatives.length) {
                        console.warn(`[AI Model] Balance Sheet forecast CASH is NEGATIVE in ${negatives.length} column(s) — (${negatives.map(x => `${periodLabels[x.i]}: ${Math.round(x.v).toLocaleString()}`).join(", ")}) — a company cannot hold negative cash; the short_term_borrowings revolver formula should prevent this. Check that row's formula on the Balance Sheet sheet.`);
                    }
                }
                if (balRange && taRange) {
                    const balVals = balRange.values[0];
                    const taVals = taRange.values[0];
                    const offCols = [];
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
            });

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
                // Key financials
                const finResp = await fetch("https://transcriptanalyser.com/goindiastock/actuals_forwards", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ fincode: comp.fincode, mode: "C", sector_type: comp.sector_type })
                });
                const finData = await finResp.json();
                const finRows = finData?.data || [];
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

        // Detects an empty/all-zero financial-statement response — used both to decide whether a
        // company has consolidated data at all (the probe below) and, later, to render "not
        // available" placeholders for sections that DID turn out empty despite being fetched.
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

        try {
            // Key Financials is mode-independent — fetch once
            const keyfResp = await fetch("https://transcriptanalyser.com/goindiastock/actuals_forwards", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ fincode, mode: "C", sector_type: sectorType })
            });
            const keyfData = await keyfResp.json();

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
                    const getHeaders = (data) => {
                        if (!data?.length) return [];
                        const headers = new Set();
                        data.forEach(row => {
                            Object.keys(row).forEach(k => {
                                if (!staticFields.includes(k) && /^[A-Z][a-z]{2}\d{4}$/.test(k)) headers.add(k);
                            });
                        });
                        return Array.from(headers).sort((a, b) => {
                            const months = { Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5, Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11 };
                            const parseDate = s => {
                                const [_, mon, year] = s.match(/^([A-Za-z]+)(\d{4})$/) || [];
                                return new Date(parseInt(year), months[mon] ?? 0);
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
const CHAT_ENABLED = false; // TEMPORARY: hidden again pre-deployment — priority is the wallet-feature update; re-enable by flipping back to true

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
    const AUTO_APPLY_KEY = "goia_chatAutoApplyEdits";
    let autoApplyEdits = localStorage.getItem(AUTO_APPLY_KEY) === "1";
    const autoApplyToggle = byId("chatAutoApply"), autoApplyLabel = byId("chatAutoApplyLabel");
    const renderAutoApplyLabel = () => {
        if (autoApplyLabel) autoApplyLabel.textContent = autoApplyEdits ? "Writing to sheets automatically" : "Ask before writing to sheets";
    };
    if (autoApplyToggle) {
        autoApplyToggle.checked = autoApplyEdits;
        renderAutoApplyLabel();
        autoApplyToggle.addEventListener("change", () => {
            autoApplyEdits = autoApplyToggle.checked;
            try { localStorage.setItem(AUTO_APPLY_KEY, autoApplyEdits ? "1" : "0"); } catch (e) { /* storage unavailable */ }
            renderAutoApplyLabel();
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
    const addMsg = (role, text) => {
        const m = document.createElement("div");
        m.className = "chat-msg " + role;
        const b = document.createElement("div");
        b.className = "chat-bubble";
        b.textContent = text;
        m.appendChild(b);
        stream.appendChild(m);
        stream.scrollTop = stream.scrollHeight;
        return b;
    };
    const addCard = (text, cls) => {
        const c = document.createElement("div");
        c.className = "chat-card " + (cls || "run");
        const t = document.createElement("span"); t.textContent = text;
        const s = document.createElement("span"); s.className = "chat-card-st";
        c.appendChild(t); c.appendChild(s);
        stream.appendChild(c);
        stream.scrollTop = stream.scrollHeight;
        return c;
    };
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

    // Gate for every workbook write: asks first unless the user has flipped on
    // "write automatically". Warns explicitly when a write would overwrite an existing sheet.
    async function maybeConfirmAndWrite(action) {
        const sheetName = String(action?.sheet || "Claude Output").slice(0, 31);
        if (!autoApplyEdits) {
            let exists = false;
            try { exists = (await toolListSheets()).sheets.some(s => s.toLowerCase() === sheetName.toLowerCase()); }
            catch (e) { /* assume new sheet */ }
            const verb = exists ? `overwrite the existing sheet "${sheetName}"` : `create a new sheet "${sheetName}"`;
            const rows = Array.isArray(action?.rows) ? action.rows.length : 0;
            const approved = await addConfirmCard(`DataGPT wants to ${verb}${rows ? ` (${rows} row${rows === 1 ? "" : "s"})` : ""}.`);
            if (!approved) return { ok: false, rejected: true, sheet: sheetName };
        }
        await writeAction(action);
        return { ok: true, sheet: sheetName };
    }

    const greet = () => addMsg("bot",
        "Hi — I'm connected to your GoIndia data and this workbook. Ask me about the selected company, or to write something into a sheet. (Experimental)");

    // Re-render whatever chat history survived from the last session.
    history.forEach(m => addMsg(m.role === "assistant" ? "bot" : "user", m.content));

    const openPanel = () => { panel.style.display = "flex"; fab.style.display = "none"; setCompany(); if (!stream.childElementCount) greet(); input.focus(); };
    const closePanel = () => { panel.style.display = "none"; fab.style.display = ""; };
    fab.addEventListener("click", openPanel);
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
    document.querySelectorAll("#chatFeature .chat-chip").forEach(ch =>
        ch.addEventListener("click", () => { input.value = ch.textContent; input.focus(); }));
    input.addEventListener("input", () => { input.style.height = "auto"; input.style.height = Math.min(96, input.scrollHeight) + "px"; });
    input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } });
    sendBtn.addEventListener("click", send);
    const ddToggle = byId("dropdownToggle");
    if (ddToggle) ddToggle.addEventListener("blur", setCompany);

    // ── GoIndia MCP access ──
    const MCP_SERVER_URL = "https://goindia-mcp.fly.dev/sse"; // the actual MCP server — see McpClient below
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
    const TEXT_CAP = 3500;
    const capText = (s, n = TEXT_CAP) => {
        s = (s || "").trim();
        return s.length > n ? s.slice(0, n) + "\n…[truncated]" : s;
    };

    const EXCEL_TOOLS = [
        { type: "function", function: { name: "list_excel_sheets", description: "List the worksheet names currently in the user's open Excel workbook.", parameters: { type: "object", properties: {} } } },
        { type: "function", function: { name: "read_excel_sheet", description: "Read a worksheet's actual used-range layout (row numbers, column letters and cell text) from the user's open workbook. ALWAYS call this before writing a formula (e.g. a VLOOKUP) that references an existing sheet — never guess row/column positions.", parameters: { type: "object", properties: { sheet_name: { type: "string" } }, required: ["sheet_name"] } } },
        { type: "function", function: { name: "write_excel_sheet", description: "Create or overwrite a worksheet in the user's workbook with a title, headers and rows. Cells may be plain values or formula strings starting with \"=\" (e.g. a VLOOKUP into a sheet you've already read).", parameters: { type: "object", properties: { sheet: { type: "string" }, title: { type: "string" }, headers: { type: "array", items: { type: "string" } }, rows: { type: "array", items: { type: "array" } } }, required: ["sheet"] } } }
    ];

    async function toolListSheets() {
        let names = [];
        await Excel.run(async (ctx) => {
            const wb = ctx.workbook;
            wb.worksheets.load("items/name");
            await ctx.sync();
            names = wb.worksheets.items.map(s => s.name);
        });
        return { sheets: names };
    }
    async function toolReadSheet({ sheet_name }) {
        const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
        let result = { error: `No sheet named "${sheet_name}" in the workbook.` };
        await Excel.run(async (ctx) => {
            const wb = ctx.workbook;
            wb.worksheets.load("items/name");
            await ctx.sync();
            const match = wb.worksheets.items.find(s => s.name.toLowerCase() === String(sheet_name || "").toLowerCase());
            if (!match) return;
            const sh = wb.worksheets.getItem(match.name);
            const used = sh.getUsedRangeOrNullObject();
            used.load(["text", "rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
            await ctx.sync();
            if (used.isNullObject) { result = { sheet: match.name, empty: true }; return; }
            const maxRows = Math.min(used.rowCount, 60);
            const maxCols = Math.min(used.columnCount, 20);
            const lines = [];
            for (let r = 0; r < maxRows; r++) {
                const rowVals = used.text[r].slice(0, maxCols);
                if (rowVals.every(v => String(v).trim() === "")) continue;
                const cells = rowVals.map((v, i) => (String(v).trim() !== "" ? `${colLetter(used.columnIndex + i + 1)}=${v}` : null)).filter(Boolean);
                lines.push(`Row ${used.rowIndex + r + 1}: ` + cells.join(" | "));
            }
            result = { sheet: match.name, address: used.address, totalRows: used.rowCount, totalCols: used.columnCount, preview: capText(lines.join("\n"), 3500) };
        });
        return result;
    }

    function toolLabel(name, args) {
        switch (name) {
            case "list_excel_sheets": return "Listing workbook sheets";
            case "read_excel_sheet": return `Reading sheet "${args.sheet_name}"`;
            default: return name;
        }
    }

    async function callTool(name, args) {
        switch (name) {
            case "list_excel_sheets": return toolListSheets();
            case "read_excel_sheet": return toolReadSheet(args);
            default: return { error: `Unknown tool "${name}"` };
        }
    }

    const padRow = (arr, n) => { const a = (arr || []).slice(0, n).map(v => (v == null ? "" : v)); while (a.length < n) a.push(""); return a; };

    // Write a simple table (the only Excel action the chat performs for now).
    async function writeAction(action) {
        const card = addCard(`Writing sheet "${action.sheet || "Claude Output"}"…`, "run");
        const st = card.querySelector(".chat-card-st");
        try {
            await Excel.run(async (ctx) => {
                const wb = ctx.workbook;
                wb.worksheets.load("items/name");
                await ctx.sync();
                const names = wb.worksheets.items.map(s => s.name);
                const sheetName = String(action.sheet || "Claude Output").slice(0, 31);
                let sh;
                if (names.includes(sheetName)) { sh = wb.worksheets.getItem(sheetName); sh.getUsedRange()?.clear(); }
                else sh = wb.worksheets.add(sheetName);
                sh.tabColor = "#D97757";
                const colL = (n) => { let s = ""; while (n > 0) { const mm = (n - 1) % 26; s = String.fromCharCode(65 + mm) + s; n = (n - mm - 1) / 26; } return s; };
                const headers = Array.isArray(action.headers) ? action.headers : [];
                const rows = Array.isArray(action.rows) ? action.rows : [];
                const ncol = Math.max(headers.length, ...rows.map(r => (r || []).length), 1);
                const last = colL(ncol);
                let r = 1;
                if (action.title) {
                    sh.getRange("A1").values = [[String(action.title)]];
                    sh.getRange("A1").format.font.bold = true;
                    sh.getRange("A1").format.font.size = 13;
                    sh.getRange("A1").format.fill.color = "#fdf3ee";
                    r = 3;
                }
                if (headers.length) {
                    const hr = sh.getRange(`A${r}:${last}${r}`);
                    hr.values = [padRow(headers, ncol)];
                    hr.format.font.bold = true; hr.format.fill.color = "#e8edf3";
                    r++;
                }
                if (rows.length) {
                    sh.getRange(`A${r}:${last}${r + rows.length - 1}`).values = rows.map(x => padRow(x, ncol));
                }
                sh.getRange("A:A").format.columnWidth = 180;
                sh.activate();
                await ctx.sync();
            });
            card.className = "chat-card done"; st.textContent = "✓";
        } catch (e) {
            console.error("[Chat] write failed", e);
            card.className = "chat-card err"; st.textContent = "✕";
        }
    }

    async function send() {
        if (busy) return;
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
                `- Before writing any formula that references an existing worksheet (e.g. a VLOOKUP into "Key Financials"), call read_excel_sheet on it first — never guess row/column positions. Use list_excel_sheets if you're unsure what sheets exist.\n` +
                `- Use write_excel_sheet to create or update sheets. Prefer live formulas (VLOOKUP/MATCH etc.) that reference a sheet you've already read; otherwise write the plain values you fetched.\n` +
                `- Every write is shown to the user for approval before it happens unless they've turned on auto-apply. If a write_excel_sheet result comes back with rejected: true, tell the user the write was skipped rather than assuming it happened — do not silently retry it.\n` +
                `- Compute derived figures yourself (YoY growth, CAGR, margins, ratios, averages) from the numbers you fetch.\n` +
                `- Be concise in your final answer.`;

            const combinedTools = EXCEL_TOOLS.concat(mcpOpenAITools);
            const convo = [{ role: "system", content: system }, ...history.slice(-8)];
            let full = "";
            for (let iter = 0; iter < 6; iter++) {
                const res = await fetch(OPENROUTER_URL, {
                    method: "POST",
                    headers: { "Content-Type": "application/json", "Authorization": `Bearer ${OPENROUTER_KEY}` },
                    body: JSON.stringify({ model: AI_MODEL, messages: convo, tools: combinedTools, temperature: 0.2, max_tokens: 3000, usage: { include: true } })
                });
                if (!res.ok) throw new Error(`OpenRouter ${res.status}: ${(await res.text()).slice(0, 160)}`);
                const data = await res.json();
                const cost = data.usage?.cost;
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
                    if (call.function.name === "write_excel_sheet") {
                        try { result = await maybeConfirmAndWrite(args); } // renders its own confirm/write card
                        catch (e) { result = { error: String(e.message || e) }; }
                    } else if (call.function.name === "list_excel_sheets" || call.function.name === "read_excel_sheet") {
                        const card = addCard(`${toolLabel(call.function.name, args)}…`, "run");
                        const st = card.querySelector(".chat-card-st");
                        try { result = await callTool(call.function.name, args); card.className = "chat-card done"; st.textContent = "✓"; }
                        catch (e) { result = { error: String(e.message || e) }; card.className = "chat-card err"; st.textContent = "✕"; }
                    } else {
                        // MCP-provided tool
                        const card = addCard(`Running "${call.function.name}"…`, "run");
                        const st = card.querySelector(".chat-card-st");
                        try {
                            const r = await mcpClient.callTool(call.function.name, args);
                            result = r.isError ? { error: r.text } : { text: capText(r.text) };
                            card.className = "chat-card " + (r.isError ? "err" : "done"); st.textContent = r.isError ? "✕" : "✓";
                        } catch (e) { result = { error: String(e.message || e) }; card.className = "chat-card err"; st.textContent = "✕"; }
                    }
                    convo.push({ role: "tool", tool_call_id: call.id, content: JSON.stringify(result).slice(0, 6000) });
                }
            }

            // Fallback: honor a fenced ```excel block if the model emitted one instead of
            // calling write_excel_sheet.
            let shown = full;
            const mm = full.match(/```excel\s*([\s\S]*?)```/i);
            if (mm) {
                try {
                    const parsed = JSON.parse(mm[1].trim());
                    for (const a of (Array.isArray(parsed) ? parsed : [parsed])) await maybeConfirmAndWrite(a);
                } catch (e) { console.warn("[Chat] bad action JSON", e); }
                shown = full.replace(/```excel[\s\S]*?```/i, "").trim();
            }

            thinking.textContent = shown || "(no response)";
            history.push({ role: "assistant", content: shown });
            persistHistory();
        } catch (e) {
            thinking.textContent = "⚠ " + e.message;
            console.error("[Chat] error", e);
        } finally {
            if (mcpClient) { try { await mcpClient.close(); } catch (e) { /* already torn down */ } }
            try { abortController.abort(); } catch (e) { /* no-op if already settled */ }
            busy = false; sendBtn.disabled = false;
        }
    }
})();
