/**
 * PrivCite v2.0 - Fully Integrated Sidebar Taskpane Controller
 * Integrates Microsoft Word Office.js bindings with the production-grade Swift engine.
 */

// Initialize the Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("PrivCite v2.0 Add-in initialized inside Microsoft Word.");
    initializeUI();
  } else {
    console.log("PrivCite v2.0 initialized in standalone web test view.");
    initializeUI();
  }
});

// Cache for active issues scanned from the document
let activeIssues = [];

function initializeUI() {
  const btnCheck = document.getElementById("btnCheckDocument");
  const pedagogicalToggle = document.getElementById("pedagogicalToggle");
  const pedagogicalBanner = document.getElementById("pedagogicalBanner");
  const trackChangesReminder = document.getElementById("trackChangesReminder");
  
  const hoverSelectionToggle = document.getElementById("hoverSelectionToggle");
  const hoverDelayInput = document.getElementById("hoverDelayInput");
  const hoverDelayVal = document.getElementById("hoverDelayVal");
  const hoverDelayRow = document.getElementById("hoverDelayRow");

  // Toggle Educational Banner based on state
  pedagogicalToggle.addEventListener("change", () => {
    if (pedagogicalToggle.checked) {
      pedagogicalBanner.style.display = "flex";
      trackChangesReminder.style.display = "none";
    } else {
      pedagogicalBanner.style.display = "none";
      trackChangesReminder.style.display = "flex";
    }
    // Dynamically re-render issue buttons to reflect standard vs pedagogical action labels
    renderIssues(activeIssues);
  });

  // Handle On-Hover auto-selection toggle
  hoverSelectionToggle.addEventListener("change", () => {
    if (hoverSelectionToggle.checked) {
      hoverDelayRow.style.opacity = "1";
      hoverDelayInput.disabled = false;
    } else {
      hoverDelayRow.style.opacity = "0.4";
      hoverDelayInput.disabled = true;
    }
  });

  // Handle hover delay slider input change
  hoverDelayInput.addEventListener("input", () => {
    hoverDelayVal.textContent = hoverDelayInput.value;
  });

  // Trigger Scanning process
  btnCheck.addEventListener("click", () => {
    btnCheck.innerHTML = `
      <svg class="logo-icon" style="animation: spin 1s linear infinite; width: 16px; height: 16px; display: inline-block; vertical-align: middle; margin-right: 6px;" viewBox="0 0 24 24">
        <circle cx="12" cy="12" r="10" stroke="rgba(255,255,255,0.2)" stroke-width="3" fill="none"/>
        <path d="M12 2a10 10 0 0 1 10 10" stroke="currentColor" stroke-width="3" stroke-linecap="round" fill="none"/>
      </svg>
      Analyzing Brief...
    `;
    
    // Trigger the real-time scanning
    if (typeof Word !== 'undefined') {
      readWordDocumentAndRender();
    } else {
      // Mock fallback for browser-only testing
      const sampleText = "We cite Map v. Ohio, 367 U.S. 643 (1961) and 543 F.Supp.2d 110.";
      readStandaloneAndRender(sampleText);
    }
  });
}

/**
 * Reads the actual active Word document text and queries our secure local Swift engine.
 */
async function readWordDocumentAndRender() {
  const btnCheck = document.getElementById("btnCheckDocument");
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      
      const docText = body.text;
      console.log("Active document text extracted. Length:", docText.length);
      
      if (!docText || docText.trim() === "") {
        activeIssues = [];
        renderIssues(activeIssues);
        resetButtonText();
        return;
      }
      // Get the spacing preference from the dropdown select setting
      const spacingPref = document.getElementById("spacingPreference").value;
      
      // Fetch analysis results from our secure Python-to-Swift API bridge
      const response = await fetch("https://localhost:3000/api/analyze", {
        method: "POST",
        headers: {
          "Content-Type": "text/plain",
          "X-PrivCite-Spacing-Preference": spacingPref
        },
        body: docText
      });
      
      if (!response.ok) {
        throw new Error(`Server returned HTTP status ${response.status}`);
      }
      
      activeIssues = await response.json();
      console.log("Swift engine completed analysis. Issues found:", activeIssues.length);
      updateStatusIndicator(true); // Local SQLite Database status: active
      renderIssues(activeIssues);
      resetButtonText();
    });
  } catch (error) {
    console.error("Office.js Word Text Scan failed: ", error);
    updateStatusIndicator(false); // Flag SQLite backend offline
    // Local fallback to client-side regex if server is unreachable
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        activeIssues = parseTextDynamically(body.text);
        renderIssues(activeIssues);
        resetButtonText();
      });
    } catch (e) {
      resetButtonText();
    }
  }
}

/**
 * Standalone browser testing route to hit the API server directly.
 */
async function readStandaloneAndRender(text) {
  try {
    const spacingPref = document.getElementById("spacingPreference").value;
    const response = await fetch("https://localhost:3000/api/analyze", {
      method: "POST",
      headers: {
        "Content-Type": "text/plain",
        "X-PrivCite-Spacing-Preference": spacingPref
      },
      body: text
    });
    
    if (!response.ok) {
      throw new Error(`HTTP status ${response.status}`);
    }
    
    activeIssues = await response.json();
    updateStatusIndicator(true);
    renderIssues(activeIssues);
    resetButtonText();
  } catch (error) {
    console.warn("Local server unreachable in browser. Falling back to local JS regex parser: ", error);
    updateStatusIndicator(false);
    activeIssues = parseTextDynamically(text);
    renderIssues(activeIssues);
    resetButtonText();
  }
}

function resetButtonText() {
  const btnCheck = document.getElementById("btnCheckDocument");
  btnCheck.innerHTML = `
    <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/>
    </svg>
    Scan Brief Citations
  `;
}

/**
 * High-precision, local Javascript citation parser running in the sidebar as a fallback.
 * @param {string} text - The raw body text extracted from Microsoft Word.
 */
function parseTextDynamically(text) {
  const issues = [];
  let id = 1;

  if (!text || text.trim() === "") {
    return issues;
  }

  // Heuristic 1: Landmark Case Spelling Typo (Map v. Ohio -> Mapp v. Ohio)
  const mappRegex = /\bMap\s+v\.\s+Ohio\b/gi;
  let match;
  while ((match = mappRegex.exec(text)) !== null) {
    issues.push({
      id: id++,
      category: "Citations",
      severity: "warning",
      targetText: match[0],
      suggestion: "Mapp v. Ohio",
      message: "Landmark case name spelling typo detected. The official registry spelling is 'Mapp v. Ohio' with two 'p's.",
      rule: "Bluebook Rule 10.2: Case names must match the official reporter listings. Common landmark cases are cross-checked locally."
    });
  }

  // Heuristic 2: Spacing inside open abbreviations (e.g., F.Supp.2d, F. Supp.2d -> F. Supp. 2d)
  const openRegex = /\b(\d+)\s+([Ff]\.[Ss]upp\.[2dDd]+|[Ff]\.\s+[Ss]upp\.[2dDd]+|[Ff]\.[Ss]upp\.\s+[2dDd]+)\s+(\d+)\b/g;
  while ((match = openRegex.exec(text)) !== null) {
    const rawMatch = match[0];
    const vol = match[1];
    const reporter = match[2];
    const page = match[3];
    
    if (reporter !== "F. Supp. 2d" && reporter !== "f. supp. 2d") {
      issues.push({
        id: id++,
        category: "Spacing",
        severity: "error",
        targetText: rawMatch,
        suggestion: `${vol} F. Supp. 2d ${page}`,
        message: "Irregular spacing inside reporter. Multi-letter abbreviations require surrounding spaces.",
        rule: "Bluebook Rule 6.1: Closed abbreviations apply only to adjacent single capital letters. Multi-letter abbreviations (like 'Supp.', 'App.', 'Decl.') must have open spaces."
      });
    }
  }

  // Heuristic 3: Spacing inside closed abbreviations (e.g. F. 3d, U. S. -> F.3d, U.S.)
  const closedRegex = /\b(\d+)\s+(F\.\s+3d|F\.\s+2d|U\.\s+S\.|S\.\s+Ct\.)\s+(\d+)\b/g;
  while ((match = closedRegex.exec(text)) !== null) {
    const rawMatch = match[0];
    const vol = match[1];
    const abbrev = match[2];
    const page = match[3];
    
    const correctedAbbrev = abbrev.replace(/\s+/g, ""); // close up spaces
    issues.push({
      id: id++,
      category: "Spacing",
      severity: "error",
      targetText: rawMatch,
      suggestion: `${vol} ${correctedAbbrev} ${page}`,
      message: "Irregular spacing inside reporter. Adjacent single-letter abbreviations must be closed up.",
      rule: "Bluebook Rule 6.1: Closed abbreviations apply to adjacent single capital letters (e.g., F.3d, U.S., S.Ct., N.Y.) with no internal spaces."
    });
  }

  return issues;
}

/**
 * Renders the parsed issue cards in the sidebar UI dynamically.
 */
function renderIssues(issues) {
  const container = document.getElementById("issueContainer");
  const badge = document.getElementById("issueBadge");
  const btnAutoFormat = document.getElementById("btnAutoFormat");
  
  container.innerHTML = "";
  badge.textContent = `${issues.length} Issues`;
  
  if (issues.length === 0) {
    if (btnAutoFormat) btnAutoFormat.style.display = "none";
    container.innerHTML = `
      <div id="emptyState" style="text-align: center; padding: 30px; color: var(--text-muted); font-size: 0.85rem; border: 1px dashed var(--border-light); border-radius: 8px; animation: slideInUp 0.3s ease;">
        🎉 No citation formatting issues found! Your open brief is structurally flawless.
      </div>
    `;
    return;
  }
  
  // Inspect the Pedagogical Toggle to determine active button label
  const isPedagogical = document.getElementById("pedagogicalToggle").checked;
  const actionLabel = isPedagogical ? "Inject Word Comment" : "Apply Redline Edit";
  
  const spacingIssues = issues.filter(i => i.category.toLowerCase() === "spacing");
  if (btnAutoFormat) {
    if (spacingIssues.length > 0 && !isPedagogical) {
      btnAutoFormat.style.display = "inline-block";
    } else {
      btnAutoFormat.style.display = "none";
    }
  }
  
  issues.forEach((issue) => {
    const card = document.createElement("div");
    card.className = `issue-card ${issue.severity}`;
    
    card.innerHTML = `
      <div class="issue-meta">
        <span>${issue.category}</span>
        <span>${issue.severity}</span>
      </div>
      <div class="issue-text">${escapeHtml(issue.targetText || "")}</div>
      <div class="issue-message">${issue.message}</div>
      <div class="issue-action-row">
        <button class="btn-action explain" onclick="explainIssue('${issue.id}')">Explain Rule</button>
        <button class="btn-action apply" onclick="applyIssue('${issue.id}')">${actionLabel}</button>
      </div>
      <div id="explanation-${issue.id}" style="display: none; font-size: 0.76rem; color: var(--text-secondary); margin-top: 8px; border-top: 1px dashed rgba(255,255,255,0.05); padding-top: 8px; line-height: 1.4;">
        ${issue.rule || "Audited and flagged by the PrivCite Swift Parser Engine."}
      </div>
    `;
    
    // Smooth, debounced on-hover document highlighting
    card.addEventListener("mouseenter", () => {
      highlightIssueInWord(issue.id);
    });
    
    // On-click automatic selection and scroll-into-view
    card.addEventListener("click", (e) => {
      if (e.target.closest(".btn-action")) {
        return; // Skip if they clicked "Explain Rule" or "Apply Redline Edit"
      }
      selectIssueInWord(issue.id);
    });
    
    container.appendChild(card);
  });

  // Dynamic Spacing Symmetrizer Command Center (reuses spacingIssues filtered above)
  const symmetrizerContainer = document.getElementById("spacingSymmetrizerContainer");
  
  if (spacingIssues.length > 0 && !isPedagogical) {
    symmetrizerContainer.innerHTML = `
      <div class="spacing-symmetrizer-card">
        <div class="spacing-symmetrizer-title">
          <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor">
            <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/>
          </svg>
          Spacing Symmetrizer Center
        </div>
        <div class="spacing-symmetrizer-desc">
          We detected <strong>${spacingIssues.length} minor spacing anomalies</strong> in this brief. To prevent warning fatigue, these have been condensed here. 
        </div>
        <button class="btn-symmetrize" onclick="formatAllSpacingToStrictBluebook()">Format All Spacing to Strict Bluebook</button>
      </div>
    `;
  } else {
    symmetrizerContainer.innerHTML = "";
  }
}

/**
 * Handles "Explain Rule" drawer toggling within the card view.
 */
window.explainIssue = function(id) {
  const drawer = document.getElementById(`explanation-${id}`);
  if (drawer.style.display === "none") {
    drawer.style.display = "block";
    drawer.style.animation = "slideInDown 0.3s ease";
  } else {
    drawer.style.display = "none";
  }
};

/**
 * Handles "Apply" button action. Toggles between auto-fix replacement and native Word comment injection.
 */
window.applyIssue = async function(id) {
  const issue = activeIssues.find(i => i.id === id);
  if (!issue) return;
  
  const isPedagogical = document.getElementById("pedagogicalToggle").checked;
  const targetTextStr = issue.targetText || issue.message; // Fallback if citationText empty
  
  // Track Changes Safety Warning Modal (Theresa / Clerk Overwrite Anxiety Mitigation)
  if (!isPedagogical) {
    const skipSafetyPrompt = sessionStorage.getItem("skipTrackChangesSafetyPrompt") === "true";
    if (!skipSafetyPrompt) {
      const confirmFix = confirm(
        "⚠️ TRACK CHANGES SAFETY REMINDER\n\n" +
        "Auto-Fix Mode will directly replace text in your document.\n\n" +
        "Please ensure 'Track Changes' is ON in Word's 'Review' ribbon tab so this correction is recorded as a tracked redline. If disabled, the edit will overwrite text silently.\n\n" +
        "Click OK to proceed with direct overwrite, or Cancel to enable Track Changes in the ribbon first."
      );
      if (!confirmFix) return;
      
      if (confirm("Would you like to suppress this Track Changes warning for the rest of this editing session?")) {
        sessionStorage.setItem("skipTrackChangesSafetyPrompt", "true");
      }
    }
  }
  
  if (typeof Word === 'undefined') {
    alert(`Standalone Browser Demo:\n` +
          `• Mode: ${isPedagogical ? "LRW Pedagogical Mode (Inject comment)" : "Standard Mode (Auto-fix text)"}\n` +
          `• Targeting: "${targetTextStr}"\n` +
          `• Correction: "${issue.suggestion}"`);
    return;
  }
  
  try {
    if (isPedagogical) {
      // Inject native Microsoft Word Comment bubble and highlight text span in yellow
      await Word.run(async (context) => {
        const body = context.document.body;
        const searchResults = body.search(targetTextStr, { matchCase: true, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        
        if (searchResults.items.length > 0) {
          const range = searchResults.items[0];
          
          const commentContent = `PrivCite Suggestion:\n` +
                                 `• Drafted: "${targetTextStr}"\n` +
                                 `• Recommends: "${issue.suggestion}"\n\n` +
                                 `Bluebook Background:\n${issue.message}`;
                                 
          range.insertComment(commentContent);
          range.font.highlightColor = "#FFFF00"; // Highlight yellow
          await context.sync();
        } else {
          // If direct search fails, add comment to selection or alert
          console.warn(`Target text "${targetTextStr}" not found during comment insertion.`);
        }
      });
    } else {
      // Direct replacement auto-fix mode: replaces text directly, generating tracked redlines if Track Changes is active
      await Word.run(async (context) => {
        const body = context.document.body;
        const searchResults = body.search(targetTextStr, { matchCase: true, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        
        if (searchResults.items.length > 0) {
          const range = searchResults.items[0];
          range.insertText(issue.suggestion, Word.InsertLocation.replace);
          await context.sync();
        } else {
          console.warn(`Target text "${targetTextStr}" not found during auto-fix.`);
        }
      });
    }
  } catch (error) {
    console.error("Action execution in Word failed: ", error);
  }
};

function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

window.highlightIssueInWord = function(id) {
  const hoverSelectionToggle = document.getElementById("hoverSelectionToggle");
  if (hoverSelectionToggle && !hoverSelectionToggle.checked) {
    return; // Hover auto-selection is turned OFF (Diana / Slower User stability focus)
  }

  if (hoverTimeout) {
    clearTimeout(hoverTimeout);
  }
  
  const hoverDelayInput = document.getElementById("hoverDelayInput");
  const delay = hoverDelayInput ? parseInt(hoverDelayInput.value) : 120;
  
  hoverTimeout = setTimeout(async () => {
    const issue = activeIssues.find(i => i.id === id);
    if (!issue) return;
    
    const targetTextStr = issue.targetText || issue.message;
    if (typeof Word === 'undefined') {
      console.log(`Standalone Browser Hover highlight targeting: "${targetTextStr}" (Delay: ${delay}ms)`);
      return;
    }
    
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const searchResults = body.search(targetTextStr, { matchCase: true, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        
        if (searchResults.items.length > 0) {
          const range = searchResults.items[0];
          range.select();
          await context.sync();
        }
      });
    } catch (error) {
      console.error("Selecting text in Word failed: ", error);
    }
  }, delay);
};

window.selectIssueInWord = async function(id) {
  const issue = activeIssues.find(i => i.id === id);
  if (!issue) return;
  
  const targetTextStr = issue.targetText || issue.message;
  if (typeof Word === 'undefined') {
    console.log(`Standalone Browser Click selection targeting: "${targetTextStr}"`);
    return;
  }
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const searchResults = body.search(targetTextStr, { matchCase: true, matchWholeWord: false });
      searchResults.load("items");
      await context.sync();
      
      if (searchResults.items.length > 0) {
        const range = searchResults.items[0];
        range.select();
        range.scrollIntoView("Inside");
        await context.sync();
      }
    });
  } catch (error) {
    console.error("Selecting and scrolling text in Word failed: ", error);
  }
};

/**
 * Spacing Command Center Bridge: formats all minor spacing anomalies
 * in the active Word document to strict Bluebook rules in a single batch pass.
 * (Charles Formatting Conformity + Matthew Warning Fatigue Resolution)
 */
window.formatAllSpacingToStrictBluebook = async function() {
  const spacingIssues = activeIssues.filter(i => i.category.toLowerCase() === "spacing");
  if (spacingIssues.length === 0) {
    alert("No spacing issues found to format!");
    return;
  }
  
  if (typeof Word === 'undefined') {
    alert(`Standalone Demo: Programmatically formatting ${spacingIssues.length} spacing issues in document!`);
    return;
  }
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      let count = 0;
      
      for (const issue of spacingIssues) {
        const targetTextStr = issue.targetText || issue.message;
        const searchResults = body.search(targetTextStr, { matchCase: true, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        
        if (searchResults.items.length > 0) {
          for (let j = 0; j < searchResults.items.length; j++) {
            searchResults.items[j].insertText(issue.suggestion, Word.InsertLocation.replace);
            count++;
          }
        }
      }
      
      await context.sync();
      alert(`Symmetrizer complete: Successfully formatted ${count} spacing occurrences to strict Bluebook!`);
      
      // Clear spacing issues from active pool and re-render
      activeIssues = activeIssues.filter(i => i.category.toLowerCase() !== "spacing");
      renderIssues(activeIssues);
    });
  } catch (error) {
    console.error("Batch spacing symmetrizer failed: ", error);
  }
};

window.autoFormatAllSpacing = window.formatAllSpacingToStrictBluebook;

/**
 * Registry Delta Sync Reconnector: attempts to ping the localized API server
 * to verify communication status and toggles indicator status.
 */
window.reconnectEngine = async function() {
  updateStatusIndicator(true, "Attempting reconnect...");
  const statusDot = document.getElementById("statusDot");
  statusDot.className = "status-dot connecting";
  
  try {
    const response = await fetch("https://localhost:3000/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: "Map v. Ohio, 367 U.S. 643 (1961)"
    });
    
    if (response.ok) {
      updateStatusIndicator(true);
    } else {
      throw new Error();
    }
  } catch (e) {
    updateStatusIndicator(false);
  }
};

/**
 * Status Indicator Controller: sets footer engine status details.
 */
function updateStatusIndicator(online, connectingText) {
  const statusDot = document.getElementById("statusDot");
  const statusText = document.getElementById("statusText");
  const btnReconnect = document.getElementById("btnReconnect");
  
  if (!statusDot || !statusText || !btnReconnect) return;
  
  if (online) {
    statusDot.className = "status-dot online";
    statusText.textContent = connectingText || "Local Engine Active (DB v2026.05.28)";
    btnReconnect.style.display = "none";
  } else {
    statusDot.className = "status-dot offline";
    statusText.textContent = "Local Engine Disconnected";
    btnReconnect.style.display = "inline-block";
  }
}
