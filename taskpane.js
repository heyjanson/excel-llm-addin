// ----------------------------------------------------------------------
// Excel LLM Assistant - taskpane.js
// Written in ES5-compatible syntax (no async/await, arrow functions,
// template literals, or fetch) because Excel 2016's task pane on
// Windows renders using Internet Explorer 11, which does not support
// modern JavaScript. Using newer syntax causes the entire script to
// fail to parse, so no click handlers get attached at all.
// ----------------------------------------------------------------------

var STORAGE_KEY = "excel_llm_openai_api_key";
var MAX_ROWS = 100; // cap to avoid huge payloads / token limits

Office.onReady(function () {
  // Restore saved API key (if any)
  var saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    document.getElementById("apiKey").value = saved;
  }

  document.getElementById("saveKey").onclick = saveApiKey;
  document.getElementById("ask").onclick = askLLM;

  var scopeEl = document.getElementById("scope");
  scopeEl.addEventListener("change", updateScopeHint);
  updateScopeHint();
});

function trim(str) {
  return str.replace(/^\s+|\s+$/g, "");
}

function saveApiKey() {
  var key = trim(document.getElementById("apiKey").value);
  if (!key) {
    setStatus("Please enter an API key before saving.");
    return;
  }
  localStorage.setItem(STORAGE_KEY, key);
  setStatus("API key saved.");
}

function setStatus(msg) {
  document.getElementById("status").innerText = msg;
}

function updateScopeHint() {
  var scope = document.getElementById("scope").value;
  var hint = document.getElementById("scopeHint");
  if (scope === "selection") {
    hint.innerText = "Will send the values from your current cell selection.";
  } else {
    hint.innerText = "Will send the active sheet's used range (capped at " + MAX_ROWS + " rows).";
  }
}

// Returns an OfficeExtension.Promise resolving to a text description
// of the relevant workbook data.
function getWorkbookContext(scope) {
  return Excel.run(function (ctx) {
    if (scope === "selection") {
      var range = ctx.workbook.getSelectedRange();
      range.load("values, address");

      return ctx.sync().then(function () {
        return "Selected range " + range.address + ":\n" + JSON.stringify(range.values);
      });
    } else {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      var usedRange = sheet.getUsedRange();
      usedRange.load("values, address, rowCount");

      return ctx.sync().then(function () {
        var values = usedRange.values;
        var totalRows = usedRange.rowCount;

        if (totalRows > MAX_ROWS) {
          values = values.slice(0, MAX_ROWS);
          return "Sheet \"" + sheet.name + "\", used range " + usedRange.address +
            " (showing first " + MAX_ROWS + " of " + totalRows + " rows):\n" +
            JSON.stringify(values);
        } else {
          return "Sheet \"" + sheet.name + "\", used range " + usedRange.address + ":\n" +
            JSON.stringify(values);
        }
      });
    }
  });
}

function askLLM() {
  var apiKey = localStorage.getItem(STORAGE_KEY);
  var question = trim(document.getElementById("question").value);
  var scope = document.getElementById("scope").value;
  var answerEl = document.getElementById("answer");

  if (!apiKey) {
    setStatus("Please enter and save your OpenAI API key first.");
    return;
  }
  if (!question) {
    setStatus("Please enter a question.");
    return;
  }

  answerEl.innerText = "";
  setStatus("Reading workbook data...");

  getWorkbookContext(scope).then(function (context) {
    setStatus("Asking the model...");

    callOpenAI(apiKey, context, question, function (errMsg, answerText) {
      if (errMsg) {
        answerEl.innerText = "";
        setStatus("Error: " + errMsg);
      } else {
        answerEl.innerText = answerText;
        setStatus("");
      }
    });
  }).catch(function (err) {
    setStatus("Error reading workbook: " + (err && err.message ? err.message : err));
  });
}

// Uses XMLHttpRequest (not fetch) for compatibility with IE11.
function callOpenAI(apiKey, context, question, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open("POST", "https://api.openai.com/v1/chat/completions", true);
  xhr.setRequestHeader("Content-Type", "application/json");
  xhr.setRequestHeader("Authorization", "Bearer " + apiKey);

  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          var data = JSON.parse(xhr.responseText);
          callback(null, data.choices[0].message.content);
        } catch (e) {
          callback("Could not parse response: " + e.message);
        }
      } else {
        callback("API error " + xhr.status + ": " + xhr.responseText);
      }
    }
  };

  xhr.onerror = function () {
    callback("Network error calling OpenAI API.");
  };

  var body = JSON.stringify({
    model: "gpt-4o-mini",
    temperature: 0.2,
    messages: [
      {
        role: "system",
        content: "You are a helpful data analyst answering questions about the user's Excel workbook. The data is provided as a JSON array of rows (each row is itself an array of cell values, in the same order as they appear on the sheet). Be concise and reference specific rows, columns, or values where relevant."
      },
      {
        role: "user",
        content: context + "\n\nQuestion: " + question
      }
    ]
  });

  xhr.send(body);
}
