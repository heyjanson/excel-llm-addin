// ----------------------------------------------------------------------
// Excel LLM Assistant - taskpane.js
// Reads data from the active workbook and sends it, along with a user
// question, to the OpenAI Chat Completions API.
// ----------------------------------------------------------------------

const STORAGE_KEY = "excel_llm_openai_api_key";
const MAX_ROWS = 100; // cap to avoid huge payloads / token limits

Office.onReady(() => {
  // Restore saved API key (if any)
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    document.getElementById("apiKey").value = saved;
  }

  document.getElementById("saveKey").onclick = saveApiKey;
  document.getElementById("ask").onclick = askLLM;

  const scopeEl = document.getElementById("scope");
  scopeEl.addEventListener("change", updateScopeHint);
  updateScopeHint();
});

function saveApiKey() {
  const key = document.getElementById("apiKey").value.trim();
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
  const scope = document.getElementById("scope").value;
  const hint = document.getElementById("scopeHint");
  hint.innerText =
    scope === "selection"
      ? "Will send the values from your current cell selection."
      : `Will send the active sheet's used range (capped at ${MAX_ROWS} rows).`;
}

// Reads data from the workbook based on the chosen scope and returns
// a text block describing it, suitable for inclusion in an LLM prompt.
async function getWorkbookContext(scope) {
  let context = "";

  await Excel.run(async (ctx) => {
    if (scope === "selection") {
      const range = ctx.workbook.getSelectedRange();
      range.load("values, address");
      await ctx.sync();

      context = `Selected range ${range.address}:\n${JSON.stringify(range.values)}`;
    } else {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const range = sheet.getUsedRange();
      range.load("values, address, rowCount");
      await ctx.sync();

      let values = range.values;
      const totalRows = range.rowCount;

      if (totalRows > MAX_ROWS) {
        values = values.slice(0, MAX_ROWS);
        context =
          `Sheet "${sheet.name}", used range ${range.address} ` +
          `(showing first ${MAX_ROWS} of ${totalRows} rows):\n` +
          JSON.stringify(values);
      } else {
        context =
          `Sheet "${sheet.name}", used range ${range.address}:\n` +
          JSON.stringify(values);
      }
    }
  });

  return context;
}

async function askLLM() {
  const apiKey = localStorage.getItem(STORAGE_KEY);
  const question = document.getElementById("question").value.trim();
  const scope = document.getElementById("scope").value;
  const answerEl = document.getElementById("answer");

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

  let context;
  try {
    context = await getWorkbookContext(scope);
  } catch (err) {
    setStatus("Error reading workbook: " + err.message);
    return;
  }

  setStatus("Asking the model...");

  try {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        temperature: 0.2,
        messages: [
          {
            role: "system",
            content:
              "You are a helpful data analyst answering questions about " +
              "the user's Excel workbook. The data is provided as a JSON " +
              "array of rows (each row is itself an array of cell values, " +
              "in the same order as they appear on the sheet). Be concise " +
              "and reference specific rows, columns, or values where relevant."
          },
          {
            role: "user",
            content: `${context}\n\nQuestion: ${question}`
          }
        ]
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`API error ${response.status}: ${errText}`);
    }

    const data = await response.json();
    answerEl.innerText = data.choices[0].message.content;
    setStatus("");
  } catch (err) {
    answerEl.innerText = "";
    setStatus("Error: " + err.message);
  }
}
