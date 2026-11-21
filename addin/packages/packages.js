const $ = id => document.getElementById(id);
const post = (action, args = {}) => chrome.webview.postMessage({ action, ...args });
const key = name => name.toLowerCase().replace(/[-_.]+/g, "-");
const tabs = [...document.querySelectorAll("[role=tab]")];
const jobs = new Map(), packageJobs = new Map(), transcriptEntries = new Map();
let installed = null, readingInventory = true;

const summaries = new Map([
  ["openpyxl", "Read and write Excel workbooks, worksheets, cells, and styles."],
  ["pandas", "Analyze and transform tables with DataFrames."],
  ["numpy", "Work with multidimensional arrays and numerical calculations."],
  ["xlsxwriter", "Create Excel workbooks with formatting, formulas, and charts."],
  ["matplotlib", "Plot data and create charts in Python."],
  ["scipy", "Use scientific computing tools for statistics, optimization, and more."]
]);
const suggested = [...summaries.keys()];
function element(tag, text, className) {
  const node = document.createElement(tag);
  if (text != null) node.textContent = text;
  if (className) node.className = className;
  return node;
}
function action(text, handler, label, className) {
  const button = element("button", text, className);
  button.onclick = handler;
  if (label) button.setAttribute("aria-label", label);
  return button;
}
function selectTab(tab) {
  for (const item of tabs) {
    const selected = item === tab;
    item.setAttribute("aria-selected", selected);
    item.tabIndex = selected ? 0 : -1;
    $(item.getAttribute("aria-controls")).hidden = !selected;
  }
}
for (const [index, tab] of tabs.entries()) {
  tab.onclick = () => selectTab(tab);
  tab.onkeydown = event => {
    const offset = { ArrowLeft: -1, ArrowRight: 1, Home: -index, End: tabs.length - 1 - index }[event.key];
    if (offset == null) return;
    event.preventDefault();
    const next = tabs[(index + tabs.length + offset) % tabs.length];
    selectTab(next);
    next.focus();
  };
}

function viewLog(job) {
  selectTab($("advanced-tab"));
  const entry = transcriptEntries.get(job.id);
  entry.focus();
  entry.scrollIntoView({ block: "center" });
}
function updateActions(row) {
  const name = row.dataset.name, id = key(name);
  const job = packageJobs.get(id);
  const version = installed?.get(id)?.version;
  const busy = job && (job.state === "queued" || job.state === "running");
  const state = busy ? job.state : version ? "installed" : job?.state === "failed" ? "failed" : readingInventory ? "checking" : "add";
  const actions = row.lastChild;
  const failedJob = job?.state === "failed" ? job : null;
  const renderKey = state + (failedJob ? failedJob.id : "");
  if (actions.dataset.state === renderKey) return;
  actions.dataset.state = renderKey;
  const add = () => post("add", { name });
  switch (state) {
    case "checking": actions.replaceChildren(element("span", "Checking…", "status")); break;
    case "queued": actions.replaceChildren(element("span", "Queued", "status")); break;
    case "running": actions.replaceChildren(element("span", "Installing", "status running")); break;
    case "installed": actions.replaceChildren(element("span", "✓ Installed", "status done")); break;
    case "failed": actions.replaceChildren(action("Retry", add, "Retry " + name)); break;
    default: actions.replaceChildren(action("Add", add, "Add " + name, "primary"));
  }
  if (failedJob) actions.prepend(action("View log", () => viewLog(failedJob), "View log for " + name));
}

let summaryRequest;
function showResults(names) {
  summaryRequest?.abort();
  summaryRequest = new AbortController();
  const signal = summaryRequest.signal;
  const rows = names.map(name => {
    const id = key(name);
    const row = element("div", null, "row");
    row.dataset.name = name;
    const info = element("div"), summary = element("p", summaries.get(id));
    info.append(element("strong", name), summary);
    const actions = element("div", null, "actions");
    actions.setAttribute("aria-live", "polite");
    row.append(info, actions);
    updateActions(row);
    if (!summaries.has(id)) {
      fetch("https://pypi.org/pypi/" + encodeURIComponent(name) + "/json", { signal })
        .then(response => response.ok ? response.json() : null)
        .then(project => {
          if (project && !signal.aborted) {
            const value = project.info.summary || "";
            summaries.set(id, value);
            summary.textContent = value;
          }
        }).catch(() => {});
    }
    return row;
  });
  $("results").replaceChildren(...rows);
}

let worker, timer;
function searchStatus(message, retry = false) {
  const status = $("search-status");
  status.textContent = message;
  status.hidden = !message;
  if (retry) status.append(" ", action("Retry search", search));
}
function search() {
  clearTimeout(timer);
  const query = $("search").value.trim();
  $("results-title").textContent = query ? "Search results" : "Suggested for Excel";
  const matches = suggested.filter(name => (name + " " + summaries.get(name)).toLowerCase().includes(query.toLowerCase()));
  showResults(matches);
  searchStatus(query ? "Searching PyPI…" : "");
  if (!query) return;
  timer = setTimeout(() => {
    if (!worker) {
      worker = new Worker("catalog.js");
      worker.onmessage = ({ data }) => {
        if (data.query !== $("search").value.trim()) return;
        if (data.error) { searchStatus(data.error, true); return; }
        const ids = new Set(data.results.map(key));
        const extras = suggested.filter(name => !ids.has(name) &&
          summaries.get(name).toLowerCase().includes(data.query.toLowerCase()));
        showResults([...data.results, ...extras]);
        searchStatus(data.count > data.results.length ? "Showing the first 12 name matches. Refine your search for more specific results."
          : data.count + extras.length === 0 ? "No matching packages found." : "");
      };
      worker.onerror = () => {
        worker.terminate(); worker = null;
        if ($("search").value.trim()) searchStatus("Could not open the package index.", true);
      };
    }
    worker.postMessage(query);
  }, 250);
}
$("search").oninput = search;

function showInventory(inventory) {
  readingInventory = !inventory.installed && !inventory.error;
  installed = inventory.installed && new Map(inventory.installed.map(p => [key(p.name), p]));
  const message = inventory.error || (installed ? (installed.size ? "" : "No packages installed.") : "Reading installed packages…");
  $("inventory-status").textContent = message;
  $("inventory-status").hidden = !message;
  const rows = (inventory.installed || []).map(p => {
    const row = element("div", null, "row");
    row.append(element("strong", p.name), element("span", p.version, "status"));
    return row;
  });
  $("distributions").replaceChildren(...rows);
}
$("refresh").onclick = () => post("refresh");

function updateJob(job) {
  jobs.set(job.id, job);
  if (job.package) packageJobs.set(key(job.package), job);
  let entry = transcriptEntries.get(job.id);
  if (!entry) {
    entry = element("article", null, "job");
    entry.tabIndex = -1;
    const heading = element("div", null, "job-heading");
    heading.append(element("code", "pip " + job.command), element("span"));
    entry.append(heading, element("pre"));
    $("transcript").append(entry);
    $("empty").hidden = true;
    transcriptEntries.set(job.id, entry);
  }
  const status = entry.firstChild.lastChild;
  const labels = { queued: "Queued", running: "Running", done: "Complete", failed: "Failed" };
  status.textContent = labels[job.state] + (job.exitCode == null ? "" : " · exit " + job.exitCode);
  status.className = "status " + job.state;
  entry.lastChild.textContent = job.output;
}

let historyOffset = 0, draft = "";
$("prompt").onsubmit = event => {
  event.preventDefault();
  const command = $("command").value.trim();
  if (command) post("run", { command });
  $("command").value = "";
  historyOffset = 0;
  $("command").focus();
};
$("command").onkeydown = event => {
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
  event.preventDefault();
  const history = [...jobs.values()];
  if (!historyOffset) draft = $("command").value;
  historyOffset = Math.max(0, Math.min(history.length, historyOffset + (event.key === "ArrowUp" ? 1 : -1)));
  $("command").value = historyOffset ? history[history.length - historyOffset].command : draft;
};
chrome.webview.addEventListener("message", ({ data }) => {
  if (data.error) { $("error").textContent = data.error; $("error").hidden = false; return; }
  $("python").textContent = data.python;
  const log = $("transcript");
  const atEnd = log.scrollHeight - log.scrollTop - log.clientHeight < 4;
  for (const job of data.jobs) updateJob(job);
  if (atEnd) log.scrollTop = log.scrollHeight;
  if (data.inventory) showInventory(data.inventory);
  for (const row of $("results").children) updateActions(row);
});
search();
post("ready");
