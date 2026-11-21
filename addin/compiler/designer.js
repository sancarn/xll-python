"use strict";

const $ = id => document.getElementById(id);
const bridge = window.chrome.webview;
let project, paths, saved, selectedFile, installed, selection, busy = false;
let reply, sequence = 0;
const images = new Map();
let macros = [];
const directoryState = new Map();

function host(action, data = {}, drop) {
  const id = ++sequence;
  return new Promise((resolve, reject) => {
    reply = {id, resolve, reject};
    const message = {action, id, ...data};
    if (drop) bridge.postMessageWithAdditionalObjects(message, [drop]);
    else bridge.postMessage(message);
  });
}
bridge.addEventListener("message", ({data}) => {
  if (data.event === "close") return run(() => leave(() => host("close")));
  if (reply && data.id === reply.id) {
    const current = reply;
    reply = null;
    if (data.error) current.reject(new Error(data.error));
    else current.resolve(data.result);
  }
});
function status(message = "", error = false) {
  $("status").classList.toggle("error", error);
  $("status").textContent = message;
}
async function run(action) {
  if (busy) return;
  busy = true;
  for (const node of document.querySelectorAll("body > :is(main,header,footer)")) node.inert = true;
  try { await action(); }
  catch (error) { status(error.message, true); }
  finally {
    busy = false;
    for (const node of document.querySelectorAll("[inert]")) node.inert = false;
  }
}
function node(tag, text, cls) {
  const el = document.createElement(tag);
  if (text !== undefined) el.textContent = text;
  if (cls) el.className = cls;
  return el;
}
function button(text, action, cls) {
  const el = node("button", text, cls);
  el.onclick = action;
  return el;
}
function startPage(welcome) {
  $("welcome").hidden = !welcome;
  for (const id of ["home", "editor", "projectActions"]) $(id).hidden = welcome;
}
async function start(action, drop) {
  const result = await host(action, {}, drop);
  if (!result) return;
  project = result.project;
  paths = result.paths;
  saved = result.manifest ? JSON.stringify(project) : null;
  selectedFile = null;
  selection = {};
  images.clear();
  directoryState.clear();
  $("name").value = project.name;
  $("optimise").checked = project.optimise === true;
  $("folder").textContent = result.root;
  $("spec").value = "";
  $("filename").textContent = "File preview";
  $("fileHint").hidden = false;
  $("roleLabel").hidden = $("source").hidden = true;
  startPage(false);
  await showTab("info");
  status(result.manifest || "Select a Python file and mark it as an entry module.");
}
async function save() {
  const path = await host("save", {project});
  if (path) {
    saved = JSON.stringify(project);
    status("Saved · " + path);
  }
  return !!path;
}
async function leave(next) {
  if ($("leave").open) return;
  if (project && JSON.stringify(project) !== saved) {
    const dialog = $("leave");
    const answer = await new Promise(resolve => {
      dialog.onclose = () => resolve(dialog.returnValue);
      dialog.returnValue = "cancel";
      dialog.showModal();
    });
    if (answer === "cancel" || (answer === "save" && !await save())) return;
  }
  await next();
}
$("cancelLeave").onclick = () => $("leave").close("cancel");
$("discard").onclick = () => $("leave").close("discard");
$("saveLeave").onclick = () => $("leave").close("save");
$("home").onclick = () => run(() => leave(() => { project = null; startPage(true); status(); }));
$("newProject").onclick = () => run(() => start("new"));
$("openProject").onclick = () => run(() => start("open"));
document.ondragover = event => {
  event.preventDefault();
  event.dataTransfer.dropEffect = project ? "none" : "copy";
  $("welcome").classList.toggle("over", !project);
};
document.ondragleave = () => $("welcome").classList.remove("over");
document.ondrop = event => {
  event.preventDefault();
  $("welcome").classList.remove("over");
  if (project) return;
  const files = event.dataTransfer.files;
  if (files.length === 1) run(() => start("new", files[0]));
  else status("Drop one project folder.", true);
};
$("save").onclick = () => run(save);
$("compile").onclick = () => run(async () => {
  if (!project.name.trim() || !project.entries.length) throw Error("Enter an add-in name and mark at least one Python file as an entry module.");
  for (const tab of project.tabs) {
    if (!tab.label.trim()) throw Error("Enter a display label for each ribbon tab.");
    for (const group of tab.groups) {
      if (!group.label.trim()) throw Error("Enter a display label for each ribbon group.");
      for (const item of group.buttons) {
        if (!item.label.trim() || !item.command.trim()) throw Error("Each ribbon button needs a label and a Python function.");
      }
    }
  }
  if (!await save()) return;
  status("Compiling " + project.name + "…");
  const path = await host("compile", {name: project.name});
  status(path ? "Built " + path : "Project saved. Compilation cancelled.");
});
$("optimise").onchange = e => { project.optimise = e.target.checked; };
$("name").oninput = e => { project.name = e.target.value; };

async function showTab(tab) {
  for (const id of ["info", "ribbon", "packages"]) $(id).hidden = id !== tab;
  for (const b of document.querySelectorAll("nav button")) {
    b.classList.toggle("active", b.dataset.tab === tab);
    if (b.dataset.tab === tab) b.setAttribute("aria-current", "page");
    else b.removeAttribute("aria-current");
  }
  if (tab === "info") renderTree();
  if (tab === "ribbon") {
    if (!selection.tab) selection = {tab: project.tabs[0]};
    macros = [];
    renderRibbon();
    macros = await host("commands", {entries: project.entries});
    for (const t of project.tabs) for (const g of t.groups) for (const b of g.buttons) {
      if (b.image && !images.has(b.image)) {
        const result = await host("image", {name: b.image});
        images.set(b.image, result.src);
      }
    }
    renderRibbon();
  }
  if (tab === "packages") {
    if (!installed) installed = await host("installed");
    $("installed").replaceChildren(...installed.map(name => node("option", name)));
    renderPackages();
  }
}
for (const b of document.querySelectorAll("nav button")) b.onclick = () => run(() => showTab(b.dataset.tab));

function role(path) {
  if (project.entries.includes(path)) return "entry";
  return project.files.includes(path) ? "helper" : "excluded";
}
function renderTree() {
  const tree = $("tree");
  tree.replaceChildren();
  const folders = new Map([["", tree]]);
  const names = [...new Set([...paths, ...project.entries, ...project.files])].sort((a,b) => a.localeCompare(b));
  for (const path of names) {
    let parent = tree, key = "";
    const parts = path.split("/");
    const name = parts.pop();
    for (const part of parts) {
      key += part + "/";
      if (!folders.has(key)) {
        const details = node("details");
        const folderKey = key;
        details.open = directoryState.get(key) ?? true;
        details.ontoggle = () => { if (details.isConnected) directoryState.set(folderKey, details.open); };
        details.append(node("summary", part));
        parent.append(details);
        folders.set(key, details);
      }
      parent = folders.get(key);
    }
    const item = button("", () => run(() => openFile(path)), "file");
    item.append(node("span", name));
    item.title = path;
    item.classList.toggle("active", selectedFile === path);
    const kind = role(path);
    if (kind !== "excluded") item.append(node("small", kind === "entry" ? "Entry" : "Included"));
    parent.append(item);
  }
  if (!names.length) tree.append(node("p", "No project files found."));
}
async function openFile(path) {
  selectedFile = path;
  renderTree();
  $("filename").textContent = path;
  $("fileHint").hidden = true;
  $("roleLabel").hidden = $("source").hidden = false;
  $("role").options[0].disabled = !path.toLowerCase().endsWith(".py");
  $("role").value = role(path);
  $("source").textContent = "";
  $("source").textContent = await host("preview", {name: path});
}
$("role").onchange = () => {
  const path = selectedFile;
  project.entries = project.entries.filter(p => p !== path);
  project.files = project.files.filter(p => p !== path);
  if ($("role").value === "entry") project.entries.push(path);
  if ($("role").value === "helper") project.files.push(path);
  renderTree();
  status("Project file selection updated.");
};
$("addFiles").onclick = () => run(async () => {
  const names = await host("addFiles");
  if (!names?.length) return;
  for (const name of names) {
    if (!paths.includes(name)) paths.push(name);
    if (role(name) === "excluded") project.files.push(name);
  }
  await openFile(names[0]);
  status("Project files included.");
});
$("refresh").onclick = () => run(async () => {
  paths = await host("refresh");
  renderTree();
  if (selectedFile) await openFile(selectedFile);
  status("Project files refreshed.");
});

function selectedNode() { return selection.item || selection.group || selection.tab; }
function select(tab, group, item) {
  selection = {tab, group, item};
  renderRibbon();
}
const glyphs = {MacroPlay:"▶", RefreshAll:"⟳", FileOpen:"📂", FileSave:"▣", TableInsert:"▦"};
function renderPreview() {
  const {tab, group, item} = selection;
  $("officeTabs").replaceChildren(...project.tabs.map(t =>
    button(t.label || "Untitled tab", () => select(t), t === tab ? "active" : "")));
  const groups = $("groups");
  groups.replaceChildren();
  if (!tab || !tab.groups.length) groups.append(node("p", tab ? "Add a group to this tab." : "Add a tab to start designing your ribbon."));
  for (const g of tab?.groups || []) {
    const panel = node("div", undefined, "group");
    const controls = node("div", undefined, "controls");
    for (const b of g.buttons) {
      const control = button("", () => select(tab, g, b), "control");
      control.classList.toggle("big", !!b.large);
      control.classList.toggle("selected", item === b);
      const src = images.get(b.image);
      const icon = src ? node("img") : node("span", glyphs[b.imageMso] || "◇");
      if (src) { icon.src = src; icon.alt = ""; }
      icon.className = "glyph";
      icon.setAttribute("aria-hidden", "true");
      control.title = b.command || "Choose a Python macro";
      control.append(icon, node("span", b.label || "Untitled button"));
      controls.append(control);
    }
    const label = button(g.label || "Untitled group", () => select(tab, g), "groupName");
    label.classList.toggle("selected", group === g && !item);
    panel.append(controls, label);
    groups.append(panel);
  }
}
function renderRibbon() {
  renderPreview();
  const {tab, group, item} = selection;
  const current = selectedNode();
  $("properties").hidden = !current;
  $("propertiesTitle").textContent = current ? (item ? "Button" : group ? "Group" : "Tab") + " properties" : "Properties";
  $("addGroup").disabled = !tab;
  $("addButton").disabled = !group;
  $("removeNode").disabled = !current;
  $("buttonFields").hidden = !item;
  if (!current) return;
  $("label").value = current.label;
  if (item) {
    const options = macros.map(m => new Option(m.label + " (" + m.target + ")", m.target));
    if (item.command && !macros.some(m => m.target === item.command))
      options.push(new Option(item.command + " (saved binding)", item.command));
    $("command").replaceChildren(new Option(macros.length ? "Choose a macro" : "No @macro functions found", ""), ...options);
    $("command").value = item.command;
    $("officeIcon").value = item.imageMso || "";
    $("large").checked = !!item.large;
    $("clearImage").hidden = !item.image;
    $("officeIcon").disabled = !!item.image;
  }
}
$("addTab").onclick = () => {
  const tab = {label:"New tab", groups:[]};
  project.tabs.push(tab); select(tab); $("label").select();
};
$("addGroup").onclick = () => {
  const group = {label:"New group", buttons:[]};
  selection.tab.groups.push(group); select(selection.tab, group); $("label").select();
};
$("addButton").onclick = () => {
  const item = {label:"New button", command:"", large:true, imageMso:"MacroPlay"};
  selection.group.buttons.push(item); select(selection.tab, selection.group, item); $("label").select();
};
$("removeNode").onclick = () => {
  const {tab, group, item} = selection;
  const list = item ? group.buttons : group ? tab.groups : project.tabs;
  list.splice(list.indexOf(selectedNode()), 1);
  select(group ? tab : project.tabs[0], item ? group : null);
};
$("command").onchange = () => {
  const item = selection.item;
  const macro = macros.find(m => m.target === $("command").value);
  if (macro && (!item.label.trim() || item.label === "New button" || item.label === macros.find(m => m.target === item.command)?.label))
    item.label = macro.label;
  item.command = $("command").value;
  item.macro = true;
  renderRibbon();
};
for (const [id, key] of [["label","label"], ["officeIcon","imageMso"], ["large","large"]]) {
  $(id).oninput = () => {
    const field = $(id);
    selectedNode()[key] = id === "large" ? field.checked : field.value;
    renderPreview();
  };
}
$("importImage").onclick = () => run(async () => {
  const result = await host("importImage");
  if (result) { images.set(result.name, result.src); selection.item.image = result.name; renderRibbon(); }
});
$("clearImage").onclick = () => { selection.item.image = ""; renderRibbon(); };

function addPackages(names) {
  for (const name of names) if (!project.packages.includes(name)) project.packages.push(name);
  renderPackages();
}
function renderPackages() {
  for (const [id, specs] of [["available", installed], ["included", project.packages]]) {
    const list = $(id);
    list.replaceChildren();
    for (const name of specs) {
      const include = id === "available";
      const action = button(include ? "Include" : "Remove", () => {
        if (include) addPackages([name]);
        else { project.packages = project.packages.filter(p => p !== name); renderPackages(); }
      });
      action.disabled = include && project.packages.includes(name);
      action.setAttribute("aria-label", action.textContent + " " + name);
      const row = node("div");
      row.append(node("span", name), action);
      list.append(row);
    }
    if (!specs.length) list.append(node("p", id === "included" ? "No packages included." : "No installed packages found."));
  }
}
$("addSpec").onclick = () => {
  const name = $("spec").value.trim();
  if (!name) return;
  addPackages([name]);
  $("spec").value = "";
};
$("spec").onkeydown = e => { if (e.key === "Enter") $("addSpec").click(); };
$("wheel").onclick = () => run(async () => addPackages(await host("wheels") || []));
run(() => host("ready"));
