"use strict";
const byId = id => document.getElementById(id);
const source = byId("source"), highlight = byId("highlight"), transcript = byId("transcript");
const run = byId("run"), clear = byId("clear"), suggestions = byId("suggestions"), status = byId("status");
const hint = status.textContent, history = [], pending = new Map();
let sequence = 0, historyIndex = 0, draft = "", completionTimer, revision = 0, selected = 0;

function element(tag, className, text) {
  const node = document.createElement(tag);
  node.className = className;
  node.textContent = text;
  return node;
}

function request(action, args, onEvents) {
  const id = ++sequence;
  return new Promise((resolve, reject) => {
    pending.set(id, {resolve, reject, onEvents});
    window.chrome.webview.postMessage({id, action, args});
  });
}
window.chrome.webview.addEventListener("message", ({data}) => {
  const call = pending.get(data.id);
  if (!call) return;
  if (data.events) { call.onEvents(data.events); return; }
  pending.delete(data.id);
  if (data.error) call.reject(new Error(data.error));
  else call.resolve(data.value);
});

const keywords = new Set(("False None True and as assert async await break class continue def del elif else " +
  "except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield " +
  "match case type").split(" "));
const tokens = /#[^\n]*|(?:[rubf]{0,2})(?:'''(?:\\[\s\S]|(?!''')[\s\S])*(?:'''|$)|"""(?:\\[\s\S]|(?!""")[\s\S])*(?:"""|$)|'(?:\\.|[^'\\\n])*(?:'|(?=\n|$))|"(?:\\.|[^"\\\n])*(?:"|(?=\n|$)))|\b(?:0[xob][\da-f_]+|\d[\d_]*(?:\.[\d_]*)?(?:e[+-]?[\d_]+)?j?)|\b[a-z_]\w*/gi;

function colorize(node, code) {
  const fragment = document.createDocumentFragment();
  let end = 0;
  for (const match of code.matchAll(tokens)) {
    const word = match[0];
    const style = word.startsWith("#") ? "comment" : /['"]/.test(word) ? "str" :
      /^\d/.test(word) ? "num" : keywords.has(word) ? "kw" : "";
    if (!style) continue;
    fragment.append(code.slice(end, match.index), element("span", style, word));
    end = match.index + word.length;
  }
  fragment.append(code.slice(end));
  node.replaceChildren(fragment);
}

function paint() {
  colorize(highlight, source.value + "\n");
  source.style.height = "auto";
  source.style.height = source.scrollHeight + "px";
  highlight.scrollTop = source.scrollTop;
}
source.addEventListener("scroll", () => { highlight.scrollTop = source.scrollTop; });
window.addEventListener("resize", paint);

function valueNode(value) {
  if (!value.handle) return document.createTextNode(value.text);
  const group = element("span", "", "");
  const toggle = element("button", "inspect", "▸ " + value.text);
  toggle.setAttribute("aria-expanded", "false");
  const children = element("span", "children", "");
  let loaded = false, offset = 0;
  async function page() {
    try {
      const result = await request("children", {handle: value.handle, offset});
      for (const item of result.items) {
        const row = element("span", "child", "");
        row.append(element("span", "label", item.label + ": "), valueNode(item));
        children.append(row);
      }
      offset += result.items.length;
      if (!offset) children.append("(no fields)");
      if (result.more) {
        const more = element("button", "more", "Show more");
        more.onclick = () => { more.remove(); page(); };
        children.append(more);
      }
    } catch (error) { children.append(element("span", "error", error.message)); }
  }
  toggle.onclick = () => {
    const open = !children.isConnected;
    if (open) group.append(children); else children.remove();
    toggle.setAttribute("aria-expanded", String(open));
    toggle.textContent = (open ? "▾ " : "▸ ") + value.text;
    if (open && !loaded) { loaded = true; page(); }
  };
  group.append(toggle);
  return group;
}

function dismiss() {
  revision++;
  clearTimeout(completionTimer);
  suggestions.hidden = true;
  suggestions.replaceChildren();
}

function tokenAtCaret() {
  const before = source.value.slice(0, source.selectionStart);
  return before.match(/[\p{ID_Continue}.]+$/u)?.[0] || "";
}

async function complete(explicit) {
  const version = revision, token = tokenAtCaret(), caret = source.selectionStart;
  if (source.readOnly || caret !== source.selectionEnd || (!explicit && !token)) return;
  try {
    const names = await request("complete", {token});
    if (version !== revision || caret !== source.selectionStart || document.activeElement !== source) return;
    const prefix = token.split(".").pop();
    if (!names.length || (names.length === 1 && names[0] === prefix)) return;
    selected = 0;
    suggestions.replaceChildren(...names.map((name, i) => {
      const option = element("button", "", name);
      option.setAttribute("role", "option");
      option.setAttribute("aria-selected", String(i === 0));
      option.tabIndex = -1;
      option.onmousedown = event => event.preventDefault();
      option.onclick = () => {
        source.setRangeText(name, caret - prefix.length, caret, "end");
        dismiss(); paint(); source.focus();
      };
      return option;
    }));
    suggestions.hidden = false;
    suggestions.scrollTop = 0;
  } catch (error) { if (version === revision) status.textContent = error.message; }
}
source.addEventListener("input", () => {
  dismiss(); paint();
  completionTimer = setTimeout(() => complete(false), 160);
});
source.addEventListener("blur", dismiss);
source.addEventListener("click", dismiss);

function newline() {
  const line = source.value.slice(0, source.selectionStart).split("\n").pop();
  const indent = line.match(/^\s*/)[0] + (line.trimEnd().endsWith(":") ? "    " : "");
  source.setRangeText("\n" + indent, source.selectionStart, source.selectionEnd, "end");
  paint();
}

async function execute() {
  const code = source.value;
  if (!code.trim() || source.readOnly) return;
  dismiss();
  source.readOnly = run.disabled = clear.disabled = true;
  status.textContent = "Running…";
  const entry = element("div", "entry", ""), command = element("div", "command", "");
  const echo = element("span", "", ""), output = element("div", "output", "");
  colorize(echo, code);
  command.append(element("span", "prompt", ">>>"), echo);
  entry.append(command, output);
  transcript.append(entry);
  transcript.scrollTop = transcript.scrollHeight;
  try {
    const result = await request("run", {source: code}, events => {
      const follow = transcript.scrollHeight - transcript.scrollTop - transcript.clientHeight < 12;
      for (const {kind, value} of events) {
        if (kind === "result" || kind === "object") {
          const node = element("span", kind, "");
          node.append(valueNode(value));
          output.append(node);
        } else output.append(element("span", kind === "stderr" ? "error" : "", value));
      }
      if (follow) transcript.scrollTop = transcript.scrollHeight;
    });
    if (result.incomplete) {
      entry.remove();
      source.setSelectionRange(code.length, code.length);
      newline();
    } else {
      for (const formula of result.worksheetFormulas || [])
        output.append(element("div", "", "Worksheet formula: " + formula));
      if (history.at(-1) !== code) history.push(code);
      if (history.length > 200) history.shift();
      if (transcript.childElementCount > 200) transcript.firstChild.remove();
      historyIndex = history.length; draft = ""; source.value = "";
    }
  } catch (error) { output.append(element("span", "error", error.message)); }
  finally {
    source.readOnly = run.disabled = clear.disabled = false;
    status.textContent = hint; paint(); source.focus();
  }
}
run.onclick = execute;
clear.onclick = async () => {
  try { await request("clear", {}); transcript.replaceChildren(); }
  catch (error) { status.textContent = error.message; }
  source.focus();
};

source.addEventListener("keydown", event => {
  if (event.isComposing || source.readOnly) return;
  const key = event.key;
  if (key === "Enter") {
    event.preventDefault(); dismiss();
    if (event.shiftKey) newline(); else execute();
  } else if (event.ctrlKey && event.code === "Space") {
    event.preventDefault(); dismiss(); complete(true);
  } else if (key === "Tab" && !event.shiftKey) {
    event.preventDefault();
    if (!suggestions.hidden) suggestions.children[selected].click();
    else { source.setRangeText("    ", source.selectionStart, source.selectionEnd, "end"); paint(); }
  } else if (key === "Escape") dismiss();
  else if (key === "ArrowUp" || key === "ArrowDown") {
    const step = key === "ArrowUp" ? -1 : 1;
    if (!suggestions.hidden) {
      event.preventDefault();
      suggestions.children[selected].setAttribute("aria-selected", "false");
      selected = (selected + step + suggestions.children.length) % suggestions.children.length;
      suggestions.children[selected].setAttribute("aria-selected", "true");
      suggestions.children[selected].scrollIntoView({block: "nearest"});
    } else if (!event.shiftKey && (step < 0 ? !source.value.slice(0, source.selectionStart).includes("\n")
                                           : !source.value.slice(source.selectionEnd).includes("\n"))) {
      event.preventDefault(); dismiss();
      if (historyIndex === history.length) draft = source.value;
      historyIndex = Math.max(0, Math.min(history.length, historyIndex + step));
      source.value = historyIndex === history.length ? draft : history[historyIndex];
      paint();
    }
  } else if (["ArrowLeft", "ArrowRight", "Home", "End"].includes(key)) dismiss();
});
paint();
