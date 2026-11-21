import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse, modify, applyEdits } from 'jsonc-parser';

function writeAtomic(file, content) {
    const temp = file + '.' + process.pid + '.tmp';
    try {
        fs.writeFileSync(temp, content);
        fs.renameSync(temp, file);
    } finally { fs.rmSync(temp, { force: true }); }
}

const name = 'Excel Python.theia-workspace';

function read(file) {
    return fs.readFileSync(file, 'utf8');
}

function folderPath(folder, directory) {
    return /^file:/i.test(folder.path) ? fileURLToPath(folder.path) : path.resolve(directory, folder.path);
}

function legacyWorkspace(directory, seedDir) {
    const exact = path.join(seedDir, name);
    if (fs.existsSync(exact)) return exact;
    // An add-in update can change the embedded interpreter paths and thus the old workspace hash.
    const recent = path.join(process.env.THEIA_CONFIG_DIR, 'recentworkspace.json');
    if (!fs.existsSync(recent)) return;
    for (const uri of parse(read(recent))?.recentRoots || []) {
        if (typeof uri !== 'string' || !uri.startsWith('file:///')) continue;
        const file = fileURLToPath(uri);
        const relative = path.relative(directory, file).split(path.sep);
        if (relative.length === 2 && /^[0-9A-F]{24}$/.test(relative[0]) && relative[1] === name && fs.existsSync(file))
            return file;
    }
}

// Theia edits the workspace itself. Apply only changes to Excel's generated defaults on reopen.
export function prepareWorkspace(defaultsFile) {
    const seedDir = path.dirname(defaultsFile);
    const directory = path.dirname(seedDir);
    const workspace = path.join(directory, name);
    const exists = fs.existsSync(workspace);
    const source = exists ? workspace : legacyWorkspace(directory, seedDir) || defaultsFile;
    const original = read(source);
    let content = original;
    const errors = [];
    const data = parse(content, errors, { allowTrailingComma: true });
    if (errors.length || !Array.isArray(data?.folders)) throw new Error('Cannot read workspace ' + source);
    const baseline = path.join(directory, 'applied-defaults.json');
    const defaults = JSON.parse(read(defaultsFile));
    const next = { folders: defaults.folders.map(folder => folder.path), settings: defaults.settings };
    const previous = fs.existsSync(baseline) ? JSON.parse(read(baseline).replace(/^\uFEFF/, '')) : { ...next };
    if (!exists && path.dirname(source) !== seedDir) {
        // Preserve custom settings while updating references to a previous embedded runtime or support directory.
        for (const [setting, value] of Object.entries(data.settings || {})) {
            const values = Array.isArray(value) ? value : [value];
            if (values.some(item => typeof item === 'string' && /[\\\\/]((runtime|support)-[0-9a-f]{16})[\\\\/]?/i.test(item)))
                previous.settings = { ...previous.settings, [setting]: value };
        }
    }
    const update = (keys, value) => {
        content = applyEdits(content, modify(content, keys, value, {
            formattingOptions: { tabSize: 2, insertSpaces: true }
        }));
    };
    const paths = seed => new Set(seed.folders.map(folder => path.resolve(folder).toLowerCase()));
    const before = paths(previous);
    const after = paths(next);
    const kept = new Set();
    const folders = [];
    for (const folder of data.folders) {
        const absolute = folderPath(folder, path.dirname(source));
        const id = absolute.toLowerCase();
        if (before.has(id) && !after.has(id)) continue;
        kept.add(id);
        folders.push(exists ? folder : { ...folder, path: absolute });
    }
    for (const folder of defaults.folders) {
        const id = folderPath(folder, directory).toLowerCase();
        if (!before.has(id) && !kept.has(id)) folders.push(folder);
    }
    if (JSON.stringify(folders) !== JSON.stringify(data.folders)) update(['folders'], folders);
    for (const [setting, value] of Object.entries(defaults.settings)) {
        const current = JSON.stringify(data.settings?.[setting]);
        if (current === JSON.stringify(previous.settings[setting]) && current !== JSON.stringify(value))
            update(['settings', setting], value);
    }
    if (!exists || content !== original) writeAtomic(workspace, content);
    if (!fs.existsSync(baseline) || JSON.stringify(previous) !== JSON.stringify(next))
        writeAtomic(baseline, JSON.stringify(next));
}
