const vscode = require('vscode');
const path = require('node:path');
const crypto = require('node:crypto');

const type = 'excel-python';
const host = (operation, args) => vscode.commands.executeCommand('excel.host', operation, args);

async function activeFile() {
    const doc = vscode.window.activeTextEditor?.document;
    if (!doc || doc.languageId !== 'python') throw new Error('Open a Python file in the editor.');
    if (doc.isUntitled) throw new Error('Save the Python file in a local folder first.');
    if (doc.isDirty && !await doc.save()) throw new Error('The Python file could not be saved.');
    const uri = doc.uri;
    if (!uri || uri.scheme !== 'file') throw new Error('Save the file in a local folder.');
    return uri.fsPath;
}

exports.activate = async context => {
    const output = vscode.window.createOutputChannel('Excel Python');
    const status = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 10);
    context.subscriptions.push(output, status);
    const environment = await host('environment');
    status.text = '$(plug) Excel ' + environment.pid + '  |  Python ' + environment.version;
    status.tooltip = environment.python;
    status.show();

    function reportError(error) {
        output.appendLine(error.message);
        vscode.window.showErrorMessage(error.message);
    }
    context.subscriptions.push(vscode.workspace.onDidSaveTextDocument(doc => {
        if (doc.uri.scheme === 'file' && path.extname(doc.uri.fsPath).toLowerCase() === '.py')
            host('refreshSaved', { path: doc.uri.fsPath }).catch(reportError);
    }));
    let folderChanges = Promise.resolve();
    context.subscriptions.push(vscode.workspace.onDidChangeWorkspaceFolders(({ added, removed }) => {
        const workspace = vscode.workspace.workspaceFile?.fsPath || null;
        for (const [folders, isAddition] of [[removed, false], [added, true]]) {
            for (const { uri } of folders) {
                if (uri.scheme !== 'file') continue;
                folderChanges = folderChanges.then(() => host('workspaceFolder', { path: uri.fsPath, added: isAddition, workspace }))
                    .catch(reportError);
            }
        }
    }));

    const commands = vscode.commands;
    const debug = vscode.debug;
    let running = false;
    async function run(file) {
        if (running) throw new Error('A Python file is already running.');
        running = true;
        output.show(true);
        output.appendLine('Running ' + file);
        try {
            const text = await host('run', { path: file });
            if (text) output.append(text);
        } catch (err) {
            output.appendLine(err.message);
            throw err;
        } finally { running = false; }
    }
    function command(id, handler) {
        context.subscriptions.push(commands.registerCommand(id, handler));
    }
    function config(currentFile) {
        return { type, request: 'attach', currentFile,
            name: currentFile ? 'Debug Current File in Excel' : 'Debug Worksheet Functions' };
    }
    command('excel.run', async () => run(await activeFile()));
    command('excel.debugFile', () => debug.startDebugging(undefined, config(true)));
    command('excel.attach', () => debug.startDebugging(undefined, config(false)));
    command('excel.refresh', async () => {
        if (!await vscode.workspace.saveAll()) throw new Error('Save the modified files before refreshing.');
        await host('refresh');
        output.appendLine('Refreshed Excel worksheet functions.');
        output.show(true);
    });
    const sessions = new Map();
    context.subscriptions.push(debug.registerDebugConfigurationProvider(type, {
        provideDebugConfigurations: () => [config(true), config(false)],
        async resolveDebugConfiguration(folder, input) {
            if (sessions.size) throw new Error('An Excel debug session is already active.');
            const cfg = { ...config(true), ...input };
            cfg.file = cfg.currentFile ? await activeFile() : undefined;
            cfg.listen = { host: '127.0.0.1', port: 0 };
            cfg.justMyCode = false;
            cfg.subProcess = false;
            return cfg;
        }
    }));
    context.subscriptions.push(debug.registerDebugAdapterDescriptorFactory(type, {
        createDebugAdapterDescriptor(session) {
            const token = crypto.randomBytes(32).toString('hex');
            sessions.set(session.id, token);
            return new vscode.DebugAdapterExecutable(environment.python,
                ['-I', path.join(environment.support, 'debugpy', 'adapter'), '--server-access-token', token]);
        }
    }));
    context.subscriptions.push(debug.registerDebugAdapterTrackerFactory(type, {
        createDebugAdapterTracker(session) {
            let connected;
            return {
                onDidSendMessage(msg) {
                    // Excel owns the debuggee. Even an explicit Stop must detach.
                    if (msg.type === 'response' && msg.command === 'initialize' && msg.body) {
                        msg.body.supportsTerminateRequest = false;
                        msg.body.supportTerminateDebuggee = false;
                    }
                    if (msg.event === 'debugpyWaitingForServer') {
                        connected = host('connect', { port: msg.body.port, token: sessions.get(session.id) });
                        connected.catch(err => { vscode.window.showErrorMessage(err.message); debug.stopDebugging(session); });
                    }
                    if (msg.type === 'response' && msg.command === 'attach' && msg.success && session.configuration.file) {
                        connected.then(() => run(session.configuration.file)).catch(err => output.appendLine(err.message))
                            .finally(() => debug.stopDebugging(session));
                    }
                },
                onWillReceiveMessage(msg) {
                    if (msg.command === 'terminate') msg.command = 'disconnect';
                    if (msg.command === 'disconnect') msg.arguments = { ...msg.arguments, terminateDebuggee: false };
                },
                onWillStopSession() { sessions.delete(session.id); }
            };
        }
    }));
    context.subscriptions.push(debug.onDidTerminateDebugSession(session => sessions.delete(session.id)));
};
