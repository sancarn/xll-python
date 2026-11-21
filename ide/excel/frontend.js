import { ContainerModule } from '@theia/core/shared/inversify';
import { FrontendApplicationContribution } from '@theia/core/lib/browser/frontend-application-contribution';
import { Saveable } from '@theia/core/lib/browser/saveable';
import { FileNavigatorContribution } from '@theia/navigator/lib/browser/navigator-contribution';
import { CommandContribution } from '@theia/core/lib/common/command';

const pending = new Map();
let sequence = 0;
window.chrome.webview.addEventListener('message', ({ data }) => {
    const request = pending.get(data.id);
    if (!request) return;
    pending.delete(data.id);
    if (data.error) request.reject(new Error(data.error));
    else request.resolve(data.value);
});

export default new ContainerModule(bind => {
    bind(FrontendApplicationContribution).toDynamicValue(({ container }) => ({
        async onDidInitializeLayout(app) {
            window.chrome.webview.addEventListener('message', async ({ data }) => {
                if (data.type === 'closing') {
                    let allow = false;
                    try { allow = !!await Saveable.confirmSaveBeforeClose(app.shell.widgets, []); }
                    finally { window.chrome.webview.postMessage({ operation: 'closeResult', allow }); }
                }
            });
            window.chrome.webview.postMessage({ operation: 'ready' });
            await container.get(FileNavigatorContribution).openView({ reveal: true });
        }
    }));
    bind(CommandContribution).toConstantValue({
        registerCommands(commands) {
            commands.registerCommand({ id: 'excel.host' }, {
                execute: (operation, args = {}) => new Promise((resolve, reject) => {
                    const id = ++sequence;
                    pending.set(id, { resolve, reject });
                    window.chrome.webview.postMessage({ id, operation, args });
                })
            });
        }
    });
});
