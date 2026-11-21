import { ContainerModule } from '@theia/core/shared/inversify';
import { BackendApplicationContribution, EarlyExpressMiddleware } from '@theia/core/lib/node/backend-application';
import { WsRequestValidatorContribution } from '@theia/core/lib/node/ws-request-validators';

import { prepareWorkspace } from './workspace';

const token = process.env.XLL_IDE_TOKEN;
if (!token) throw new Error('Missing Excel IDE session');

function authorized(req) {
    const host = process.env.XLL_IDE_HOST + ':' + req.socket.localPort;
    return req.headers.host === host &&
        (!req.headers.origin || req.headers.origin === 'http://' + host) &&
        req.headers.cookie?.split(/;\s*/).includes('xll_' + req.socket.localPort + '=' + token);
}

export default new ContainerModule(bind => {
    bind(BackendApplicationContribution).toDynamicValue(({ container }) => ({
        initialize() {
            prepareWorkspace(process.env.XLL_IDE_DEFAULTS);
            container.get(EarlyExpressMiddleware).handlers.push((req, res, next) => {
                if (authorized(req)) next();
                else res.sendStatus(403);
            });
        },
        onStart(server) {
            const announce = () => process.stdout.write('XLL_IDE_PORT=' + server.address().port + '\n');
            if (server.listening) announce();
            else server.once('listening', announce);
        }
    }));
    bind(WsRequestValidatorContribution).toConstantValue({ allowWsUpgrade: authorized });
});
