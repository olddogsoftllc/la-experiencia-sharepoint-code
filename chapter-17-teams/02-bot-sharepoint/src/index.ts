// src/index.ts — Bot de Teams que lee SharePoint via Graph (cap17)
import 'dotenv/config';
import * as restify from 'restify';
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  ActivityHandler,
} from 'botbuilder';
import { getGraphClient, IGraphSiteCollection, IGraphDrive, IGraphDriveItem } from './graphAuth';

// 1) Credenciales del bot = App Registration de Entra ID
const appId = process.env.MICROSOFT_APP_ID;
const appPassword = process.env.MICROSOFT_APP_PASSWORD;
if (!appId || !appPassword) {
  throw new Error('MICROSOFT_APP_ID / MICROSOFT_APP_PASSWORD are missing. Copy .env.example to .env.');
}

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: appId,
  MicrosoftAppPassword: appPassword,
  MicrosoftAppType: 'MultiTenant',
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

// 2) Handler
class SharePointBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const text = (context.activity.text || '').trim().toLowerCase();
      const graph = getGraphClient();

      try {
        if (text === 'sites' || text === 'list sites') {
          const sites = await graph.api('/sites?$top=10').get() as IGraphSiteCollection;
          const list = sites.value
            .map((s) => `• ${s.displayName} — ${s.webUrl}`)
            .join('\n');
          await context.sendActivity(
            sites.value.length
              ? `Sites (${sites.value.length}):\n${list}`
              : 'No sites found.'
          );
        } else if (text.startsWith('docs ')) {
          const siteName = text.slice(5).trim();
          const search = await graph.api(`/sites?$search="${siteName}"`).get() as IGraphSiteCollection;
          if (!search.value.length) {
            await context.sendActivity(`I didn't find site "${siteName}".`);
          } else {
            const site = search.value[0];
            const drives = await graph.api(`/sites/${site.id}/drives`).get() as { value: IGraphDrive[] };
            if (!drives.value.length) {
              await context.sendActivity(`"${site.displayName}" has no libraries.`);
            } else {
              const items = await graph
                .api(`/drives/${drives.value[0].id}/root/children?$top=10`)
                .get() as { value: IGraphDriveItem[] };
              const list = items.value.map((d) => `• ${d.name}`).join('\n');
              await context.sendActivity(
                `Documents in "${site.displayName}":\n${list || '(empty)'}`
              );
            }
          }
        } else if (text === 'help' || text === '') {
          await context.sendActivity(
            'Commands:\n• `sites` — list sites\n• `docs <site>` — documents in a site\n• `help` — this help'
          );
        } else {
          await context.sendActivity(
            `I didn't recognize "${text}". Type \`help\` to see commands.`
          );
        }
      } catch (err: unknown) {
        const e = err as { statusCode?: number; message?: string };
        await context.sendActivity(`Error Graph: ${e.statusCode ?? '?'} ${e.message ?? err}`);
      }
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const members = context.activity.membersAdded ?? [];
      for (const member of members) {
        if (member.id !== context.activity.recipient?.id) {
          await context.sendActivity(
            `Hola. Soy el bot de SharePoint. Escribe \`help\` para ver comandos.`
          );
        }
      }
      await next();
    });
  }
}

const bot = new SharePointBot();

// 3) restify server with the messaging endpoint
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.PORT ? Number(process.env.PORT) : 3978, () => {
  console.log(`Bot listening on ${server.url}/api/messages`);
});

server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async (context: TurnContext) => {
    await bot.run(context);
  });
});

server.get('/', (_req, res) => {
  res.send(200, 'SharePoint Teams Bot — endpoint: /api/messages');
});