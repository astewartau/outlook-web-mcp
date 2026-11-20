#!/usr/bin/env node
// Call one tool on the server, for testing: node scripts/call.mjs <tool> '<json args>'   (or: list)
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
const c = new Client({ name: 'call', version: '1' });
await c.connect(new StdioClientTransport({ command: 'node', args: [new URL('../src/server.mjs', import.meta.url).pathname], env: process.env }));
if (process.argv[2] === 'list') console.log((await c.listTools()).tools.map(t => t.name).join('\n'));
else {
  const r = await c.callTool({ name: process.argv[2], arguments: JSON.parse(process.argv[3] || '{}') }, undefined, { timeout: 600000 });
  for (const x of r.content) console.log(x.type === 'text' ? x.text : `[${x.type}, ${x.data?.length} base64 chars]`);
  if (r.isError) process.exitCode = 1;
}
await c.close();
