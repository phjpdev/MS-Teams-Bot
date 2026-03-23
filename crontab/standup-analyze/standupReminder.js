#!/usr/bin/env node
/**
 * Standup reminder cron script.
 * Call the bot's /api/standup/trigger endpoint (ask or analyze).
 *
 * Usage:
 *   node standupReminder.js           → action=ask (default)
 *   node standupReminder.js analyze  → action=analyze
 *
 * Env (set in Azure App Settings or server):
 *   BOT_BASE_URL             - e.g. https://your-app.azurewebsites.net (no trailing slash)
 *   STANDUP_TRIGGER_SECRET   - same value as in the bot app settings
 */

const https = require('https');
const http = require('http');

const action = (process.argv[2] || 'ask').toLowerCase() === 'analyze' ? 'analyze' : 'ask';
const baseUrl = process.env.BOT_BASE_URL
  || (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : null)
  || 'http://localhost:' + (process.env.PORT || '8080');
const secret = process.env.STANDUP_TRIGGER_SECRET || '';

const url = new URL('/api/standup/trigger', baseUrl);
url.searchParams.set('action', action);
if (secret) url.searchParams.set('secret', secret);

const client = url.protocol === 'https:' ? https : http;

const req = client.get(url.toString(), (res) => {
  let body = '';
  res.on('data', (chunk) => { body += chunk; });
  res.on('end', () => {
    if (res.statusCode === 200) {
      console.log(`[${new Date().toISOString()}] Standup ${action} OK`);
    } else {
      console.error(`[${new Date().toISOString()}] Standup ${action} HTTP ${res.statusCode}: ${body || res.statusMessage}`);
    }
  });
});

req.on('error', (err) => {
  console.error(`[${new Date().toISOString()}] Standup ${action} error:`, err.message);
  process.exit(1);
});

req.setTimeout(30000, () => {
  req.destroy();
  console.error(`[${new Date().toISOString()}] Standup ${action} timeout`);
  process.exit(1);
});
