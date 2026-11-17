import { chromium } from 'playwright';
import fs from 'node:fs';
const OUT = '/tmp/claude-1000/-home-local/4bd4dbda-ccf9-4e6b-a738-bbbc9b7bbfcc/scratchpad/shots2/out/';
const BASE = 'http://127.0.0.1:5175';
const SESS = JSON.parse(fs.readFileSync(new URL('./sess.json', import.meta.url)));
const only = process.argv.slice(2);
const b = await chromium.launch({ args: ['--use-fake-ui-for-media-stream', '--use-fake-device-for-media-stream', '--use-file-for-fake-video-capture=/tmp/claude-1000/-home-local/4bd4dbda-ccf9-4e6b-a738-bbbc9b7bbfcc/scratchpad/shots2/face.y4m'] });
async function page(ls, { w = 1440, h = 900, gw } = {}) {
  const ctx = await b.newContext({ viewport: { width: w, height: h }, deviceScaleFactor: 1.5, locale: 'vi-VN', timezoneId: 'Asia/Ho_Chi_Minh', permissions: ['camera'] });
  await ctx.addInitScript((kv) => { if (!sessionStorage.getItem('_init')) { sessionStorage.setItem('_init', 1); for (const [k, v] of Object.entries(kv)) localStorage.setItem(k, v); } }, ls);
  // AI Gateway hiển thị là aigate.is-tech.vn nhưng mọi yêu cầu được chuyển về máy giả lập cục bộ (không gọi máy chủ thật)
  if (gw) await ctx.route(gw + '/**', async (r) => { const q = r.request(); const res = await r.fetch({ url: q.url().replace(gw, BASE) }); await r.fulfill({ response: res }); });
  const p = await ctx.newPage();
  p.on('pageerror', e => console.log('ERR', e.message));
  return p;
}
const shot = (p, n) => p.screenshot({ path: OUT + n + '.png' });
const S = { ...SESS };
const jobs = {
  async tl1_01_login() { const p = await page({ shot_web: '1' }); await p.goto(BASE + '/'); await p.waitForTimeout(2500); await shot(p, 'tl1_01_login'); },
  async tl1_02_faceid() { const p = await page({ shot_web: '1', shot_view: 'FACEID' }); await p.goto(BASE + '/'); await p.waitForTimeout(3500); await shot(p, 'tl1_02_faceid'); },
  async tl1_03_workspace() { const p = await page(S); await p.goto(BASE + '/'); await p.waitForTimeout(2500); await shot(p, 'tl1_03_workspace'); },
  async tl1_04_public() {
    const p = await page(S); await p.goto(BASE + '/'); await p.waitForTimeout(2000);
    await p.getByText('Mở Kho phân vùng').click(); await p.waitForTimeout(1200);
    const box = p.getByLabel('Tìm phân vùng Public'); await box.fill('LAB'); await p.waitForTimeout(1500);
    await shot(p, 'tl1_04_public');
  },
  async tl1_05_history() {
    const p = await page(S); await p.goto(BASE + '/history'); await p.waitForTimeout(1500);
    await p.getByText('Quy chế thực tập năm 2026 quy định').first().click(); await p.waitForTimeout(1500); await shot(p, 'tl1_05_history');
  },
  async tl1_06_identity() {
    const p = await page(S); await p.goto(BASE + '/identity'); await p.waitForTimeout(1500);
    await p.getByText('Nguyễn Văn An').first().click(); await p.waitForTimeout(1500); await shot(p, 'tl1_06_identity');
  },
  async tl1_07_answer() {
    const p = await page(S); await p.goto(BASE + '/'); await p.waitForTimeout(2000);
    await p.selectOption('#sidebar-store-select', 'Public'); await p.waitForTimeout(1500);
    const ta = p.locator('textarea').first(); await ta.fill('Quy chế thực tập năm 2026 quy định những gì?'); await ta.press('Enter');
    await p.waitForTimeout(5000); await shot(p, 'tl1_07_answer');
  },
  async tl1_08_comm() {
    const p = await page(S); await p.goto(BASE + '/comm'); await p.waitForTimeout(1500);
    await p.getByText('Nhóm Thực tập 2026').first().click(); await p.waitForTimeout(2000); await shot(p, 'tl1_08_comm');
  },
  async tl2_01_login_gw() {
    const gw = 'https://aigate.is-tech.vn';
    const p = await page({ tara_gateway: gw, tara_gateways: JSON.stringify([{ url: gw }]) }, { gw }); await p.goto(BASE + '/'); await p.waitForTimeout(2500); await shot(p, 'tl2_01_login_gw');
  },
  async tl2_02_server() { const p = await page(S); await p.goto(BASE + '/analytics'); await p.waitForTimeout(2500); await shot(p, 'tl2_02_server'); },
  async tl2_03_platform() { const p = await page(S, { h: 1100 }); await p.goto(BASE + '/analytics?tab=platform'); await p.waitForTimeout(3000); await shot(p, 'tl2_03_platform'); },
  async tl2_04_clouds() {
    const p = await page(S); await p.goto(BASE + '/identity'); await p.waitForTimeout(1500);
    await p.getByText('Phân vùng', { exact: true }).first().click(); await p.waitForTimeout(1800); await shot(p, 'tl2_04_clouds');
  },
};
for (const [n, f] of Object.entries(jobs)) if (!only.length || only.includes(n)) { try { await f(); console.log('ok', n); } catch (e) { console.log('FAIL', n, e.message.split('\n')[0]); } }
await b.close();
