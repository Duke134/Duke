#!/usr/bin/env node
// 카드뉴스 렌더러: content.json → out/card_01.png … + contact_sheet.png + caption.txt
// 사용: node render.js content.json [출력폴더]   (Playwright + Chromium 필요)
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const [,, contentPath, outDirArg] = process.argv;
if (!contentPath) { console.error('사용: node render.js content.json [출력폴더]'); process.exit(1); }
const c = JSON.parse(fs.readFileSync(contentPath, 'utf8'));
const outDir = path.resolve(outDirArg || 'out');
fs.mkdirSync(outDir, { recursive: true });

const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const rich = s => esc(s).replace(/&lt;br&gt;/g, '<br>').replace(/&lt;em&gt;/g, '<em>').replace(/&lt;\/em&gt;/g, '</em>');
const total = c.items.length + 2;
const foot = i => `<div class="foot"><span>${esc(c.brand)}</span><span>${i} / ${total}</span></div>`;

const sections = [];
sections.push(`<section class="card cover"><div class="tag">${esc(c.tag)}</div><h1>${rich(c.cover.title)}</h1><p>${rich(c.cover.subtitle)}</p>${foot(1)}</section>`);
c.items.forEach((it, idx) => {
  const n = String(idx + 1).padStart(2, '0');
  sections.push(`<section class="card"><div class="num">${n}<small>${esc(it.label)}</small></div><h2>${rich(it.title)}</h2><ul class="body">${it.points.map(p => `<li>${esc(p)}</li>`).join('')}</ul><div class="key">핵심: ${esc(it.key)}</div>${foot(idx + 2)}</section>`);
});
sections.push(`<section class="card last"><h2>${rich(c.last.title)}</h2><div class="quote">${esc(c.last.experience)}<span class="who">— ${esc(c.last.who)}</span></div><div class="cta">${esc(c.last.cta)}<span>${esc(c.last.cta_sub)}</span></div>${foot(total)}</section>`);

const cssPath = path.join(__dirname, 'template.css');
const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><link rel="stylesheet" href="file://${cssPath}"></head><body>${sections.join('\n')}</body></html>`;
const htmlPath = path.join(outDir, 'cards.html');
fs.writeFileSync(htmlPath, html);

// 캡션
const cap = c.caption || {};
const list = c.items.map((it, i) => `${i + 1}. ${it.title.replace(/<br>/g, ' ')}`).join('\n');
const caption = [cap.lead, '', list, '', `[Duke 경험 한 줄 — 마지막 카드와 같은 내용]`, '', cap.outro, '', (cap.hashtags || []).map(h => '#' + h).join(' ')].join('\n');
fs.writeFileSync(path.join(outDir, 'caption.txt'), caption);

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1350 } });
  await page.goto('file://' + htmlPath);
  await page.evaluate(() => document.fonts.ready);
  const cards = page.locator('section.card');
  const n = await cards.count();
  for (let i = 0; i < n; i++) {
    await cards.nth(i).screenshot({ path: path.join(outDir, `card_${String(i + 1).padStart(2, '0')}.png`), type: 'png' });
  }
  await page.addStyleTag({ content: 'body{display:grid;grid-template-columns:repeat(4,1080px);gap:40px;padding:40px;width:max-content}' });
  const box = await page.locator('body').boundingBox();
  await page.setViewportSize({ width: Math.ceil(box.width), height: Math.ceil(box.height) });
  await page.screenshot({ path: path.join(outDir, 'contact_sheet.png'), fullPage: true });
  await browser.close();
  console.log(`카드 ${n}장 + contact_sheet.png + caption.txt → ${outDir}`);
})().catch(e => { console.error(e); process.exit(1); });
