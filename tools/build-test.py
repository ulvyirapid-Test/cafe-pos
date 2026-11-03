#!/usr/bin/env python3
"""Builds the TEST copy of the POS into ../test/ from the live files.

The test copy runs on the same website (github.io), so it must never touch the
live POS data. This script changes, in the copy only:
  * IndexedDB name, sessionStorage key, localStorage prefix  -> separate data
  * Service-worker cache name and app manifest              -> separate app
  * Telegram messages start with "[TEST]"
  * An orange TEST banner with "demo data" and "reset" buttons

Run after every change to index.html:  python3 tools/build-test.py
"""
import pathlib, re, shutil, sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / 'test'


def rep(s, old, new, count=1):
    n = s.count(old)
    if n != count:
        sys.exit(f'build-test: expected {count}x, found {n}x: {old[:70]!r}')
    return s.replace(old, new)


html = (ROOT / 'index.html').read_text(encoding='utf-8')

# --- separate storage -------------------------------------------------------
html = rep(html, "const IDB_NAME = 'angkor-cafe-pos',", "const IDB_NAME = 'angkor-cafe-pos-TEST',")
html = html.replace("'cafe_session'", "'cafe_test_session'")
html = re.sub(r'\blocalStorage\b', 'testLS', html)
html = rep(html, "        // ==================== STORAGE (IndexedDB, testLS fallback) ====================", """        // ==================== TEST MODE: localStorage keys get a "test_" prefix ====================
        const testLS = {
            _keys() { return Object.keys(localStorage).filter(k => k.startsWith('test_')); },
            get length() { return this._keys().length; },
            key(i) { const k = this._keys()[i]; return k ? k.slice(5) : null; },
            getItem(k) { return localStorage.getItem('test_' + k); },
            setItem(k, v) { localStorage.setItem('test_' + k, v); },
            removeItem(k) { localStorage.removeItem('test_' + k); }
        };
        // ==================== STORAGE (IndexedDB, testLS fallback) ====================""")

# --- look & labels ----------------------------------------------------------
html = rep(html, '<title>', '<title>🧪 TEST — ')
html = rep(html, "document.title = `${settings.shopName} — POS`;", "document.title = `🧪 TEST — ${settings.shopName}`;")
html = rep(html, '<header class="bg-amber-950', '<header class="bg-orange-800')
html = rep(html, '<meta name="theme-color" content="#451a03">', '<meta name="theme-color" content="#c2410c">')
html = html.replace('`☕ <b>${tgEsc(settings.shopName)}</b>', '`🧪 <b>[TEST] ${tgEsc(settings.shopName)}</b>')

banner = '''
    <!-- ================= TEST MODE BANNER ================= -->
    <div id="test-banner" class="fixed bottom-2 left-2 z-[100] flex items-center gap-2 bg-orange-600 text-white text-[11px] font-bold px-3 py-1.5 rounded-full shadow-lg">
        <span>🧪 TEST MODE — ទិន្នន័យដាច់ពី POS ពិត</span>
        <button onclick="seedDemoData()" class="bg-white/20 hover:bg-white/30 px-2 py-0.5 rounded-full">+ ទិន្នន័យគំរូ</button>
        <button onclick="resetTestData()" class="bg-white/20 hover:bg-white/30 px-2 py-0.5 rounded-full">លុបទាំងអស់</button>
    </div>
'''
html = re.sub(r'(<body[^>]*>)', lambda m: m.group(1) + banner, html, count=1)

test_js = '''
        // ==================== TEST MODE TOOLS ====================
        async function resetTestData() {
            if (!confirm('លុបទិន្នន័យសាកល្បងទាំងអស់? (POS ពិតមិនត្រូវបានប៉ះពាល់ទេ)')) return;
            try { if (Store.db) Store.db.close(); } catch (e) { }
            await new Promise(r => { const q = indexedDB.deleteDatabase(IDB_NAME); q.onsuccess = q.onerror = q.onblocked = () => r(); });
            testLS._keys().forEach(k => localStorage.removeItem(k));
            try { sessionStorage.removeItem('cafe_test_session'); } catch (e) { }
            location.reload();
        }
        // Creates realistic sales for the last N days so reports, archive and analytics can be tried
        function seedDemoData(days = 120, perDay = 30) {
            if (!currentUser) { showToast('សូម Login ជាមុន', 'warning'); return; }
            if (!confirm(`បង្កើតវិក្កយបត្រគំរូ ~${days * perDay} សម្រាប់ ${days} ថ្ងៃចុងក្រោយ?`)) return;
            const items = menuItems.filter(m => !m.soldOut);
            if (!items.length) { showToast('គ្មាន Menu', 'warning'); return; }
            const pick = a => a[Math.floor(Math.random() * a.length)];
            const methods = ['cash', 'cash', 'cash', 'khqr', 'khqr', 'card'];
            let invNo = Math.max(Number(loadJSON('cafe_last_inv', 1000)) || 1000, ...salesHistory.map(s => s.invNo || 0));
            const now = Date.now(), out = [];
            for (let d = days; d >= 1; d--) {
                const n = Math.round(perDay * (0.6 + Math.random() * 0.8));
                for (let i = 0; i < n; i++) {
                    const t = new Date(now - d * 86400000);
                    t.setHours(7 + Math.floor(Math.random() * 13), Math.floor(Math.random() * 60), 0, 0);
                    const cart = [];
                    for (let k = 0, lines = 1 + Math.floor(Math.random() * 3); k < lines; k++) {
                        const m = pick(items), large = hasOptions(m) && Math.random() < 0.3;
                        cart.push({ cartUniqueId: now + out.length * 10 + k, itemId: m.id, name: m.name, unitPrice: round2(m.price + (large ? settings.largeExtra : 0)),
                            qty: 1 + (Math.random() < 0.2 ? 1 : 0), size: hasOptions(m) ? (large ? 'Large' : 'Regular') : '-',
                            sugar: hasOptions(m) ? pick(['0%', '25%', '50%', '100%']) : '-', ice: hasOptions(m) ? 'Normal Ice' : '-', toppings: [], note: '' });
                    }
                    const subtotal = round2(cart.reduce((s, c) => s + c.unitPrice * c.qty, 0));
                    const pct = Math.random() < 0.08 ? 10 : 0, discount = round2(subtotal * pct / 100);
                    const { vatUSD, totalUSD } = computeVat(subtotal - discount);
                    const method = pick(methods), consumption = computeConsumption(cart);
                    out.push({ id: t.getTime() * 100 + i, invNo: ++invNo, invId: `#TEST-${invNo}`, table: Math.random() < 0.4 ? 'ខ្ចប់ (Takeaway)' : `តុ ${1 + Math.floor(Math.random() * 12)}`,
                        timestamp: t.toISOString(), items: cart, subtotal, discount, discountPct: pct, discountBy: pct ? currentUser.name : null,
                        totalUSD, paymentMethod: method, paidUSD: totalUSD, changeUSD: 0, paidCashUSD: 0, paidCashKHR: 0, ref: method === 'khqr' ? String(100000 + i) : '',
                        rate: rate(), cashierId: currentUser.id, cashierName: currentUser.name, orderType: Math.random() < 0.4 ? 'takeaway' : 'dine-in',
                        payments: [{ method, amountUSD: totalUSD }], redeemUSD: 0, redeemPoints: 0, vatUSD, vatRate: vatOn() ? settings.vat.rate : 0,
                        vatInclusive: vatOn() ? !!settings.vat.inclusive : null, customerId: null, customerName: '', customerPhone: '', pointsEarned: 0,
                        cogs: costOfConsumption(consumption), consumption, voided: Math.random() < 0.01, demo: true });
                }
            }
            salesHistory = salesHistory.concat(out).sort((a, b) => a.timestamp.localeCompare(b.timestamp));
            saveSalesHistory();
            saveJSON('cafe_last_inv', invNo);
            archiveOldSales().then(() => showToast(`បានបង្កើត ${out.length} វិក្កយបត្រគំរូ`, 'success'));
        }
'''
html = rep(html, '        // ==================== OFFLINE (PWA) ====================', test_js + '        // ==================== OFFLINE (PWA) ====================')

OUT.mkdir(exist_ok=True)
(OUT / 'index.html').write_text(html, encoding='utf-8')

sw = (ROOT / 'sw.js').read_text(encoding='utf-8')
sw = re.sub(r"const CACHE = '([^']+)';", r"const CACHE = 'TEST-\1';", sw, count=1)
(OUT / 'sw.js').write_text(sw, encoding='utf-8')

manifest = (ROOT / 'manifest.json').read_text(encoding='utf-8')
manifest = manifest.replace('"Angkor Café POS"', '"Café POS — TEST"').replace('"Café POS"', '"POS TEST"').replace('#451a03', '#c2410c')
(OUT / 'manifest.json').write_text(manifest, encoding='utf-8')
shutil.copy(ROOT / 'icon.svg', OUT / 'icon.svg')
print('Built', OUT)
