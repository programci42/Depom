// Modern JS: fetch + async/await + template literal + ?. + ??  (IE11/HTA'da CALISMAZ)

// /api ucuna istek atan kucuk yardimci -> VBScript'ten JSON doner
const api = async (action, params = {}) => {
  const qs = new URLSearchParams({ action, ...params });
  const res = await fetch(`/api?${qs}`);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
};

const $ = (sel) => document.querySelector(sel);

const escapeHtml = (s) =>
  String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

async function loadInfo() {
  try {
    const i = await api('info');
    $('#info').innerHTML = `
      <dl class="kv">
        <dt>Bilgisayar</dt><dd>${escapeHtml(i.computer ?? '-')}</dd>
        <dt>Kullanıcı</dt><dd>${escapeHtml(i.user ?? '-')}</dd>
        <dt>İşletim Sistemi</dt><dd>${escapeHtml(i.os || '-')}</dd>
        <dt>Arka uç</dt><dd>${escapeHtml(i.engine ?? '-')}</dd>
      </dl>`;
  } catch (e) {
    $('#info').innerHTML = `<span class="no">Hata: ${escapeHtml(e.message)}</span>`;
  }
}

async function loadUsers() {
  try {
    const users = await api('list');
    $('#users').innerHTML = users.length
      ? users.map((u) => `
          <tr>
            <td>${u.id}</td>
            <td>${escapeHtml(u.name)}</td>
            <td>${escapeHtml(u.role)}</td>
            <td><button class="link danger" data-del="${u.id}">Sil</button></td>
          </tr>`).join('')
      : `<tr><td colspan="4" class="muted">Henüz kayıt yok</td></tr>`;
  } catch (e) {
    $('#users').innerHTML = `<tr><td colspan="4" class="no">Hata: ${escapeHtml(e.message)}</td></tr>`;
  }
}

$('#addForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const name = $('#name').value.trim();
  const role = $('#role').value;
  if (!name) return;
  await api('add', { p1: name, p2: role });   // -> backend.vbs AddUser
  e.target.reset();
  $('#name').focus();
  loadUsers();
});

$('#users').addEventListener('click', async (e) => {
  const id = e.target.dataset?.del;
  if (!id) return;
  if (!confirm('Bu kayıt silinsin mi?')) return;
  await api('del', { p1: id });               // -> backend.vbs DelUser
  loadUsers();
});

// "Gercekten modern motor mu?" kaniti
$('#engine').textContent = navigator.userAgent;
const mark = (sel, ok) => {
  const el = $(sel);
  el.textContent = ok ? 'evet ✓' : 'hayır ✗';
  el.className = ok ? 'ok' : 'no';
};
mark('#grid', CSS.supports('display', 'grid'));
mark('#has',  CSS.supports('selector(:has(*))'));
mark('#opt',  (() => { try { const o = {}; return o?.x === undefined; } catch { return false; } })());

loadInfo();
loadUsers();
