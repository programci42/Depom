# Modern Tarayıcı ↔ VBScript Köprüsü (Node yok · WebView2 yok)

Bu örnek, **modern HTML5/CSS/JS arayüzünü** (gerçek Chromium motoru) çalıştırırken
**VBScript iş mantığınızı** (ADODB, Access, dosya işlemleri, COM/WMI…) korumanın
yolunu gösterir. Aradaki köprü, **Windows'ta zaten kurulu olan PowerShell** ile kurulur.

```
┌────────────────────┐   fetch /api    ┌──────────────────────┐   cscript    ┌───────────────┐
│  Modern arayüz     │ ───────────────▶│  PowerShell köprüsü  │ ───────────▶ │  backend.vbs  │
│  (Edge/Chrome)     │                 │  server.ps1          │              │  (VBScript)   │
│  HTML5 · CSS Grid  │◀─────────────── │  HttpListener        │◀─────────────│  ADODB/FSO... │
│  ES2023 · fetch    │   JSON yanıt    │  localhost:8765      │   UTF-8 JSON │  COM/WMI      │
└────────────────────┘                 └──────────────────────┘              └───────────────┘
        sandbox                          Windows'ta hazır gelir                  sizin kodunuz
```

## Neden bu mimari?

- VBScript yalnızca eski MSHTML/IE motorunda **sayfa içinde** çalışır; modern tarayıcılar
  onu tamamen kaldırdı. Bu yüzden VBScript'i **sayfanın dışına** (cscript) taşıyıp
  ağ üzerinden çağırıyoruz.
- Tarayıcı sayfası kum havuzundadır; doğrudan `cscript` çalıştıramaz. Bu nedenle
  araya **yerel bir HTTP dinleyici** gerekir. Node yerine **PowerShell** kullanıyoruz —
  ekstra kurulum yok, `System.Net.HttpListener` ile gerçek bir sunucu olur.
- Arayüz ile `/api` aynı origin'den (`localhost:8765`) servis edildiği için **CORS derdi yok**.

## Çalıştırma

1. Bu klasörü Windows'ta bir yere koyun.
2. **`baslat.vbs`** dosyasına çift tıklayın. Bu:
   - `server.ps1`'i ayrı bir PowerShell penceresinde başlatır (logları orada görürsünüz),
   - Edge/Chrome'u **uygulama modunda** (`--app=`) kenarlıksız açar.
3. Kapatmak için PowerShell penceresinde **Ctrl+C**.

> Sadece köprüyü test etmek isterseniz `server.ps1`'i elle çalıştırıp tarayıcınızdan
> `http://localhost:8765/` adresini açabilirsiniz.

## Dosyalar

| Dosya | Görevi |
|---|---|
| `baslat.vbs` | Başlatıcı: PowerShell sunucusunu açar + tarayıcıyı app modunda açar |
| `server.ps1` | **Köprü.** Statik dosyaları servis eder; `/api`'yi `cscript backend.vbs`'e bağlar |
| `backend.vbs` | **Sizin VBScript iş mantığınız.** `info / list / add / del / files` action'ları |
| `web/index.html` · `style.css` · `app.js` | Modern arayüz (CSS Grid, `:has()`, `fetch`, `async/await`, `?.`) |
| `data/` | Çalışma zamanı verisi (örnekte `users.txt`) |

## Akış (örnek: kullanıcı ekleme)

1. Arayüz: `fetch('/api?action=add&p1=Ayşe&p2=Yönetici')`
2. `server.ps1`: `cscript //nologo backend.vbs <tmp> add "Ayşe" "Yönetici"`
3. `backend.vbs`: `AddUser` → `FileSystemObject` ile yazar → UTF-8 JSON döndürür
4. `server.ps1`: sonucu okuyup tarayıcıya `application/json` olarak döner
5. Arayüz: listeyi tazeler

## Kendi projenize uyarlama

`backend.vbs` içindeki `ListUsers` / `AddUser` (FileSystemObject deposu) yerine
**kendi Access/ADODB kodunuzu** koyun. Dosya sonunda hazır bir şablon var:

```vbscript
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fso.BuildPath(dataDir, "app.mdb")
Set rs = conn.Execute("SELECT id, name, role FROM Users")
' ... JSON kur ...
```

Frontend tarafında değişiklik gerekmez; aynı `/api?action=...` çağrıları çalışır.

## Önemli notlar

- **Jet 4.0 (`.mdb`) 32-bit'tir.** 64-bit Windows'ta `server.ps1` içindeki
  `$cscript` değerini `"$env:WINDIR\SysWOW64\cscript.exe"` yapın; ya da `.accdb` +
  64-bit ACE OLEDB sağlayıcısını kurun.
- **Türkçe karakterler:** Köprü, sonucu ADODB.Stream ile **UTF-8 (BOM'suz)** üretir
  ve veri dosyasını FSO Unicode (UTF-16) tutar; "Ayşe Yılmaz" gibi girdiler bozulmaz.
- **PowerShell yürütme politikası:** `baslat.vbs` zaten `-ExecutionPolicy Bypass`
  kullanır; sistem politikanızı kalıcı değiştirmeniz gerekmez.
- **Güvenlik:** Sunucu yalnızca `localhost`'a bağlanır, `/api` yalnızca sabit
  action listesini (Select Case) çalıştırır ve `cscript` bir kabuk araya girmeden
  doğrudan çalıştırılır (komut enjeksiyonu yok). Yine de üretimde gelen parametreleri
  doğrulayın.
- Bu yaklaşım VBScript'e bağımlıdır; Microsoft VBScript'i aşamalı kaldırıyor. Uzun
  vadede iş mantığını PowerShell'e taşımak köprüyü "geleceğe dayanıklı" hale getirir
  (arayüz ve köprü aynı kalır, yalnızca `backend.vbs` → `backend.ps1` olur).
