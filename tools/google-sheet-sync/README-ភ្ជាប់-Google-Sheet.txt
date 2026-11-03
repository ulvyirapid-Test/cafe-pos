☕ Cafee Ankor POS — ភ្ជាប់ Cloud Sync ទៅ Google Sheet
=====================================================

ក. បង្កើត Google Sheet
  1. បើក https://sheets.new (Login Gmail របស់ហាង)
  2. ដាក់ឈ្មោះ៖ Cafee Ankor POS Data

ខ. ដាក់ Apps Script
  3. ម៉ឺនុយ Extensions (ផ្នែកបន្ថែម) → Apps Script
  4. លុបកូដចាស់ទាំងអស់ (Ctrl+A → Delete)
  5. បើកឯកសារ Code.gs (ក្នុង Folder នេះ) ដោយ Notepad → Ctrl+A → Ctrl+C
     → ត្រឡប់ទៅ Apps Script → Ctrl+V → ចុច 💾 Save (Ctrl+S)

គ. Run setup (បង្កើត Sheet និង API Key)
  6. នៅលើរបារខាងលើ ជ្រើស function "setup" → ចុច ▶ Run
  7. លោតផ្ទាំង "Authorization required" → Review permissions → ជ្រើស Gmail
     → បើឃើញ "Google hasn't verified this app" → Advanced → Go to ... (unsafe) → Allow
     (នេះជាកូដរបស់អ្នកផ្ទាល់ ដូច្នេះមានសុវត្ថិភាព)
  8. ខាងក្រោមលោត Execution log៖  ✅ Setup done. API KEY = xxxxxxxx
     → Copy API KEY ទុក

ឃ. Deploy ជា Web App
  9. ចុច Deploy (ខាងស្តាំលើ) → New deployment → ⚙️ → Web app
       Execute as: Me
       Who has access: Anyone
     → Deploy → Copy "Web app URL" (ចប់ដោយ /exec)

ង. ភ្ជាប់ក្នុង POS
 10. POS → ការកំណត់ → Cloud Sync (Google Sheets)
       Web App URL  = URL ពីជំហាន 9
       API Key      = Key ពីជំហាន 8
       លេខកូដម៉ាស៊ីន = A (ម៉ាស៊ីនទី២ ដាក់ B ...)
     → ធីក "បើក Sync" → ចុច "រក្សាទុក & សាកល្បង"
     → ឃើញ "តភ្ជាប់ Google Sheet ជោគជ័យ!" ✅
 11. ចុច "ផ្ញើប្រវត្តិលក់ចាស់ឡើង" ម្តង ដើម្បីផ្ញើការលក់ចាស់ៗទៅ Sheet

Sheet ដែលនឹងមាន
  Summary    — សរុបប្រចាំថ្ងៃ + មុខម្ហូបលក់ដាច់ (គណនាស្វ័យប្រវត្តិ)
  Sales      — វិក្កយបត្រម្តងមួយជួរ
  SaleItems  — មុខម្ហូបនីមួយៗ
  Voids      — ការ Void (អ្នកធ្វើ / អ្នកអនុម័ត / មូលហេតុ)
  Shifts     — បិទវេន (X/Z)
  Inventory  — ស្តុកបច្ចុប្បន្ន (ប្រើរួមគ្រប់ម៉ាស៊ីន)
  StockLog   — ទិញចូល / បន្ថែមស្តុក / រាប់ស្តុក
  (Config និង Processed ត្រូវបានលាក់ — កុំកែ)

សុវត្ថិភាព
  • កុំ Share Google Sheet នេះជាមួយអ្នកផ្សេង (មាន PIN អ៊ិនគ្រីប, Token Telegram)
  • កុំផ្ញើ API Key ជាមួយអ្នកផ្សេង
  • ភ្លេច Key → ក្នុង Apps Script ជ្រើស "showApiKey" → Run
  • ចង់ប្តូរ Key → ជ្រើស "newApiKey" → Run → ដាក់ Key ថ្មីក្នុង POS គ្រប់ម៉ាស៊ីន
  • កែ Code.gs ពេលក្រោយ → Deploy → Manage deployments → ✏️ → Version: New version → Deploy
    (URL នៅដដែល)
