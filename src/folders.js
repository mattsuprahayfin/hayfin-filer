// ─── KNOWN FOLDERS ───────────────────────────────────────────────────────────
// Generated from live MS365 folder structure.
// To refresh: open a Claude conversation and run the folder-refresh prompt.
const KNOWN_FOLDERS = [
  // 1. Admin
  { name: "1. Admin",                          id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAADcg4RAAA=" },
  { name: "1. Admin / Arctic Project",         id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAF9pu_eAAA=" },
  { name: "1. Admin / Expenses",               id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J0pAAA=" },
  { name: "1. Admin / GWI",                   id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAANruPcsAAA=" },
  { name: "1. Admin / Training",              id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClSWAAA=" },
  // 2. Personal
  { name: "2. Personal",                      id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAADcg4SAAA=" },
  // 3. PSG
  { name: "3. PSG",                           id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClTSAAA=" },
  { name: "3. PSG / APAC",                    id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J0rAAA=" },
  { name: "3. PSG / Co-investments",          id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClTVAAA=" },
  { name: "3. PSG / DealCloud",               id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J2qAAA=" },
  { name: "3. PSG / DLF",                     id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClUhAAA=" },
  { name: "3. PSG / ESG",                     id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAF9pu_PAAA=" },
  { name: "3. PSG / HOF",                     id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClUzAAA=" },
  { name: "3. PSG / HTS-SOF",                id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClU1AAA=" },
  { name: "3. PSG / HYSL",                    id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAADcg4UAAA=" },
  { name: "3. PSG / Interval Fund",           id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JzCAAA=" },
  { name: "3. PSG / LGPS",                    id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClVlAAA=" },
  { name: "3. PSG / Maritime",                id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClVjAAA=" },
  { name: "3. PSG / PES",                     id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAOdClU4AAA=" },
  // 4. Multi-strat SMAs
  { name: "4. Multi-strat SMAs",              id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05Jy-AAA=" },
  { name: "4. Multi-strat SMAs / ART",        id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JzPAAA=" },
  { name: "4. Multi-strat SMAs / Big Cypress",id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J2jAAA=" },
  { name: "4. Multi-strat SMAs / Chief Illinois",id:"AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J2zAAA=" },
  { name: "4. Multi-strat SMAs / Future Fund", id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JzjAAA=" },
  { name: "4. Multi-strat SMAs / Hostplus",   id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JzuAAA=" },
  { name: "4. Multi-strat SMAs / HSBC",       id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAbbnl5fAAA=" },
  { name: "4. Multi-strat SMAs / Other OZ Accounts",id:"AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J0DAAA=" },
  { name: "4. Multi-strat SMAs / OTPP",       id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J3GAAA=" },
  { name: "4. Multi-strat SMAs / QIC",        id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JzTAAA=" },
  { name: "4. Multi-strat SMAs / REST",       id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JzUAAA=" },
  // 5. Research team
  { name: "5. Research team",                 id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAADcg4lAAA=" },
  // 6. External research
  { name: "6. External research",             id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAADcg4mAAA=" },
  { name: "6. External research / Octus",     id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05J2pAAA=" },
  { name: "6. External research / Reading",   id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAX05JztAAA=" },
  // 7. Old-Archive — never suggested by AI
  { name: "7. Old-Archive",                   id: "AAMkADkzM2ZmNmEyLWY2YTktNDM2YS1hY2NjLWQ5NWUxOWVlYjU3OAAuAAAAAAD7ma-xOdfUToQR3vJX4_-uAQBlHhb84Js_TIaV1hMreBscAAbbnl6kAAA=" },
];

// Folders available for AI suggestions (exclude archive)
const SUGGESTABLE_FOLDERS = KNOWN_FOLDERS.filter(f => f.name !== "7. Old-Archive");
