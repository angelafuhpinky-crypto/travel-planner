function updateTripEuropeFamily2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const checklist = ss.getSheetByName("待辦清單");
  const itinerary = ss.getSheetByName("每日行程");
  const transport = ss.getSheetByName("交通與門票");
  const checks = ss.getSheetByName("檢查紀錄");
  const sources = ss.getSheetByName("資料來源");

  if (!checklist || !itinerary || !transport || !checks || !sources) {
    SpreadsheetApp.getUi().alert("找不到必要分頁，請確認有：待辦清單、每日行程、交通與門票、檢查紀錄、資料來源");
    return;
  }

  updateFullItineraryEurope2026_(itinerary);
  updateTransportAndTicketsEurope2026_(transport);
  updateChecklistTasksEurope2026_(checklist);
  updateCheckRecordsEurope2026_(checks);
  updateSourcesEurope2026_(sources);
  formatTripSheets_(checklist, itinerary, transport, checks, sources);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("已更新 2026/8/4–8/17 荷蘭/比利時/德國行程、預訂清單與檢查紀錄。");
}


function updateTripVanGoghArrivalWindmill() {
  updateTripEuropeFamily2026();
}


function d_(year, month, day) {
  return new Date(year, month - 1, day);
}


function updateFullItineraryEurope2026_(sheet) {
  const rows = [
    [
      d_(2026, 8, 4),
      "Day 1",
      "抵達阿姆斯特丹＋Zaanse Schans 風車村",
      "阿姆斯特丹",
      "07:00 抵達 Schiphol；08:30–09:45 飯店寄放行李；10:15–11:15 前往 Zaanse Schans；11:15–14:00 風車村；下午回阿姆斯特丹 check-in、休息、晚餐、附近運河區散步。",
      "Schiphol → Amsterdam 用 NS 火車；Amsterdam Centraal → Zaanse Schans 可搭火車到 Zaandijk Zaanse Schans 後步行約 15 分鐘，或查 9292 / 當日巴士。交通不用預約，用 OVpay / NS App / 9292 查詢即可。",
      "第一天早上 7 點抵達，放完行李去風車村可行，但晚上不要排正式景點。若航班延誤或爸媽累，風車村停留時間可縮短。"
    ],
    [
      d_(2026, 8, 5),
      "Day 2",
      "Van Gogh Museum＋阿姆斯特丹運河＋Anne Frank House",
      "阿姆斯特丹",
      "09:30–12:00 Van Gogh Museum；12:00–13:30 Museumplein / De Pijp 午餐；14:00–16:00 Jordaan、九小街與運河區散步；16:30–17:45 運河遊船；18:00–19:00 簡單晚餐；19:30–19:45 Anne Frank House 入場。",
      "阿姆斯特丹市區用 GVB / OVpay；Van Gogh Museum 建議搭電車或步行到 Museumplein；Anne Frank House 可從 Jordaan / 市中心步行前往。",
      "Van Gogh Museum 需先買指定時段票，建議選 09:30–10:30 入場。Anne Frank House 已訂 19:30–19:45。不要同天再塞 Rijksmuseum。"
    ],
    [
      d_(2026, 8, 6),
      "Day 3",
      "Gouda 起司市集＋Delft 放行李＋IKEA",
      "Delft / 海牙",
      "08:00–09:30 阿姆斯特丹退房並前往 Gouda；10:00–12:30 Gouda Cheese Market；12:30–14:00 Gouda 午餐與老城散步；14:30–15:30 前往 Delft；15:30–16:30 宿舍 / 住宿寄放行李；16:30–18:30 IKEA Delft 採買生活用品。",
      "Amsterdam → Gouda → Delft 用 NS 火車；Delft 市內可用公車、步行或計程車。",
      "2026/08/06 是星期四，符合 Gouda 起司市集時間。大行李建議這天直接帶去 Delft，不要再回阿姆斯特丹。"
    ],
    [
      d_(2026, 8, 7),
      "Day 4",
      "Delft 老城＋TU Delft＋海牙海灘",
      "Delft / 海牙",
      "09:30–11:30 Delft 老城、新教堂廣場、運河、市集區；12:00–13:30 TU Delft 周邊；14:00–16:30 海牙市區；17:00–19:30 Scheveningen 海灘散步與晚餐。",
      "Delft ↔ Den Haag 用 NS 火車或電車；海牙市區到 Scheveningen 用電車。",
      "這天讓爸媽看你未來生活圈，不要排太遠。若 Mauritshuis 想入內，海灘時間需縮短。"
    ],
    [
      d_(2026, 8, 8),
      "Day 5",
      "鹿特丹＋Kinderdijk 小孩堤防 → 布魯塞爾",
      "布魯塞爾",
      "09:00–09:45 Delft/海牙 → Rotterdam；10:00–11:30 Markthal、方塊屋；12:00–13:00 前往 Kinderdijk；13:00–15:15 Kinderdijk 小孩堤防；15:15–16:30 回 Rotterdam Centraal；17:00 後搭國際火車去 Brussels；晚上住宿附近簡單晚餐。",
      "Delft/Den Haag → Rotterdam 用 OVpay；Rotterdam ↔ Kinderdijk 優先查 Waterbus / WaterShuttle；Rotterdam → Brussels 國際火車要先訂，建議 17:00 後。",
      "這是全程最緊的一天。加入 Kinderdijk 後，鹿特丹市區只能簡短停留。大行李留 Delft，只帶輕便小行李。"
    ],
    [
      d_(2026, 8, 9),
      "Day 6",
      "布魯塞爾市區一日",
      "布魯塞爾",
      "09:30 Grand Place；10:30 皇家聖休伯特拱廊；12:00 尿尿小童與市中心午餐；14:00 Mont des Arts 與皇宮周邊；16:30 咖啡、巧克力與休息；18:30 淡菜鍋或比利時料理晚餐。",
      "布魯塞爾市中心以步行為主；需要時用 STIB 地鐵／電車或感應付款。",
      "住宿建議 Brussels Central / Grand Place 附近，不建議住 Midi 車站正旁邊。"
    ],
    [
      d_(2026, 8, 10),
      "Day 7",
      "布魯日住一晚",
      "布魯日",
      "09:00–10:30 布魯塞爾前往布魯日；10:30–11:30 飯店寄放行李；12:00–14:00 Markt 廣場、Belfry 外觀或視體力登塔；14:30–17:00 Rozenhoedkaai、布魯日運河與老城散步；晚上布魯日夜景散步。",
      "Brussels → Bruges 用 SNCB 比利時國鐵；布魯日車站到老城可搭公車或計程車。",
      "首選把行李寄放布魯日飯店；備案是 Brugge Station 置物櫃。住一晚的重點是晚上遊客少、氣氛更好。"
    ],
    [
      d_(2026, 8, 11),
      "Day 8",
      "根特半日＋回布魯塞爾",
      "布魯塞爾",
      "09:00–10:30 布魯日退房後回布魯塞爾寄放行李；11:30–12:15 輕裝前往 Ghent；12:45 根特老城午餐；14:00–16:30 Graslei、Korenlei 與根特老城散步；17:00–18:00 回布魯塞爾。",
      "Bruges → Brussels → Ghent → Brussels 用 SNCB 比利時國鐵；Ghent 車站進城可搭電車或計程車。",
      "不建議拖行李玩根特，先回布魯塞爾飯店寄放會比較輕鬆。這樣稍微繞，但爸媽同行更舒服。"
    ],
    [
      d_(2026, 8, 12),
      "Day 9",
      "布魯塞爾 → 科隆",
      "科隆",
      "10:00–11:00 布魯塞爾退房；12:00–14:00 國際火車到 Köln Hbf；14:30–15:30 飯店 check-in；16:00–17:30 科隆大教堂；17:30–19:00 Hohenzollern Bridge 與萊茵河畔；晚上科隆老城晚餐。",
      "Bruxelles-Midi / Brussel-Zuid → Köln Hbf，建議先訂 Eurostar / ICE / DB / SNCB International。",
      "住宿建議 Köln Hbf / Cologne Cathedral / Altstadt-Nord 附近。抵達日不要再安排其他城市。"
    ],
    [
      d_(2026, 8, 13),
      "Day 10",
      "科隆完整一日",
      "科隆",
      "09:30–11:30 科隆大教堂深度參觀；12:00–13:30 老城午餐；14:00–16:30 Museum Ludwig 或巧克力博物館二選一；17:00–18:30 萊茵河畔散步；晚上科隆市區晚餐。",
      "住車站或大教堂附近多數可步行；需要時用 KVB / VRS。",
      "這天完整玩科隆，不建議再塞其他城市。若登大教堂塔，下午行程放慢。"
    ],
    [
      d_(2026, 8, 14),
      "Day 11",
      "科隆彈性日",
      "科隆",
      "09:30–11:00 科隆市區悠閒早餐；11:30–14:00 Schildergasse 購物街、咖啡、伴手禮；14:30–18:00 彈性選項：留在科隆休息，或爸媽體力好就去杜塞道夫半日；晚上科隆最後晚餐。",
      "科隆市區步行或 KVB / VRS；若去杜塞道夫則搭區域火車。",
      "建議以休息為主，因為前後都有跨國移動。杜塞道夫是備案，不是必走。"
    ],
    [
      d_(2026, 8, 15),
      "Day 12",
      "科隆 → Delft",
      "Delft",
      "10:00–11:00 科隆退房；11:30–15:30 搭國際火車回 Delft，可能經 Utrecht / Rotterdam / Amsterdam 轉車；16:00–18:00 回宿舍整理、洗衣、補買生活用品；晚上休息。",
      "Köln Hbf → Delft，可經 Utrecht / Rotterdam / Amsterdam 轉車；建議先訂國際段。",
      "這天不要排正式景點，重點是回宿舍整理。週六移動日，需留轉車緩衝。"
    ],
    [
      d_(2026, 8, 16),
      "Day 13",
      "Utrecht 米飛兔博物館＋Delft 生活準備",
      "Delft",
      "09:00–10:15 Delft 前往 Utrecht；10:30–12:00 Miffy Museum / Nijntje Museum；12:00–14:00 Utrecht 市中心、米飛兔紅綠燈、午餐；14:00–15:15 回 Delft；15:30–18:00 採買、洗衣、整理宿舍；晚上 Delft 老城晚餐。",
      "Delft → Utrecht Centraal 搭 NS 火車；Utrecht 市區步行或公車；回 Delft 後用步行／公車處理採買。",
      "Keukenhof 8 月未開放，所以米飛兔行程改排 Utrecht Miffy Museum。Miffy Museum 需線上預訂時段票。"
    ],
    [
      d_(2026, 8, 17),
      "Day 14",
      "Delft 彈性日／送機",
      "Delft / Schiphol",
      "09:30–11:30 最後整理宿舍、確認文件與交通；12:00–14:00 若爸媽航班較晚，可安排 Delft 或海牙輕鬆散步；14:30–16:30 視航班時間前往 Schiphol 送機；晚上行程結束。",
      "Delft → Schiphol 用 NS 火車；Delft / 海牙市區用 OVpay。",
      "保留彈性，不要再安排遠距離景點。送機時間依實際航班調整，至少預留充足緩衝。"
    ]
  ];

  rows.forEach(function(row) {
    upsertItineraryByDate_(sheet, row[0], row);
  });
}


function updateTransportAndTicketsEurope2026_(sheet) {
  const rows = [
    ["不用訂", d_(2026, 8, 4), "8/4 Schiphol → Amsterdam", "Amsterdam Airport Schiphol → Amsterdam", "07:00 抵達後", "NS App / OVpay", "https://www.ns.nl/en", "抵達後搭 NS 火車進阿姆斯特丹。若飯店不能提早入住，先寄放行李。"],
    ["不用訂", d_(2026, 8, 4), "Zaanse Schans 風車村交通", "Amsterdam Centraal ↔ Zaandijk Zaanse Schans / Zaanse Schans", "10:15–15:00 建議", "NS App / 9292 / OVpay", "https://www.dezaanseschans.nl/en/plan-your-visit/accessibility/", "第一天抵達後放完行李去風車村。可搭火車到 Zaandijk Zaanse Schans 後步行約 15 分鐘，或依 9292 查當日巴士。"],
    ["選購", d_(2026, 8, 4), "Ticket Zaanse Schans / 風車入內票", "Zaanse Schans", "中午前後", "Zaanse Schans 官方票務 / 現場購票", "https://tickets.dezaanseschans.nl/en/", "村落本身免費；若想進 Zaans Museum、Windmill Museum 或風車內部，再買 Ticket Zaanse Schans 或單項票。"],
    ["必訂", d_(2026, 8, 5), "Van Gogh Museum 梵谷博物館門票", "Museumplein, Amsterdam", "建議 09:30–10:30 入場", "Van Gogh Museum 官方網站", "https://www.vangoghmuseum.nl/en/visit/tickets-and-ticket-prices", "需線上購買指定入場時間票。建議 8/5 上午入場，下午留給運河遊船與 Anne Frank House。"],
    ["建議先訂", d_(2026, 8, 5), "阿姆斯特丹運河遊船", "Amsterdam Canal Cruise", "建議 16:30–17:45", "Blue Boat / Lovers / Klook / GetYourGuide / Tiqets", "", "排在 Van Gogh Museum 與散步之後、Anne Frank House 之前。建議不要選太晚，避免趕不上 19:30 入場。"],
    ["已完成", d_(2026, 8, 5), "Anne Frank House", "Anne Frank House, Amsterdam", "19:30–19:45", "Anne Frank House 官方網站", "https://www.annefrank.org/en/museum/tickets/", "你已經訂好 19:30–19:45，請把票存到手機，也備份 PDF。"],
    ["不用訂", d_(2026, 8, 6), "Amsterdam → Gouda → Delft NS 火車", "Amsterdam Centraal → Gouda → Delft", "08:00–15:30", "NS App / OVpay", "https://www.ns.nl/en", "荷蘭境內火車不用提前訂票，但當天請用 NS App 查即時班次與轉車。"],
    ["不用訂", d_(2026, 8, 6), "Gouda Cheese Market", "Markt, Gouda", "10:00–12:30", "Welcome in Gouda 官方活動頁", "https://welkomingouda.nl/en/", "2026/08/06 是星期四，可排 Gouda 起司市集。建議 09:30 前到，並在出發前確認 2026 官方場次。"],
    ["不用訂", d_(2026, 8, 7), "Delft ↔ The Hague / Scheveningen", "Delft → Den Haag → Scheveningen", "下午至晚上", "NS / HTM / 9292 / OVpay", "https://9292.nl/en/", "Delft 到海牙可搭火車或電車；海牙市區到 Scheveningen 搭電車。"],
    ["必訂", d_(2026, 8, 8), "Rotterdam → Brussels 國際火車", "Rotterdam Centraal → Bruxelles-Midi / Brussel-Zuid", "建議 17:00 後", "NS International / SNCB International / Eurostar", "https://www.nsinternational.com/en", "因 8/8 加入 Kinderdijk，Rotterdam → Brussels 不要訂太早。建議 17:00 後，並預留回 Rotterdam Centraal 的緩衝。"],
    ["建議先訂", d_(2026, 8, 8), "Kinderdijk 小孩堤防門票", "Kinderdijk UNESCO World Heritage", "13:00–15:15", "Kinderdijk 官方票務", "https://kinderdijk.com/tickets/", "建議線上買票。若只想拍外景可不買完整票，但爸媽同行建議買完整票比較順。"],
    ["建議先訂", d_(2026, 8, 8), "Rotterdam ↔ Kinderdijk 水上交通", "Rotterdam → Kinderdijk → Rotterdam", "12:00–16:30", "Waterbus / WaterShuttle / Kinderdijk 官方網站", "https://www.waterbus.nl/", "Waterbus / WaterShuttle 班次需出發前確認。若水上交通不順，改查 9292 的火車加巴士備案。"],
    ["不用訂", d_(2026, 8, 10), "Brussels → Bruges 比利時國鐵", "Brussels Central / Midi → Bruges", "09:00–10:30", "SNCB / Belgian Train", "https://www.belgiantrain.be/en", "比利時國內火車通常不用太早訂，當天查班次即可。"],
    ["不用訂", d_(2026, 8, 11), "Bruges → Brussels → Ghent → Brussels", "Bruges / Brussels / Gent-Sint-Pieters", "整天", "SNCB / Belgian Train", "https://www.belgiantrain.be/en", "建議先回布魯塞爾放行李，再輕裝去根特。"],
    ["必訂", d_(2026, 8, 12), "Brussels → Cologne 國際火車", "Bruxelles-Midi / Brussel-Zuid → Köln Hbf", "建議中午前後", "Eurostar / DB / SNCB International", "https://www.b-europe.com/EN", "建議先訂票，避免旺季價格上升。"],
    ["可選", d_(2026, 8, 13), "Museum Ludwig 或巧克力博物館", "Cologne", "14:00–16:30", "各館官方網站 / 現場", "", "二選一即可。若遇到雨天，這段很適合作為室內備案。"],
    ["可選", d_(2026, 8, 14), "Cologne ↔ Düsseldorf 區域火車", "Köln Hbf ↔ Düsseldorf Hbf", "下午備案", "DB / VRS / VRR", "https://int.bahn.de/en", "杜塞道夫是彈性備案，不是必走。爸媽體力普通就留科隆休息。"],
    ["必訂", d_(2026, 8, 15), "Cologne → Delft 國際火車", "Köln Hbf → Delft", "建議上午晚一點或中午", "DB / NS International", "https://www.nsinternational.com/en", "週六移動日，建議先訂國際段並預留轉車緩衝。"],
    ["必訂", d_(2026, 8, 16), "Miffy Museum / Nijntje Museum 門票", "Agnietenstraat 2, Utrecht", "建議 10:30–12:00", "Miffy Museum 官方網站", "https://nijntjemuseum.nl/en/times-and-prices/", "需線上預訂時段票。建議選上午或中午前後，下午回 Delft 整理。"],
    ["不適用", d_(2026, 8, 16), "Keukenhof 檢查", "Lisse / Keukenhof", "8 月不開放", "Keukenhof 官方網站", "https://keukenhof.nl/en/", "Keukenhof 是春季花園，8/4–8/17 期間不適合排入主行程。"],
    ["不建議主行程", d_(2026, 8, 17), "Giethoorn 羊角村備案", "Delft / Amsterdam → Steenwijk → Giethoorn", "需整天", "9292 / NS App", "https://9292.nl/en/", "不建議放入主行程。交通時間太長，會壓縮 Delft 整理與送機彈性。"]
  ];

  rows.forEach(function(row) {
    upsertRowByValue_(sheet, 3, row[2], row);
  });
}


function updateChecklistTasksEurope2026_(sheet) {
  const tasks = [
    [false, "高", d_(2026, 7, 3), "景點門票", "預訂 8/5 Van Gogh Museum 梵谷博物館時段票", "Van Gogh Museum 官方網站", "https://www.vangoghmuseum.nl/en/visit/tickets-and-ticket-prices", "建議選 8/5 上午 09:30–10:30 入場，下午留給運河遊船與 Anne Frank House。", "待確認"],
    [false, "中", d_(2026, 7, 5), "交通", "確認 8/4 抵達後前往 Zaanse Schans 交通", "NS App / 9292 / Zaanse Schans 官方交通頁", "https://www.dezaanseschans.nl/en/plan-your-visit/accessibility/", "第一天 07:00 到 Schiphol，先到飯店寄放行李，再搭火車或巴士去風車村。", "待確認"],
    [false, "低", d_(2026, 7, 8), "景點門票", "決定是否購買 Zaanse Schans 風車入內票", "Zaanse Schans 官方票務", "https://tickets.dezaanseschans.nl/en/", "村落免費；若要進博物館或風車內部才需要買票。", "待確認"],
    [false, "中", d_(2026, 7, 8), "景點門票", "預訂 8/5 阿姆斯特丹運河遊船", "Blue Boat / Lovers / Klook / GetYourGuide / Tiqets", "", "建議選 16:30–17:45，不要太晚，避免趕不上 Anne Frank House 19:30–19:45。", "待確認"],
    [true, "高", d_(2026, 6, 27), "景點門票", "Anne Frank House 19:30–19:45 已完成", "Anne Frank House 官方網站", "https://www.annefrank.org/en/museum/tickets/", "已訂 8/5 19:30–19:45。請將票存手機與雲端備份。", "OK"],
    [false, "中", d_(2026, 7, 10), "行程確認", "出發前確認 Gouda Cheese Market 2026 官方場次", "Welcome in Gouda 官方網站", "https://welkomingouda.nl/en/", "2026/08/06 是星期四，日期合理；仍建議出發前確認 2026 活動場次與時間。", "待確認"],
    [false, "高", d_(2026, 7, 10), "交通", "預訂 8/8 Rotterdam → Brussels 國際火車 17:00 後班次", "NS International / SNCB International / Eurostar", "https://www.nsinternational.com/en", "加入 Kinderdijk 後，Rotterdam → Brussels 不要訂太早。若已訂早班，請評估是否改票。", "待確認"],
    [false, "高", d_(2026, 7, 10), "景點門票", "預訂 8/8 Kinderdijk 小孩堤防門票", "Kinderdijk 官方票務", "https://kinderdijk.com/tickets/", "建議線上買票；若只拍外景可不買完整票，但爸媽同行建議買完整票比較順。", "待確認"],
    [false, "高", d_(2026, 7, 10), "交通", "預訂或確認 8/8 Rotterdam ↔ Kinderdijk 水上交通", "Waterbus / Kinderdijk WaterShuttle", "https://www.waterbus.nl/", "Waterbus / WaterShuttle 班次需出發前確認。若班次不順，改查 9292 備案。", "待確認"],
    [false, "高", d_(2026, 7, 12), "住宿", "確認 8/8–8/10 布魯塞爾住宿位置", "Booking / Google Maps / 飯店官網", "", "建議住 Brussels Central / Grand Place 附近，不建議住 Midi 車站正旁邊。", "待確認"],
    [false, "中", d_(2026, 7, 12), "住宿", "確認 8/10 布魯日飯店可寄放行李", "飯店訊息 / Booking", "", "首選把行李寄放布魯日飯店；備案是 Brugge Station 置物櫃。", "待確認"],
    [false, "高", d_(2026, 7, 15), "交通", "預訂 8/12 Brussels → Cologne 國際火車", "Eurostar / DB / SNCB International", "https://www.b-europe.com/EN", "建議中午前後班次，抵達後可輕鬆看科隆大教堂與河畔。", "待確認"],
    [false, "高", d_(2026, 7, 15), "交通", "預訂 8/15 Cologne → Delft 國際火車", "DB / NS International", "https://www.nsinternational.com/en", "週六移動日，建議先訂國際段並預留轉車緩衝。", "待確認"],
    [false, "高", d_(2026, 7, 18), "景點門票", "預訂 8/16 Miffy Museum / Nijntje Museum 時段票", "Miffy Museum 官方網站", "https://nijntjemuseum.nl/en/times-and-prices/", "需線上預訂時段票。建議選上午或中午前後。", "待確認"],
    [false, "中", d_(2026, 7, 20), "行李", "確認 8/8 Rotterdam 小行李寄放方式", "Rotterdam Centraal / 飯店 / 置物櫃", "", "這天大行李留 Delft，只帶小行李。若去 Kinderdijk，建議小行李寄放 Rotterdam Centraal 或盡量輕便。", "待確認"],
    [false, "中", d_(2026, 7, 25), "送機", "確認 8/17 爸媽航班時間與 Delft → Schiphol 出發時間", "NS App / 航空公司", "https://www.ns.nl/en", "依實際航班時間倒推，預留火車延誤、退稅、行李與安檢時間。", "待確認"],
    [false, "低", d_(2026, 7, 25), "備案", "羊角村 Giethoorn 不放入主行程", "9292 / NS App", "https://9292.nl/en/", "交通時間太長，會壓縮 Delft 整理日與送機彈性；目前不建議加入。", "待確認"]
  ];

  tasks.forEach(function(row) {
    upsertChecklistTask_(sheet, row);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow >= 8) {
    try {
      sheet.getRange(8, 1, lastRow - 7, 1).insertCheckboxes();
    } catch (e) {}
  }
}


function updateCheckRecordsEurope2026_(sheet) {
  const rows = [
    ["日期檢查", "本版行程日期為 2026/08/04–2026/08/17，共 14 天。2026/08/06 是星期四，因此 Gouda 起司市集排 Day 3 合理。", "OK"],
    ["8/4 抵達後風車村安排檢查", "07:00 抵達 Schiphol 後，先到阿姆斯特丹飯店寄放行李，再去 Zaanse Schans 可行；但晚上不排正式景點。", "OK"],
    ["8/5 Van Gogh Museum 安排檢查", "Van Gogh Museum 排上午，下午散步與運河遊船，晚上 Anne Frank House 19:30–19:45，時間順。", "OK"],
    ["8/6 Gouda 起司市集安排檢查", "2026/08/06 是星期四，可排 Gouda 起司市集 10:00–12:30；之後去 Delft 放行李與 IKEA。", "OK"],
    ["8/7 Delft / 海牙安排檢查", "Delft 老城、TU Delft、海牙與 Scheveningen 同一天可行，但 Mauritshuis 若入內，海灘時間需縮短。", "OK"],
    ["Kinderdijk 小孩堤防安排檢查", "Kinderdijk 排在 8/8 Rotterdam → Brussels 當天可行，但必須壓縮鹿特丹市區，並將 Rotterdam → Brussels 火車訂在 17:00 後。", "偏緊但可行"],
    ["比利時段行李檢查", "8/11 先從布魯日回布魯塞爾寄放行李，再輕裝去根特，交通稍繞但對爸媽同行較舒服。", "OK"],
    ["科隆段節奏檢查", "8/12 抵達科隆後只排大教堂、橋與河畔；8/13 完整玩科隆；8/14 彈性休息，節奏合理。", "OK"],
    ["Keukenhof 時間檢查", "Keukenhof 是春季花園，與 8/4–8/17 行程不符，因此不排入主行程。", "不排入"],
    ["米飛兔替代方案檢查", "米飛兔行程改為 8/16 Utrecht Miffy Museum / Nijntje Museum，需線上預訂時段票。", "OK"],
    ["Giethoorn 羊角村安排檢查", "羊角村交通時間過長，與目前荷蘭南下比利時德國動線不順，不建議加入主行程。", "不建議"],
    ["送機日檢查", "8/17 應保留 Delft / 海牙輕鬆散步與送機彈性，不建議再安排遠距離景點。", "OK"]
  ];

  rows.forEach(function(row) {
    upsertRowByValue_(sheet, 1, row[0], row);
  });
}


function updateSourcesEurope2026_(sheet) {
  const rows = [
    ["NS Netherlands", "https://www.ns.nl/en", "荷蘭境內火車查詢與票務。"],
    ["9292 Journey Planner", "https://9292.nl/en/", "荷蘭境內大眾運輸查詢，含電車、公車、轉乘。"],
    ["Zaanse Schans Accessibility", "https://www.dezaanseschans.nl/en/plan-your-visit/accessibility/", "Zaanse Schans 官方交通資訊。"],
    ["Ticket Zaanse Schans", "https://tickets.dezaanseschans.nl/en/", "Zaanse Schans 官方票務與風車、博物館票券資訊。"],
    ["Van Gogh Museum Tickets", "https://www.vangoghmuseum.nl/en/visit/tickets-and-ticket-prices", "Van Gogh Museum 官方門票與指定時段預訂。"],
    ["Van Gogh Museum Opening Hours", "https://www.vangoghmuseum.nl/en/visit/address-and-opening-hours", "Van Gogh Museum 官方開放時間。"],
    ["Anne Frank House Tickets", "https://www.annefrank.org/en/museum/tickets/", "Anne Frank House 官方票務資訊。"],
    ["Welcome in Gouda", "https://welkomingouda.nl/en/", "Gouda 官方旅遊與活動資訊，出發前確認 Cheese Market 2026 場次。"],
    ["Kinderdijk Tickets", "https://kinderdijk.com/tickets/", "Kinderdijk 官方票務資訊。"],
    ["Waterbus", "https://www.waterbus.nl/", "Rotterdam / Dordrecht / Kinderdijk 周邊水上交通查詢。"],
    ["Belgian Train / SNCB", "https://www.belgiantrain.be/en", "比利時國內火車查詢與票務。"],
    ["NS International", "https://www.nsinternational.com/en", "荷蘭出發國際火車查詢與票務。"],
    ["SNCB International", "https://www.b-europe.com/EN", "比利時出發國際火車查詢與票務。"],
    ["DB Bahn", "https://int.bahn.de/en", "德國與跨國火車查詢與票務。"],
    ["Miffy Museum Times and Prices", "https://nijntjemuseum.nl/en/times-and-prices/", "Miffy Museum 開放時間與線上預約規則。"],
    ["Keukenhof", "https://keukenhof.nl/en/", "Keukenhof 官方網站，用於確認非 8 月開放景點。"]
  ];

  rows.forEach(function(row) {
    upsertRowByValue_(sheet, 1, row[0], row);
  });
}


function upsertItineraryByDate_(sheet, targetDate, rowValues) {
  const lastRow = sheet.getLastRow();

  for (let r = 2; r <= lastRow; r++) {
    const value = sheet.getRange(r, 1).getValue();
    if (isSameDate_(value, targetDate)) {
      sheet.getRange(r, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }

  sheet.appendRow(rowValues);
}


function upsertChecklistTask_(sheet, rowValues) {
  const taskName = rowValues[4];
  const startRow = 8;
  const lastRow = sheet.getLastRow();

  for (let r = startRow; r <= lastRow; r++) {
    const existingTask = String(sheet.getRange(r, 5).getValue()).trim();
    if (existingTask === taskName) {
      sheet.getRange(r, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }

  sheet.appendRow(rowValues);
}


function upsertRowByValue_(sheet, searchColumn, searchValue, rowValues) {
  const lastRow = sheet.getLastRow();

  for (let r = 2; r <= lastRow; r++) {
    const existingValue = String(sheet.getRange(r, searchColumn).getValue()).trim();
    if (existingValue === searchValue) {
      sheet.getRange(r, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }

  sheet.appendRow(rowValues);
}


function isSameDate_(value, targetDate) {
  if (value instanceof Date) {
    return value.getFullYear() === targetDate.getFullYear() &&
           value.getMonth() === targetDate.getMonth() &&
           value.getDate() === targetDate.getDate();
  }

  const text = String(value).trim();
  const yyyy = targetDate.getFullYear();
  const mm = String(targetDate.getMonth() + 1).padStart(2, "0");
  const dd = String(targetDate.getDate()).padStart(2, "0");

  return text === `${yyyy}-${mm}-${dd}` ||
         text === `${yyyy}/${mm}/${dd}` ||
         text === `${yyyy}.${mm}.${dd}`;
}


function formatTripSheets_(checklist, itinerary, transport, checks, sources) {
  checklist.getRange("C8:C300").setNumberFormat("yyyy/mm/dd");
  itinerary.getRange("A2:A100").setNumberFormat("yyyy/mm/dd");
  transport.getRange("B2:B300").setNumberFormat("yyyy/mm/dd");

  checklist.getRange("E8:H300").setWrap(true);
  itinerary.getRange("E2:G100").setWrap(true);
  transport.getRange("C2:H300").setWrap(true);
  checks.getRange("A2:C300").setWrap(true);
  sources.getRange("A2:C300").setWrap(true);

  checklist.setColumnWidth(1, 90);
  checklist.setColumnWidth(2, 70);
  checklist.setColumnWidth(3, 110);
  checklist.setColumnWidth(4, 110);
  checklist.setColumnWidth(5, 390);
  checklist.setColumnWidth(6, 260);
  checklist.setColumnWidth(7, 390);
  checklist.setColumnWidth(8, 520);
  checklist.setColumnWidth(9, 90);

  itinerary.setColumnWidth(1, 110);
  itinerary.setColumnWidth(2, 80);
  itinerary.setColumnWidth(3, 320);
  itinerary.setColumnWidth(4, 150);
  itinerary.setColumnWidth(5, 620);
  itinerary.setColumnWidth(6, 560);
  itinerary.setColumnWidth(7, 560);

  transport.setColumnWidth(1, 90);
  transport.setColumnWidth(2, 110);
  transport.setColumnWidth(3, 360);
  transport.setColumnWidth(4, 360);
  transport.setColumnWidth(5, 190);
  transport.setColumnWidth(6, 320);
  transport.setColumnWidth(7, 390);
  transport.setColumnWidth(8, 560);

  const normalSheets = [itinerary, transport, checks, sources];

  normalSheets.forEach(function(sh) {
    const lastCol = sh.getLastColumn();
    sh.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setBackground("#E8F0FE")
      .setHorizontalAlignment("center");
    sh.setFrozenRows(1);
  });

  checklist.getRange("A7:I7")
    .setFontWeight("bold")
    .setBackground("#E8F0FE")
    .setHorizontalAlignment("center");

  checklist.setFrozenRows(7);
}


const TRAVEL_PLANNER_SYNC_SHEET = "TravelPlannerSync";


function doGet(e) {
  try {
    const docId = getSyncDocId_(e);
    const record = readTravelPlannerRecord_(docId);

    if (!record) {
      return travelPlannerJson_({
        ok: true,
        docId: docId,
        data: null,
        message: "No travel planner data has been saved for this document yet."
      });
    }

    return travelPlannerJson_({
      ok: true,
      docId: docId,
      updatedAt: record.updatedAt,
      data: JSON.parse(record.json)
    });
  } catch (error) {
    return travelPlannerJson_({ ok: false, error: error.message });
  }
}


function doPost(e) {
  try {
    const docId = getSyncDocId_(e);
    const body = e && e.postData && e.postData.contents ? e.postData.contents : "";
    const data = JSON.parse(body);

    if (!data || typeof data !== "object") {
      throw new Error("Request body must be a JSON object.");
    }

    const updatedAt = data.updatedAt || new Date().toISOString();
    writeTravelPlannerRecord_(docId, JSON.stringify(data), updatedAt);

    return travelPlannerJson_({
      ok: true,
      docId: docId,
      updatedAt: updatedAt
    });
  } catch (error) {
    return travelPlannerJson_({ ok: false, error: error.message });
  }
}


function getSyncDocId_(e) {
  const docId = e && e.parameter && e.parameter.doc ? String(e.parameter.doc).trim() : "default";
  return docId || "default";
}


function getTravelPlannerSyncSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TRAVEL_PLANNER_SYNC_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(TRAVEL_PLANNER_SYNC_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([["docId", "updatedAt", "json"]]);
    sheet.hideSheet();
  }

  return sheet;
}


function readTravelPlannerRecord_(docId) {
  const sheet = getTravelPlannerSyncSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === docId) {
      return {
        row: i + 2,
        updatedAt: values[i][1],
        json: values[i][2]
      };
    }
  }

  return null;
}


function writeTravelPlannerRecord_(docId, json, updatedAt) {
  const sheet = getTravelPlannerSyncSheet_();
  const existing = readTravelPlannerRecord_(docId);
  const row = existing ? existing.row : sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 3).setValues([[docId, updatedAt, json]]);
}


function travelPlannerJson_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
