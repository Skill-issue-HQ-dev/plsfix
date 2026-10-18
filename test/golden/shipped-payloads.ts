// Payloads written by the codec that shipped as v2.9.25 (61f6a19): byte for byte
// what `encodePayload` produced for a picture link with a chart. Shared by the
// tests that must read them exactly as that version did.

export const SHIPPED_STACKED =
  '{"v":1,"kind":"picture","mime":"image/png","width":800,"height":400,"png":"iVBORw0KGgo=","src":{"workbook":"Model_v4.xlsx","sheet":"Model","ref":"B4:F12","anchor":"PLSFIX_LINK_00000000"},"pushedAt":"2026-09-12T08:00:00.000Z","hash":"0000000000000000000000000000000000000000000000000000000000000000","chart":{"v":1,"kind":"stackedColumn","title":"Revenue by segment","categories":["2024A","2025E","2026E"],"series":[{"name":"Retail","values":[60,70,80],"labels":["60","70","80"],"colors":["#2EC4B6","#2EC4B6","#2EC4B6"]},{"name":"Corporate","values":[40,50,-10],"labels":["40","50","(10)"],"colors":["#14213D","#14213D","#14213D"]}],"font":"Aptos Narrow","ink":"#282623","titleColor":"#14213D"}}';
export const SHIPPED_TORNADO =
  '{"v":1,"kind":"picture","mime":"image/png","width":800,"height":400,"png":"iVBORw0KGgo=","src":{"workbook":"Model_v4.xlsx","sheet":"Model","ref":"B4:F12","anchor":"PLSFIX_LINK_00000000"},"pushedAt":"2026-09-12T08:00:00.000Z","hash":"0000000000000000000000000000000000000000000000000000000000000000","chart":{"v":1,"kind":"bar","title":null,"categories":["Low","High"],"series":[{"name":"Delta","values":[-40,40],"labels":["(40)","40"],"colors":["#B27E54","#B27E54"]}],"font":"Aptos Narrow","ink":"#282623","titleColor":"#14213D","overlap":true}}';
