// 產生旅程路線地圖用的 SVG 國界路徑與城市座標（輸出 JSON 給 build_site.py 用）
import fs from 'fs';
import { feature } from 'topojson-client';
import { geoMercator, geoPath } from 'd3-geo';

const topo = JSON.parse(fs.readFileSync('node_modules/world-atlas/countries-50m.json'));
const all = feature(topo, topo.objects.countries).features;
const W = 760, H = 420;
// 顯示範圍（經度 9.6–20.6、緯度 45.9–49.4）
const box = { type: 'Feature', geometry: { type: 'Polygon', coordinates: [[[9.6, 45.9], [9.6, 49.4], [20.6, 49.4], [20.6, 45.9], [9.6, 45.9]]] } };
const proj = geoMercator().fitExtent([[0, 0], [W, H]], box);
const path = geoPath(proj);
const NAMES = { Austria: '奧地利', Germany: '德國', Hungary: '匈牙利', Czechia: '捷克', Slovakia: '斯洛伐克', Slovenia: '斯洛維尼亞', Italy: '義大利', Switzerland: '瑞士', Croatia: '克羅埃西亞', Liechtenstein: '' };
const LABEL = { Austria: [14.6, 47.55], Germany: [11.9, 48.3], Hungary: [18.9, 46.75], Czechia: [15.2, 49.05], Slovakia: [18.7, 48.9], Slovenia: [14.8, 46.12], Italy: [12.5, 46.3], Switzerland: [9.9, 46.7] };
const countries = all.filter(f => f.properties.name in NAMES).map(f => ({
  name: f.properties.name, zh: NAMES[f.properties.name], d: path(f),
  label: LABEL[f.properties.name] ? proj(LABEL[f.properties.name]).map(Math.round) : null,
}));
const cities = JSON.parse(fs.readFileSync(process.argv[2]));
for (const c of cities) c.xy = proj([c.lon, c.lat]).map(v => Math.round(v * 10) / 10);
fs.writeFileSync(process.argv[3], JSON.stringify({ W, H, countries, cities }));
console.log(countries.map(c => c.name).join(','), cities.map(c => c.xy));
