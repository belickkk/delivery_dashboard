import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  ResponsiveContainer, LineChart, Line, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ComposedChart, ScatterChart, Scatter, ZAxis, Cell
} from "recharts";
import { UploadCloud, Truck, Clock, ChefHat, AlertTriangle, Timer, Settings2, MapPin, Receipt, Percent } from "lucide-react";

// ---------- Restaurant locations ----------

const LOCATIONS = {
  loc1: { name: "Локал 1", address: "al. Jana Pawła II 45A (Wola/Centrum)", color: "#E8432C" },
  loc2: { name: "Локал 2", address: "Al. Rzeczypospolitej 8 (Wilanów)", color: "#5B7FBF" },
  loc3: { name: "Локал 3", address: "ul. Remiszewska 14a (Targówek)", color: "#4C7A5C" },
};
const UNKNOWN_COLOR = "#6B6E76";

// Zone -> nearest location + straight-line distance (km). Used as a fallback for
// direct-courier deliveries where the real customer address isn't the restaurant's own.
const DEFAULT_ZONE_MAP = {
  "Zoliborz": { loc: "loc1", km: 2.9 },
  "Wilanów": { loc: "loc2", km: 2.6 },
  "Mokotów": { loc: "loc2", km: 4.7 },
  "Ochota": { loc: "loc1", km: 4.0 },
  "Ząbki-Marki": { loc: "loc3", km: 3.9 },
  "Ursynów": { loc: "loc2", km: 3.9 },
  "Targówek": { loc: "loc3", km: 1.4 },
  "Środmeście": { loc: "loc1", km: 2.2 },
  "Ursus-piastów": { loc: "loc1", km: 9.1 },
  "Praga połnoc": { loc: "loc3", km: 2.2 },
  "Wola": { loc: "loc1", km: 2.6 },
  "Praga południe-Gocław": { loc: "loc3", km: 5.6 },
  "Białołęka": { loc: "loc3", km: 7.3 },
  "Bemowo": { loc: "loc1", km: 5.4 },
  "Bielany": { loc: "loc1", km: 6.9 },
  "Włochy-Raszyn": { loc: "loc1", km: 7.7 },
  "Wawer-Radośc-Miedzylesie (wk)": { loc: "loc2", km: 8.2 },
  "Wołomin (wk)": { loc: "loc3", km: 14.7 },
  "Rembertów-Wesoła": { loc: "loc3", km: 6.0 },
};

const DISTANCE_BUCKETS = [
  { max: 3, label: "0–3 км" },
  { max: 6, label: "3–6 км" },
  { max: 10, label: "6–10 км" },
  { max: Infinity, label: "10+ км" },
];

const PLATFORMS = {
  DIRECT: { name: "Прямая доставка", color: "#E8432C" },
  GLOVO: { name: "Glovo", color: "#D98F3D" },
  UBER: { name: "Uber Eats", color: "#5B7FBF" },
  BOLT: { name: "Bolt Food", color: "#4C7A5C" },
  WOLT: { name: "Wolt", color: "#8A7FE0" },
};
const PAYMENT_LABELS = {
  cash: "Наличные",
  card_on_delivery: "Карта при получении",
  online: "Онлайн",
  aggregator: "Через агрегатора",
};
const PAYMENT_COLORS = {
  cash: "#D98F3D", card_on_delivery: "#5B7FBF", online: "#4C7A5C", aggregator: "#8A7FE0",
};

// ---------- Timings-file parsing ----------

const MONTHS = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
  sty: 0, lut: 1, kwi: 3, maj: 4, cze: 5, lip: 6, sie: 7, wrz: 8, "paź": 9, lis: 10, gru: 11,
  янв: 0, фев: 1, мар: 2, апр: 3, май: 4, июн: 5, июл: 6, авг: 7, сен: 8, окт: 9, ноя: 10, дек: 11
};

function parseTimeDur(str) {
  if (!str || typeof str !== "string") return { time: null, min: null };
  const m = str.match(/(\d{1,2}):(\d{2})\s*\(\s*(-|\d+)\s*(?:мин)?\s*\)/i);
  if (!m) return { time: null, min: null };
  return { time: `${m[1]}:${m[2]}`, min: m[3] === "-" ? null : parseInt(m[3], 10) };
}
function parseTotalMin(str) {
  if (!str || typeof str !== "string") return null;
  const m = str.match(/(\d+)\s*мин/i);
  return m ? parseInt(m[1], 10) : null;
}
function parseOrderNo(str) {
  if (!str || typeof str !== "string") return { id: null };
  const m = str.match(/#?(\S+)/);
  return { id: m ? m[1] : str };
}
function parseDateField(str) {
  if (!str || typeof str !== "string") return { date: null, isPreorder: false };
  const isPreorder = /предзаказ/i.test(str);
  const cleaned = str.replace(/,?\s*предзаказ/i, "").trim();
  const m = cleaned.match(/^(\d{1,2})\s+([a-zA-Zа-яёА-ЯЁźżąćęłńóśŹŻĄĆĘŁŃÓŚ]+)\.?\s+(\d{1,2}):(\d{2})$/i);
  if (!m) return { date: null, isPreorder };
  const day = parseInt(m[1], 10);
  const month = MONTHS[m[2].toLowerCase()];
  if (month === undefined) return { date: null, isPreorder };
  const hh = parseInt(m[3], 10), mm = parseInt(m[4], 10);
  const now = new Date();
  let year = now.getFullYear();
  let candidate = new Date(year, month, day, hh, mm);
  if (candidate.getTime() > now.getTime() + 24 * 3600 * 1000) {
    year -= 1;
    candidate = new Date(year, month, day, hh, mm);
  }
  return { date: candidate, isPreorder };
}
function normalizeRow(row) {
  const out = {};
  Object.keys(row).forEach(k => {
    const nk = k.replace(/^\uFEFF/, "").trim();
    out[nk] = typeof row[k] === "string" ? row[k].trim() : row[k];
  });
  return out;
}
function looksGarbled(text) {
  const firstLine = (text.split(/\r?\n/)[0] || "");
  const hasReplacement = (firstLine.match(/\uFFFD/g) || []).length > 2;
  const hasKnownHeader = /дата|заказ|кухня|курьер|сумма/i.test(firstLine);
  return hasReplacement || !hasKnownHeader;
}
async function readRowsObjects(file) {
  const buf = await file.arrayBuffer();
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv")) {
    let text = new TextDecoder("utf-8").decode(buf);
    if (looksGarbled(text)) { try { text = new TextDecoder("windows-1251").decode(buf); } catch (e) {} }
    const delim = name.endsWith(".tsv") ? "\t" : undefined;
    return Papa.parse(text, { header: true, skipEmptyLines: true, delimiter: delim }).data;
  }
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: null });
}
async function readRowsArray(file) {
  const buf = await file.arrayBuffer();
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv")) {
    let text = new TextDecoder("utf-8").decode(buf);
    if (looksGarbled(text)) { try { text = new TextDecoder("windows-1251").decode(buf); } catch (e) {} }
    const delim = name.endsWith(".tsv") ? "\t" : undefined;
    return Papa.parse(text, { header: false, skipEmptyLines: true, delimiter: delim }).data;
  }
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
}

function buildTimingRecord(rawRow) {
  const row = normalizeRow(rawRow);
  const dateInfo = parseDateField(row["Дата"]);
  const orderInfo = parseOrderNo(row["Заказ"]);
  const accepted = parseTimeDur(row["Принят"]);
  const kitchen = parseTimeDur(row["Кухня"]);
  const courier = parseTimeDur(row["Курьер"]);
  const completedRaw = row["Завершён"];
  const completedMatch = typeof completedRaw === "string" ? completedRaw.match(/^(\d{1,2}:\d{2})/) : null;
  const isCompleted = !!completedMatch;
  const totalMin = parseTotalMin(row["Затрачено времени"]);
  const zoneRaw = row["Зона доставки"];
  const zone = zoneRaw && String(zoneRaw).trim() ? String(zoneRaw).trim() : null;
  const isPickup = !zone;

  let transitMin = null;
  if (totalMin != null && accepted.min != null && kitchen.min != null && courier.min != null) {
    const rem = totalMin - accepted.min - kitchen.min - courier.min;
    if (rem >= 0) transitMin = rem;
  }
  let pickupWaitMin = null;
  if (isPickup && totalMin != null && accepted.min != null && kitchen.min != null) {
    const rem = totalMin - accepted.min - kitchen.min;
    if (rem >= 0) pickupWaitMin = rem;
  }
  const valid = !!(dateInfo.date && orderInfo.id);
  return {
    valid, date: dateInfo.date, isPreorder: dateInfo.isPreorder, orderId: orderInfo.id,
    acceptMin: accepted.min, kitchenMin: kitchen.min, courierMin: courier.min,
    transitMin, pickupWaitMin, totalMin, isCompleted, zone, isPickup,
  };
}

// ---------- Carts-file (detailed report) parsing ----------

function parseCartNum(v) {
  if (v == null) return null;
  const s = String(v).trim();
  if (s === "") return null;
  const n = parseFloat(s.replace(/\s/g, "").replace(",", "."));
  return isNaN(n) ? null : n;
}
function parseCartDate(s) {
  if (!s) return null;
  const m = String(s).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const [, d, mo, y, h, mi] = m;
  return new Date(parseInt(y, 10), parseInt(mo, 10) - 1, parseInt(d, 10), parseInt(h, 10), parseInt(mi, 10));
}
function detectPlatform(name) {
  if (!name) return null;
  const m = String(name).match(/^(GLOVO|UBER|WOLT|BOLT)\s*#/i);
  return m ? m[1].toUpperCase() : null;
}
function resolvePaymentMethod(o) {
  if (o.cash != null) return { method: "cash", amount: o.cash };
  if (o.cardOnDelivery != null) return { method: "card_on_delivery", amount: o.cardOnDelivery };
  if (o.online1 != null || o.online2 != null) return { method: "online", amount: (o.online1 || 0) + (o.online2 || 0) };
  if (o.portali != null) return { method: "aggregator", amount: o.portali };
  return { method: null, amount: null };
}
function normalizeAddr(s) {
  if (!s) return "";
  return String(s).toLowerCase()
    .replace(/ą/g, "a").replace(/ć/g, "c").replace(/ę/g, "e").replace(/ł/g, "l")
    .replace(/ń/g, "n").replace(/ó/g, "o").replace(/ś/g, "s").replace(/ź/g, "z").replace(/ż/g, "z")
    .replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}
function matchRestaurantByAddress(address) {
  const n = normalizeAddr(address);
  if (!n) return null;
  if (n.includes("remiszewska")) return "loc3";
  if (n.includes("rzeczypospolitej")) return "loc2";
  if (n.includes("jana pawla")) return "loc1";
  return null;
}
function buildCartOrders(rows) {
  const orders = [];
  let current = null;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const numRaw = row[0];
    if (numRaw && String(numRaw).trim() !== "") {
      if (current) orders.push(current);
      current = {
        orderId: parseInt(String(numRaw).trim(), 10), dateRaw: row[1],
        sum: parseCartNum(row[2]), discount: parseCartNum(row[3]), deliveryCostCharged: parseCartNum(row[4]),
        cash: parseCartNum(row[5]), cardOnDelivery: parseCartNum(row[6]),
        online1: parseCartNum(row[7]), online2: parseCartNum(row[8]), portali: parseCartNum(row[9]),
        phone: row[10] || null, name: row[11] || null, address: row[12] || null,
        bonus: parseCartNum(row[16]), items: [],
      };
      const dish = row[13], qty = row[14], price = row[15];
      if (dish) current.items.push({ dish: String(dish).trim(), qty: parseCartNum(qty) || 0, price: parseCartNum(price) || 0 });
    } else if (current) {
      const dish = row[13], qty = row[14], price = row[15];
      if (dish && String(dish).trim() !== "") {
        current.items.push({ dish: String(dish).trim(), qty: parseCartNum(qty) || 0, price: parseCartNum(price) || 0 });
      }
    }
  }
  if (current) orders.push(current);
  return orders;
}
function finalizeCartOrder(o) {
  const date = parseCartDate(o.dateRaw);
  const platform = detectPlatform(o.name);
  const pay = resolvePaymentMethod(o);
  const itemCount = o.items.reduce((s, it) => s + (it.qty || 0), 0);
  return {
    orderId: o.orderId, date, sum: o.sum, discount: o.discount, deliveryCostCharged: o.deliveryCostCharged,
    paymentMethod: pay.method, paymentAmount: pay.amount, platform, phone: o.phone, address: o.address,
    items: o.items, itemCount, bonus: o.bonus, isDuplicate: false,
  };
}
function markDuplicates(cartOrders) {
  const byPhone = new Map();
  cartOrders.forEach(o => {
    if (!o.phone) return;
    if (!byPhone.has(o.phone)) byPhone.set(o.phone, []);
    byPhone.get(o.phone).push(o);
  });
  let count = 0;
  byPhone.forEach(list => {
    if (list.length < 2) return;
    list.sort((a, b) => (a.date && b.date ? a.date - b.date : 0));
    for (let i = 0; i < list.length - 1; i++) {
      const a = list[i], b = list[i + 1];
      if (a.sum != null && b.sum != null && Math.abs(a.sum - b.sum) < 0.5 && a.date && b.date) {
        const diffMin = Math.abs(a.date - b.date) / 60000;
        if (diffMin <= 20 && !a.isDuplicate && !b.isDuplicate) {
          const toExclude = (a.platform && !b.platform) ? b : ((b.platform && !a.platform) ? a : b);
          toExclude.isDuplicate = true;
          count++;
        }
      }
    }
  });
  return count;
}

// ---------- Small stat utils ----------
const avg = arr => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null);
const sum = arr => arr.reduce((a, b) => a + b, 0);
const fmtMin = v => (v == null ? "—" : `${Math.round(v)} мин`);
const fmtKm = v => (v == null ? "—" : `${v.toFixed(1)} км`);
const fmtPln = v => (v == null ? "—" : `${Math.round(v).toLocaleString("ru-RU")} zł`);
const fmtDate = d => d.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit" });
const pad2 = n => String(n).padStart(2, "0");
const locColor = loc => (loc && LOCATIONS[loc] ? LOCATIONS[loc].color : UNKNOWN_COLOR);
const locName = loc => (loc && LOCATIONS[loc] ? LOCATIONS[loc].name : "Без точки");

const ACCEPT_COLOR = "#D98F3D";
const KITCHEN_COLOR = "#E8432C";
const COURIER_COLOR = "#5B7FBF";
const TRANSIT_COLOR = "#3E6B4F";

// ---------- UI subcomponents ----------

function Kpi({ icon: Icon, label, value, sub, accent }) {
  return (
    <div className="dod-card" style={{ padding: "18px 20px" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <span className="dod-eyebrow">{label}</span>
        {Icon && <Icon size={16} color={accent || "var(--muted)"} strokeWidth={2} />}
      </div>
      <div className="dod-kpi-value" style={{ color: accent || "var(--text)" }}>{value}</div>
      {sub && <div className="dod-kpi-sub">{sub}</div>}
    </div>
  );
}
function SectionTitle({ children, note }) {
  return (
    <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 6 }}>
      <h2 className="dod-h2">{children}</h2>
      {note && <span className="dod-note">{note}</span>}
    </div>
  );
}
const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className="dod-tooltip">
      <div className="dod-tooltip-label">{label}</div>
      {payload.map((p, i) => (
        p.value != null && p.value !== 0 ? (
          <div key={i} className="dod-tooltip-row">
            <span style={{ color: p.color || p.fill }}>{p.name}</span>
            <span>{typeof p.value === "number" ? Math.round(p.value * 100) / 100 : p.value}</span>
          </div>
        ) : null
      ))}
    </div>
  );
};
const ZoneScatterTooltip = ({ active, payload }) => {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <div className="dod-tooltip">
      <div className="dod-tooltip-label">{d.zone}</div>
      <div className="dod-tooltip-row"><span>Локал</span><span>{locName(d.loc)}</span></div>
      <div className="dod-tooltip-row"><span>Расстояние</span><span>{fmtKm(d.distanceKm)}</span></div>
      <div className="dod-tooltip-row"><span>Ср. время</span><span>{fmtMin(d.avgTime)}</span></div>
      <div className="dod-tooltip-row"><span>Заказов</span><span>{d.count}</span></div>
    </div>
  );
};

// ---------- Main component ----------

export default function DeliveryOpsDashboard() {
  const [timingRows, setTimingRows] = useState([]);
  const [timingFileName, setTimingFileName] = useState(null);
  const [timingSkipped, setTimingSkipped] = useState(0);
  const [timingError, setTimingError] = useState(null);
  const [timingLoading, setTimingLoading] = useState(false);

  const [cartOrders, setCartOrders] = useState([]);
  const [cartFileName, setCartFileName] = useState(null);
  const [cartError, setCartError] = useState(null);
  const [cartLoading, setCartLoading] = useState(false);
  const [dupCount, setDupCount] = useState(0);

  const [channelFilter, setChannelFilter] = useState("all");
  const [locationFilter, setLocationFilter] = useState("all");
  const [zoneOverrides, setZoneOverrides] = useState({});
  const [showMapping, setShowMapping] = useState(false);

  const handleTimingFile = useCallback(async (file) => {
    setTimingLoading(true); setTimingError(null);
    try {
      const rows = await readRowsObjects(file);
      const built = rows.map(buildTimingRecord);
      const valid = built.filter(r => r.valid);
      setTimingSkipped(built.length - valid.length);
      if (valid.length === 0) {
        setTimingError("Не удалось распознать файл таймингов. Проверьте столбцы: Дата, Заказ, Принят, Кухня, Курьер, Завершён, Затрачено времени, Зона доставки.");
      } else {
        setTimingRows(valid); setTimingFileName(file.name);
      }
    } catch (e) {
      setTimingError("Ошибка чтения файла: " + e.message);
    } finally { setTimingLoading(false); }
  }, []);

  const handleCartFile = useCallback(async (file) => {
    setCartLoading(true); setCartError(null);
    try {
      const rows = await readRowsArray(file);
      const raw = buildCartOrders(rows);
      const finalized = raw.map(finalizeCartOrder);
      const dc = markDuplicates(finalized);
      if (finalized.length === 0) {
        setCartError("Не удалось распознать файл детального отчёта. Проверьте, что это тот же шаблон (колонки #, Дата, Сумма, ... Блюда, Кол-во, Цена).");
      } else {
        setCartOrders(finalized); setCartFileName(file.name); setDupCount(dc);
      }
    } catch (e) {
      setCartError("Ошибка чтения файла: " + e.message);
    } finally { setCartLoading(false); }
  }, []);

  const onTimingInput = e => { const f = e.target.files && e.target.files[0]; if (f) handleTimingFile(f); e.target.value = ""; };
  const onCartInput = e => { const f = e.target.files && e.target.files[0]; if (f) handleCartFile(f); e.target.value = ""; };

  const cartByOrderId = useMemo(() => {
    const m = new Map();
    cartOrders.forEach(o => { if (!m.has(o.orderId)) m.set(o.orderId, o); });
    return m;
  }, [cartOrders]);

  const zonesInData = useMemo(() => {
    const set = new Set();
    timingRows.forEach(r => { if (!r.isPickup && r.zone) set.add(r.zone); });
    return [...set].sort();
  }, [timingRows]);

  const setZoneOverride = (zone, field, value) => {
    setZoneOverrides(prev => { const cur = prev[zone] || {}; return { ...prev, [zone]: { ...cur, [field]: value } }; });
  };

  // enrich timing records with location + distance + matched cart data
  const enriched = useMemo(() => {
    return timingRows.map(r => {
      const orderIdNum = parseInt(r.orderId, 10);
      const cart = !isNaN(orderIdNum) ? cartByOrderId.get(orderIdNum) : undefined;
      const cartActive = cart && !cart.isDuplicate ? cart : null;

      let location = null, distanceKm = null;
      const addrLoc = cart ? matchRestaurantByAddress(cart.address) : null;
      if (addrLoc) {
        location = addrLoc; // definitive: pickup counter or aggregator hand-off point
      } else if (!r.isPickup && r.zone) {
        const override = zoneOverrides[r.zone];
        const base = DEFAULT_ZONE_MAP[r.zone];
        const loc = override && override.loc !== undefined ? override.loc : (base ? base.loc : null);
        const km = override && override.km !== undefined ? override.km : (base ? base.km : null);
        location = loc || null;
        distanceKm = km != null ? Number(km) : null;
      }
      const date = (cart && cart.date) ? cart.date : r.date;
      return { ...r, date, location, distanceKm, cart: cart || null, cartActive };
    });
  }, [timingRows, cartByOrderId, zoneOverrides]);

  const filtered = useMemo(() => {
    return enriched.filter(r => {
      if (channelFilter === "delivery" && r.isPickup) return false;
      if (channelFilter === "pickup" && !r.isPickup) return false;
      if (locationFilter !== "all") {
        if (locationFilter === "unknown") { if (r.location) return false; }
        else if (r.location !== locationFilter) return false;
      }
      return true;
    });
  }, [enriched, channelFilter, locationFilter]);

  const hasTimings = timingRows.length > 0;
  const hasCarts = cartOrders.length > 0;

  const stats = useMemo(() => {
    if (!hasTimings || !filtered.length) return null;
    const total = filtered.length;
    const deliveryCount = filtered.filter(r => !r.isPickup).length;
    const pickupCount = total - deliveryCount;
    const completed = filtered.filter(r => r.isCompleted).length;
    const noTotalMark = total - completed;
    const preorders = filtered.filter(r => r.isPreorder).length;

    const avgTotal = avg(filtered.filter(r => !r.isPreorder && r.totalMin != null).map(r => r.totalMin));
    const avgAccept = avg(filtered.filter(r => !r.isPreorder && r.acceptMin != null).map(r => r.acceptMin));
    const avgKitchen = avg(filtered.filter(r => r.kitchenMin != null).map(r => r.kitchenMin));
    const avgCourier = avg(filtered.filter(r => !r.isPickup && r.courierMin != null).map(r => r.courierMin));
    const avgTransit = avg(filtered.filter(r => !r.isPickup && r.transitMin != null).map(r => r.transitMin));
    const avgDistance = avg(filtered.filter(r => r.distanceKm != null).map(r => r.distanceKm));

    // revenue (only orders with an active, non-duplicate cart match)
    const withRevenue = filtered.filter(r => r.cartActive && r.cartActive.sum != null);
    const totalRevenue = sum(withRevenue.map(r => r.cartActive.sum));
    const avgOrderValue = withRevenue.length ? totalRevenue / withRevenue.length : null;
    const discounted = filtered.filter(r => r.cartActive && r.cartActive.discount != null && r.cartActive.discount > 0);
    const totalDiscount = sum(discounted.map(r => r.cartActive.discount));
    const avgDeliveryFee = avg(filtered.filter(r => !r.isPickup && r.cartActive && r.cartActive.deliveryCostCharged != null).map(r => r.cartActive.deliveryCostCharged));
    const unmatchedRevenue = filtered.filter(r => !r.cart).length;

    // daily trend: count + revenue
    const byDay = new Map();
    filtered.forEach(r => {
      if (!r.date) return;
      const key = `${r.date.getFullYear()}-${pad2(r.date.getMonth() + 1)}-${pad2(r.date.getDate())}`;
      if (!byDay.has(key)) byDay.set(key, { key, date: r.date, count: 0, revenue: 0 });
      const d = byDay.get(key);
      d.count += 1;
      if (r.cartActive && r.cartActive.sum != null) d.revenue += r.cartActive.sum;
    });
    const daily = [...byDay.values()].sort((a, b) => a.date - b.date)
      .map(d => ({ label: fmtDate(d.date), count: d.count, revenue: Math.round(d.revenue) }));

    // hourly load per location
    const byHourLoc = Array.from({ length: 24 }, (_, h) => ({ hour: pad2(h), loc1: 0, loc2: 0, loc3: 0, unknown: 0 }));
    filtered.forEach(r => {
      if (!r.date) return;
      const key = r.location && LOCATIONS[r.location] ? r.location : "unknown";
      byHourLoc[r.date.getHours()][key] += 1;
    });

    // per-location summary
    const byLocation = {};
    Object.keys(LOCATIONS).forEach(k => { byLocation[k] = { count: 0, totalSum: 0, totalN: 0, distSum: 0, distN: 0, zones: new Set(), revSum: 0, revN: 0 }; });
    filtered.forEach(r => {
      if (!r.location || !byLocation[r.location]) return;
      const b = byLocation[r.location];
      b.count += 1;
      if (r.zone) b.zones.add(r.zone);
      if (!r.isPreorder && r.totalMin != null) { b.totalSum += r.totalMin; b.totalN += 1; }
      if (r.distanceKm != null) { b.distSum += r.distanceKm; b.distN += 1; }
      if (r.cartActive && r.cartActive.sum != null) { b.revSum += r.cartActive.sum; b.revN += 1; }
    });
    const unassignedCount = filtered.filter(r => !r.location).length;

    // zones
    const zoneMap = new Map();
    filtered.filter(r => !r.isPickup && r.zone).forEach(r => {
      if (!zoneMap.has(r.zone)) zoneMap.set(r.zone, { zone: r.zone, count: 0, totalSum: 0, totalN: 0, loc: r.location, distanceKm: r.distanceKm });
      const z = zoneMap.get(r.zone);
      z.count += 1;
      if (!r.isPreorder && r.totalMin != null) { z.totalSum += r.totalMin; z.totalN += 1; }
    });
    const zoneByCount = [...zoneMap.values()].sort((a, b) => b.count - a.count).slice(0, 12)
      .map(z => ({ zone: z.zone, count: z.count, loc: z.loc }));
    const zoneByTime = [...zoneMap.values()].filter(z => z.totalN >= 3)
      .map(z => ({ zone: z.zone, avgTime: Math.round(z.totalSum / z.totalN), loc: z.loc }))
      .sort((a, b) => b.avgTime - a.avgTime).slice(0, 12);
    const zoneScatter = [...zoneMap.values()].filter(z => z.totalN >= 2 && z.distanceKm != null)
      .map(z => ({ zone: z.zone, distanceKm: z.distanceKm, avgTime: Math.round(z.totalSum / z.totalN), count: z.count, loc: z.loc }));

    // distance buckets
    const distBuckets = DISTANCE_BUCKETS.map(b => ({ ...b, sum: 0, n: 0, count: 0 }));
    filtered.forEach(r => {
      if (r.distanceKm == null) return;
      const bucket = distBuckets.find(b => r.distanceKm <= b.max);
      if (!bucket) return;
      bucket.count += 1;
      if (!r.isPreorder && r.totalMin != null) { bucket.sum += r.totalMin; bucket.n += 1; }
    });
    const distanceStats = distBuckets.map(b => ({ label: b.label, avgTime: b.n ? Math.round(b.sum / b.n) : null, count: b.count }));

    // stage breakdown
    const stageBars = [{
      name: "Доставка",
      "Принятие": Math.round(avgAccept || 0), "Кухня": Math.round(avgKitchen || 0),
      "Ожидание курьера": Math.round(avgCourier || 0), "В пути": Math.round(avgTransit || 0),
    }];
    const pickupAvgAccept = avg(filtered.filter(r => r.isPickup && !r.isPreorder && r.acceptMin != null).map(r => r.acceptMin));
    const pickupAvgKitchen = avg(filtered.filter(r => r.isPickup && r.kitchenMin != null).map(r => r.kitchenMin));
    const pickupAvgWait = avg(filtered.filter(r => r.isPickup && r.pickupWaitMin != null).map(r => r.pickupWaitMin));
    if (pickupCount > 0) {
      stageBars.push({
        name: "Самовывоз", "Принятие": Math.round(pickupAvgAccept || 0), "Кухня": Math.round(pickupAvgKitchen || 0),
        "Ожидание курьера": 0, "В пути": 0, "Ожидание забора": Math.round(pickupAvgWait || 0),
      });
    }

    // channel / platform mix
    const platformMap = {};
    filtered.forEach(r => {
      let key;
      if (r.isPickup) key = "PICKUP";
      else key = (r.cartActive && r.cartActive.platform) ? r.cartActive.platform : "DIRECT";
      if (!platformMap[key]) platformMap[key] = { key, count: 0, revenue: 0, revN: 0 };
      platformMap[key].count += 1;
      if (r.cartActive && r.cartActive.sum != null) { platformMap[key].revenue += r.cartActive.sum; platformMap[key].revN += 1; }
    });
    const platformStats = Object.values(platformMap).map(p => ({
      ...p, name: p.key === "PICKUP" ? "Самовывоз" : (PLATFORMS[p.key] ? PLATFORMS[p.key].name : p.key),
      color: p.key === "PICKUP" ? UNKNOWN_COLOR : (PLATFORMS[p.key] ? PLATFORMS[p.key].color : UNKNOWN_COLOR),
      aov: p.revN ? Math.round(p.revenue / p.revN) : null,
    })).sort((a, b) => b.count - a.count);

    // payment method mix
    const payMap = {};
    filtered.forEach(r => {
      if (!r.cartActive || !r.cartActive.paymentMethod) return;
      const k = r.cartActive.paymentMethod;
      if (!payMap[k]) payMap[k] = { key: k, count: 0, amount: 0 };
      payMap[k].count += 1;
      payMap[k].amount += r.cartActive.paymentAmount || 0;
    });
    const paymentStats = Object.values(payMap).map(p => ({
      ...p, name: PAYMENT_LABELS[p.key] || p.key, color: PAYMENT_COLORS[p.key] || UNKNOWN_COLOR, amount: Math.round(p.amount),
    })).sort((a, b) => b.count - a.count);

    // top dishes
    const dishMap = new Map();
    filtered.forEach(r => {
      if (!r.cartActive) return;
      r.cartActive.items.forEach(it => {
        if (!it.dish) return;
        if (!dishMap.has(it.dish)) dishMap.set(it.dish, { dish: it.dish, qty: 0, revenue: 0 });
        const d = dishMap.get(it.dish);
        d.qty += it.qty || 0;
        d.revenue += (it.qty || 0) * (it.price || 0);
      });
    });
    const topDishes = [...dishMap.values()].sort((a, b) => b.qty - a.qty).slice(0, 10);

    // slowest orders
    const slowest = filtered.filter(r => !r.isPreorder && r.totalMin != null).sort((a, b) => b.totalMin - a.totalMin).slice(0, 12);

    return {
      total, deliveryCount, pickupCount, completed, noTotalMark, preorders, avgDistance,
      avgTotal, avgAccept, avgKitchen, avgCourier, avgTransit,
      totalRevenue, avgOrderValue, totalDiscount, discountedCount: discounted.length, avgDeliveryFee, unmatchedRevenue,
      daily, byHourLoc, byLocation, unassignedCount,
      zoneByCount, zoneByTime, zoneScatter, distanceStats, stageBars, slowest, platformStats, paymentStats, topDishes,
    };
  }, [filtered, hasTimings]);

  // Lightweight stats when only the carts file is loaded (no timings uploaded yet)
  const cartOnlyStats = useMemo(() => {
    if (hasTimings || !hasCarts) return null;
    const active = cartOrders.filter(o => !o.isDuplicate);
    const withSum = active.filter(o => o.sum != null);
    const totalRevenue = sum(withSum.map(o => o.sum));
    const avgOrderValue = withSum.length ? totalRevenue / withSum.length : null;
    const discounted = active.filter(o => o.discount != null && o.discount > 0);
    const totalDiscount = sum(discounted.map(o => o.discount));

    const platformMap = {};
    active.forEach(o => {
      const key = o.platform || "DIRECT";
      if (!platformMap[key]) platformMap[key] = { key, count: 0, revenue: 0, revN: 0 };
      platformMap[key].count += 1;
      if (o.sum != null) { platformMap[key].revenue += o.sum; platformMap[key].revN += 1; }
    });
    const platformStats = Object.values(platformMap).map(p => ({
      ...p, name: PLATFORMS[p.key] ? PLATFORMS[p.key].name : p.key,
      color: PLATFORMS[p.key] ? PLATFORMS[p.key].color : UNKNOWN_COLOR,
      aov: p.revN ? Math.round(p.revenue / p.revN) : null,
    })).sort((a, b) => b.count - a.count);

    const payMap = {};
    active.forEach(o => {
      if (!o.paymentMethod) return;
      if (!payMap[o.paymentMethod]) payMap[o.paymentMethod] = { key: o.paymentMethod, count: 0, amount: 0 };
      payMap[o.paymentMethod].count += 1;
      payMap[o.paymentMethod].amount += o.paymentAmount || 0;
    });
    const paymentStats = Object.values(payMap).map(p => ({
      ...p, name: PAYMENT_LABELS[p.key] || p.key, color: PAYMENT_COLORS[p.key] || UNKNOWN_COLOR, amount: Math.round(p.amount),
    })).sort((a, b) => b.count - a.count);

    const dishMap = new Map();
    active.forEach(o => o.items.forEach(it => {
      if (!it.dish) return;
      if (!dishMap.has(it.dish)) dishMap.set(it.dish, { dish: it.dish, qty: 0, revenue: 0 });
      const d = dishMap.get(it.dish);
      d.qty += it.qty || 0; d.revenue += (it.qty || 0) * (it.price || 0);
    }));
    const topDishes = [...dishMap.values()].sort((a, b) => b.qty - a.qty).slice(0, 10);

    const byDay = new Map();
    active.forEach(o => {
      if (!o.date) return;
      const key = `${o.date.getFullYear()}-${pad2(o.date.getMonth() + 1)}-${pad2(o.date.getDate())}`;
      if (!byDay.has(key)) byDay.set(key, { key, date: o.date, count: 0, revenue: 0 });
      const d = byDay.get(key); d.count += 1; d.revenue += o.sum || 0;
    });
    const daily = [...byDay.values()].sort((a, b) => a.date - b.date).map(d => ({ label: fmtDate(d.date), count: d.count, revenue: Math.round(d.revenue) }));

    return { total: active.length, totalRevenue, avgOrderValue, totalDiscount, discountedCount: discounted.length, platformStats, paymentStats, topDishes, daily };
  }, [hasTimings, hasCarts, cartOrders]);

  return (
    <div className="dod-root">
      <style>{`
        .dod-root {
          --ink: #121316; --card: #1B1D22; --card-border: #2A2D34; --text: #F2EEE4; --muted: #8D8B85;
          --vermillion: #E8432C; --vermillion-dim: rgba(232,67,44,0.14); --nori: #4C7A5C; --amber: #D98F3D;
          --mono: ui-monospace, "SF Mono", "Roboto Mono", Menlo, Consolas, monospace;
          --sans: ui-sans-serif, -apple-system, "Helvetica Neue", Arial, sans-serif;
          background: var(--ink); color: var(--text); font-family: var(--sans); min-height: 100%; padding: 28px; border-radius: 16px;
        }
        .dod-h2 { font-family: var(--sans); font-weight: 800; font-size: 13px; letter-spacing: 0.08em; text-transform: uppercase; color: var(--text); margin: 0; }
        .dod-note { font-family: var(--mono); font-size: 11px; color: var(--muted); }
        .dod-eyebrow { font-family: var(--sans); font-size: 11px; letter-spacing: 0.08em; text-transform: uppercase; color: var(--muted); font-weight: 700; }
        .dod-card { background: var(--card); border: 1px solid var(--card-border); border-radius: 12px; }
        .dod-kpi-value { font-family: var(--mono); font-size: 26px; font-weight: 600; line-height: 1.1; }
        .dod-kpi-sub { font-family: var(--mono); font-size: 11px; color: var(--muted); margin-top: 6px; }
        .dod-tooltip { background: #0E0F12; border: 1px solid var(--card-border); border-radius: 8px; padding: 8px 10px; font-family: var(--mono); font-size: 12px; }
        .dod-tooltip-label { color: var(--muted); margin-bottom: 4px; }
        .dod-tooltip-row { display: flex; justify-content: space-between; gap: 16px; }
        .dod-filter-btn { font-family: var(--sans); font-size: 12px; font-weight: 700; letter-spacing: 0.03em; padding: 7px 14px; border-radius: 8px; border: 1px solid var(--card-border); background: transparent; color: var(--muted); cursor: pointer; transition: all .15s ease; }
        .dod-filter-btn.active { background: var(--vermillion-dim); border-color: var(--vermillion); color: var(--vermillion); }
        .dod-filter-btn:hover { color: var(--text); }
        .dod-drop { border: 1.5px dashed var(--card-border); border-radius: 12px; padding: 20px; text-align: center; cursor: pointer; transition: border-color .15s ease; }
        .dod-drop:hover { border-color: var(--vermillion); }
        .dod-ticket { font-family: var(--mono); font-size: 12px; }
        .dod-ticket-row { display: grid; grid-template-columns: 80px 1fr 100px 55px; gap: 10px; align-items: center; padding: 9px 0; border-bottom: 1px dashed var(--card-border); }
        .dod-ticket-row:last-child { border-bottom: none; }
        .dod-seg-bar { display: flex; height: 6px; width: 100%; border-radius: 3px; overflow: hidden; background: #26282f; }
        .dod-stamp { width: 10px; height: 10px; background: var(--vermillion); border-radius: 2px; display: inline-block; margin-right: 10px; flex-shrink: 0; }
        .dod-loc-card { padding: 16px 18px; }
        .dod-loc-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; margin-right: 8px; flex-shrink: 0; }
        .dod-map-row { display: grid; grid-template-columns: 1fr 140px 100px; gap: 10px; align-items: center; padding: 7px 0; border-bottom: 1px solid var(--card-border); font-size: 12px; }
        .dod-map-row:last-child { border-bottom: none; }
        .dod-select, .dod-num { font-family: var(--mono); font-size: 12px; background: #16171c; border: 1px solid var(--card-border); color: var(--text); border-radius: 6px; padding: 5px 8px; width: 100%; }
        .dod-dish-row { display: flex; justify-content: space-between; padding: 7px 0; border-bottom: 1px dashed var(--card-border); font-family: var(--mono); font-size: 12px; }
        .dod-dish-row:last-child { border-bottom: none; }
        input[type=file] { display: none; }
      `}</style>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 14 }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          <span className="dod-stamp" />
          <div>
            <h1 style={{ fontFamily: "var(--sans)", fontWeight: 800, fontSize: 20, letterSpacing: "0.02em", textTransform: "uppercase", margin: 0 }}>Пульс доставки</h1>
            <div style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)", marginTop: 2 }}>
              {hasTimings ? `${timingFileName} · ${timingRows.length} заказов${timingSkipped ? ` · ${timingSkipped} не распознано` : ""}` : "тайминги не загружены"}
              {hasCarts ? ` · ${cartFileName} · ${cartOrders.length} строк отчёта${dupCount ? ` · ${dupCount} дублей` : ""}` : ""}
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {hasTimings && (
            <button className="dod-filter-btn" onClick={() => setShowMapping(s => !s)} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <Settings2 size={13} /> Точки и зоны
            </button>
          )}
          <label className="dod-filter-btn" style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <UploadCloud size={14} /> {hasTimings ? "Заменить тайминги" : "Загрузить тайминги"}
            <input type="file" accept=".csv,.tsv,.xlsx,.xls" onChange={onTimingInput} />
          </label>
          <label className="dod-filter-btn" style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <Receipt size={14} /> {hasCarts ? "Заменить отчёт" : "Загрузить детальный отчёт"}
            <input type="file" accept=".csv,.tsv,.xlsx,.xls" onChange={onCartInput} />
          </label>
        </div>
      </div>

      {hasTimings && (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
          <button className={`dod-filter-btn ${channelFilter === "all" ? "active" : ""}`} onClick={() => setChannelFilter("all")}>Все</button>
          <button className={`dod-filter-btn ${channelFilter === "delivery" ? "active" : ""}`} onClick={() => setChannelFilter("delivery")}>Доставка</button>
          <button className={`dod-filter-btn ${channelFilter === "pickup" ? "active" : ""}`} onClick={() => setChannelFilter("pickup")}>Самовывоз</button>
          <span style={{ width: 1, background: "var(--card-border)", margin: "0 4px" }} />
          <button className={`dod-filter-btn ${locationFilter === "all" ? "active" : ""}`} onClick={() => setLocationFilter("all")}>Все точки</button>
          {Object.keys(LOCATIONS).map(k => (
            <button key={k} className={`dod-filter-btn ${locationFilter === k ? "active" : ""}`} onClick={() => setLocationFilter(k)}>{LOCATIONS[k].name}</button>
          ))}
          <button className={`dod-filter-btn ${locationFilter === "unknown" ? "active" : ""}`} onClick={() => setLocationFilter("unknown")}>Без точки</button>
        </div>
      )}

      {showMapping && hasTimings && (
        <div className="dod-card" style={{ padding: 20, marginBottom: 22 }}>
          <SectionTitle note="влияет на все графики ниже">Привязка зон к точкам</SectionTitle>
          <div style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)", marginBottom: 12 }}>
            Точки определяются автоматически: для самовывоза и заказов через агрегатора — по адресу пикапа из детального отчёта (если он загружен). Для обычной доставки — по зоне, ближайшая точка по прямой. Можно поправить вручную.
          </div>
          <div className="dod-map-row" style={{ color: "var(--muted)", fontSize: 10, textTransform: "uppercase", letterSpacing: "0.05em" }}>
            <span>Зона</span><span>Точка</span><span>Км</span>
          </div>
          {zonesInData.map(zone => {
            const override = zoneOverrides[zone] || {};
            const base = DEFAULT_ZONE_MAP[zone] || {};
            const loc = override.loc !== undefined ? override.loc : (base.loc || "");
            const km = override.km !== undefined ? override.km : (base.km != null ? base.km : "");
            return (
              <div className="dod-map-row" key={zone}>
                <span style={{ color: "var(--text)" }}>{zone}</span>
                <select className="dod-select" value={loc} onChange={e => setZoneOverride(zone, "loc", e.target.value || null)}>
                  <option value="">Без точки</option>
                  {Object.keys(LOCATIONS).map(k => <option key={k} value={k}>{LOCATIONS[k].name}</option>)}
                </select>
                <input className="dod-num" type="number" step="0.1" value={km}
                  onChange={e => setZoneOverride(zone, "km", e.target.value === "" ? null : parseFloat(e.target.value))} />
              </div>
            );
          })}
        </div>
      )}

      {(timingError || cartError) && (
        <div className="dod-card" style={{ padding: 14, marginBottom: 20, borderColor: "var(--vermillion)", display: "flex", flexDirection: "column", gap: 8 }}>
          {timingError && <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}><AlertTriangle size={16} color="var(--vermillion)" style={{ flexShrink: 0, marginTop: 2 }} /><span style={{ fontSize: 13 }}>{timingError}</span></div>}
          {cartError && <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}><AlertTriangle size={16} color="var(--vermillion)" style={{ flexShrink: 0, marginTop: 2 }} /><span style={{ fontSize: 13 }}>{cartError}</span></div>}
        </div>
      )}

      {!hasTimings && !hasCarts && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }} className="dod-grid-2">
          <div className="dod-drop" onClick={() => document.querySelectorAll('.dod-root input[type=file]')[0].click()}>
            <UploadCloud size={24} color="var(--muted)" style={{ marginBottom: 8 }} />
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>{timingLoading ? "Читаю…" : "Файл таймингов"}</div>
            <div style={{ fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)" }}>Дата, Заказ, Принят, Кухня, Курьер, Завершён, Затрачено времени, Зона доставки</div>
          </div>
          <div className="dod-drop" onClick={() => document.querySelectorAll('.dod-root input[type=file]')[1].click()}>
            <Receipt size={24} color="var(--muted)" style={{ marginBottom: 8 }} />
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>{cartLoading ? "Читаю…" : "Детальный отчёт (корзины)"}</div>
            <div style={{ fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)" }}>#, Дата, Сумма, Скидка, Стоимость доставки, оплата, Адрес, Блюда...</div>
          </div>
        </div>
      )}

      {/* Cart-only lightweight view */}
      {cartOnlyStats && (
        <>
          <div style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)", marginBottom: 16 }}>
            Загрузите ещё и файл таймингов — появятся точки, зоны, время выполнения и расстояния.
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 26 }}>
            <Kpi icon={Receipt} label="Заказов" value={cartOnlyStats.total} />
            <Kpi icon={Truck} label="Выручка" value={fmtPln(cartOnlyStats.totalRevenue)} accent="var(--vermillion)" />
            <Kpi icon={Clock} label="Средний чек" value={fmtPln(cartOnlyStats.avgOrderValue)} />
            <Kpi icon={Percent} label="Скидки выдано" value={fmtPln(cartOnlyStats.totalDiscount)} sub={`${cartOnlyStats.discountedCount} заказов со скидкой`} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 26 }} className="dod-grid-2">
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note="кол-во и выручка">Каналы</SectionTitle>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={cartOnlyStats.platformStats} layout="vertical" margin={{ left: 10, right: 20 }}>
                  <CartesianGrid stroke="#26282f" horizontal={false} />
                  <XAxis type="number" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fill: "#F2EEE4", fontSize: 11 }} axisLine={false} tickLine={false} width={110} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="count" name="Заказы" radius={[0, 3, 3, 0]}>
                    {cartOnlyStats.platformStats.map((p, i) => <Cell key={i} fill={p.color} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note="топ-10 по кол-ву">Популярные блюда</SectionTitle>
              {cartOnlyStats.topDishes.map((d, i) => (
                <div className="dod-dish-row" key={i}><span>{d.dish}</span><span style={{ color: "var(--muted)" }}>{d.qty} шт · {fmtPln(d.revenue)}</span></div>
              ))}
            </div>
          </div>
        </>
      )}

      {stats && (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 20 }}>
            <Kpi icon={Truck} label="Заказов всего" value={stats.total} sub={`${stats.deliveryCount} доставка · ${stats.pickupCount} самовывоз`} />
            <Kpi icon={Clock} label="Ср. общее время" value={fmtMin(stats.avgTotal)} accent="var(--vermillion)" />
            <Kpi icon={Timer} label="Ср. принятие" value={fmtMin(stats.avgAccept)} />
            <Kpi icon={ChefHat} label="Ср. время кухни" value={fmtMin(stats.avgKitchen)} />
            <Kpi icon={MapPin} label="Ср. расстояние" value={fmtKm(stats.avgDistance)} sub="по прямой, зона → точка" />
            <Kpi icon={AlertTriangle} label="Без отметки завершения" value={stats.noTotalMark} accent={stats.noTotalMark > 0 ? "var(--amber)" : undefined} />
          </div>

          {hasCarts && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 20 }}>
              <Kpi icon={Receipt} label="Выручка" value={fmtPln(stats.totalRevenue)} accent="var(--vermillion)" />
              <Kpi icon={Clock} label="Средний чек" value={fmtPln(stats.avgOrderValue)} />
              <Kpi icon={Percent} label="Скидки выдано" value={fmtPln(stats.totalDiscount)} sub={`${stats.discountedCount} заказов со скидкой`} />
              <Kpi icon={Truck} label="Ср. стоимость доставки" value={fmtPln(stats.avgDeliveryFee)} sub="то, что оплатил клиент" />
            </div>
          )}

          {(stats.preorders > 0 || (hasCarts && stats.unmatchedRevenue > 0) || dupCount > 0) && (
            <div style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)", marginTop: -10, marginBottom: 20, lineHeight: 1.6 }}>
              {stats.preorders > 0 && <div>{stats.preorders} предзаказов исключены из «ср. принятие» и «ср. общее время»</div>}
              {hasCarts && stats.unmatchedRevenue > 0 && <div>{stats.unmatchedRevenue} заказов не найдены в детальном отчёте — выручка по ним не учтена</div>}
              {dupCount > 0 && <div>{dupCount} вероятных дублей (агрегаторы, тот же телефон и сумма, ≤20 мин) исключены из выручки; в кол-ве заказов и времени кухни оставлены — неясно, готовилось ли блюдо дважды</div>}
            </div>
          )}

          {/* Per-location cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 26 }}>
            {Object.keys(LOCATIONS).map(k => {
              const b = stats.byLocation[k];
              const loc = LOCATIONS[k];
              const avgTime = b.totalN ? Math.round(b.totalSum / b.totalN) : null;
              const avgDist = b.distN ? b.distSum / b.distN : null;
              const avgRev = b.revN ? b.revSum / b.revN : null;
              return (
                <div className="dod-card dod-loc-card" key={k}>
                  <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
                    <span className="dod-loc-dot" style={{ background: loc.color }} />
                    <span style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.04em" }}>{loc.name}</span>
                  </div>
                  <div style={{ fontFamily: "var(--mono)", fontSize: 10, color: "var(--muted)", marginBottom: 10 }}>{loc.address}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12 }}><span style={{ color: "var(--muted)" }}>Заказов</span><span>{b.count}</span></div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12, marginTop: 4 }}><span style={{ color: "var(--muted)" }}>Ср. время</span><span>{fmtMin(avgTime)}</span></div>
                  {hasCarts && <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12, marginTop: 4 }}><span style={{ color: "var(--muted)" }}>Ср. чек</span><span>{fmtPln(avgRev)}</span></div>}
                  <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12, marginTop: 4 }}><span style={{ color: "var(--muted)" }}>Ср. расстояние</span><span>{fmtKm(avgDist)}</span></div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12, marginTop: 4 }}><span style={{ color: "var(--muted)" }}>Зон</span><span>{b.zones.size}</span></div>
                </div>
              );
            })}
            <div className="dod-card dod-loc-card" style={{ borderStyle: "dashed" }}>
              <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
                <span className="dod-loc-dot" style={{ background: UNKNOWN_COLOR }} />
                <span style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.04em", color: "var(--muted)" }}>Без точки</span>
              </div>
              <div style={{ fontFamily: "var(--mono)", fontSize: 10, color: "var(--muted)", marginBottom: 10 }}>зоны без привязки</div>
              <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12 }}><span style={{ color: "var(--muted)" }}>Заказов</span><span>{stats.unassignedCount}</span></div>
            </div>
          </div>

          {/* Daily trend + hourly load per location */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 16, marginBottom: 26 }} className="dod-grid-2">
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note={hasCarts ? "заказы + выручка" : "заказов в день"}>Динамика по дням</SectionTitle>
              <ResponsiveContainer width="100%" height={220}>
                <ComposedChart data={stats.daily} margin={{ left: -20, right: 10 }}>
                  <CartesianGrid stroke="#26282f" vertical={false} />
                  <XAxis dataKey="label" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={{ stroke: "#26282f" }} tickLine={false} interval="preserveStartEnd" />
                  <YAxis yAxisId="left" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} width={30} />
                  {hasCarts && <YAxis yAxisId="right" orientation="right" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} width={40} />}
                  <Tooltip content={<CustomTooltip />} />
                  <Line yAxisId="left" type="monotone" dataKey="count" name="Заказы" stroke="#E8432C" strokeWidth={2} dot={false} />
                  {hasCarts && <Line yAxisId="right" type="monotone" dataKey="revenue" name="Выручка, zł" stroke="#5B7FBF" strokeWidth={2} dot={false} />}
                </ComposedChart>
              </ResponsiveContainer>
            </div>
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note="по часам, по точкам">Нагрузка по часам</SectionTitle>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={stats.byHourLoc} margin={{ left: -20, right: 10 }}>
                  <CartesianGrid stroke="#26282f" vertical={false} />
                  <XAxis dataKey="hour" tick={{ fill: "#8D8B85", fontSize: 10, fontFamily: "ui-monospace" }} axisLine={{ stroke: "#26282f" }} tickLine={false} interval={2} />
                  <YAxis tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="loc1" name={LOCATIONS.loc1.name} stackId="h" fill={LOCATIONS.loc1.color} />
                  <Bar dataKey="loc2" name={LOCATIONS.loc2.name} stackId="h" fill={LOCATIONS.loc2.color} />
                  <Bar dataKey="loc3" name={LOCATIONS.loc3.name} stackId="h" fill={LOCATIONS.loc3.color} />
                  <Bar dataKey="unknown" name="Без точки" stackId="h" fill={UNKNOWN_COLOR} radius={[3, 3, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginTop: 8, fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>
                {Object.keys(LOCATIONS).map(k => (<span key={k}><span style={{ color: LOCATIONS[k].color }}>■</span> {LOCATIONS[k].name}</span>))}
                <span><span style={{ color: UNKNOWN_COLOR }}>■</span> Без точки</span>
              </div>
            </div>
          </div>

          {/* Channels & payments */}
          {hasCarts && (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 26 }} className="dod-grid-2">
              <div className="dod-card" style={{ padding: 20 }}>
                <SectionTitle note="заказы · средний чек">Каналы доставки</SectionTitle>
                <ResponsiveContainer width="100%" height={Math.max(160, stats.platformStats.length * 34)}>
                  <BarChart data={stats.platformStats} layout="vertical" margin={{ left: 10, right: 20 }}>
                    <CartesianGrid stroke="#26282f" horizontal={false} />
                    <XAxis type="number" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" tick={{ fill: "#F2EEE4", fontSize: 11 }} axisLine={false} tickLine={false} width={120} />
                    <Tooltip content={<CustomTooltip />} />
                    <Bar dataKey="count" name="Заказы" radius={[0, 3, 3, 0]}>
                      {stats.platformStats.map((p, i) => <Cell key={i} fill={p.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="dod-card" style={{ padding: 20 }}>
                <SectionTitle note="способ оплаты">Оплата</SectionTitle>
                <ResponsiveContainer width="100%" height={Math.max(160, stats.paymentStats.length * 34)}>
                  <BarChart data={stats.paymentStats} layout="vertical" margin={{ left: 10, right: 20 }}>
                    <CartesianGrid stroke="#26282f" horizontal={false} />
                    <XAxis type="number" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" tick={{ fill: "#F2EEE4", fontSize: 11 }} axisLine={false} tickLine={false} width={140} />
                    <Tooltip content={<CustomTooltip />} />
                    <Bar dataKey="count" name="Заказы" radius={[0, 3, 3, 0]}>
                      {stats.paymentStats.map((p, i) => <Cell key={i} fill={p.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          {/* Distance analytics */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 26 }} className="dod-grid-2">
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note="время + объём">Время в пути по дистанции</SectionTitle>
              <ResponsiveContainer width="100%" height={220}>
                <ComposedChart data={stats.distanceStats} margin={{ left: -20, right: 10 }}>
                  <CartesianGrid stroke="#26282f" vertical={false} />
                  <XAxis dataKey="label" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={{ stroke: "#26282f" }} tickLine={false} />
                  <YAxis yAxisId="left" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} width={30} />
                  <YAxis yAxisId="right" orientation="right" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar yAxisId="left" dataKey="avgTime" name="Ср. время, мин" fill="#E8432C" radius={[3, 3, 0, 0]} />
                  <Line yAxisId="right" type="monotone" dataKey="count" name="Заказов" stroke="#5B7FBF" strokeWidth={2} dot={{ r: 3 }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note="пузырь = кол-во заказов">Зоны: расстояние × время</SectionTitle>
              {stats.zoneScatter.length > 0 ? (
                <ResponsiveContainer width="100%" height={220}>
                  <ScatterChart margin={{ left: -10, right: 10, top: 10, bottom: 0 }}>
                    <CartesianGrid stroke="#26282f" />
                    <XAxis type="number" dataKey="distanceKm" name="Расстояние" unit=" км" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={{ stroke: "#26282f" }} tickLine={false} />
                    <YAxis type="number" dataKey="avgTime" name="Время" unit=" мин" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} width={34} />
                    <ZAxis type="number" dataKey="count" range={[40, 300]} />
                    <Tooltip content={<ZoneScatterTooltip />} cursor={{ strokeDasharray: "3 3" }} />
                    {Object.keys(LOCATIONS).map(k => (<Scatter key={k} data={stats.zoneScatter.filter(z => z.loc === k)} fill={LOCATIONS[k].color} />))}
                  </ScatterChart>
                </ResponsiveContainer>
              ) : <div style={{ color: "var(--muted)", fontSize: 12, fontFamily: "var(--mono)" }}>недостаточно данных</div>}
            </div>
          </div>

          {/* Zones */}
          {stats.zoneByCount.length > 0 && (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 26 }} className="dod-grid-2">
              <div className="dod-card" style={{ padding: 20 }}>
                <SectionTitle note="цвет = точка">Заказы по зонам</SectionTitle>
                <ResponsiveContainer width="100%" height={Math.max(180, stats.zoneByCount.length * 28)}>
                  <BarChart data={stats.zoneByCount} layout="vertical" margin={{ left: 10, right: 20 }}>
                    <CartesianGrid stroke="#26282f" horizontal={false} />
                    <XAxis type="number" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="zone" tick={{ fill: "#F2EEE4", fontSize: 11 }} axisLine={false} tickLine={false} width={130} />
                    <Tooltip content={<CustomTooltip />} />
                    <Bar dataKey="count" name="Заказы" radius={[0, 3, 3, 0]}>{stats.zoneByCount.map((z, i) => <Cell key={i} fill={locColor(z.loc)} />)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="dod-card" style={{ padding: 20 }}>
                <SectionTitle note="где ≥ 3 заказов">Ср. время по зонам</SectionTitle>
                {stats.zoneByTime.length > 0 ? (
                  <ResponsiveContainer width="100%" height={Math.max(180, stats.zoneByTime.length * 28)}>
                    <BarChart data={stats.zoneByTime} layout="vertical" margin={{ left: 10, right: 20 }}>
                      <CartesianGrid stroke="#26282f" horizontal={false} />
                      <XAxis type="number" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} />
                      <YAxis type="category" dataKey="zone" tick={{ fill: "#F2EEE4", fontSize: 11 }} axisLine={false} tickLine={false} width={130} />
                      <Tooltip content={<CustomTooltip />} />
                      <Bar dataKey="avgTime" name="Минут" radius={[0, 3, 3, 0]}>{stats.zoneByTime.map((z, i) => <Cell key={i} fill={locColor(z.loc)} />)}</Bar>
                    </BarChart>
                  </ResponsiveContainer>
                ) : <div style={{ color: "var(--muted)", fontSize: 12, fontFamily: "var(--mono)" }}>недостаточно данных</div>}
              </div>
            </div>
          )}

          {/* Stage breakdown */}
          <div className="dod-card" style={{ padding: 20, marginBottom: 26 }}>
            <SectionTitle note="средние минуты на этап">Куда уходит время</SectionTitle>
            <ResponsiveContainer width="100%" height={140}>
              <BarChart data={stats.stageBars} layout="vertical" margin={{ left: 10, right: 20 }} stackOffset="sign">
                <CartesianGrid stroke="#26282f" horizontal={false} />
                <XAxis type="number" tick={{ fill: "#8D8B85", fontSize: 11, fontFamily: "ui-monospace" }} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="name" tick={{ fill: "#F2EEE4", fontSize: 12, fontWeight: 700 }} axisLine={false} tickLine={false} width={90} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="Принятие" stackId="s" fill={ACCEPT_COLOR} />
                <Bar dataKey="Кухня" stackId="s" fill={KITCHEN_COLOR} />
                <Bar dataKey="Ожидание курьера" stackId="s" fill={COURIER_COLOR} />
                <Bar dataKey="В пути" stackId="s" fill={TRANSIT_COLOR} radius={[0, 3, 3, 0]} />
                <Bar dataKey="Ожидание забора" stackId="s" fill={COURIER_COLOR} radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginTop: 10, fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>
              <span><span style={{ color: ACCEPT_COLOR }}>■</span> Принятие</span>
              <span><span style={{ color: KITCHEN_COLOR }}>■</span> Кухня</span>
              <span><span style={{ color: COURIER_COLOR }}>■</span> Ожидание курьера / забора</span>
              <span><span style={{ color: TRANSIT_COLOR }}>■</span> В пути</span>
            </div>
          </div>

          {/* Top dishes */}
          {hasCarts && stats.topDishes.length > 0 && (
            <div className="dod-card" style={{ padding: 20, marginBottom: 26 }}>
              <SectionTitle note="топ-10 по количеству">Популярные блюда</SectionTitle>
              {stats.topDishes.map((d, i) => (
                <div className="dod-dish-row" key={i}><span>{d.dish}</span><span style={{ color: "var(--muted)" }}>{d.qty} шт · {fmtPln(d.revenue)}</span></div>
              ))}
            </div>
          )}

          {/* Slowest orders */}
          {stats.slowest.length > 0 && (
            <div className="dod-card" style={{ padding: 20 }}>
              <SectionTitle note={`топ-${stats.slowest.length}`}>Самые долгие заказы</SectionTitle>
              <div className="dod-ticket">
                <div className="dod-ticket-row" style={{ borderBottom: "1px solid var(--card-border)", color: "var(--muted)", fontSize: 10, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                  <span>Заказ</span><span>Распределение времени</span><span>Зона / адрес</span><span style={{ textAlign: "right" }}>Итого</span>
                </div>
                {stats.slowest.map((r, i) => {
                  const segs = [
                    { v: r.acceptMin || 0, c: ACCEPT_COLOR }, { v: r.kitchenMin || 0, c: KITCHEN_COLOR },
                    { v: (r.courierMin || 0) + (r.pickupWaitMin || 0), c: COURIER_COLOR }, { v: r.transitMin || 0, c: TRANSIT_COLOR },
                  ];
                  const segSum = segs.reduce((a, s) => a + s.v, 0) || 1;
                  const label = r.isPickup ? "самовывоз" : (r.cart && r.cart.address ? r.cart.address : (r.zone || "—"));
                  return (
                    <div className="dod-ticket-row" key={i}>
                      <span style={{ color: "var(--text)" }}>#{r.orderId}</span>
                      <span className="dod-seg-bar">{segs.map((s, j) => s.v > 0 && <span key={j} style={{ width: `${(s.v / segSum) * 100}%`, background: s.c }} />)}</span>
                      <span style={{ color: "var(--muted)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={label}>{label}</span>
                      <span style={{ textAlign: "right", color: "var(--vermillion)", fontWeight: 700 }}>{r.totalMin} мин</span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </>
      )}

      <style>{`
        @media (max-width: 720px) {
          .dod-grid-2 { grid-template-columns: 1fr !important; }
          .dod-ticket-row { grid-template-columns: 55px 1fr 45px !important; }
          .dod-ticket-row span:nth-child(3) { display: none; }
          .dod-map-row { grid-template-columns: 1fr 110px 70px; }
        }
      `}</style>
    </div>
  );
}
