import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Package, ShoppingCart, Receipt, BarChart3, Plus, Trash2,
  AlertTriangle, Bot, Send, Loader2, TrendingUp, TrendingDown, RotateCcw,
  Mic, Camera, Check, Ban, ClipboardCheck,
  Store, Maximize2, X, ImageOff, History, Download, Upload, Palette, Lock
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid
} from "recharts";

const STORAGE_KEY = "kiosco-store-v1";
const THEME_KEY = "kiosco-theme-v1";
// Código para recuperar la clave de administrador si se olvida. Cambialo por uno propio y guardalo.
const RECOVERY_CODE = "RECREO-4827-QX";

const THEMES = {
  clasico: {
    label: "Clásico", swatch: "#0F5C4E",
    paper: "#FAF6EF", paperDark: "#F0EAD9", ink: "#21261F", inkSoft: "#5B6157",
    teal: "#0F5C4E", tealDark: "#0A4238", mustard: "#E8A33D", red: "#C1443A",
    line: "#DCD3BE", surfaceDark: "#1F2420", cardBg: "#FFFFFF",
  },
  oceano: {
    label: "Océano", swatch: "#1B6E93",
    paper: "#F2F7FA", paperDark: "#E3EEF3", ink: "#122430", inkSoft: "#4F6672",
    teal: "#1B6E93", tealDark: "#12506C", mustard: "#F2A65A", red: "#D1495B",
    line: "#CADDE4", surfaceDark: "#0E2733", cardBg: "#FFFFFF",
  },
  menta: {
    label: "Menta", swatch: "#1F8F5D",
    paper: "#F6FAF7", paperDark: "#E7F3EA", ink: "#14261D", inkSoft: "#516459",
    teal: "#1F8F5D", tealDark: "#166D46", mustard: "#E9B44C", red: "#C1443A",
    line: "#CFE6D6", surfaceDark: "#0E2018", cardBg: "#FFFFFF",
  },
  frutal: {
    label: "Frutal", swatch: "#E2632A",
    paper: "#FFF8F0", paperDark: "#FCEADA", ink: "#3A2113", inkSoft: "#8B5D46",
    teal: "#E2632A", tealDark: "#B94E1E", mustard: "#FFC93C", red: "#D1495B",
    line: "#F4D6BC", surfaceDark: "#2B160B", cardBg: "#FFFFFF",
  },
  oscuro: {
    label: "Oscuro", swatch: "#22A281",
    paper: "#1C1E1A", paperDark: "#262924", ink: "#F1EEE3", inkSoft: "#ACA99C",
    teal: "#22A281", tealDark: "#188066", mustard: "#F0B85E", red: "#E4776C",
    line: "#3A3D34", surfaceDark: "#0F110D", cardBg: "#252822",
  },
};

const fmt = (n) =>
  (Number(n) || 0).toLocaleString("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  });

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

// Achica y comprime la foto elegida por el usuario para que ocupe poco espacio guardada
function compressImageFile(file, maxDim = 480, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxDim) {
          height = Math.round((height * maxDim) / width);
          width = maxDim;
        } else if (height >= width && height > maxDim) {
          width = Math.round((width * maxDim) / height);
          height = maxDim;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => reject(new Error("No se pudo leer la imagen"));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsDataURL(file);
  });
}

function todayKey(d) {
  return new Date(d).toDateString();
}

function exportExcel(store) {
  const wb = XLSX.utils.book_new();

  const ventasSheet = XLSX.utils.json_to_sheet(
    store.sales.map((s) => ({
      Fecha: new Date(s.date).toLocaleString("es-AR"),
      Producto: s.productName,
      Cantidad: s.quantity,
      "Precio unitario": s.price,
      "Costo unitario": s.cost || 0,
      Total: s.total,
      Ganancia: (s.price - (s.cost || 0)) * s.quantity,
    }))
  );
  XLSX.utils.book_append_sheet(wb, ventasSheet, "Ventas");

  const gastosSheet = XLSX.utils.json_to_sheet(
    store.expenses.map((e) => ({
      Fecha: new Date(e.date).toLocaleString("es-AR"),
      Descripción: e.description,
      Monto: e.amount,
    }))
  );
  XLSX.utils.book_append_sheet(wb, gastosSheet, "Gastos");

  const historialSheet = XLSX.utils.json_to_sheet(
    (store.priceHistory || []).map((h) => ({
      Fecha: new Date(h.date).toLocaleString("es-AR"),
      Producto: h.productName,
      Campo: h.field === "price" ? "Precio de venta" : "Costo",
      Antes: h.oldValue,
      Después: h.newValue,
      "Variación %": h.oldValue ? (((h.newValue - h.oldValue) / h.oldValue) * 100).toFixed(1) : "",
    }))
  );
  XLSX.utils.book_append_sheet(wb, historialSheet, "Historial de precios");

  const productosSheet = XLSX.utils.json_to_sheet(
    store.products.map((p) => ({
      Producto: p.name,
      Categoría: p.category,
      "Precio actual": p.price,
      "Costo actual": p.cost,
      "Margen %": p.price > 0 && p.cost > 0 ? (((p.price - p.cost) / p.price) * 100).toFixed(0) : "",
      Stock: p.stock,
    }))
  );
  XLSX.utils.book_append_sheet(wb, productosSheet, "Productos");

  XLSX.writeFile(wb, `quiosco-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

export default function App() {
  const [store, setStore] = useState({ products: [], sales: [], expenses: [], priceHistory: [] });
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState(false);
  const [tab, setTab] = useState("catalogo");
  const [cart, setCart] = useState({});
  const setCartQty = (id, qty) =>
    setCart((prev) => {
      const next = { ...prev };
      if (qty <= 0) delete next[id];
      else next[id] = qty;
      return next;
    });
  const [presentMode, setPresentMode] = useState(false);
  const [themeKey, setThemeKey] = useState("clasico");
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [adminPin, setAdminPin] = useState(null);
  const [pinInput, setPinInput] = useState("");
  const [pinError, setPinError] = useState("");
  const [recoveryMode, setRecoveryMode] = useState(false);
  const [recoveryInput, setRecoveryInput] = useState("");
  const [recoveryError, setRecoveryError] = useState("");
  const [themePickerOpen, setThemePickerOpen] = useState(false);
  const feedRef = useRef(null);
  const theme = THEMES[themeKey] || THEMES.clasico;

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(THEME_KEY, false);
        if (res && res.value && THEMES[res.value]) setThemeKey(res.value);
      } catch (e) {}
    })();
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("kiosco-admin-pin", true);
        if (res && res.value) setAdminPin(res.value);
      } catch (e) {}
      try {
        const res2 = await window.storage.get("kiosco-admin-unlocked", false);
        if (res2 && res2.value === "true") setAdminUnlocked(true);
      } catch (e) {}
    })();
  }, []);

  async function tryUnlock() {
    setPinError("");
    if (!pinInput.trim()) return;
    if (!adminPin) {
      // primera vez: se define la clave
      if (pinInput.trim().length < 4) {
        setPinError("Usá al menos 4 caracteres para la clave.");
        return;
      }
      try {
        await window.storage.set("kiosco-admin-pin", pinInput.trim(), true);
      } catch (e) {}
      setAdminPin(pinInput.trim());
      setAdminUnlocked(true);
      try { await window.storage.set("kiosco-admin-unlocked", "true", false); } catch (e) {}
      setPinInput("");
      setTab("productos");
    } else if (pinInput.trim() === adminPin) {
      setAdminUnlocked(true);
      try { await window.storage.set("kiosco-admin-unlocked", "true", false); } catch (e) {}
      setPinInput("");
      setTab("productos");
    } else {
      setPinError("Clave incorrecta.");
    }
  }

  async function lockAdmin() {
    setAdminUnlocked(false);
    try { await window.storage.set("kiosco-admin-unlocked", "false", false); } catch (e) {}
    setTab("catalogo");
  }

  async function recoverPin() {
    setRecoveryError("");
    if (recoveryInput.trim() !== RECOVERY_CODE) {
      setRecoveryError("Código de recuperación incorrecto.");
      return;
    }
    try { await window.storage.delete("kiosco-admin-pin", true); } catch (e) {}
    setAdminPin(null);
    setRecoveryMode(false);
    setRecoveryInput("");
    setPinInput("");
    setPinError("");
  }

  useEffect(() => {
    (async () => {
      try {
        await window.storage.set(THEME_KEY, themeKey, false);
      } catch (e) {}
    })();
  }, [themeKey]);

  // ---- load / persist ----
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, true);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setStore({
            products: parsed.products || [],
            sales: parsed.sales || [],
            expenses: parsed.expenses || [],
            priceHistory: parsed.priceHistory || [],
          });
        }
      } catch (e) {
        // no existe todavía, arrancamos vacío
      }
      setLoading(false);
    })();
  }, []);

  useEffect(() => {
    if (loading) return;
    (async () => {
      try {
        const res = await window.storage.set(STORAGE_KEY, JSON.stringify(store), true);
        setSaveError(!res);
      } catch (e) {
        setSaveError(true);
      }
    })();
  }, [store, loading]);

  // ---- mutations ----
  function addProduct(p) {
    setStore((prev) => ({
      ...prev,
      products: [
        ...prev.products,
        {
          id: uid(),
          name: p.name,
          category: p.category || "General",
          price: Number(p.price) || 0,
          cost: Number(p.cost) || 0,
          stock: Number(p.stock) || 0,
          minStock: Number(p.minStock) || 5,
          imageUrl: p.imageUrl || "",
        },
      ],
    }));
  }

  function updateProduct(id, patch) {
    setStore((prev) => {
      const product = prev.products.find((p) => p.id === id);
      if (!product) return prev;
      const entries = [];
      if (patch.price != null && Number(patch.price) !== product.price) {
        entries.push({ id: uid(), date: new Date().toISOString(), productId: id, productName: product.name, field: "price", oldValue: product.price, newValue: Number(patch.price) });
      }
      if (patch.cost != null && Number(patch.cost) !== product.cost) {
        entries.push({ id: uid(), date: new Date().toISOString(), productId: id, productName: product.name, field: "cost", oldValue: product.cost, newValue: Number(patch.cost) });
      }
      return {
        ...prev,
        products: prev.products.map((p) => (p.id === id ? { ...p, ...patch } : p)),
        priceHistory: entries.length ? [...entries, ...(prev.priceHistory || [])] : prev.priceHistory,
      };
    });
  }

  function deleteProduct(id) {
    setStore((prev) => ({ ...prev, products: prev.products.filter((p) => p.id !== id) }));
  }

  function recordSale({ productId, quantity, priceOverride }) {
    setStore((prev) => {
      const product = prev.products.find((p) => p.id === productId);
      if (!product || !quantity || quantity <= 0) return prev;
      const price = priceOverride != null ? Number(priceOverride) : product.price;
      const total = price * quantity;
      const products = prev.products.map((p) =>
        p.id === productId ? { ...p, stock: Math.max(0, p.stock - quantity) } : p
      );
      const sale = {
        id: uid(),
        date: new Date().toISOString(),
        productId,
        productName: product.name,
        quantity: Number(quantity),
        price,
        cost: product.cost,
        total,
      };
      return { ...prev, products, sales: [sale, ...prev.sales] };
    });
  }

  function deleteSale(id) {
    setStore((prev) => {
      const sale = prev.sales.find((s) => s.id === id);
      if (!sale) return prev;
      const products = prev.products.map((p) =>
        p.id === sale.productId ? { ...p, stock: p.stock + sale.quantity } : p
      );
      return { ...prev, products, sales: prev.sales.filter((s) => s.id !== id) };
    });
  }

  function addExpense(e) {
    setStore((prev) => ({
      ...prev,
      expenses: [
        { id: uid(), date: new Date().toISOString(), description: e.description || "Gasto", amount: Number(e.amount) || 0, type: e.type === "fijo" ? "fijo" : "variable" },
        ...prev.expenses,
      ],
    }));
  }

  function updateExpense(id, patch) {
    setStore((prev) => ({
      ...prev,
      expenses: prev.expenses.map((e) => (e.id === id ? { ...e, ...patch } : e)),
    }));
  }

  function deleteExpense(id) {
    setStore((prev) => ({ ...prev, expenses: prev.expenses.filter((e) => e.id !== id) }));
  }

  function restock(productId, quantity, cost) {
    setStore((prev) => {
      const product = prev.products.find((p) => p.id === productId);
      if (!product) return prev;
      const entries = [];
      if (cost != null && Number(cost) !== product.cost) {
        entries.push({ id: uid(), date: new Date().toISOString(), productId, productName: product.name, field: "cost", oldValue: product.cost, newValue: Number(cost) });
      }
      return {
        ...prev,
        products: prev.products.map((p) =>
          p.id === productId
            ? { ...p, stock: p.stock + Number(quantity || 0), cost: cost != null ? Number(cost) : p.cost }
            : p
        ),
        priceHistory: entries.length ? [...entries, ...(prev.priceHistory || [])] : prev.priceHistory,
      };
    });
  }

  async function resetAll() {
    if (!window.confirm("¿Borrar todos los datos del quiosco? Esta acción no se puede deshacer.")) return;
    try {
      await window.storage.delete(STORAGE_KEY, true);
    } catch (e) {}
    setStore({ products: [], sales: [], expenses: [], priceHistory: [] });
  }

  // ---- derived data ----
  const today = todayKey(new Date());
  const todaysSales = store.sales.filter((s) => todayKey(s.date) === today);
  const todaysExpenses = store.expenses.filter((e) => todayKey(e.date) === today);
  const todaysSalesTotal = todaysSales.reduce((a, s) => a + s.total, 0);
  const todaysExpensesTotal = todaysExpenses.reduce((a, e) => a + e.amount, 0);
  const lowStock = store.products.filter((p) => p.stock <= (p.minStock ?? 5));

  const totalSalesAll = store.sales.reduce((a, s) => a + s.total, 0);
  const totalExpensesAll = store.expenses.reduce((a, e) => a + e.amount, 0);
  const totalCostOfGoods = store.sales.reduce((a, s) => a + (s.cost || 0) * s.quantity, 0);
  const grossProfit = totalSalesAll - totalCostOfGoods - totalExpensesAll;

  const salesByProduct = useMemo(() => {
    const map = {};
    store.sales.forEach((s) => {
      if (!map[s.productId]) map[s.productId] = { name: s.productName, total: 0, quantity: 0 };
      map[s.productId].total += s.total;
      map[s.productId].quantity += s.quantity;
    });
    return Object.values(map).sort((a, b) => b.total - a.total);
  }, [store.sales]);

  // ---- agent ----
  const [agentMessages, setAgentMessages] = useState([
    { role: "agent", text: "Contame qué pasó y lo cargo yo. Por ejemplo: \"vendí 2 alfajores y 1 coca de 500\", o \"gasto de luz 8000\". También podés dictarlo con el micrófono o sacarle una foto a una factura." },
  ]);
  const [agentInput, setAgentInput] = useState("");
  const [agentBusy, setAgentBusy] = useState(false);
  const [listening, setListening] = useState(false);
  const [pendingReceipt, setPendingReceipt] = useState(null); // { actions, reply, kind }
  const recognitionRef = useRef(null);
  const receiptInputRef = useRef(null);

  const speechSupported =
    typeof window !== "undefined" && (window.SpeechRecognition || window.webkitSpeechRecognition);

  function toggleListening() {
    if (!speechSupported) return;
    if (listening) {
      recognitionRef.current && recognitionRef.current.stop();
      return;
    }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    const recognition = new SR();
    recognition.lang = "es-AR";
    recognition.interimResults = true;
    recognition.continuous = false;
    recognition.onstart = () => setListening(true);
    recognition.onerror = () => setListening(false);
    recognition.onend = () => setListening(false);
    recognition.onresult = (event) => {
      let transcript = "";
      for (let i = 0; i < event.results.length; i++) transcript += event.results[i][0].transcript;
      setAgentInput(transcript);
    };
    recognitionRef.current = recognition;
    recognition.start();
  }

  async function handleReceiptFile(e) {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setAgentMessages((prev) => [...prev, { role: "user", text: "📷 Subió la foto de un comprobante." }]);
    setAgentBusy(true);
    try {
      const dataUrl = await compressImageFile(file, 1400, 0.85);
      const base64 = dataUrl.split(",")[1];
      const productList = store.products.map((p) => ({ id: p.id, name: p.name }));
      const system = `Sos el asistente de un quiosco argentino. Te paso la foto de un comprobante: puede ser una factura de SERVICIO (luz, gas, internet, alquiler, impuestos) o una factura/remito de MERCADERÍA de un proveedor (productos para reponer stock).
Lista de productos existentes (JSON): ${JSON.stringify(productList)}
Devolvé EXCLUSIVAMENTE un JSON válido (sin texto adicional, sin markdown) con esta forma:
{
  "kind": "servicio" | "mercaderia" | "desconocido",
  "actions": [
    {"type":"expense","description":"<texto>","amount":<numero>},
    {"type":"restock","productId":"<id existente o null>","productName":"<nombre le\u00eddo>","quantity":<numero>,"cost":<costo unitario o null>},
    {"type":"new_product","name":"<nombre>","price":null,"cost":<costo unitario o null>,"stock":<numero>,"category":null}
  ],
  "reply": "<resumen breve en espa\u00f1ol de lo que se leyó, aclarando si algo no se pudo leer con confianza>"
}
Reglas: si es un comprobante de servicio, devolvé una sola acción "expense" con el monto total. Si es mercadería, devolvé una acción por cada producto de la lista del comprobante (restock si coincide con un producto existente, new_product si no). Nunca inventes montos o cantidades que no puedas leer: si hay dudas, dejalo afuera y aclaralo en "reply". No agregues texto fuera del JSON.`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1200,
          system,
          messages: [
            {
              role: "user",
              content: [
                { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
                { type: "text", text: "Leé este comprobante y devolveme el JSON pedido." },
              ],
            },
          ],
        }),
      });
      const data = await response.json();
      const raw = (data.content || []).map((b) => b.text || "").join("");
      const clean = raw.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      if (!parsed.actions || parsed.actions.length === 0) {
        setAgentMessages((prev) => [...prev, { role: "agent", text: parsed.reply || "No pude leer datos claros en la foto. ¿Podés sacarla de nuevo, más de cerca?" }]);
      } else {
        setPendingReceipt({ actions: parsed.actions, reply: parsed.reply, kind: parsed.kind });
        setAgentMessages((prev) => [...prev, { role: "agent", text: (parsed.reply || "Esto es lo que leí:") + " Revisá abajo y confirmá si está bien." }]);
      }
    } catch (err) {
      setAgentMessages((prev) => [...prev, { role: "agent", text: "No pude leer la foto. Probá con más luz o más cerca del texto." }]);
    } finally {
      setAgentBusy(false);
    }
  }

  function confirmReceipt() {
    if (!pendingReceipt) return;
    applyActions(pendingReceipt.actions);
    setPendingReceipt(null);
    setAgentMessages((prev) => [...prev, { role: "agent", text: "Listo, lo cargué." }]);
  }

  function cancelReceipt() {
    setPendingReceipt(null);
    setAgentMessages((prev) => [...prev, { role: "agent", text: "Ok, no cargué nada. Podés volver a intentar con otra foto." }]);
  }

  function describeAction(a) {
    if (a.type === "expense") return `Gasto: ${a.description} — ${fmt(a.amount)}`;
    if (a.type === "restock") return `Reponer: ${a.productName || "producto"} x${a.quantity}${a.cost ? ` (costo ${fmt(a.cost)} c/u)` : ""}`;
    if (a.type === "new_product") return `Producto nuevo: ${a.name}${a.cost ? ` (costo ${fmt(a.cost)})` : ""}`;
    return "Acción";
  }

  useEffect(() => {
    if (feedRef.current) feedRef.current.scrollTop = feedRef.current.scrollHeight;
  }, [agentMessages, agentBusy]);

  function applyActions(actions) {
    let anyUnresolved = false;
    const warnings = [];
    (actions || []).forEach((a) => {
      if (a.type === "sale") {
        if (a.productId) {
          if ((a.quantity || 1) >= 20) warnings.push(`vendiste ${a.quantity} de "${a.productName}", revisá que esté bien`);
          recordSale({ productId: a.productId, quantity: a.quantity || 1 });
        } else anyUnresolved = true;
      } else if (a.type === "expense") {
        if (a.amount) {
          if (a.amount >= 100000) warnings.push(`cargaste un gasto de ${fmt(a.amount)}, revisá que esté bien`);
          addExpense({ description: a.description || "Gasto", amount: a.amount });
        } else anyUnresolved = true;
      } else if (a.type === "restock") {
        if (a.productId) {
          if ((a.quantity || 0) >= 50) warnings.push(`repusiste ${a.quantity} unidades, revisá que esté bien`);
          restock(a.productId, a.quantity || 0, a.cost);
        } else anyUnresolved = true;
      } else if (a.type === "new_product") {
        if (a.name && a.price) {
          addProduct({ name: a.name, price: a.price, cost: a.cost || 0, stock: a.stock || 0, category: a.category || "General", minStock: 5 });
        } else {
          anyUnresolved = true;
        }
      }
    });
    return { anyUnresolved, warnings };
  }

  async function handleAgentSubmit() {
    const text = agentInput.trim();
    if (!text || agentBusy) return;
    setAgentInput("");
    setAgentMessages((prev) => [...prev, { role: "user", text }]);
    setAgentBusy(true);
    try {
      const productList = store.products.map((p) => ({ id: p.id, name: p.name, price: p.price, stock: p.stock }));
      const system = `Sos el asistente de un quiosco argentino y ayudás a cargar datos a partir de mensajes en lenguaje natural.
Lista de productos existentes (JSON): ${JSON.stringify(productList)}

Devolvé EXCLUSIVAMENTE un JSON válido (sin texto adicional, sin bloques de código markdown) con esta forma exacta:
{
  "actions": [
    {"type":"sale","productId":"<id existente o null>","productName":"<nombre mencionado>","quantity":<numero>},
    {"type":"expense","description":"<texto>","amount":<numero>},
    {"type":"restock","productId":"<id existente o null>","productName":"<nombre>","quantity":<numero>,"cost":<numero o null>},
    {"type":"new_product","name":"<nombre>","price":<numero o null>,"cost":<numero o null>,"stock":<numero o null>,"category":<texto o null>}
  ],
  "reply": "<respuesta breve en español confirmando qué se hizo, o pidiendo el dato que falta>"
}
Reglas:
- Un mensaje puede incluir varias acciones (ej: vender dos productos distintos).
- Usá el productId existente cuando el nombre mencionado coincida, aunque sea parcialmente, con uno de la lista.
- Si se menciona un producto que NO está en la lista para una venta o reposición, dejá productId en null y explicá en "reply" que hay que cargarlo primero como producto nuevo.
- Nunca inventes precios ni montos: si falta un dato imprescindible dejalo en null y pedilo en "reply".
- Si el mensaje no tiene relación con ventas, gastos, stock o productos, devolvé "actions": [] y una "reply" aclarando que no entendiste el pedido.
- No agregues texto fuera del JSON.`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          system,
          messages: [{ role: "user", content: text }],
        }),
      });
      const data = await response.json();
      const raw = (data.content || []).map((b) => b.text || "").join("");
      const clean = raw.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      const { warnings } = applyActions(parsed.actions || []);
      const replyText = (parsed.reply || "Listo, lo cargué.") + (warnings.length ? ` ⚠️ Ojo: ${warnings.join("; ")}.` : "");
      setAgentMessages((prev) => [...prev, { role: "agent", text: replyText }]);
    } catch (e) {
      setAgentMessages((prev) => [...prev, { role: "agent", text: "No pude procesar ese pedido. ¿Podés reformularlo con más detalle?" }]);
    } finally {
      setAgentBusy(false);
    }
  }

  if (loading) {
    return (
      <div className="kiosco-app flex items-center justify-center" style={{ minHeight: "480px" }}>
        <Loader2 className="animate-spin" size={28} style={{ color: "var(--teal)" }} />
      </div>
    );
  }

  if (presentMode) {
    return <CatalogoFullscreen products={store.products} onClose={() => setPresentMode(false)} theme={theme} whatsappNumber={store.whatsappNumber} cart={cart} onQty={setCartQty} onClear={() => setCart({})} />;
  }

  return (
    <div className="kiosco-app">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Manrope:wght@400;500;600;700;800&display=swap');
        .kiosco-app {
          --paper: ${theme.paper};
          --paper-dark: ${theme.paperDark};
          --ink: ${theme.ink};
          --ink-soft: ${theme.inkSoft};
          --teal: ${theme.teal};
          --teal-dark: ${theme.tealDark};
          --mustard: ${theme.mustard};
          --red: ${theme.red};
          --line: ${theme.line};
          --surface-dark: ${theme.surfaceDark};
          --card-bg: ${theme.cardBg};
          font-family: 'Manrope', sans-serif;
          background: var(--paper);
          color: var(--ink);
          min-height: 100%;
          padding: 20px 16px 0;
        }
        .kiosco-app .mono { font-family: 'JetBrains Mono', monospace; }
        .kiosco-app input, .kiosco-app select {
          font-family: 'Manrope', sans-serif;
          background: var(--card-bg);
          border: 1px solid var(--line);
          border-radius: 8px;
          padding: 7px 10px;
          font-size: 14px;
          color: var(--ink);
          outline: none;
        }
        .kiosco-app input:focus, .kiosco-app select:focus {
          border-color: var(--teal);
          box-shadow: 0 0 0 2px rgba(15,92,78,0.15);
        }
        .kiosco-app .btn-primary {
          background: var(--teal);
          color: white;
          border: none;
          border-radius: 8px;
          padding: 8px 14px;
          font-weight: 600;
          font-size: 14px;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          cursor: pointer;
          transition: background 0.15s ease;
        }
        .kiosco-app .btn-primary:hover { background: var(--teal-dark); }
        .kiosco-app .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
        .kiosco-app .btn-ghost {
          background: transparent;
          border: 1px solid var(--line);
          color: var(--ink-soft);
          border-radius: 8px;
          padding: 7px 12px;
          font-size: 13px;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          cursor: pointer;
        }
        .kiosco-app .btn-ghost:hover { border-color: var(--ink-soft); color: var(--ink); }
        .kiosco-app .icon-btn {
          background: transparent;
          border: none;
          cursor: pointer;
          color: var(--ink-soft);
          padding: 5px;
          border-radius: 6px;
          display: inline-flex;
        }
        .kiosco-app .icon-btn:hover { background: var(--paper-dark); color: var(--red); }
        .kiosco-app .card {
          background: var(--card-bg);
          border: 1px solid var(--line);
          border-radius: 12px;
        }
        .kiosco-app .tab-btn {
          display: flex; align-items: center; gap: 6px;
          padding: 9px 14px;
          border-radius: 9px;
          font-size: 13.5px;
          font-weight: 600;
          border: none;
          background: transparent;
          color: var(--ink-soft);
          cursor: pointer;
          white-space: nowrap;
        }
        .kiosco-app .tab-btn.active { background: var(--teal); color: white; }
        .kiosco-app table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
        .kiosco-app th {
          text-align: left; font-size: 11px; letter-spacing: 0.04em; text-transform: uppercase;
          color: var(--ink-soft); font-weight: 700; padding: 8px 10px; border-bottom: 1px solid var(--line);
        }
        .kiosco-app td { padding: 8px 10px; border-bottom: 1px solid var(--paper-dark); vertical-align: middle; }
        .kiosco-app tr:last-child td { border-bottom: none; }
        .kiosco-app .cell-input { width: 100%; padding: 4px 6px; font-size: 13.5px; border-radius: 6px; }
        .kiosco-app .ticket {
          position: relative;
          background: var(--surface-dark);
          color: white;
          border-radius: 12px 12px 0 0;
          padding: 18px 20px 24px;
        }
        .kiosco-app .ticket::after {
          content: "";
          position: absolute; left: 0; right: 0; bottom: -1px; height: 14px;
          background-image:
            linear-gradient(-45deg, var(--paper) 7px, transparent 0),
            linear-gradient(45deg, var(--paper) 7px, transparent 0);
          background-size: 14px 14px;
          background-repeat: repeat-x;
          background-position: bottom;
        }
        .kiosco-app .ticket-row { display: flex; justify-content: space-between; padding: 3px 0; font-size: 13.5px; }
        .kiosco-app .ticket-dash { border-top: 1px dashed rgba(255,255,255,0.35); margin: 8px 0; }
        .kiosco-app .low-stock-badge {
          background: rgba(193,68,58,0.1); color: var(--red); font-weight: 700; font-size: 11px;
          padding: 2px 7px; border-radius: 999px; display: inline-flex; align-items: center; gap: 4px;
        }
        .kiosco-app .agent-bar {
          position: sticky; bottom: 0; margin: 20px -16px 0; padding: 12px 16px;
          background: linear-gradient(to bottom, transparent, var(--paper) 20%);
        }
        .kiosco-app .agent-inner {
          background: var(--surface-dark); border-radius: 14px; padding: 10px 12px;
        }
        .kiosco-app .agent-feed { max-height: 130px; overflow-y: auto; margin-bottom: 8px; }
        .kiosco-app .agent-msg { font-size: 13px; padding: 5px 0; color: #E9E5DA; line-height: 1.4; }
        .kiosco-app .agent-msg.user { color: var(--mustard); font-weight: 600; }
        .kiosco-app .agent-input-row { display: flex; gap: 8px; }
        .kiosco-app .agent-input-row input {
          flex: 1; background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.15);
          color: white; border-radius: 9px;
        }
        .kiosco-app .agent-input-row input::placeholder { color: rgba(255,255,255,0.45); }
        .kiosco-app .agent-icon-btn {
          background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.15); border-radius: 9px; width: 38px;
          display: flex; align-items: center; justify-content: center; cursor: pointer; color: #E9E5DA; flex-shrink: 0;
        }
        .kiosco-app .agent-icon-btn:hover { background: rgba(255,255,255,0.16); }
        .kiosco-app .agent-icon-btn.listening { background: var(--red); border-color: var(--red); color: white; animation: pulseMic 1.1s infinite; }
        @keyframes pulseMic { 0%,100% { opacity: 1; } 50% { opacity: 0.55; } }
        .kiosco-app .receipt-draft { background: rgba(255,255,255,0.06); border: 1px dashed rgba(255,255,255,0.25); border-radius: 10px; padding: 8px 10px; margin-bottom: 8px; }
        .kiosco-app .receipt-confirm, .kiosco-app .receipt-cancel {
          border: none; border-radius: 7px; padding: 5px 10px; font-size: 11.5px; font-weight: 700; cursor: pointer;
          display: inline-flex; align-items: center; gap: 4px;
        }
        .kiosco-app .receipt-confirm { background: var(--mustard); color: var(--surface-dark); }
        .kiosco-app .receipt-cancel { background: transparent; color: #E9E5DA; border: 1px solid rgba(255,255,255,0.25); }
        .kiosco-app .agent-send {
          background: var(--mustard); border: none; border-radius: 9px; width: 38px;
          display: flex; align-items: center; justify-content: center; cursor: pointer; color: var(--surface-dark);
        }
        .kiosco-app .agent-send:disabled { opacity: 0.5; cursor: not-allowed; }
        .kiosco-app .stat-card { background: var(--card-bg); border: 1px solid var(--line); border-radius: 12px; padding: 14px 16px; }
        .kiosco-app .stat-label { font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); font-weight: 700; }
        .kiosco-app .stat-value { font-family: 'JetBrains Mono', monospace; font-size: 22px; font-weight: 700; margin-top: 4px; }
      `}</style>

      {/* Header / ticket */}
      <div className="max-w-4xl mx-auto">
        <div className="flex items-center justify-between mb-4">
          <div>
            <h1 className="text-xl font-extrabold" style={{ letterSpacing: "-0.01em" }}>
              {adminUnlocked ? "Control del quiosco" : "Nuestro catálogo"}
            </h1>
            {adminUnlocked && (
              <p className="text-xs" style={{ color: "var(--ink-soft)" }}>
                {saveError ? "No se pudo guardar el último cambio — revisá tu conexión." : "Los datos se guardan solos, en este dispositivo y en cualquiera que abra este mismo enlace."}
              </p>
            )}
          </div>
          <div className="flex items-center gap-2" style={{ position: "relative" }}>
            <button className="btn-ghost" onClick={() => setThemePickerOpen((v) => !v)}>
              <Palette size={14} /> Tema
            </button>
            {themePickerOpen && (
              <div className="card" style={{ position: "absolute", top: "110%", right: 0, zIndex: 20, padding: 10, display: "flex", gap: 8, boxShadow: "0 6px 20px rgba(0,0,0,0.12)" }}>
                {Object.entries(THEMES).map(([key, t]) => (
                  <button
                    key={key}
                    title={t.label}
                    onClick={() => { setThemeKey(key); setThemePickerOpen(false); }}
                    style={{
                      width: 30, height: 30, borderRadius: "50%", background: t.swatch, cursor: "pointer",
                      border: themeKey === key ? "3px solid var(--ink)" : "1px solid var(--line)",
                      padding: 0,
                    }}
                  />
                ))}
              </div>
            )}
            {adminUnlocked ? (
              <>
                <button className="btn-ghost" onClick={() => exportExcel(store)}>
                  <Download size={14} /> Exportar Excel
                </button>
                <button className="btn-ghost" onClick={resetAll}>
                  <RotateCcw size={14} /> Reiniciar
                </button>
                <button className="btn-ghost" onClick={lockAdmin}>
                  <Lock size={14} /> Bloquear
                </button>
              </>
            ) : (
              <button className="btn-ghost" onClick={() => setTab("admin-gate")}>
                <Lock size={14} /> Modo administrador
              </button>
            )}
          </div>
        </div>

        {adminUnlocked && (
        <div className="card ticket mono">
          <div className="ticket-row"><span>Fecha</span><span>{new Date().toLocaleDateString("es-AR", { weekday: "long", day: "numeric", month: "long" })}</span></div>
          <div className="ticket-dash" />
          <div className="ticket-row"><span>Ventas de hoy</span><span>{fmt(todaysSalesTotal)}</span></div>
          <div className="ticket-row"><span>Gastos de hoy</span><span>{fmt(todaysExpensesTotal)}</span></div>
          <div className="ticket-row" style={{ fontWeight: 700 }}><span>Neto de hoy</span><span>{fmt(todaysSalesTotal - todaysExpensesTotal)}</span></div>
          {lowStock.length > 0 && (
            <>
              <div className="ticket-dash" />
              <div className="ticket-row"><span>⚠ Stock bajo</span><span>{lowStock.length} producto{lowStock.length > 1 ? "s" : ""}</span></div>
            </>
          )}
        </div>
        )}

        {/* Tabs */}
        <div className="flex gap-1 mt-5 mb-4 overflow-x-auto">
          <button className={`tab-btn ${tab === "catalogo" ? "active" : ""}`} onClick={() => setTab("catalogo")}><Store size={15} /> Catálogo</button>
          {adminUnlocked && (
            <>
              <button className={`tab-btn ${tab === "productos" ? "active" : ""}`} onClick={() => setTab("productos")}><Package size={15} /> Productos</button>
              <button className={`tab-btn ${tab === "ventas" ? "active" : ""}`} onClick={() => setTab("ventas")}><ShoppingCart size={15} /> Ventas</button>
              <button className={`tab-btn ${tab === "cierre" ? "active" : ""}`} onClick={() => setTab("cierre")}><ClipboardCheck size={15} /> Cierre</button>
              <button className={`tab-btn ${tab === "gastos" ? "active" : ""}`} onClick={() => setTab("gastos")}><Receipt size={15} /> Gastos</button>
              <button className={`tab-btn ${tab === "resumen" ? "active" : ""}`} onClick={() => setTab("resumen")}><BarChart3 size={15} /> Resumen</button>
              <button className={`tab-btn ${tab === "historial" ? "active" : ""}`} onClick={() => setTab("historial")}><History size={15} /> Historial</button>
            </>
          )}
        </div>

        <div className="pb-4">
          {tab === "admin-gate" && (
            <div className="card p-6 max-w-sm mx-auto text-center">
              <Lock size={22} style={{ color: "var(--teal)", margin: "0 auto 10px" }} />
              <h3 className="text-sm font-bold mb-1">
                {adminPin ? "Ingresá la clave de administrador" : "Definí una clave de administrador"}
              </h3>
              <p className="text-xs mb-3" style={{ color: "var(--ink-soft)" }}>
                {adminPin ? "Para editar productos, ventas, gastos y el resumen." : "La vas a usar de ahora en más para entrar al panel de administración. Guardala en un lugar seguro."}
              </p>
              <input
                type="password"
                value={pinInput}
                onChange={(e) => setPinInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && tryUnlock()}
                placeholder="Clave"
                className="w-full text-center mb-2"
                autoFocus
              />
              {pinError && <p className="text-xs mb-2" style={{ color: "var(--red)" }}>{pinError}</p>}
              <div className="flex gap-2 justify-center">
                <button className="btn-primary" onClick={tryUnlock}>Entrar</button>
                <button className="btn-ghost" onClick={() => { setTab("catalogo"); setPinInput(""); setPinError(""); setRecoveryMode(false); }}>Cancelar</button>
              </div>
              {adminPin && !recoveryMode && (
                <button className="text-xs mt-3" style={{ background: "none", border: "none", color: "var(--teal)", cursor: "pointer", textDecoration: "underline" }} onClick={() => setRecoveryMode(true)}>
                  Olvidé mi clave
                </button>
              )}
              {adminPin && recoveryMode && (
                <div className="mt-3 pt-3" style={{ borderTop: "1px dashed var(--line)" }}>
                  <p className="text-xs mb-2" style={{ color: "var(--ink-soft)" }}>Ingresá tu código de recuperación para definir una clave nueva.</p>
                  <input
                    value={recoveryInput}
                    onChange={(e) => setRecoveryInput(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && recoverPin()}
                    placeholder="Código de recuperación"
                    className="w-full text-center mb-2"
                  />
                  {recoveryError && <p className="text-xs mb-2" style={{ color: "var(--red)" }}>{recoveryError}</p>}
                  <div className="flex gap-2 justify-center">
                    <button className="btn-primary" onClick={recoverPin}>Recuperar</button>
                    <button className="btn-ghost" onClick={() => { setRecoveryMode(false); setRecoveryInput(""); setRecoveryError(""); }}>Volver</button>
                  </div>
                </div>
              )}
            </div>
          )}
          {tab === "productos" && adminUnlocked && (
            <ProductosTab products={store.products} onAdd={addProduct} onUpdate={updateProduct} onDelete={deleteProduct} onRestock={restock} />
          )}
          {tab === "ventas" && adminUnlocked && (
            <VentasTab products={store.products} sales={store.sales} onSale={recordSale} onDelete={deleteSale} />
          )}
          {tab === "cierre" && adminUnlocked && (
            <CierreTab products={store.products} onSale={recordSale} />
          )}
          {tab === "gastos" && adminUnlocked && (
            <GastosTab expenses={store.expenses} onAdd={addExpense} onUpdate={updateExpense} onDelete={deleteExpense} />
          )}
          {tab === "resumen" && adminUnlocked && (
            <ResumenTab
              totalSalesAll={totalSalesAll}
              totalExpensesAll={totalExpensesAll}
              grossProfit={grossProfit}
              salesByProduct={salesByProduct}
              lowStock={lowStock}
              sales={store.sales}
              expenses={store.expenses}
            />
          )}
          {tab === "catalogo" && adminUnlocked && (
            <CatalogoTab products={store.products} onUpdate={updateProduct} onPresent={() => setPresentMode(true)} whatsappNumber={store.whatsappNumber} onWhatsapp={(v) => setStore((prev) => ({ ...prev, whatsappNumber: v }))} />
          )}
          {tab === "catalogo" && !adminUnlocked && (
            <CatalogoPublico products={store.products} whatsappNumber={store.whatsappNumber} onPresent={() => setPresentMode(true)} cart={cart} onQty={setCartQty} onClear={() => setCart({})} />
          )}
          {tab === "historial" && adminUnlocked && <HistorialTab priceHistory={store.priceHistory} />}
        </div>
      </div>

      {/* Agent bar */}
      {adminUnlocked && (
      <div className="agent-bar max-w-4xl mx-auto" style={{ left: 0, right: 0 }}>
        <div className="agent-inner">
          <div className="flex items-center gap-2 mb-2" style={{ color: "var(--mustard)" }}>
            <Bot size={15} />
            <span className="text-xs font-bold uppercase" style={{ letterSpacing: "0.04em" }}>Agente del quiosco</span>
          </div>
          <div className="agent-feed" ref={feedRef}>
            {agentMessages.map((m, i) => (
              <div key={i} className={`agent-msg ${m.role === "user" ? "user" : ""}`}>
                {m.role === "user" ? "Vos: " : "Agente: "}{m.text}
              </div>
            ))}
            {agentBusy && <div className="agent-msg">Agente: pensando…</div>}
          </div>
          {pendingReceipt && (
            <div className="receipt-draft">
              <div className="text-xs font-bold uppercase mb-1" style={{ color: "var(--mustard)" }}>Confirmá lo que leí de la foto</div>
              <ul style={{ margin: 0, paddingLeft: 16 }}>
                {pendingReceipt.actions.map((a, i) => (
                  <li key={i} style={{ fontSize: 12.5, color: "#E9E5DA", padding: "2px 0" }}>{describeAction(a)}</li>
                ))}
              </ul>
              <div className="flex gap-2 mt-2">
                <button className="receipt-confirm" onClick={confirmReceipt}><Check size={13} /> Confirmar y cargar</button>
                <button className="receipt-cancel" onClick={cancelReceipt}><Ban size={13} /> Descartar</button>
              </div>
            </div>
          )}
          <div className="agent-input-row">
            <input
              value={agentInput}
              onChange={(e) => setAgentInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleAgentSubmit()}
              placeholder={listening ? "Escuchando…" : 'Ej: "vendí 3 alfajores y 1 gaseosa de litro"'}
              disabled={agentBusy}
            />
            <input ref={receiptInputRef} type="file" accept="image/*" capture="environment" onChange={handleReceiptFile} style={{ display: "none" }} />
            <button
              className="agent-icon-btn"
              onClick={() => receiptInputRef.current && receiptInputRef.current.click()}
              disabled={agentBusy}
              title="Sacarle foto a una factura o comprobante"
            >
              <Camera size={16} />
            </button>
            {speechSupported && (
              <button
                className={`agent-icon-btn ${listening ? "listening" : ""}`}
                onClick={toggleListening}
                disabled={agentBusy}
                title="Dictar por voz"
              >
                <Mic size={16} />
              </button>
            )}
            <button className="agent-send" onClick={handleAgentSubmit} disabled={agentBusy || !agentInput.trim()}>
              {agentBusy ? <Loader2 size={16} className="animate-spin" /> : <Send size={16} />}
            </button>
          </div>
        </div>
      </div>
      )}
    </div>
  );
}

// Miniatura clickeable que abre el selector de archivos y sube+comprime la foto elegida
function ImageUploadCell({ imageUrl, onImage, size = 32, rounded = 6, iconSize = 14 }) {
  const inputRef = useRef(null);
  const [busy, setBusy] = useState(false);

  async function handleChange(e) {
    const file = e.target.files[0];
    if (!file) return;
    setBusy(true);
    try {
      const dataUrl = await compressImageFile(file);
      onImage(dataUrl);
    } catch (err) {
      console.error(err);
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  }

  return (
    <button
      type="button"
      onClick={() => inputRef.current && inputRef.current.click()}
      title="Subir foto"
      style={{
        width: size, height: size, borderRadius: rounded, border: "1px dashed var(--line)",
        background: imageUrl ? "transparent" : "var(--paper-dark)",
        display: "flex", alignItems: "center", justifyContent: "center",
        color: "var(--ink-soft)", cursor: "pointer", padding: 0, overflow: "hidden", position: "relative",
      }}
    >
      <input ref={inputRef} type="file" accept="image/*" onChange={handleChange} style={{ display: "none" }} />
      {busy ? (
        <Loader2 size={iconSize} className="animate-spin" />
      ) : imageUrl ? (
        <img src={imageUrl} alt="" style={{ width: "100%", height: "100%", objectFit: "cover", objectPosition: "center" }} onError={(e) => { e.target.style.display = "none"; }} />
      ) : (
        <Upload size={iconSize} />
      )}
    </button>
  );
}

// Campo numérico que sólo confirma el cambio (y por lo tanto el historial) al salir del campo
function EditableNumberCell({ value, onCommit, ...props }) {
  const [val, setVal] = useState(value);
  useEffect(() => setVal(value), [value]);
  return (
    <input
      {...props}
      type="number"
      value={val}
      onChange={(e) => setVal(e.target.value)}
      onBlur={() => {
        const num = Number(val);
        if (!isNaN(num) && num !== value) onCommit(num);
        else setVal(value);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") e.target.blur();
      }}
    />
  );
}

// ---------------- Productos ----------------
function ProductosTab({ products, onAdd, onUpdate, onDelete, onRestock }) {
  const [form, setForm] = useState({ name: "", category: "", price: "", cost: "", stock: "", minStock: "5", imageUrl: "" });
  const [restockFor, setRestockFor] = useState(null);
  const [restockQty, setRestockQty] = useState("");

  function submit() {
    if (!form.name.trim() || !form.price) return;
    onAdd(form);
    setForm({ name: "", category: "", price: "", cost: "", stock: "", minStock: "5", imageUrl: "" });
  }

  const byCategory = useMemo(() => {
    const map = {};
    products.forEach((p) => {
      const cat = p.category || "General";
      if (!map[cat]) map[cat] = [];
      map[cat].push(p);
    });
    return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
  }, [products]);

  return (
    <div className="space-y-4">
      <div className="card p-3">
        <div className="grid grid-cols-2 md:grid-cols-6 gap-2">
          <input placeholder="Nombre" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className="col-span-2 md:col-span-2" />
          <input placeholder="Categoría" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} />
          <input placeholder="Precio venta" type="number" value={form.price} onChange={(e) => setForm({ ...form, price: e.target.value })} />
          <input placeholder="Costo" type="number" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} />
          <input placeholder="Stock inicial" type="number" value={form.stock} onChange={(e) => setForm({ ...form, stock: e.target.value })} />
          <input placeholder="URL de imagen (opcional)" value={form.imageUrl} onChange={(e) => setForm({ ...form, imageUrl: e.target.value })} className="col-span-2 md:col-span-3" />
        </div>
        <div className="flex justify-end mt-2">
          <button className="btn-primary" onClick={submit} disabled={!form.name.trim() || !form.price}>
            <Plus size={15} /> Agregar producto
          </button>
        </div>
      </div>

      <div className="card overflow-x-auto">
        <table>
          <thead>
            <tr>
              <th></th><th>Producto</th><th>Categoría</th><th>Precio</th><th>Costo</th><th>Margen</th><th>Stock</th><th>Mín.</th><th></th>
            </tr>
          </thead>
          <tbody>
            {products.length === 0 && (
              <tr><td colSpan={9} style={{ textAlign: "center", color: "var(--ink-soft)", padding: "24px" }}>Todavía no cargaste productos. Sumá el primero arriba.</td></tr>
            )}
            {byCategory.map(([cat, items]) => (
              <React.Fragment key={cat}>
                <tr>
                  <td colSpan={9} style={{ background: "var(--paper-dark)", fontWeight: 700, fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.04em", color: "var(--ink-soft)", padding: "6px 10px", borderBottom: "1px solid var(--line)" }}>
                    {cat} <span style={{ fontWeight: 500, textTransform: "none" }}>({items.length})</span>
                  </td>
                </tr>
                {items.map((p) => {
                  const margin = p.price > 0 && p.cost > 0 ? ((p.price - p.cost) / p.price) * 100 : null;
                  return (
                  <tr key={p.id}>
                <td>
                  <ImageUploadCell imageUrl={p.imageUrl} onImage={(dataUrl) => onUpdate(p.id, { imageUrl: dataUrl })} size={32} />
                </td>
                <td><input className="cell-input" value={p.name} onChange={(e) => onUpdate(p.id, { name: e.target.value })} /></td>
                <td><input className="cell-input" value={p.category} onChange={(e) => onUpdate(p.id, { category: e.target.value })} /></td>
                <td><EditableNumberCell className="cell-input mono" value={p.price} onCommit={(v) => onUpdate(p.id, { price: v })} /></td>
                <td><EditableNumberCell className="cell-input mono" value={p.cost} onCommit={(v) => onUpdate(p.id, { cost: v })} /></td>
                <td className="mono font-bold" style={{ color: margin == null ? "var(--ink-soft)" : margin < 15 ? "var(--red)" : "var(--teal)" }}>
                  {margin == null ? "—" : `${margin.toFixed(0)}%`}
                </td>
                <td>
                  <div className="flex items-center gap-1">
                    <input className="cell-input mono" style={{ width: "60px" }} type="number" value={p.stock} onChange={(e) => onUpdate(p.id, { stock: Number(e.target.value) })} />
                    {p.stock <= (p.minStock ?? 5) && <span className="low-stock-badge"><AlertTriangle size={11} />bajo</span>}
                  </div>
                </td>
                <td><input className="cell-input mono" style={{ width: "50px" }} type="number" value={p.minStock} onChange={(e) => onUpdate(p.id, { minStock: Number(e.target.value) })} /></td>
                <td>
                  <div className="flex items-center gap-1">
                    {restockFor === p.id ? (
                      <>
                        <input autoFocus className="cell-input mono" style={{ width: "55px" }} type="number" placeholder="cant." value={restockQty} onChange={(e) => setRestockQty(e.target.value)} />
                        <button className="icon-btn" style={{ color: "var(--teal)" }} onClick={() => { onRestock(p.id, Number(restockQty) || 0); setRestockFor(null); setRestockQty(""); }}><Plus size={15} /></button>
                      </>
                    ) : (
                      <button className="btn-ghost" style={{ padding: "4px 8px" }} onClick={() => setRestockFor(p.id)}>Reponer</button>
                    )}
                    <button className="icon-btn" onClick={() => onDelete(p.id)}><Trash2 size={15} /></button>
                  </div>
                </td>
              </tr>
              );
                })}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------------- Ventas ----------------
function VentasTab({ products, sales, onSale, onDelete }) {
  const [productId, setProductId] = useState("");
  const [qty, setQty] = useState("1");

  useEffect(() => {
    if (!productId && products.length) setProductId(products[0].id);
  }, [products]);

  const selected = products.find((p) => p.id === productId);
  const total = selected ? selected.price * (Number(qty) || 0) : 0;

  function submit() {
    if (!productId || !qty) return;
    onSale({ productId, quantity: Number(qty) });
    setQty("1");
  }

  const dailyTotals = useMemo(() => {
    const map = {};
    sales.forEach((s) => {
      const key = todayKey(s.date);
      if (!map[key]) map[key] = { date: s.date, total: 0, count: 0 };
      map[key].total += s.total;
      map[key].count += 1;
    });
    return Object.values(map).sort((a, b) => new Date(b.date) - new Date(a.date));
  }, [sales]);

  return (
    <div className="space-y-4">
      <div className="card p-3">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2 items-center">
          <select className="col-span-2" value={productId} onChange={(e) => setProductId(e.target.value)}>
            {products.length === 0 && <option value="">Cargá un producto primero</option>}
            {products.map((p) => (
              <option key={p.id} value={p.id}>{p.name} (stock: {p.stock})</option>
            ))}
          </select>
          <input type="number" min="1" value={qty} onChange={(e) => setQty(e.target.value)} placeholder="Cantidad" />
          <div className="mono text-sm font-bold flex items-center justify-end" style={{ color: "var(--teal)" }}>{fmt(total)}</div>
        </div>
        <div className="flex justify-end mt-2">
          <button className="btn-primary" onClick={submit} disabled={!productId || !qty}>
            <Plus size={15} /> Registrar venta
          </button>
        </div>
      </div>

      <div className="card p-4">
        <h3 className="text-sm font-bold mb-3">Venta total por día</h3>
        {dailyTotals.length === 0 ? (
          <p className="text-sm" style={{ color: "var(--ink-soft)" }}>Todavía no hay ventas registradas.</p>
        ) : (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
            {dailyTotals.map((d) => (
              <div key={d.date + d.total} className="stat-card" style={{ padding: "10px 12px" }}>
                <div className="stat-label" style={{ fontSize: 10.5 }}>
                  {new Date(d.date).toLocaleDateString("es-AR", { weekday: "short", day: "2-digit", month: "2-digit" })}
                </div>
                <div className="mono font-extrabold" style={{ fontSize: 17, color: "var(--teal)" }}>{fmt(d.total)}</div>
                <div className="text-xs" style={{ color: "var(--ink-soft)" }}>{d.count} venta{d.count > 1 ? "s" : ""}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="card overflow-x-auto">
        <table>
          <thead>
            <tr><th>Fecha</th><th>Producto</th><th>Cant.</th><th>Precio</th><th>Total</th><th></th></tr>
          </thead>
          <tbody>
            {sales.length === 0 && (
              <tr><td colSpan={6} style={{ textAlign: "center", color: "var(--ink-soft)", padding: "24px" }}>Todavía no registraste ventas.</td></tr>
            )}
            {sales.map((s) => (
              <tr key={s.id}>
                <td className="mono" style={{ fontSize: "12.5px" }}>{new Date(s.date).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}</td>
                <td>{s.productName}</td>
                <td className="mono">{s.quantity}</td>
                <td className="mono">{fmt(s.price)}</td>
                <td className="mono font-bold">{fmt(s.total)}</td>
                <td><button className="icon-btn" onClick={() => onDelete(s.id)}><Trash2 size={15} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------------- Cierre de recreo ----------------
function CierreTab({ products, onSale }) {
  const [counts, setCounts] = useState({});
  const [cash, setCash] = useState("");
  const [done, setDone] = useState(false);

  const rows = [...products]
    .sort((a, b) => (a.category || "").localeCompare(b.category || "") || a.name.localeCompare(b.name))
    .map((p) => {
      const raw = counts[p.id];
      const counted = raw === undefined || raw === "" ? null : Number(raw);
      const invalid = counted != null && counted > p.stock;
      const sold = counted == null || invalid ? 0 : p.stock - counted;
      return { p, counted, invalid, sold, subtotal: sold * p.price };
    });
  const totalExpected = rows.reduce((a, r) => a + r.subtotal, 0);
  const totalUnits = rows.reduce((a, r) => a + r.sold, 0);
  const hasInvalid = rows.some((r) => r.invalid);
  const diff = cash !== "" ? Number(cash) - totalExpected : null;

  function apply() {
    rows.forEach((r) => {
      if (r.sold > 0) onSale({ productId: r.p.id, quantity: r.sold });
    });
    setCounts({});
    setCash("");
    setDone(true);
  }

  return (
    <div className="space-y-4">
      <div className="card p-4">
        <h3 className="text-sm font-bold mb-1">Cierre de recreo</h3>
        <p className="text-sm" style={{ color: "var(--ink-soft)" }}>
          Cuando termine el recreo, contá lo que te quedó de cada producto y anotalo en "Quedó". La app calcula cuánto vendiste de cada uno (stock anterior menos lo que quedó) y carga todas las ventas juntas. Dejá en blanco los productos que no se movieron.
        </p>
      </div>

      {done && (
        <div className="card p-3 text-sm font-bold" style={{ color: "var(--teal)" }}>
          Listo: se cargaron las ventas del recreo y el stock quedó actualizado.
        </div>
      )}

      <div className="card overflow-x-auto">
        <table>
          <thead>
            <tr><th>Producto</th><th>Antes</th><th>Quedó</th><th>Vendidas</th><th>Subtotal</th></tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--ink-soft)", padding: "24px" }}>Cargá productos primero.</td></tr>
            )}
            {rows.map((r) => (
              <tr key={r.p.id}>
                <td>{r.p.name}</td>
                <td className="mono">{r.p.stock}</td>
                <td>
                  <input
                    className="cell-input mono"
                    style={{ width: 70, borderColor: r.invalid ? "var(--red)" : undefined }}
                    type="number"
                    min="0"
                    value={counts[r.p.id] ?? ""}
                    onChange={(e) => { setCounts({ ...counts, [r.p.id]: e.target.value }); setDone(false); }}
                  />
                  {r.invalid && <div style={{ fontSize: 10.5, color: "var(--red)" }}>Más que el stock cargado: reponé primero</div>}
                </td>
                <td className="mono font-bold">{r.sold}</td>
                <td className="mono">{fmt(r.subtotal)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="card p-4 space-y-3">
        <div className="grid grid-cols-2 gap-3">
          <div>
            <div className="stat-label">Unidades vendidas</div>
            <div className="mono font-extrabold" style={{ fontSize: 20 }}>{totalUnits}</div>
          </div>
          <div>
            <div className="stat-label">Total esperado</div>
            <div className="mono font-extrabold" style={{ fontSize: 20, color: "var(--teal)" }}>{fmt(totalExpected)}</div>
          </div>
        </div>
        <div>
          <label className="stat-label">Plata que juntaste (opcional, para comparar)</label>
          <input type="number" value={cash} onChange={(e) => setCash(e.target.value)} placeholder="Ej: 45000" style={{ display: "block", marginTop: 4, maxWidth: 220 }} />
          {diff != null && (
            <p className="text-sm font-bold mt-2" style={{ color: diff === 0 ? "var(--teal)" : "var(--red)" }}>
              {diff === 0 ? "Coincide justo con lo esperado." : diff < 0 ? `Faltan ${fmt(Math.abs(diff))} respecto a lo esperado.` : `Sobran ${fmt(diff)} respecto a lo esperado.`}
            </p>
          )}
        </div>
        <div className="flex justify-end">
          <button className="btn-primary" onClick={apply} disabled={totalUnits === 0 || hasInvalid}>
            <Check size={15} /> Cargar ventas del recreo
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------- Gastos ----------------
function GastosTab({ expenses, onAdd, onUpdate, onDelete }) {
  const [form, setForm] = useState({ description: "", amount: "", type: "variable" });

  function submit() {
    if (!form.description.trim() || !form.amount) return;
    onAdd(form);
    setForm({ description: "", amount: "", type: "variable" });
  }

  const dailyTotals = useMemo(() => {
    const map = {};
    expenses.forEach((e) => {
      const key = todayKey(e.date);
      if (!map[key]) map[key] = { date: e.date, total: 0, count: 0 };
      map[key].total += e.amount;
      map[key].count += 1;
    });
    return Object.values(map).sort((a, b) => new Date(b.date) - new Date(a.date));
  }, [expenses]);

  return (
    <div className="space-y-4">
      <div className="card p-3">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
          <input className="md:col-span-2" placeholder="Descripción (ej: luz, mercadería, alquiler)" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />
          <input type="number" placeholder="Monto" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} />
          <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
            <option value="variable">Variable</option>
            <option value="fijo">Fijo</option>
          </select>
        </div>
        <div className="flex justify-end mt-2">
          <button className="btn-primary" onClick={submit} disabled={!form.description.trim() || !form.amount}>
            <Plus size={15} /> Agregar gasto
          </button>
        </div>
      </div>

      <div className="card p-4">
        <h3 className="text-sm font-bold mb-3">Gasto total por día</h3>
        {dailyTotals.length === 0 ? (
          <p className="text-sm" style={{ color: "var(--ink-soft)" }}>Todavía no hay gastos registrados.</p>
        ) : (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
            {dailyTotals.map((d) => (
              <div key={d.date + d.total} className="stat-card" style={{ padding: "10px 12px" }}>
                <div className="stat-label" style={{ fontSize: 10.5 }}>
                  {new Date(d.date).toLocaleDateString("es-AR", { weekday: "short", day: "2-digit", month: "2-digit" })}
                </div>
                <div className="mono font-extrabold" style={{ fontSize: 17, color: "var(--red)" }}>{fmt(d.total)}</div>
                <div className="text-xs" style={{ color: "var(--ink-soft)" }}>{d.count} gasto{d.count > 1 ? "s" : ""}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="card overflow-x-auto">
        <table>
          <thead><tr><th>Fecha</th><th>Descripción</th><th>Tipo</th><th>Monto</th><th></th></tr></thead>
          <tbody>
            {expenses.length === 0 && (
              <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--ink-soft)", padding: "24px" }}>Todavía no cargaste gastos.</td></tr>
            )}
            {expenses.map((e) => (
              <tr key={e.id}>
                <td className="mono" style={{ fontSize: "12.5px" }}>{new Date(e.date).toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit" })}</td>
                <td><input className="cell-input" value={e.description} onChange={(ev) => onUpdate(e.id, { description: ev.target.value })} /></td>
                <td>
                  <select className="cell-input" value={e.type || "variable"} onChange={(ev) => onUpdate(e.id, { type: ev.target.value })}>
                    <option value="variable">Variable</option>
                    <option value="fijo">Fijo</option>
                  </select>
                </td>
                <td><input className="cell-input mono" type="number" value={e.amount} onChange={(ev) => onUpdate(e.id, { amount: Number(ev.target.value) })} /></td>
                <td><button className="icon-btn" onClick={() => onDelete(e.id)}><Trash2 size={15} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------------- Resumen ----------------
function ResumenTab({ totalSalesAll, totalExpensesAll, grossProfit, salesByProduct, lowStock, sales, expenses }) {
  const weekly = useMemo(() => {
    const DAY = 24 * 60 * 60 * 1000;
    const now = new Date();
    const startThis = new Date(now.getTime() - 6 * DAY);
    startThis.setHours(0, 0, 0, 0);
    const startPrev = new Date(startThis.getTime() - 7 * DAY);
    let thisWeek = 0;
    let prevWeek = 0;
    (sales || []).forEach((s) => {
      const d = new Date(s.date);
      if (d >= startThis) thisWeek += s.total;
      else if (d >= startPrev && d < startThis) prevWeek += s.total;
    });
    const diff = prevWeek > 0 ? ((thisWeek - prevWeek) / prevWeek) * 100 : null;
    return { thisWeek, prevWeek, diff };
  }, [sales]);

  const breakeven = useMemo(() => {
    const fixed = (expenses || []).filter((e) => e.type === "fijo").reduce((a, e) => a + e.amount, 0);
    const variable = (expenses || []).filter((e) => e.type !== "fijo").reduce((a, e) => a + e.amount, 0);
    // días distintos con al menos una venta, para estimar un promedio de venta diaria
    const daysWithSales = new Set((sales || []).map((s) => todayKey(s.date))).size || 1;
    const totalSales = (sales || []).reduce((a, s) => a + s.total, 0);
    const avgDailySales = totalSales / daysWithSales;
    const dailyTargetToCoverFixed = avgDailySales > 0 ? fixed / daysWithSales : null; // referencia simple
    return { fixed, variable, avgDailySales, daysWithSales };
  }, [expenses, sales]);

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
        <div className="stat-card">
          <div className="stat-label">Ventas totales</div>
          <div className="stat-value" style={{ color: "var(--teal)" }}>{fmt(totalSalesAll)}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Gastos totales</div>
          <div className="stat-value" style={{ color: "var(--red)" }}>{fmt(totalExpensesAll)}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Ganancia estimada</div>
          <div className="stat-value flex items-center gap-1" style={{ color: grossProfit >= 0 ? "var(--teal)" : "var(--red)" }}>
            {grossProfit >= 0 ? <TrendingUp size={18} /> : <TrendingDown size={18} />} {fmt(grossProfit)}
          </div>
        </div>
      </div>

      <div className="card p-4">
        <h3 className="text-sm font-bold mb-3">Esta semana vs. la anterior</h3>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <div className="stat-label">Últimos 7 días</div>
            <div className="mono font-extrabold" style={{ fontSize: 20, color: "var(--teal)" }}>{fmt(weekly.thisWeek)}</div>
          </div>
          <div>
            <div className="stat-label">7 días anteriores</div>
            <div className="mono font-extrabold" style={{ fontSize: 20, color: "var(--ink-soft)" }}>{fmt(weekly.prevWeek)}</div>
          </div>
        </div>
        {weekly.diff != null && (
          <div className="mt-2 text-sm font-bold flex items-center gap-1" style={{ color: weekly.diff >= 0 ? "var(--teal)" : "var(--red)" }}>
            {weekly.diff >= 0 ? <TrendingUp size={14} /> : <TrendingDown size={14} />}
            {weekly.diff >= 0 ? "+" : ""}{weekly.diff.toFixed(1)}% respecto a la semana anterior
          </div>
        )}
      </div>

      <div className="card p-4">
        <h3 className="text-sm font-bold mb-1">Punto de equilibrio (estimado)</h3>
        <p className="text-xs mb-3" style={{ color: "var(--ink-soft)" }}>
          Es una referencia simple, no un cálculo contable exacto: te muestra cuánto de tus gastos son fijos y cuál es tu promedio de venta diaria actual.
        </p>
        <div className="grid grid-cols-2 gap-3 mb-2">
          <div>
            <div className="stat-label">Gastos fijos (total)</div>
            <div className="mono font-bold" style={{ color: "var(--red)" }}>{fmt(breakeven.fixed)}</div>
          </div>
          <div>
            <div className="stat-label">Gastos variables (total)</div>
            <div className="mono font-bold" style={{ color: "var(--red)" }}>{fmt(breakeven.variable)}</div>
          </div>
        </div>
        <div>
          <div className="stat-label">Promedio de venta por día ({breakeven.daysWithSales} día{breakeven.daysWithSales !== 1 ? "s" : ""} con ventas)</div>
          <div className="mono font-bold" style={{ color: "var(--teal)" }}>{fmt(breakeven.avgDailySales)}</div>
        </div>
        {breakeven.fixed > 0 && breakeven.avgDailySales > 0 && (
          <p className="text-xs mt-2" style={{ color: "var(--ink-soft)" }}>
            Repartiendo tus gastos fijos entre los días con ventas, necesitás vender en promedio <strong className="mono" style={{ color: "var(--ink)" }}>{fmt(breakeven.fixed / breakeven.daysWithSales)}</strong> por día solo para cubrirlos.
          </p>
        )}
      </div>

      <div className="card p-4">
        <h3 className="text-sm font-bold mb-3">Ventas por producto</h3>
        {salesByProduct.length === 0 ? (
          <p className="text-sm" style={{ color: "var(--ink-soft)" }}>Todavía no hay ventas para mostrar.</p>
        ) : (
          <div style={{ width: "100%", height: Math.max(180, salesByProduct.length * 34) }}>
            <ResponsiveContainer>
              <BarChart data={salesByProduct} layout="vertical" margin={{ left: 10, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F0EAD9" horizontal={false} />
                <XAxis type="number" tickFormatter={(v) => `$${v}`} fontSize={11} />
                <YAxis type="category" dataKey="name" width={110} fontSize={12} />
                <Tooltip formatter={(v) => fmt(v)} />
                <Bar dataKey="total" fill="#0F5C4E" radius={[0, 6, 6, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      <div className="card p-4">
        <h3 className="text-sm font-bold mb-3 flex items-center gap-2"><AlertTriangle size={15} style={{ color: "var(--red)" }} /> Stock bajo</h3>
        {lowStock.length === 0 ? (
          <p className="text-sm" style={{ color: "var(--ink-soft)" }}>Todo el stock está por encima del mínimo. 👍</p>
        ) : (
          <ul className="space-y-1">
            {lowStock.map((p) => (
              <li key={p.id} className="flex justify-between text-sm">
                <span>{p.name}</span>
                <span className="mono font-bold" style={{ color: "var(--red)" }}>{p.stock} und. (mín. {p.minStock})</span>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

// ---------------- Historial de precios ----------------
function HistorialTab({ priceHistory }) {
  const rows = priceHistory || [];
  return (
    <div className="space-y-4">
      <div className="card p-4">
        <p className="text-sm" style={{ color: "var(--ink-soft)" }}>
          Cada vez que cambiás el precio de venta o el costo de un producto, queda registrado acá con fecha. Las ventas ya realizadas guardan el precio y el costo que tenían en ese momento, así que la ganancia de cada venta no se recalcula si después cambiás el precio.
        </p>
      </div>
      <div className="card overflow-x-auto">
        <table>
          <thead>
            <tr><th>Fecha</th><th>Producto</th><th>Campo</th><th>Antes</th><th>Después</th><th>Variación</th></tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr><td colSpan={6} style={{ textAlign: "center", color: "var(--ink-soft)", padding: "24px" }}>Todavía no hubo cambios de precio o costo.</td></tr>
            )}
            {rows.map((h) => {
              const variation = h.oldValue ? ((h.newValue - h.oldValue) / h.oldValue) * 100 : null;
              const up = h.newValue > h.oldValue;
              return (
                <tr key={h.id}>
                  <td className="mono" style={{ fontSize: "12.5px" }}>{new Date(h.date).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "2-digit", hour: "2-digit", minute: "2-digit" })}</td>
                  <td>{h.productName}</td>
                  <td>{h.field === "price" ? "Precio de venta" : "Costo"}</td>
                  <td className="mono">{fmt(h.oldValue)}</td>
                  <td className="mono font-bold">{fmt(h.newValue)}</td>
                  <td className="mono font-bold" style={{ color: variation == null ? "var(--ink-soft)" : up ? "var(--teal)" : "var(--red)" }}>
                    {variation == null ? "—" : `${up ? "+" : ""}${variation.toFixed(1)}%`}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------------- Catálogo (vista embebida en el panel) ----------------
function CatalogoTab({ products, onUpdate, onPresent, whatsappNumber, onWhatsapp }) {
  const byCategory = useMemo(() => {
    const map = {};
    products.forEach((p) => {
      const cat = p.category || "General";
      if (!map[cat]) map[cat] = [];
      map[cat].push(p);
    });
    return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
  }, [products]);

  return (
    <div className="space-y-4">
      <div className="card p-3 flex items-center justify-between flex-wrap gap-2">
        <p className="text-sm" style={{ color: "var(--ink-soft)" }}>
          Tocá el recuadro de cada producto para subir su foto desde el celular o la compu. También podés pegar una URL de imagen si preferís.
        </p>
        <button className="btn-primary" onClick={onPresent} disabled={products.length === 0}>
          <Maximize2 size={15} /> Pantalla completa
        </button>
      </div>

      <div className="card p-3">
        <label className="text-xs font-bold uppercase" style={{ color: "var(--ink-soft)", letterSpacing: "0.04em" }}>WhatsApp para pedidos</label>
        <p className="text-xs mb-2" style={{ color: "var(--ink-soft)" }}>
          Cargá tu número con código de país, sin espacios ni signos (ej: 5493704123456). Los clientes van a poder armar su pedido en el catálogo (con cantidades) y enviártelo ya armado por WhatsApp.
        </p>
        <input
          value={whatsappNumber || ""}
          onChange={(e) => onWhatsapp(e.target.value.replace(/[^0-9]/g, ""))}
          placeholder="5493704123456"
          style={{ maxWidth: 260 }}
        />
      </div>

      {products.length === 0 ? (
        <div className="card p-6 text-center text-sm" style={{ color: "var(--ink-soft)" }}>
          Cargá productos en la pestaña "Productos" para que aparezcan acá.
        </div>
      ) : (
        byCategory.map(([cat, items]) => (
          <div key={cat}>
            <h3 className="text-xs font-bold uppercase mb-2" style={{ letterSpacing: "0.05em", color: "var(--ink-soft)" }}>{cat} <span style={{ fontWeight: 500, textTransform: "none" }}>({items.length})</span></h3>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {items.map((p) => (
                <div key={p.id} className="card p-2">
                  <div style={{ aspectRatio: "1", width: "100%" }}>
                    <ImageUploadCell
                      imageUrl={p.imageUrl}
                      onImage={(dataUrl) => onUpdate(p.id, { imageUrl: dataUrl })}
                      size={"100%"}
                      rounded={8}
                      iconSize={22}
                    />
                  </div>
                  <div className="mt-2 text-sm font-semibold truncate">{p.name}</div>
                  <div className="mono font-bold text-sm" style={{ color: "var(--teal)" }}>{fmt(p.price)}</div>
                  <input
                    className="cell-input mt-1"
                    placeholder="o pegá una URL de imagen"
                    value={p.imageUrl && p.imageUrl.startsWith("data:") ? "" : p.imageUrl || ""}
                    onChange={(e) => onUpdate(p.id, { imageUrl: e.target.value })}
                    style={{ fontSize: "11px" }}
                  />
                </div>
              ))}
            </div>
          </div>
        ))
      )}
    </div>
  );
}

// ---------------- Pedido por WhatsApp ----------------
function QtyControl({ value, max, onChange, compact }) {
  const btn = {
    width: 28, height: 28, borderRadius: 8, border: "1px solid var(--line)", background: "var(--card-bg)",
    color: "var(--ink)", cursor: "pointer", fontWeight: 800, fontSize: 16, lineHeight: 1,
    display: "flex", alignItems: "center", justifyContent: "center", padding: 0,
  };
  if (!value) {
    return (
      <button
        type="button"
        onClick={() => onChange(1)}
        style={{ marginTop: 6, width: "100%", padding: "6px 8px", borderRadius: 8, border: "none", background: "var(--teal)", color: "white", fontWeight: 700, fontSize: 12, cursor: "pointer" }}
      >
        {compact ? "Agregar" : "Agregar al pedido"}
      </button>
    );
  }
  return (
    <div style={{ marginTop: 6, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 }}>
      <button type="button" style={btn} onClick={() => onChange(value - 1)}>−</button>
      <span className="mono" style={{ fontWeight: 800, fontSize: 14 }}>{value}</span>
      <button type="button" style={{ ...btn, opacity: value >= max ? 0.4 : 1 }} disabled={value >= max} onClick={() => onChange(value + 1)}>+</button>
    </div>
  );
}

function CartBar({ products, cart, onQty, onClear, whatsappNumber }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState("");
  const [note, setNote] = useState("");

  const items = products
    .filter((p) => cart[p.id] > 0)
    .map((p) => ({ p, qty: Math.min(cart[p.id], p.stock) }))
    .filter((i) => i.qty > 0);
  if (!whatsappNumber || items.length === 0) return null;

  const total = items.reduce((a, i) => a + i.p.price * i.qty, 0);
  const units = items.reduce((a, i) => a + i.qty, 0);
  const lines = items.map((i) => `• ${i.qty} x ${i.p.name} (${fmt(i.p.price * i.qty)})`);
  const text =
    `Hola! Quiero hacer un pedido:\n${lines.join("\n")}\nTotal: ${fmt(total)}` +
    (name.trim() ? `\nA nombre de: ${name.trim()}` : "") +
    (note.trim() ? `\nAclaración: ${note.trim()}` : "");
  const href = `https://wa.me/${whatsappNumber}?text=${encodeURIComponent(text)}`;

  const inputStyle = {
    width: "100%", padding: "8px 10px", borderRadius: 8, border: "1px solid var(--line)", background: "var(--card-bg)",
    color: "var(--ink)", fontSize: 14, fontFamily: "inherit", boxSizing: "border-box", marginTop: 6,
  };

  return (
    <div style={{ position: "sticky", bottom: 0, zIndex: 30, marginTop: 16 }}>
      <div style={{ background: "var(--surface-dark)", color: "white", borderRadius: 14, padding: 12, boxShadow: "0 -4px 20px rgba(0,0,0,0.18)" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <div>
            <div style={{ fontSize: 12, opacity: 0.75 }}>Tu pedido: {units} unidad{units !== 1 ? "es" : ""}</div>
            <div className="mono" style={{ fontWeight: 800, fontSize: 18 }}>{fmt(total)}</div>
          </div>
          <button
            type="button"
            onClick={() => setOpen((v) => !v)}
            style={{ background: "var(--mustard)", color: "var(--surface-dark)", border: "none", borderRadius: 10, padding: "8px 14px", fontWeight: 800, fontSize: 13, cursor: "pointer" }}
          >
            {open ? "Ocultar" : "Ver pedido"}
          </button>
        </div>

        {open && (
          <div style={{ marginTop: 10, borderTop: "1px dashed rgba(255,255,255,0.25)", paddingTop: 10 }}>
            {items.map((i) => (
              <div key={i.p.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "4px 0" }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{i.p.name}</div>
                  <div className="mono" style={{ fontSize: 12, opacity: 0.75 }}>{fmt(i.p.price * i.qty)}</div>
                </div>
                <div style={{ width: 96, color: "var(--ink)" }}>
                  <QtyControl value={i.qty} max={i.p.stock} onChange={(q) => onQty(i.p.id, q)} compact />
                </div>
              </div>
            ))}
            <input style={inputStyle} placeholder="Tu nombre" value={name} onChange={(e) => setName(e.target.value)} />
            <input style={inputStyle} placeholder="Aclaración (ej: curso, hora del recreo)" value={note} onChange={(e) => setNote(e.target.value)} />
            <a
              href={href}
              target="_blank"
              rel="noopener noreferrer"
              style={{ display: "block", textAlign: "center", background: "#25D366", color: "#0b2e13", fontWeight: 800, borderRadius: 10, padding: 10, textDecoration: "none", fontSize: 14, marginTop: 10 }}
            >
              Enviar pedido por WhatsApp
            </a>
            <button
              type="button"
              onClick={onClear}
              style={{ display: "block", margin: "8px auto 0", background: "none", border: "none", color: "white", opacity: 0.7, fontSize: 12, textDecoration: "underline", cursor: "pointer" }}
            >
              Vaciar pedido
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------------- Catálogo público (solo lectura, para clientes) ----------------
function CatalogoPublico({ products, whatsappNumber, onPresent, cart, onQty, onClear }) {
  const byCategory = useMemo(() => {
    const map = {};
    products.forEach((p) => {
      const cat = p.category || "General";
      if (!map[cat]) map[cat] = [];
      map[cat].push(p);
    });
    return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
  }, [products]);

  if (products.length === 0) {
    return (
      <div className="card p-6 text-center text-sm" style={{ color: "var(--ink-soft)" }}>
        Todavía no hay productos cargados en el catálogo.
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex justify-end">
        <button className="btn-primary" onClick={onPresent}>
          <Maximize2 size={15} /> Pantalla completa
        </button>
      </div>
      {byCategory.map(([cat, items]) => (
        <div key={cat}>
          <h3 className="text-xs font-bold uppercase mb-2" style={{ letterSpacing: "0.05em", color: "var(--ink-soft)" }}>{cat}</h3>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {items.map((p) => (
              <div key={p.id} className="card p-2" style={{ opacity: p.stock > 0 ? 1 : 0.55 }}>
                <div style={{ aspectRatio: "1", width: "100%", borderRadius: 8, overflow: "hidden", background: "var(--paper-dark)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {p.imageUrl ? (
                    <img src={p.imageUrl} alt={p.name} style={{ width: "100%", height: "100%", objectFit: "cover", objectPosition: "center" }} onError={(e) => { e.target.style.display = "none"; }} />
                  ) : (
                    <ImageOff size={20} style={{ color: "var(--ink-soft)" }} />
                  )}
                </div>
                <div className="mt-2 text-sm font-semibold truncate">{p.name}</div>
                <div className="mono font-bold text-sm" style={{ color: "var(--teal)" }}>{fmt(p.price)}</div>
                {p.stock <= 0 && <div style={{ fontSize: 10.5, fontWeight: 700, color: "var(--mustard)" }}>Sin stock</div>}
                {whatsappNumber && p.stock > 0 && (
                  <QtyControl value={cart[p.id] || 0} max={p.stock} onChange={(q) => onQty(p.id, q)} />
                )}
              </div>
            ))}
          </div>
        </div>
      ))}
      <CartBar products={products} cart={cart} onQty={onQty} onClear={onClear} whatsappNumber={whatsappNumber} />
    </div>
  );
}

// ---------------- Catálogo en pantalla completa (para mostrar al cliente) ----------------
function CatalogoFullscreen({ products, onClose, theme, whatsappNumber, cart, onQty, onClear }) {
  const t = theme || THEMES.clasico;
  const byCategory = useMemo(() => {
    const map = {};
    products.forEach((p) => {
      const cat = p.category || "General";
      if (!map[cat]) map[cat] = [];
      map[cat].push(p);
    });
    return map;
  }, [products]);

  return (
    <div
      className="kiosco-app"
      style={{ position: "fixed", inset: 0, zIndex: 50, overflowY: "auto", padding: "24px 20px 40px" }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Manrope:wght@400;500;600;700;800&display=swap');
        .kiosco-app { --paper:${t.paper}; --paper-dark:${t.paperDark}; --ink:${t.ink}; --ink-soft:${t.inkSoft}; --teal:${t.teal}; --mustard:${t.mustard}; --line:${t.line}; --surface-dark:${t.surfaceDark}; --card-bg:${t.cardBg}; font-family:'Manrope',sans-serif; background:var(--paper); color:var(--ink); min-height:100%; }
        .kiosco-app .mono { font-family:'JetBrains Mono', monospace; }
        .kiosco-app .wa-btn { display:block; text-align:center; margin-top:6px; font-size:10.5px; font-weight:700; padding:4px 6px; border-radius:6px; border:1px solid var(--line); color:var(--ink-soft); text-decoration:none; }
        .kiosco-app .wa-btn:hover { border-color: var(--teal); color: var(--teal); }
      `}</style>
      <div className="max-w-5xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-2" style={{ color: "var(--teal)" }}>
            <Store size={22} />
            <h1 className="text-2xl font-extrabold" style={{ letterSpacing: "-0.01em" }}>Nuestros productos</h1>
          </div>
          <button
            onClick={onClose}
            style={{ background: "var(--surface-dark)", color: "white", border: "none", borderRadius: 10, padding: "8px 14px", display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}
          >
            <X size={15} /> Cerrar
          </button>
        </div>

        {Object.entries(byCategory).map(([cat, items]) => (
          <div key={cat} className="mb-8">
            <h2 className="text-sm font-bold uppercase mb-3" style={{ letterSpacing: "0.05em", color: "var(--ink-soft)" }}>{cat}</h2>
            <div className="grid grid-cols-4 md:grid-cols-5 gap-3">
              {items.map((p) => (
                <div key={p.id} style={{ background: "var(--card-bg)", border: "1px solid var(--line)", borderRadius: 12, overflow: "hidden", opacity: p.stock > 0 ? 1 : 0.5 }}>
                  <div style={{ aspectRatio: "1", background: "var(--paper-dark)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {p.imageUrl ? (
                      <img src={p.imageUrl} alt={p.name} style={{ width: "100%", height: "100%", objectFit: "cover", objectPosition: "center" }} onError={(e) => { e.target.style.display = "none"; }} />
                    ) : (
                      <ImageOff size={20} style={{ color: "var(--ink-soft)" }} />
                    )}
                  </div>
                  <div style={{ padding: "8px 9px" }}>
                    <div className="font-semibold truncate" style={{ fontSize: 12.5 }}>{p.name}</div>
                    <div className="mono font-extrabold" style={{ fontSize: 14.5, color: "var(--teal)" }}>{fmt(p.price)}</div>
                    {p.stock <= 0 && <div style={{ fontSize: 10.5, fontWeight: 700, color: "var(--mustard)" }}>Sin stock</div>}
                    {whatsappNumber && p.stock > 0 && (
                      <QtyControl value={cart[p.id] || 0} max={p.stock} onChange={(q) => onQty(p.id, q)} compact />
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
        <CartBar products={products} cart={cart} onQty={onQty} onClear={onClear} whatsappNumber={whatsappNumber} />
      </div>
    </div>
  );
}
