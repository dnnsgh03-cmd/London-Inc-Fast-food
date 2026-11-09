const DEFAULT_MENU = [
    ["Salchipapas", [["Salchipapa Casera Inglesa", 16000], ["Salchipapa de Chorizo Inglesa", 21000], ["Salchipapa Big Ben (para 4 personas)", 50000]]],
    ["Hamburguesas", [["Hamburguesa de Carne S", 15000], ["Hamburguesa de Carne M", 25000], ["Hamburguesa de Carne L", 38000], ["Hamburguesa de Carne XL", 50000], ["Hamburguesa de Pollo M", 26000], ["Hamburguesa Vegetariana M", 22000]]],
    ["Perros Calientes", [["Perro Sencillo", 8000], ["Perro Ranchero", 12000], ["Perro Suizo", 14000]]],
    ["Desgranados", [["Chuzos Desgranados de Pollo", 21000], ["Mazorcas Desgranadas", 23000]]],
    ["Especialidades de la Casa", [["Brochets Breaded with Chips (Pincho inglés con papas inglesas)", 20000], ["English Fish & Chips (Pescado apanado inglés con papas inglesas)", 30000]]],
    ["Menú Infantil (Little Angels Menu)", [["Mini Pinchos con porción de papas inglesa", 8000], ["Mini Salchipapas", 10000], ["Mini Pescado Inglés con papas inglesa", 4000], ["Hamburguesa de Carne XS", 10000], ["Hamburguesa de Pollo XS", 11000]]],
    ["Adicionales", [["Pinchos de Salchicha y Queso", 8000], ["Papas Inglesa", 8000], ["Papas a la Francesa", 6000], ["Queso Costeño", 2000], ["Queso Mozzarella", 3000]]],
    ["Bebidas", [["Agua Brisa Sabor Manzana", 3000], ["Coca Cola 400 ml Personal", 4000], ["Sprite Personal 400 ml", 4000], ["Coca Cola Zero Personal", 4000], ["Quatro Personal 400 ml", 4000], ["Jugos Hit Personal", 4000], ["Jugos de Sábila Personal", 4000], ["Coca Cola 1.5 L", 4000]]],
];

const LEGACY_MENU = [
    ["Salchipapas", [["Salchipapa casera inglesa", 16000], ["Salchipapa de chorizo casera inglesa", 21000], ["Salchipapa Big Ben (4 personas)", 45000]]],
    ["Extras", [["Pinchos de salchicha y queso", 8000], ["Papas inglesas", 8000], ["Papas a la francesa", 6000]]],
    ["Adicional", [["Queso costeño", 2000], ["Queso mozzarella", 3000]]],
    ["Bebidas", [["Agua Brisa manzana", 3000], ["Coca-Cola 400 ml", 4000], ["Sprite personal", 4000], ["Coca-Cola Zero personal", 4000], ["Quatro personal", 4000], ["Jugos Hit personal", 4000], ["Jugo de sábila personal", 4000], ["Coca-Cola 1,7 L", 8000]]],
];

const DEFAULT_STAFF = [
    ["Nathan Heilbron Ruff", "Gerente"],
    ["Luis/Lucho", "Administrador"],
    ["Hanny", "Administradora"],
    ["Juand David", "Domicilios y atención"],
    ["Daniela", "Administración y caja"],
    ["Tania", "Cocina"],
    ["Asistente de cocina (por decidir)", "Asistente de cocina"],
    ["Mesero asistente (por decidir)", "Mesero"],
    ["Fercho", "Domiciliario"],
    ["Domiciliario (por decidir)", "Domicilios"],
];
const LEGACY_STAFF = [["Nathan", "Administrador"], ["Hanny", "Cocina"], ["Jesus", "Domicilios"]];
const PREVIOUS_STAFF = [
    ["Nathan Heilbron Ruff", "Representante legal, dueño y gerente; contaduría y pedidos de materia prima/productos"],
    ["Luis/Lucho", "Administrador de turno y jefe de pedidos en sitio"],
    ["Hanny", "Administradora de turno algunos días; asistente preparadora y contadora algunos días; administradora de London Inc Delicatessen"],
    ["Juand David", "Jefe de despacho de domicilios y mesero"],
    ["Daniela", "Administradora y cajera de pedidos en sitio y domicilios; contaduría y pedidos de materia prima/productos"],
    ["Tania", "Jefa de cocina"],
    ["Asistente de cocina (por decidir)", "Asistente de cocina"],
    ["Mesero asistente (por decidir)", "Mesero asistente"],
    ["Fercho", "Domiciliario"],
    ["Domiciliario (por decidir)", "Domiciliario"],
];

let MENU = DEFAULT_MENU;
let P = [];
let apiConnected = false;
let excelWritable = true;
let excelFilename = "londoninc.xlsx";
let saveQueue = Promise.resolve(true);
let connectionNoticeShown = false;
let pendingSaves = 0;
let filterOrder = "all"; // 'all', 'today', 'pending'

// Variables para sincronización en tiempo real y Discord
let currentServerVersion = 0;
let isBackgroundSyncing = false;
let isSoundEnabled = localStorage.getItem("licSound") !== "false";
let discordStatus = { configured: false, masked_url: "", source: "none" };

const TABS = [
    ["menu", "Menú QR"],
    ["ped", "Pedidos"],
    ["per", "Personal"],
    ["caja", "Caja"],
    ["inv", "Inventario"],
    ["xls", "Excel"],
];

const todayDate = () => new Date().toLocaleDateString("es-CO", { day: "2-digit", month: "2-digit", year: "numeric" });
const todayIso = () => {
    const date = new Date();
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
};
const now = () => new Date().toLocaleTimeString("es-CO", { hour: "2-digit", minute: "2-digit" });

const def = () => ({
    tab: "ped",
    cart: {},
    editingOrderId: null,
    editingUnavailableItems: [],
    f: { n: "", address: "", t: "Mesa", p: "Efectivo", comments: "" },
    orders: [],
    pays: [],
    moves: [],
    open: 100000,
    url: "https://londoninc.com/menu",
    menu: DEFAULT_MENU,
    staff: DEFAULT_STAFF,
    balanceItems: [],
    inv: [["Papa (kg)", 20, 8], ["Salchicha (und)", 80, 30], ["Chorizo (und)", 60, 20], ["Queso costeño (kg)", 3, 2], ["Queso mozzarella (kg)", 2, 2], ["Coca-Cola 400 ml (und)", 24, 10], ["Agua Brisa (und)", 18, 10], ["Aceite (L)", 10, 4]],
});

let S = def();
try {
    const rawLocal = localStorage.getItem("lic");
    if (rawLocal) {
        S = Object.assign(def(), JSON.parse(rawLocal));
    }
} catch (e) {
    console.error("No se pudieron leer los datos guardados localmente.", e);
}
migrateLegacyMenu();
migrateLegacyStaff();

function normalizeStateDates() {
    const today = todayDate();
    if (Array.isArray(S.orders)) {
        S.orders.forEach(o => {
            if (!o.fec) o.fec = (o.id === 1 || o.id === 2) ? "20/10/2026" : today;
        });
    }
    if (Array.isArray(S.pays)) {
        S.pays.forEach(p => {
            if (!p.fec) p.fec = today;
        });
    }
    if (Array.isArray(S.moves)) {
        S.moves.forEach(m => {
            if (!m.fec) m.fec = today;
        });
    }
}
normalizeStateDates();

function updateMenu() {
    MENU = Array.isArray(S.menu) ? S.menu : DEFAULT_MENU;
    P = [];
    MENU.forEach(([c, a]) => {
        if (Array.isArray(a)) {
            a.forEach(([n, p]) => P.push({ id: P.length, c, n, p: Number(p) || 0 }));
        }
    });
}

function migrateLegacyMenu(state = S) {
    if (JSON.stringify(state.menu) === JSON.stringify(LEGACY_MENU)) {
        state.menu = DEFAULT_MENU;
        return true;
    }
    return false;
}

function migrateLegacyStaff(state = S) {
    const staff = JSON.stringify(state.staff);
    if (staff === JSON.stringify(LEGACY_STAFF) || staff === JSON.stringify(PREVIOUS_STAFF)) {
        state.staff = DEFAULT_STAFF;
        return true;
    }
    return false;
}

async function apiRequest(url, options = {}, timeoutMs = 0) {
    const controller = timeoutMs ? new AbortController() : null;
    const timeoutId = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
    try {
        const response = await fetch(url, {
            ...options,
            headers: { "Content-Type": "application/json", ...(options.headers || {}) },
            ...(controller ? { signal: controller.signal } : {}),
        });
        const contentType = response.headers.get("content-type") || "";
        const isJson = contentType.includes("application/json");
        let body = null;
        if (isJson) {
            try {
                body = await response.json();
            } catch (error) {
                throw new Error("El servidor devolvió una respuesta JSON inválida.");
            }
        }
        if (!response.ok) {
            let message = body?.error || `Error del servidor (${response.status})`;
            if (response.status === 404 && url === "/api/discord/order/update") {
                message = "El servidor todavía no tiene la ruta para avisar ediciones a Discord. Detén y vuelve a iniciar app.py.";
            } else if (response.redirected || contentType.includes("text/html")) {
                message = "El servidor devolvió una página HTML en lugar de la respuesta de la API. Reinicia app.py y vuelve a iniciar sesión.";
            }
            const error = new Error(message);
            error.status = response.status;
            throw error;
        }
        if (!isJson) {
            throw new Error("El servidor devolvió una respuesta inesperada para la API.");
        }
        return body;
    } catch (error) {
        if (controller && error.name === "AbortError") {
            throw new Error("El servidor tardó demasiado en cargar Excel. Verifica que app.py siga activo y que el libro esté disponible.");
        }
        throw error;
    } finally {
        if (timeoutId) clearTimeout(timeoutId);
    }
}

function updateSaveStatus(message, state) {
    const element = document.getElementById("sync-status");
    if (!element) return;
    element.textContent = message;
    element.className = `sync-status ${state}`;
}

function save() {
    normalizeStateDates();
    const serialized = JSON.stringify(S);
    try {
        localStorage.setItem("lic", serialized);
        localStorage.setItem("licPending", serialized);
    } catch (error) {
        console.error("No se pudo guardar la copia local.", error);
    }
    pendingSaves++;
    updateSaveStatus("Guardando cambios en Excel...", "saving");
    saveQueue = saveQueue
        .then(async () => {
            const saveRes = await apiRequest("/api/state", { method: "POST", body: serialized });
            apiConnected = true;
            excelWritable = true;
            connectionNoticeShown = false;
            if (saveRes && typeof saveRes.version === "number") {
                currentServerVersion = saveRes.version;
            }
            try {
                if (localStorage.getItem("licPending") === serialized) {
                    localStorage.removeItem("licPending");
                }
            } catch (error) {
                console.error("No se pudo limpiar el indicador de guardado pendiente.", error);
            }
            pendingSaves--;
            if (!pendingSaves) {
                updateSaveStatus(`Guardado en ${excelFilename} · ${now()}`, "saved");
            }
            return true;
        })
        .catch(error => {
            pendingSaves--;
            if (error.status === 423) {
                apiConnected = true;
                excelWritable = false;
                updateSaveStatus(`Datos protegidos en el navegador. Cierra ${excelFilename} para volcarlos al archivo`, "saving");
                if (!connectionNoticeShown) {
                    toast(`Excel está abierto. Tus datos están seguros aquí; se guardarán en el archivo al cerrar Excel.`);
                    connectionNoticeShown = true;
                }
            } else {
                apiConnected = false;
                excelWritable = false;
                updateSaveStatus(`Error de conexión con Python: ${error.message}`, "error");
                if (!connectionNoticeShown) {
                    toast(`No se pudo conectar con el servidor: ${error.message}`);
                    connectionNoticeShown = true;
                }
            }
            return false;
        });
    return saveQueue;
}

async function manualSaveAll() {
    const btn = document.getElementById("btn-quick-sync");
    if (btn) btn.textContent = "Guardando todo...";
    connectionNoticeShown = false;
    updateSaveStatus("Forzando guardado completo en Excel...", "saving");
    const ok = await save();
    if (ok) {
        toast(`Todos los datos han quedado guardados con éxito en ${excelFilename}.`);
        updateSaveStatus(`Guardado completo en ${excelFilename} · ${now()}`, "saved");
        if (btn) {
            btn.textContent = "Guardado con éxito";
            setTimeout(() => { btn.textContent = "Guardar todo en Excel ahora"; }, 2500);
        }
    } else {
        if (!excelWritable) {
            toast(`${excelFilename} está abierto en Excel. Tus datos están seguros aquí. Cierra Excel y vuelve a pulsar el botón.`);
            if (btn) {
                btn.textContent = "Excel abierto (datos seguros)";
                setTimeout(() => { btn.textContent = "Guardar todo en Excel ahora"; }, 3500);
            }
        } else {
            toast("No hay conexión con el servidor Python (http://127.0.0.1:5000).");
            if (btn) btn.textContent = "Guardar todo en Excel ahora";
        }
    }
    render();
}

async function loadState() {
    try {
        const result = await apiRequest("/api/state", {}, 15000);
        const local = JSON.parse(localStorage.getItem("lic") || "null");
        apiConnected = true;
        excelWritable = result.writable !== false;
        excelFilename = result.workbook || excelFilename;
        if (result && typeof result.version === "number") {
            currentServerVersion = result.version;
        }
        const pending = JSON.parse(localStorage.getItem("licPending") || "null");

        if (pending) {
            S = Object.assign(def(), pending);
            normalizeStateDates();
            migrateLegacyMenu();
            migrateLegacyStaff();
            if (excelWritable) {
                await save();
            } else {
                updateSaveStatus(`Hay cambios pendientes guardados localmente; cierra ${excelFilename} para sincronizar`, "error");
            }
        } else if (!result.initialized && local && (local.orders?.length || local.pays?.length || local.moves?.length)) {
            S = Object.assign(def(), local);
            normalizeStateDates();
            migrateLegacyMenu();
            migrateLegacyStaff();
            if (excelWritable) await save();
        } else {
            S = Object.assign(def(), result.state);
            normalizeStateDates();
            const menuMigrated = migrateLegacyMenu();
            const staffMigrated = migrateLegacyStaff();
            if ((menuMigrated || staffMigrated) && excelWritable) await save();
        }
    } catch (error) {
        apiConnected = false;
        updateSaveStatus(`Excel no disponible: ${error.message}`, "error");
        console.error("No se pudo cargar el estado desde Python/Excel.", error);
        toast(`Servidor no disponible: ${error.message}`);
    }
    if (apiConnected && !pendingSaves && !localStorage.getItem("licPending")) {
        updateSaveStatus(
            excelWritable ? `Conectado a ${excelFilename} (todo al día)` : `${excelFilename} abierto en Excel (datos seguros)`,
            excelWritable ? "saved" : "saving"
        );
    }
    updateMenu();
    render();
    const soundBtn = document.getElementById("btn-sound-toggle");
    if (soundBtn) {
        soundBtn.textContent = isSoundEnabled ? "🔔 Sonido activado" : "🔕 Sonido silenciado";
    }
    loadDiscordStatus();
    startRealtimeSync();
}

function playOrderAlertSound() {
    if (!isSoundEnabled) return;
    try {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        if (!AudioCtx) return;
        const ctx = new AudioCtx();
        if (ctx.state === "suspended") {
            ctx.resume();
        }
        const nowTime = ctx.currentTime;
        
        // Tone 1: 587.33 Hz (D5)
        const osc1 = ctx.createOscillator();
        const gain1 = ctx.createGain();
        osc1.type = "sine";
        osc1.frequency.setValueAtTime(587.33, nowTime);
        gain1.gain.setValueAtTime(0, nowTime);
        gain1.gain.linearRampToValueAtTime(0.25, nowTime + 0.04);
        gain1.gain.exponentialRampToValueAtTime(0.001, nowTime + 0.35);
        osc1.connect(gain1);
        gain1.connect(ctx.destination);
        osc1.start(nowTime);
        osc1.stop(nowTime + 0.35);

        // Tone 2: 880 Hz (A5)
        const osc2 = ctx.createOscillator();
        const gain2 = ctx.createGain();
        osc2.type = "sine";
        osc2.frequency.setValueAtTime(880, nowTime + 0.15);
        gain2.gain.setValueAtTime(0, nowTime + 0.15);
        gain2.gain.linearRampToValueAtTime(0.3, nowTime + 0.19);
        gain2.gain.exponentialRampToValueAtTime(0.001, nowTime + 0.65);
        osc2.connect(gain2);
        gain2.connect(ctx.destination);
        osc2.start(nowTime + 0.15);
        osc2.stop(nowTime + 0.65);
    } catch (e) {
        console.warn("Audio chime no disponible:", e);
    }
}

function toggleSound() {
    isSoundEnabled = !isSoundEnabled;
    try {
        localStorage.setItem("licSound", isSoundEnabled ? "true" : "false");
    } catch (e) {}
    const btn = document.getElementById("btn-sound-toggle");
    if (btn) {
        btn.textContent = isSoundEnabled ? "🔔 Sonido activado" : "🔕 Sonido silenciado";
        btn.title = isSoundEnabled ? "Sonido activo al recibir nuevos pedidos" : "Sonido desactivado";
    }
    if (isSoundEnabled) {
        playOrderAlertSound();
        toast("Sonido de pedidos activado.");
    } else {
        toast("Sonido de pedidos desactivado.");
    }
}

function updateLiveIndicator(isLive) {
    const el = document.getElementById("live-indicator");
    if (!el) return;
    if (isLive) {
        el.className = "live-pill";
        el.innerHTML = '<span class="live-dot"></span> EN VIVO';
        el.title = "Sincronización en tiempo real activa entre dispositivos";
    } else {
        el.className = "live-pill offline";
        el.innerHTML = '<span class="live-dot"></span> DESCONECTADO';
        el.title = "Sin conexión en tiempo real con el servidor";
    }
}

async function loadDiscordStatus() {
    try {
        const res = await apiRequest("/api/discord/status");
        if (res && res.ok) {
            discordStatus = {
                configured: !!res.configured,
                masked_url: res.masked_url || "",
                source: res.source || "none"
            };
            if (S.tab === "xls") render();
        }
    } catch (e) {
        console.warn("No se pudo cargar estado de Discord:", e);
    }
}

async function testDiscordConnection() {
    const btn = document.getElementById("btn-test-discord");
    if (btn) {
        btn.disabled = true;
        btn.textContent = "Probando conexión...";
    }
    try {
        const res = await apiRequest("/api/discord/test", { method: "POST" });
        if (res && res.ok) {
            toast("✅ Mensaje de prueba enviado con éxito a Discord!");
            discordStatus.configured = true;
            render();
        } else {
            toast("❌ " + (res.error || "No se pudo conectar con Discord"));
        }
    } catch (err) {
        toast("❌ Error probando Discord: " + err.message);
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.textContent = "Enviar mensaje de prueba";
        }
    }
}

async function saveDiscordWebhook() {
    const input = document.getElementById("discord-webhook-input");
    if (!input) return;
    const url = input.value.trim();
    if (!url) {
        return toast("Ingresa la URL del Webhook de Discord.");
    }
    if (!url.startsWith("https://discord.com/api/webhooks/") && !url.startsWith("https://discordapp.com/api/webhooks/")) {
        return toast("La URL debe ser un webhook válido de Discord (https://discord.com/api/webhooks/...)");
    }
    const btn = document.getElementById("btn-save-discord");
    if (btn) {
        btn.disabled = true;
        btn.textContent = "Guardando...";
    }
    try {
        const res = await apiRequest("/api/discord/config", {
            method: "POST",
            body: JSON.stringify({ webhook_url: url })
        });
        if (res && res.ok) {
            discordStatus.configured = true;
            discordStatus.masked_url = res.masked_url;
            discordStatus.source = "excel";
            toast("✅ Webhook de Discord guardado correctamente.");
            render();
        } else {
            toast("Error guardando webhook: " + (res.error || "Desconocido"));
        }
    } catch (err) {
        toast("Error guardando webhook: " + err.message);
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.textContent = "Guardar Webhook";
        }
    }
}

let realtimeSyncInterval = null;
function startRealtimeSync() {
    if (realtimeSyncInterval) return;
    updateLiveIndicator(true);

    const checkVersion = async () => {
        try {
            const data = await apiRequest("/api/state/version", {}, 4000);
            updateLiveIndicator(true);
            if (data && typeof data.version === "number") {
                if (currentServerVersion === 0) {
                    currentServerVersion = data.version;
                } else if (data.version > currentServerVersion) {
                    syncRemoteState(data.version);
                }
            }
        } catch (err) {
            updateLiveIndicator(false);
        }
    };

    realtimeSyncInterval = setInterval(checkVersion, 2500);

    window.addEventListener("focus", () => {
        checkVersion();
    });
}

function syncRemoteState(newVersion) {
    if (isBackgroundSyncing) return;
    isBackgroundSyncing = true;
    apiRequest("/api/state", {}, 6000)
        .then(result => {
            if (!result || !result.state) return;
            const remoteOrders = result.state.orders || [];
            const previousOrders = S.orders || [];
            
            // Check for new orders
            const prevIds = new Set(previousOrders.map(o => o.id));
            const newOrders = remoteOrders.filter(o => !prevIds.has(o.id));

            // Update core data
            S.orders = remoteOrders;
            S.pays = result.state.pays || [];
            S.moves = result.state.moves || [];
            S.inv = result.state.inv || S.inv;
            S.menu = result.state.menu || S.menu;
            S.staff = result.state.staff || S.staff;
            S.balanceItems = result.state.balanceItems || [];
            currentServerVersion = result.version || newVersion;

            try {
                localStorage.setItem("lic", JSON.stringify(S));
            } catch (e) {}

            updateMenu();
            updateLiveIndicator(true);

            if (newOrders.length > 0) {
                playOrderAlertSound();
                const latest = newOrders[0];
                toast(`🔔 ¡Nuevo pedido #${latest.id} (${latest.n || 'Cliente'})! Total: ${F(latest.total)}`);
            }

            const activeTag = document.activeElement ? document.activeElement.tagName : "";
            const isTyping = activeTag === "INPUT" || activeTag === "TEXTAREA";
            if (!isTyping) {
                render();
            }
        })
        .catch(err => {
            console.warn("Fallo leve en sincronización remota:", err);
            updateLiveIndicator(false);
        })
        .finally(() => {
            isBackgroundSyncing = false;
        });
}

const F = n => "$" + Number(n).toLocaleString("es-CO");
const sum = a => a.reduce((s, x) => s + x, 0);

function toast(t) {
    const e = document.getElementById("toast");
    if (!e) return;
    e.textContent = t;
    e.className = "on";
    setTimeout(() => { e.className = ""; }, 2800);
}

function go(t) {
    S.tab = t;
    render();
}

const opt = (a, v) => a.map(x => `<option${x == v ? " selected" : ""}>${x}</option>`).join("");
const escapeHtml = value => String(value).replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
}[character]));

function add(i, d) {
    S.cart[i] = (S.cart[i] || 0) + d;
    if (S.cart[i] <= 0) delete S.cart[i];
    save();
    render();
}

function editOrder(id) {
    const order = S.orders.find(item => item.id === id);
    if (!order) return toast(`No se encontró el pedido #${id}.`);
    if (Object.keys(S.cart).length && !confirm("Al editar este pedido se reemplazará el borrador actual. ¿Deseas continuar?")) return;

    const cart = {};
    const unavailableItems = [];
    for (const item of order.items || []) {
        if (!Array.isArray(item) || item.length < 2) {
            return toast(`No se puede editar: el detalle del pedido #${id} tiene un producto inválido.`);
        }
        const productIndex = P.findIndex(product => product.n === item[0]);
        if (productIndex < 0) {
            unavailableItems.push([item[0], Number(item[1]) || 0, Number(item[2]) || 0]);
            continue;
        }
        cart[productIndex] = (cart[productIndex] || 0) + (Number(item[1]) || 0);
    }

    S.editingOrderId = order.id;
    S.cart = cart;
    S.editingUnavailableItems = unavailableItems;
    S.f = {
        n: order.n === "Cliente general" ? "" : order.n || "",
        address: order.address || "",
        t: order.t || "Mesa",
        p: order.p || "Efectivo",
        comments: order.comments || "",
    };
    save();
    render();
    const form = document.getElementById("order-form");
    if (form) form.scrollIntoView({ behavior: "smooth", block: "start" });
}

function cancelOrderEdit() {
    if (!confirm("¿Cancelar la edición y descartar los cambios del pedido?")) return;
    S.editingOrderId = null;
    S.cart = {};
    S.editingUnavailableItems = [];
    S.f = { n: "", address: "", t: "Mesa", p: "Efectivo", comments: "" };
    save();
    render();
}

async function deleteOrder(id) {
    const orderIndex = S.orders.findIndex(order => order.id === id);
    if (orderIndex < 0) return toast(`No se encontró el pedido #${id}.`);
    if (!confirm(`¿Eliminar el pedido #${id}? Esta acción no se puede deshacer.`)) return;

    S.orders.splice(orderIndex, 1);
    if (S.editingOrderId === id) {
        S.editingOrderId = null;
        S.cart = {};
        S.editingUnavailableItems = [];
        S.f = { n: "", address: "", t: "Mesa", p: "Efectivo", comments: "" };
    }

    const persisted = save();
    render();
    if (await persisted) toast(`Pedido #${id} eliminado.`);
}

async function reg() {
    const k = Object.keys(S.cart);
    if (!k.length && !S.editingUnavailableItems.length) return toast("Agrega al menos un producto tocando el menú.");
    if (S.f.t === "Domicilio" && !(S.f.address || "").trim()) {
        return toast("Escribe la dirección para registrar el domicilio.");
    }
    const items = [
        ...k.map(i => [P[i].n, S.cart[i], P[i].p]),
        ...(S.editingOrderId === null ? [] : S.editingUnavailableItems),
    ];
    const editingOrder = S.editingOrderId === null
        ? null
        : S.orders.find(order => order.id === S.editingOrderId);
    if (S.editingOrderId !== null && !editingOrder) {
        S.editingOrderId = null;
        return toast("No se encontró el pedido que intentabas editar.");
    }
    const previousOrder = editingOrder
        ? { ...editingOrder, items: (editingOrder.items || []).map(item => [...item]) }
        : null;
    const nextId = S.orders.length ? Math.max(...S.orders.map(o => o.id || 0)) + 1 : 1;
    const updatedOrder = {
        ...(editingOrder || {}),
        id: editingOrder ? editingOrder.id : nextId,
        fec: editingOrder ? editingOrder.fec : todayDate(),
        h: editingOrder ? editingOrder.h : now(),
        n: S.f.n || "Cliente general",
        address: (S.f.address || (editingOrder ? editingOrder.address : "") || "").trim(),
        t: S.f.t,
        p: S.f.p,
        comments: (S.f.comments || "").trim(),
        items,
        total: sum(items.map(x => x[1] * x[2])),
        st: editingOrder ? editingOrder.st : "Pendiente",
    };
    if (editingOrder) {
        const orderIndex = S.orders.indexOf(editingOrder);
        S.orders[orderIndex] = updatedOrder;
    } else {
        S.orders.unshift(updatedOrder);
    }
    S.cart = {};
    S.editingOrderId = null;
    S.editingUnavailableItems = [];
    S.f.n = "";
    S.f.address = "";
    S.f.comments = "";
    const persisted = save();
    render();
    if (editingOrder) {
        if (await persisted) {
            try {
                await apiRequest("/api/discord/order/update", {
                    method: "POST",
                    body: JSON.stringify({ previous: previousOrder, order: updatedOrder }),
                });
                toast(`Pedido #${updatedOrder.id} actualizado y notificado en Discord.`);
            } catch (error) {
                toast(`Pedido #${updatedOrder.id} guardado, pero falló el aviso de edición: ${error.message}`);
            }
        }
        return;
    }
    if (await persisted) {
        try {
            await apiRequest("/api/discord/order", {
                method: "POST",
                body: JSON.stringify(updatedOrder),
            });
            toast(`Pedido #${updatedOrder.id} guardado y notificado en Discord.`);
        } catch (error) {
            toast(`Pedido #${updatedOrder.id} guardado, pero falló la notificación: ${error.message}`);
        }
    }
}

function nextSt(i) {
    const L = ["Pendiente", "Preparando", "Entregado"];
    const o = S.orders[i];
    if (!o) return;
    const currentIndex = L.indexOf(o.st);
    if (currentIndex < 2) {
        const prevSt = o.st;
        o.st = L[currentIndex + 1];
        save();
        render();
        toast(`Pedido #${o.id} actualizado a: ${o.st}`);
        apiRequest("/api/discord/order/status", {
            method: "POST",
            body: JSON.stringify({ order: o, previous_status: prevSt, new_status: o.st }),
        }).catch(err => console.log("Aviso de cambio de estado a Discord opcional:", err));
    }
}

async function payStaff() {
    const n = v("pn");
    const a = +v("pa");
    if (!a) return toast("Escribe el monto a pagar.");
    S.pays.unshift({ fec: todayDate(), h: now(), n, c: v("pc") || "Pago de turno", a });
    const persisted = save();
    render();
    if (await persisted) toast("Pago registrado y guardado en Excel.");
}

async function spend() {
    const a = +v("ea");
    if (!a) return toast("Escribe el monto del gasto.");
    S.moves.unshift({
        fec: todayDate(),
        h: now(),
        c: v("ec") || "Gasto operativo",
        a,
        type: v("expense-type") || "liability",
    });
    const persisted = save();
    render();
    if (await persisted) toast("Gasto registrado y guardado en Excel.");
}

function stock(i, d) {
    S.inv[i][1] = Math.max(0, S.inv[i][1] + d);
    save();
    render();
}

function addInv() {
    const n = v("in");
    if (!n) return toast("Escribe el nombre del ingrediente o producto.");
    S.inv.push([n, +v("iq") || 0, +v("im") || 0]);
    save();
    render();
    toast(`"${n}" agregado al inventario.`);
}

function setBalanceImpactFromCategory() {
    const category = v("balance-category");
    const impact = document.getElementById("balance-impact");
    if (impact) impact.value = ["Deuda", "Gasto"].includes(category) ? "false" : "true";
}

function addBalanceItem() {
    const concept = v("balance-concept");
    const category = v("balance-category");
    const value = Number(v("balance-value"));
    if (!concept) return toast("Escribe el concepto o ítem.");
    if (!Number.isFinite(value) || value <= 0) return toast("Escribe un valor mayor que cero.");
    S.balanceItems = Array.isArray(S.balanceItems) ? S.balanceItems : [];
    S.balanceItems.push({
        concept,
        category,
        value,
        impact: v("balance-impact") === "true",
    });
    save();
    render();
    toast(`"${concept}" agregado al balance.`);
}

function removeBalanceItem(index) {
    const item = S.balanceItems[index];
    if (!item || !confirm(`¿Eliminar "${item.concept}" del balance?`)) return;
    S.balanceItems.splice(index, 1);
    save();
    render();
    toast("Ítem eliminado del balance.");
}

function toggleBalanceImpact(index) {
    const item = S.balanceItems[index];
    if (!item) return;
    item.impact = item.impact !== true;
    save();
    render();
    toast(`Impacto actualizado: ${item.impact ? "Suma al negocio" : "No suma / Resta"}.`);
}

async function setOpen() {
    S.open = +v("op") || 0;
    const persisted = save();
    render();
    if (await persisted) toast("Base de caja guardada en Excel.");
}

function v(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : "";
}

function qr() {
    S.url = v("qu") || S.url;
    save();
    drawQR();
    toast("Enlace del QR actualizado");
}

function drawQR() {
    const e = document.getElementById("qr");
    if (!e) return;
    e.innerHTML = "";
    try {
        new QRCode(e, { text: S.url, width: 170, height: 170, colorDark: "#0b1f4d", colorLight: "#ffffff" });
    } catch (x) {
        e.textContent = "QR (se generará aquí)";
    }
}

async function testConnection() {
    try {
        const result = await apiRequest("/api/health");
        apiConnected = true;
        excelWritable = result.writable !== false;
        excelFilename = result.workbook || excelFilename;
        if (excelWritable) {
            updateSaveStatus(`Conectado a ${excelFilename} (todo listo)`, "saved");
            toast(`Conexión correcta: ${excelFilename} listo para guardar.`);
        } else {
            updateSaveStatus(`${excelFilename} está abierto en Excel (cierra el libro para guardar los cambios)`, "saving");
            toast(`Conectado, pero ${excelFilename} está abierto en Excel. Tus datos están seguros.`);
        }
        render();
    } catch (error) {
        apiConnected = false;
        excelWritable = false;
        updateSaveStatus(`Excel no disponible: ${error.message}`, "error");
        render();
        toast(`No se pudo conectar: ${error.message}`);
    }
}

async function importExcel() {
    if (localStorage.getItem("licPending") && !confirm("Hay cambios locales que todavía no se guardaron en Excel. Al cargar los datos de Excel se sobrescribirán. ¿Deseas continuar?")) return;
    try {
        await saveQueue;
        updateSaveStatus("Cargando datos desde Excel...", "saving");
        const result = await apiRequest("/api/import", { method: "POST" });
        S = Object.assign(def(), result.state);
        normalizeStateDates();
        updateMenu();
        apiConnected = true;
        excelWritable = result.writable !== false;
        saveQueue = Promise.resolve(true);
        try {
            localStorage.setItem("lic", JSON.stringify(S));
            localStorage.removeItem("licPending");
        } catch (error) {
            console.error("No se pudo actualizar la copia local.", error);
        }
        const msg = excelWritable
            ? `Datos cargados desde ${excelFilename} · ${now()}`
            : `Datos cargados desde ${excelFilename} (archivo abierto en Excel)`;
        updateSaveStatus(msg, "saved");
        render();
        toast(`Menú, precios e inventario cargados desde ${excelFilename}.`);
    } catch (error) {
        if (error.status === 423) {
            apiConnected = true;
            excelWritable = false;
        }
        updateSaveStatus(`Error al cargar de Excel: ${error.message}`, "error");
        toast(`No se pudo cargar el Excel: ${error.message}`);
    }
}

function downloadExcel() {
    window.location.href = "/api/excel/download";
}

function updateReportRange() {
    const allDates = v("report-range") === "all";
    const dateInput = document.getElementById("report-date");
    if (dateInput) dateInput.disabled = allDates;
}

async function downloadReport() {
    const range = v("report-range");
    const date = range === "all" ? "all" : v("report-date");
    if (!date) return toast("Selecciona la fecha del informe.");
    const button = document.querySelector(".report-controls button");
    if (button) {
        button.disabled = true;
        button.textContent = "Preparando informe...";
    }
    try {
        const response = await fetch(`/api/reports/download?date=${encodeURIComponent(date)}`);
        if (!response.ok) {
            let message = `Error del servidor (${response.status})`;
            try {
                const body = await response.json();
                message = body.error || message;
            } catch (error) {
                console.error("No se pudo leer el error del servidor al generar el informe.", error);
            }
            if (response.status === 404) {
                message = "El servidor abierto no reconoce la descarga del informe. Detén python app.py y vuelve a iniciarlo para cargar la versión actualizada.";
            }
            throw new Error(message);
        }

        const file = await response.blob();
        if (!file.size || !file.type.includes("spreadsheetml")) {
            throw new Error("El servidor no devolvió un archivo Excel válido.");
        }
        const url = URL.createObjectURL(file);
        const link = document.createElement("a");
        link.href = url;
        link.download = `informe-${date === "all" ? "historico" : date}.xlsx`;
        link.style.display = "none";
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 60000);
        toast("Informe Excel descargado.");
    } catch (error) {
        toast(`No se pudo descargar el informe: ${error.message}`);
    } finally {
        if (button) {
            button.disabled = false;
            button.textContent = "Descargar informe Excel";
        }
    }
}

async function resetAll() {
    if (!confirm("ATENCIÓN: ¿Estás seguro de que deseas borrar los datos actuales y restaurar los de prueba? Esta acción no se puede deshacer.")) return;
    if (!confirm("Confirma nuevamente: ¿Deseas restablecer todos los pedidos, pagos y gastos a los valores iniciales?")) return;
    try {
        const result = await apiRequest("/api/reset", { method: "POST" });
        S = Object.assign(def(), result.state);
        normalizeStateDates();
        updateMenu();
        apiConnected = true;
        excelWritable = true;
        saveQueue = Promise.resolve(true);
        try {
            localStorage.setItem("lic", JSON.stringify(S));
            localStorage.removeItem("licPending");
        } catch (error) {
            console.error("No se pudo actualizar la copia local.", error);
        }
        updateSaveStatus(`Restaurado a valores de prueba en ${excelFilename}`, "saved");
        render();
        toast("Datos de prueba restaurados");
    } catch (error) {
        if (error.status === 423) {
            apiConnected = true;
            excelWritable = false;
        }
        updateSaveStatus(`Error al restablecer Excel: ${error.message}`, "error");
        toast(`No se pudieron restaurar los datos: ${error.message}`);
    }
}

function setOrderFilter(f) {
    filterOrder = f;
    render();
}

window.addEventListener("focus", async () => {
    if (localStorage.getItem("licPending") && !pendingSaves) {
        try {
            const h = await apiRequest("/api/health");
            if (h.connected && h.writable) {
                await save();
                render();
            }
        } catch (e) {
            // Silently ignore background check
        }
    }
});

function stChip(s) {
    return `<span class="chip ${s == "Entregado" ? "ok" : s == "Preparando" ? "pre" : ""}">${s}</span>`;
}

const V = {
    menu() {
        return `<div class="two">
<div>
    <h2>Menú digital para tus clientes</h2>
    <p class="mut" style="margin-top:-6px;margin-bottom:12px">Los precios mostrados aquí son los que verán tus clientes al escanear el QR.</p>
    ${MENU.map(([c, a]) => `
    <div class="card">
        <b style="color:var(--red);font-size:16px">${c}</b>
        ${a.map(([n, p]) => `
        <div class="ml">
            <span>${n}</span>
            <b style="color:var(--navy)">${F(p)}</b>
        </div>`).join("")}
    </div>`).join("")}
</div>
<div>
    <h2>Código QR para mesas</h2>
    <div class="card">
        <div id="qr"></div>
        <p class="mut" style="margin:10px 0 4px">Enlace al que lleva el QR:</p>
        <input id="qu" value="${S.url}">
        <div class="row noprint" style="margin-top:8px">
            <button class="btn red" onclick="qr()">Generar QR</button>
            <button class="btn line" onclick="print()">Imprimir QR</button>
        </div>
    </div>
    <div class="card mut" style="font-size:13px">
        <b>Consejo:</b> Imprime este código QR y colócalo en las mesas o mostrador. Los clientes lo escanean con su cámara y ven este menú en su celular.
    </div>
</div>
</div>`;
    },

    ped() {
        const k = Object.keys(S.cart);
        const unavailableItems = S.editingUnavailableItems || [];
        const tot = sum(k.map(i => S.cart[i] * P[i].p))
            + sum(unavailableItems.map(item => item[1] * item[2]));
        const todayStr = todayDate();

        // Filter orders
        const allOrders = S.orders || [];
        const todayOrders = allOrders.filter(o => (o.fec || todayStr) === todayStr);
        const pendingOrders = allOrders.filter(o => o.st !== "Entregado");

        let visibleOrders = allOrders;
        if (filterOrder === "today") visibleOrders = todayOrders;
        if (filterOrder === "pending") visibleOrders = pendingOrders;

        const totalVentasAcumuladas = sum(allOrders.map(o => o.total || 0));

        return `<div class="two">
<div>
    <h2>1. Elige los productos</h2>
    <p class="mut" style="margin-top:-6px;margin-bottom:10px">Toca un producto para sumarlo al pedido. Si te equivocas, puedes ajustar la cantidad en el ticket de la derecha.</p>
    ${MENU.map(([c]) => `
    <div class="card" style="margin-bottom:12px">
        <b style="color:var(--red);font-size:15px">${c}</b>
        <div class="grid" style="margin-top:8px">
            ${P.filter(p => p.c == c).map(p => {
                const countInCart = S.cart[p.id] || 0;
                return `
                <button class="prod ${countInCart ? 'selected' : ''}" onclick="add(${p.id},1)">
                    <span style="font-weight:600">${p.n}</span>
                    <b>${F(p.p)}</b>
                    ${countInCart ? `<span class="cart-count">${countInCart} en ticket</span>` : ''}
                </button>`;
            }).join("")}
        </div>
    </div>`).join("")}
</div>

<div>
    <h2>2. ${S.editingOrderId === null ? "Datos del pedido" : `Editando pedido #${S.editingOrderId}`}</h2>
    <div class="card" id="order-form" style="position:sticky;top:60px">
        <label style="font-size:12px;font-weight:600;color:var(--mut);display:block;margin-bottom:4px">Cliente o Mesa</label>
        <input placeholder="Ejemplo: Mesa 4, Domicilio Carlos, etc." value="${escapeHtml(S.f.n || "")}" oninput="S.f.n=this.value;save()">
        
        <div class="row" style="margin-bottom:8px">
            <div style="flex:1">
                <label style="font-size:12px;font-weight:600;color:var(--mut);display:block;margin-bottom:4px">Tipo</label>
                <select onchange="S.f.t=this.value;save();render()">${opt(["Mesa", "Para llevar", "Domicilio"], S.f.t)}</select>
            </div>
            <div style="flex:1">
                <label style="font-size:12px;font-weight:600;color:var(--mut);display:block;margin-bottom:4px">Medio de pago</label>
                <select onchange="S.f.p=this.value;save()">${opt(["Efectivo", "Nequi", "Daviplata", "Transfiya", "Llave/Bre-B", "Ahorros Davivienda"], S.f.p)}</select>
            </div>
        </div>

        ${S.f.t === "Domicilio" ? `
        <label style="font-size:12px;font-weight:600;color:var(--mut);display:block;margin-bottom:4px">Dirección de entrega</label>
        <input placeholder="Ejemplo: Calle 72 # 45-20 Apto 402" value="${escapeHtml(S.f.address || "")}" oninput="S.f.address=this.value;save()">
        ` : ""}

        <label style="font-size:12px;font-weight:600;color:var(--mut);display:block;margin-bottom:4px">Comentarios o instrucciones</label>
        <textarea rows="3" placeholder="Ejemplo: sin salsa, sin chongo o sin verduras" oninput="S.f.comments=this.value;save()">${escapeHtml(S.f.comments || "")}</textarea>

        <div style="border-top:1px dashed var(--line);margin:10px 0;padding-top:10px">
            <b style="font-size:13px;color:var(--navy);display:block;margin-bottom:6px">Productos en el pedido:</b>
            ${k.length || unavailableItems.length ? `${k.map(i => `
            <div class="row" style="margin-bottom:6px;background:var(--bg);padding:6px 10px;border-radius:8px">
                <span style="flex:1;font-size:13px"><b>${S.cart[i]}×</b> ${P[i].n}</span>
                <span style="font-weight:600;margin-right:8px">${F(S.cart[i] * P[i].p)}</span>
                <span style="white-space:nowrap">
                    <button class="btn line sm" onclick="add(${i},-1)">−</button>
                    <button class="btn line sm" onclick="add(${i},1)">+</button>
                </span>
            </div>`).join("")}${unavailableItems.map(item => `
            <div class="row mut" style="margin-bottom:6px;background:var(--bg);padding:6px 10px;border-radius:8px">
                <span style="flex:1;font-size:13px"><b>${item[1]}×</b> ${escapeHtml(item[0])} <small>(producto anterior)</small></span>
                <span style="font-weight:600;margin-right:8px">${F(item[1] * item[2])}</span>
            </div>`).join("")}` : '<p class="mut" style="padding:10px 0;text-align:center">Toca los productos de la izquierda para agregarlos al ticket.</p>'}
        </div>

        <div class="row" style="margin:12px 0 14px">
            <span style="font-size:16px;font-weight:600">Total a cobrar:</span>
            <span class="tot">${F(tot)}</span>
        </div>
        <button class="btn red" style="width:100%;padding:14px;font-size:16px" onclick="reg()">${S.editingOrderId === null ? "Registrar y cobrar pedido" : "Guardar cambios del pedido"}</button>
        ${S.editingOrderId !== null ? `<button class="btn line" style="width:100%;margin-top:8px" onclick="deleteOrder(${Number(S.editingOrderId)})">Eliminar pedido</button>
        <button class="btn line" style="width:100%;margin-top:8px" onclick="cancelOrderEdit()">Cancelar edición</button>` : ""}
    </div>
</div>
</div>

<div style="margin-top:20px">
    <div class="row" style="flex-wrap:wrap;margin-bottom:8px">
        <div>
            <h2>Historial de pedidos</h2>
            <p class="mut" style="margin-top:-6px;margin-bottom:6px">Todos los pedidos quedan registrados de forma permanente con su <b>fecha y hora</b>. No se borra nada.</p>
        </div>
        <div class="row" style="gap:6px;margin-bottom:8px">
            <button class="btn-pill ${filterOrder === 'all' ? 'on' : ''}" onclick="setOrderFilter('all')">Todos (${allOrders.length})</button>
            <button class="btn-pill ${filterOrder === 'today' ? 'on' : ''}" onclick="setOrderFilter('today')">Solo hoy (${todayOrders.length})</button>
            <button class="btn-pill ${filterOrder === 'pending' ? 'on' : ''}" onclick="setOrderFilter('pending')">Pendientes (${pendingOrders.length})</button>
        </div>
    </div>

    <div class="card wrap">
        <table>
            <tr>
                <th>#</th>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Cliente</th>
                <th>Dirección</th>
                <th>Tipo</th>
                <th>Pago</th>
                <th>Detalle de productos</th>
                <th>Comentarios</th>
                <th>Total</th>
                <th>Estado</th>
                <th>Acción</th>
            </tr>
            ${visibleOrders.map((o, i) => {
                const globalIndex = S.orders.indexOf(o);
                return `
                <tr>
                    <td><b>#${o.id}</b></td>
                    <td style="white-space:nowrap;font-weight:600;color:var(--navy)">${o.fec || todayStr}</td>
                    <td style="white-space:nowrap;color:var(--mut)">${o.h}</td>
                    <td><b>${o.n}</b></td>
                    <td>${escapeHtml(o.address || "")}</td>
                    <td><span class="chip" style="background:#eef2ff;color:#2b59c3">${o.t}</span></td>
                    <td>${o.p}</td>
                    <td>${(o.items || []).map(x => x[1] + "× " + x[0]).join(", ")}</td>
                    <td style="white-space:pre-wrap">${escapeHtml(o.comments || "")}</td>
                    <td><b style="color:var(--red)">${F(o.total)}</b></td>
                    <td>${stChip(o.st)}</td>
                    <td>
                        <button class="btn line sm" onclick="editOrder(${Number(o.id)})">Editar</button>
                        ${o.st !== "Entregado" ? `
                        <button class="btn sm" onclick="nextSt(${globalIndex})">
                            ${o.st === "Pendiente" ? "Preparar" : "Entregar"}
                        </button>` : '<span style="color:#1f7a42;font-weight:600">Listo</span>'}
                    </td>
                </tr>`;
            }).join("") || '<tr><td colspan="12" class="mut" style="text-align:center;padding:16px">No hay pedidos registrados en este filtro.</td></tr>'}
        </table>
    </div>
</div>`;
    },

    per() {
        const todayStr = todayDate();
        return `<h2>Pagos y turnos del personal</h2>
<p class="mut" style="margin-top:-6px;margin-bottom:12px">Lleva el control de cuánto pagas por turnos, propinas o domicilios a cada empleado.</p>
<div class="two">
<div class="card wrap">
    <table>
        <tr>
            <th>Fecha</th>
            <th>Hora</th>
            <th>Empleado</th>
            <th>Concepto</th>
            <th>Monto</th>
        </tr>
        ${(S.pays || []).map(p => `
        <tr>
            <td style="white-space:nowrap;font-weight:600">${p.fec || todayStr}</td>
            <td style="color:var(--mut)">${p.h}</td>
            <td><b>${p.n}</b></td>
            <td>${p.c}</td>
            <td><b style="color:var(--red)">${F(p.a)}</b></td>
        </tr>`).join("") || '<tr><td colspan="5" class="mut">Sin pagos registrados.</td></tr>'}
    </table>
</div>

<div class="card">
    <b style="font-size:15px">Registrar un pago al empleado</b><br><br>
    <label style="font-size:12px;font-weight:600;color:var(--mut)">¿A quién le vas a pagar?</label>
    <select id="pn">${(S.staff || []).filter(s => !s[0].includes("(por decidir)")).map(s => `<option>${s[0]}</option>`).join("")}</select>
    <label style="font-size:12px;font-weight:600;color:var(--mut)">Concepto</label>
    <input id="pc" placeholder="Ejemplo: Turno de noche, Domicilio #3, Adelanto">
    <label style="font-size:12px;font-weight:600;color:var(--mut)">Monto a pagar</label>
    <input id="pa" type="number" placeholder="Ejemplo: 40000">
    <button class="btn red" style="width:100%;margin-top:6px" onclick="payStaff()">Guardar pago</button>
</div>
</div>

<div class="card">
    <b>Equipo de trabajo registrado</b>
    <div style="margin-top:8px">
        ${(S.staff || []).map(s => `
        <div class="ml">
            <span><b>${s[0]}</b></span>
            <span class="chip" style="background:#eef2ff;color:#2b59c3;white-space:normal">${s[1]}</span>
        </div>`).join("")}
    </div>
</div>`;
    },

    caja() {
        const todayStr = todayDate();
        const allOrders = S.orders || [];
        const todayOrders = allOrders.filter(o => (o.fec || todayStr) === todayStr);

        const efTotal = sum(allOrders.filter(o => o.p === "Efectivo").map(o => o.total));
        const efHoy = sum(todayOrders.filter(o => o.p === "Efectivo").map(o => o.total));

        const dgTotal = sum(allOrders.filter(o => o.p !== "Efectivo").map(o => o.total));
        const dgHoy = sum(todayOrders.filter(o => o.p !== "Efectivo").map(o => o.total));

        const sp = sum((S.pays || []).map(p => p.a));
        const gs = sum((S.moves || []).map(p => p.a));
        const efectivoEsperado = S.open + efTotal - sp - gs;

        return `<h2>Control de caja</h2>
<p class="mut" style="margin-top:-6px;margin-bottom:12px">Controla cuánto dinero en efectivo debe haber en caja y el total de ventas digitales.</p>

<div class="kpis">
    <div class="kpi"><small>Base inicial de caja</small><div>${F(S.open)}</div></div>
    <div class="kpi"><small>Efectivo hoy (${todayOrders.filter(o => o.p === "Efectivo").length} ped.)</small><div>${F(efHoy)}</div></div>
    <div class="kpi"><small>Digital hoy (Nequi/etc.)</small><div>${F(dgHoy)}</div></div>
    <div class="kpi"><small>Pagos y gastos totales</small><div>${F(sp + gs)}</div></div>
    <div class="kpi highlight"><small>Efectivo esperado en caja</small><div>${F(efectivoEsperado)}</div></div>
</div>

<div class="two">
<div class="card">
    <b>Gastos operativos registrados</b>
    <div class="wrap" style="margin-top:8px">
        <table>
            <tr>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Concepto</th>
                <th>Monto</th>
            </tr>
            ${(S.moves || []).map(p => {
                const expenseTypes = {
                    current_asset: { label: "Activo Corriente", className: "current" },
                    fixed_asset: { label: "Activo Fijo", className: "fixed" },
                    liability: { label: "Pasivo", className: "liability" },
                };
                const expenseType = expenseTypes[p.type] || expenseTypes.liability;
                return `
            <tr>
                <td style="white-space:nowrap;font-weight:600">${p.fec || todayStr}</td>
                <td style="color:var(--mut)">${p.h}</td>
                <td>${p.c}</td>
                <td><b style="color:var(--red)">-${F(p.a)}</b> <span class="expense-badge ${expenseType.className}">${expenseType.label}</span></td>
            </tr>`;
            }).join("") || '<tr><td colspan="4" class="mut">Sin gastos registrados.</td></tr>'}
        </table>
    </div>
</div>

<div>
    <div class="card">
        <b>Actualizar base de caja</b>
        <p class="mut" style="font-size:12px;margin:4px 0 8px">El dinero con el que inicias el día en caja física.</p>
        <input id="op" type="number" value="${S.open}">
        <button class="btn" style="width:100%" onclick="setOpen()">Actualizar base</button>
    </div>

    <div class="card">
        <b>Registrar un gasto rápido</b>
        <p class="mut" style="font-size:12px;margin:4px 0 8px">Compras imprevistas (hielo, gas, bolsas, etc.)</p>
        <input id="ec" placeholder="¿En qué se gastó? (ej. Hielo, Gas)">
        <input id="ea" type="number" placeholder="Monto gastado">
        <label for="expense-type" style="font-size:12px;font-weight:600;color:var(--mut)">Clasificación contable</label>
        <select id="expense-type">
            <option value="current_asset">🟢 Activo Corriente</option>
            <option value="fixed_asset">🔵 Activo Fijo</option>
            <option value="liability" selected>🔴 Pasivo</option>
        </select>
        <button class="btn red" style="width:100%" onclick="spend()">Guardar gasto</button>
    </div>
</div>
</div>`;
    },

    inv() {
        return `<h2>Inventario y Balance Simple</h2>

<h3>Control de existencias</h3>
<p class="mut" style="margin-top:-6px;margin-bottom:12px">Controla las cantidades disponibles de ingredientes y bebidas.</p>
<div class="card wrap">
    <table>
        <tr>
            <th>Producto / Ingrediente</th>
            <th>Stock actual</th>
            <th>Mínimo deseado</th>
            <th>Estado</th>
            <th>Sumar / Restar</th>
        </tr>
        ${(S.inv || []).map((x, i) => `
        <tr>
            <td><b>${x[0]}</b></td>
            <td><b style="font-size:16px">${x[1]}</b></td>
            <td style="color:var(--mut)">${x[2]}</td>
            <td>${x[1] <= x[2] ? '<span class="chip bad">Por reponer</span>' : '<span class="chip ok">Suficiente</span>'}</td>
            <td>
                <button class="btn line sm" onclick="stock(${i},-1)" title="Restar 1">−</button>
                <button class="btn line sm" onclick="stock(${i},1)" title="Sumar 1">+</button>
                <button class="btn line sm" onclick="stock(${i},10)" title="Sumar 10">+10</button>
            </td>
        </tr>`).join("")}
    </table>
</div>

<div class="card">
    <b>Agregar un nuevo ingrediente o producto al inventario</b><br><br>
    <div class="row" style="flex-wrap:wrap">
        <div style="flex:2;min-width:180px">
            <label style="font-size:12px;font-weight:600;color:var(--mut)">Nombre y unidad</label>
            <input id="in" placeholder="Ej. Tocineta (kg), Servilletas (paq)">
        </div>
        <div style="flex:1;min-width:120px">
            <label style="font-size:12px;font-weight:600;color:var(--mut)">Cantidad actual</label>
            <input id="iq" type="number" placeholder="Ej. 10">
        </div>
        <div style="flex:1;min-width:120px">
            <label style="font-size:12px;font-weight:600;color:var(--mut)">Avisar si baja de</label>
            <input id="im" type="number" placeholder="Ej. 3">
        </div>
        <button class="btn red" style="align-self:flex-end;margin-bottom:8px;padding:10px 18px" onclick="addInv()">+ Agregar</button>
    </div>
</div>`;
    },

    xls() {
        return `<h2>Conexión con tu archivo Excel</h2>
<p class="mut" style="margin-top:-6px;margin-bottom:12px">Aquí puedes verificar la conexión con tu archivo de Excel, forzar un guardado manual o cargar cambios que hayas hecho en el archivo.</p>

${!apiConnected ? `
<div class="status-box err">
    <div>
        <h3>Servidor desconectado</h3>
        <p>No se pudo conectar con el servidor local de Python. Asegúrate de tener la consola ejecutando <code>python app.py</code>.</p>
    </div>
</div>` : excelWritable ? `
<div class="status-box ok">
    <div>
        <h3>¡Excel conectado y al día!</h3>
        <p>Archivo vinculado: <b>${excelFilename}</b>. Todos los pedidos, pagos, gastos e inventario se guardan automáticamente en tu Excel.</p>
    </div>
</div>` : `
<div class="status-box warn">
    <div>
        <h3>El archivo ${excelFilename} está abierto en Microsoft Excel</h3>
        <p>Todos tus pedidos y cambios están <b>100% seguros y guardados</b> en la plataforma. Cuando termines de ver tu Excel y lo cierres, se guardarán en el archivo automáticamente.</p>
    </div>
</div>`}

<div class="card" style="padding:18px">
    <div class="row" style="flex-wrap:wrap">
        <div style="flex:2;min-width:240px">
            <b style="font-size:17px;color:#0f5132">Guardar todo en Excel ahora</b>
            <p class="mut" style="margin:4px 0 0;font-size:13px">La plataforma guarda automáticamente en cada acción, pero puedes presionar este botón en cualquier momento para forzar el guardado de todos los datos en tu archivo Excel.</p>
        </div>
        <button class="btn green lg" style="margin-top:8px" onclick="manualSaveAll()">Guardar todo en Excel</button>
    </div>
</div>

<div class="card discord-card" style="padding:18px">
    <div class="discord-header">
        <div>
            <b style="font-size:17px;color:#5865F2">Conexión con Discord (Alertas en tiempo real)</b>
            <p class="mut" style="margin:4px 0 0;font-size:13px">Recibe alertas automáticas con sonido e información completa en tu canal de Discord cada vez que haya un nuevo pedido, cambio de estado o edición.</p>
        </div>
        <span class="discord-badge ${discordStatus.configured ? '' : 'off'}">
            ${discordStatus.configured ? '● Discord Conectado' : '○ Discord no configurado'}
        </span>
    </div>
    <div style="margin-top:10px">
        <label style="font-size:12px;font-weight:600;color:var(--mut);display:block;margin-bottom:4px">URL del Webhook de Discord</label>
        <div class="discord-input-group">
            <input id="discord-webhook-input" placeholder="https://discord.com/api/webhooks/..." value="${escapeHtml(discordStatus.masked_url || '')}">
            <button id="btn-save-discord" class="btn line sm" onclick="saveDiscordWebhook()">Guardar Webhook</button>
            <button id="btn-test-discord" class="btn sm" style="background:#5865F2;color:#fff" onclick="testDiscordConnection()">Enviar mensaje de prueba</button>
        </div>
        <small class="mut" style="display:block;margin-top:8px">
            💡 <b>¿Cómo obtener la URL?</b> En Discord: clic derecho en tu canal de pedidos &gt; <i>Editar canal</i> &gt; <i>Integraciones</i> &gt; <i>Webhooks</i> &gt; <i>Nuevo Webhook</i> &gt; Copiar URL del Webhook.
        </small>
    </div>
</div>

<div class="card">
    <b style="font-size:15px">Descargar informe</b>
    <p class="mut" style="margin:4px 0 0;font-size:13px">El archivo Excel incluye un resumen y el detalle de pedidos, pagos al personal y gastos registrados.</p>
    <div class="report-controls">
        <div>
            <label for="report-range">Periodo</label>
            <select id="report-range" onchange="updateReportRange()">
                <option value="day">Un día</option>
                <option value="all">Todo el historial</option>
            </select>
        </div>
        <div>
            <label for="report-date">Fecha</label>
            <input id="report-date" type="date" value="${todayIso()}" max="${todayIso()}">
        </div>
        <button class="btn red" onclick="downloadReport()">Descargar informe Excel</button>
    </div>
</div>

<div class="card">
    <b style="font-size:15px">Otras acciones con Excel:</b>
    <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin-top:10px">
        <button class="btn line" onclick="importExcel()" style="text-align:left;padding:12px">
            <b style="color:var(--navy);font-size:14px">Cargar cambios de Excel</b><br>
            <small class="mut">Si editaste precios o inventario directamente en Excel, pulsa aquí para cargarlos en pantalla.</small>
        </button>
        <button class="btn line" onclick="testConnection()" style="text-align:left;padding:12px">
            <b style="color:var(--navy);font-size:14px">Comprobar conexión</b><br>
            <small class="mut">Verifica en 1 segundo que Excel y el servidor estén comunicándose bien.</small>
        </button>
        <button class="btn line" onclick="downloadExcel()" style="text-align:left;padding:12px">
            <b style="color:var(--navy);font-size:14px">Descargar copia de Excel</b><br>
            <small class="mut">Descarga una copia de seguridad de ${excelFilename} en tu carpeta de descargas.</small>
        </button>
    </div>
</div>

<div class="card">
    <b>¿Cómo funciona? (Guía súper sencilla para cualquier persona)</b>
    <div class="step-grid">
        <div class="step-card">
            <span class="step-num">1</span><b>Todo se guarda solo</b>
            <p class="mut" style="font-size:12px;margin:4px 0 0">Cada vez que tomas un pedido, cobras o anotas un gasto, el sistema lo guarda solo en <b>${excelFilename}</b>.</p>
        </div>
        <div class="step-card">
            <span class="step-num">2</span><b>Revisar en Excel</b>
            <p class="mut" style="font-size:12px;margin:4px 0 0">Abre <b>${excelFilename}</b> en tu computador cuando quieras para ver tus tablas, sumar o imprimir tus ventas.</p>
        </div>
        <div class="step-card">
            <span class="step-num">3</span><b>Cambiar precios</b>
            <p class="mut" style="font-size:12px;margin:4px 0 0">Edita los precios en Excel, presiona Guardar (Ctrl + S), cierra Excel y pulsa <b>Cargar cambios de Excel</b>.</p>
        </div>
    </div>
</div>

<div class="card wrap">
    <b>Hojas dentro de tu archivo Excel</b>
    <table style="margin-top:8px">
        <tr><th>Pestaña en Excel</th><th>¿Qué guarda?</th><th>Columnas guardadas</th></tr>
        <tr><td><b>Pedidos</b></td><td>Historial completo de clientes, ventas e instrucciones</td><td>ID, Fecha, Hora, Cliente, Tipo, Pago, Detalle, Total, Estado, Comentarios</td></tr>
        <tr><td><b>Menu</b></td><td>Productos, categorías y precios</td><td>Categoría, Producto, Precio</td></tr>
        <tr><td><b>Pagos</b></td><td>Pagos y turnos del personal</td><td>Fecha, Hora, Empleado, Concepto, Monto</td></tr>
        <tr><td><b>Gastos</b></td><td>Salidas de dinero de caja</td><td>Fecha, Hora, Concepto, Monto</td></tr>
        <tr><td><b>Inventario</b></td><td>Control de existencias e insumos</td><td>Producto, Stock, Mínimo</td></tr>
        <tr><td><b>Configuración</b></td><td>Base de caja y enlace del QR</td><td>Clave, Valor</td></tr>
    </table>
</div>

<div style="text-align:right;margin-top:14px">
    <button class="btn line sm" style="opacity:0.6" onclick="resetAll()">Restaurar datos de prueba iniciales</button>
</div>`;
    },
};

function render() {
    document.getElementById("nav").innerHTML = TABS.map(t => `<button class="${S.tab == t[0] ? "on" : ""}" onclick="go('${t[0]}')">${t[1]}</button>`).join("");
    document.getElementById("m").innerHTML = V[S.tab]();
    if (S.tab == "menu") drawQR();
}

loadState();
