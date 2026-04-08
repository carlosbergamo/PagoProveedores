// ════════════════════════════════════════════════════════════════
//  CONFIGURACIÓN — Editá solo estas líneas
// ════════════════════════════════════════════════════════════════
const EMAIL_NOTIFICACION = "carlosebergamo@gmail.com"; // ← tu email
const NOMBRE_SISTEMA     = "Gestión de Pagos";

// Columnas de la hoja Compras (base 0)
const COL_ID          = 0;
const COL_REGISTRADO  = 1;
const COL_PROVEEDOR   = 2;
const COL_MONTO       = 3;
const COL_DESCRIPCION = 4;
const COL_FECHA_COMPRA= 5;
const COL_FECHA_VENC  = 6;
const COL_PLAZO       = 7;
const COL_ESTADO      = 8;  // "Pendiente" | "Preparado" | "Pagado" | "Automático"
const COL_FECHA_PAGO  = 9;  // fecha efectiva de pago
const COL_TIPO        = 10; // "Sin Factura" | "Con Factura"
const COL_MODALIDAD   = 11; // "Efectivo" | "Cheque" | "Efectivo / Cheque" | "Transferencia"
const COL_TIPO_CHEQUE = 12; // "Electrónico" | "Físico"  (solo cuando modalidad contiene "Cheque" y tipo es "Con Factura")
const COL_NUM_CHEQUE  = 13; // Número de cheque
const COL_BANCO       = 14; // Banco emisor

// ════════════════════════════════════════════════════════════════
//  doGet
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === "getProveedores")       return getProveedores();
  if (action === "getComprasPendientes") return getComprasPendientes();
  if (action === "getHistorial")         return getHistorial();
  if (action === "getModalidades")       return getModalidades();
  if (action === "getBancos")            return getBancos();
  return jsonResponse({ success: false, error: "Acción no válida." });
}

// ════════════════════════════════════════════════════════════════
//  doPost
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || "registrarCompra";
    if (action === "registrarCompra")  return jsonResponse(registrarCompra(data));
    if (action === "marcarPreparado")  return jsonResponse(marcarEstado(data, "Preparado"));
    if (action === "marcarPagado")     return jsonResponse(marcarEstado(data, "Pagado"));
    if (action === "validarLogin")     return jsonResponse(validarLogin(data));
    if (action === "eliminarCompra")   return jsonResponse(eliminarCompra(data));
    return jsonResponse({ success: false, error: "Acción no válida." });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
//  getProveedores — solo nombres (el plazo lo ingresa el usuario)
// ════════════════════════════════════════════════════════════════
function getProveedores() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Proveedores");
    if (!sheet) return jsonResponse({ success: false, error: "Hoja 'Proveedores' no encontrada." });
    const rows        = sheet.getDataRange().getValues();
    const proveedores = rows.slice(1)
      .filter(r => r[0])
      .map(r => ({ nombre: r[0].toString().trim() }));
    return jsonResponse({ success: true, proveedores });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
//  getModalidades — lee de Config_Tablas
// ════════════════════════════════════════════════════════════════
function getModalidades() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Config_Tablas");
    if (!sheet) return jsonResponse({ success: false, error: "Hoja 'Config_Tablas' no encontrada." });
    const rows       = sheet.getDataRange().getValues();
    const modalidades = rows.slice(1).filter(r => r[0]).map(r => r[0].toString().trim());
    return jsonResponse({ success: true, modalidades });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
//  getBancos — lee de la hoja "Bancos"
// ════════════════════════════════════════════════════════════════
function getBancos() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Bancos");
    if (!sheet) return jsonResponse({ success: false, error: "Hoja 'Bancos' no encontrada." });
    const rows  = sheet.getDataRange().getValues();
    const bancos = rows.slice(1).filter(r => r[0]).map(r => r[0].toString().trim());
    return jsonResponse({ success: true, bancos });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
//  getComprasPendientes — Solo Sin Factura, no pagadas
// ════════════════════════════════════════════════════════════════
function getComprasPendientes() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Compras");
    if (!sheet) return jsonResponse({ success: false, error: "Hoja 'Compras' no encontrada." });

    const rows = sheet.getDataRange().getValues();
    const hoy  = new Date(); hoy.setHours(0, 0, 0, 0);

    const pendientes = rows.slice(1)
      .filter(r =>
        r[COL_ID] &&
        r[COL_TIPO] === "Sin Factura" &&   // ← solo sin factura
        r[COL_ESTADO] !== "Pagado"          // ← no pagados
      )
      .map(r => {
        const fechaVenc = new Date(r[COL_FECHA_VENC]);
        fechaVenc.setHours(0, 0, 0, 0);
        const diff = Math.round((fechaVenc - hoy) / 86400000);
        return {
          id:            r[COL_ID].toString(),
          proveedor:     r[COL_PROVEEDOR],
          monto:         Number(r[COL_MONTO]),
          descripcion:   r[COL_DESCRIPCION] || "",
          fechaCompra:   formatFecha(new Date(r[COL_FECHA_COMPRA])),
          fechaVenc:     formatFecha(fechaVenc),
          estado:        r[COL_ESTADO] || "Pendiente",
          modalidad:     r[COL_MODALIDAD] || "",
          diasRestantes: diff
        };
      })
      .sort((a, b) => a.diasRestantes - b.diasRestantes);

    return jsonResponse({ success: true, pendientes });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
//  getHistorial — Todas las compras para la pestaña Historial
// ════════════════════════════════════════════════════════════════
function getHistorial() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Compras");
    if (!sheet) return jsonResponse({ success: false, error: "Hoja 'Compras' no encontrada." });

    const rows = sheet.getDataRange().getValues();
    const hoy  = new Date(); hoy.setHours(0, 0, 0, 0);

    const compras = rows.slice(1)
      .filter(r => r[COL_ID])
      .map(r => {
        const fechaVenc = new Date(r[COL_FECHA_VENC]);
        fechaVenc.setHours(0, 0, 0, 0);
        const diff = Math.round((fechaVenc - hoy) / 86400000);
        return {
          id:           r[COL_ID].toString(),
          fechaCompra:  formatFecha(new Date(r[COL_FECHA_COMPRA])),
          proveedor:    r[COL_PROVEEDOR],
          descripcion:  r[COL_DESCRIPCION] || "",
          monto:        Number(r[COL_MONTO]),
          fechaVenc:    formatFecha(fechaVenc),
          tipo:         r[COL_TIPO] || "Sin Factura",
          estado:       r[COL_ESTADO] || "Pendiente",
          modalidad:    r[COL_MODALIDAD] || "",
          fechaPago:    r[COL_FECHA_PAGO] ? formatFecha(new Date(r[COL_FECHA_PAGO])) : "",
          tipoCheque:   r[COL_TIPO_CHEQUE] || "",
          numCheque:    r[COL_NUM_CHEQUE]  || "",
          banco:        r[COL_BANCO]       || "",
          diasRestantes: diff
        };
      })
      .sort((a, b) => {
        // Ordenar: primero los pendientes urgentes, después los pagados/automáticos
        const aActivo = a.estado === "Pendiente";
        const bActivo = b.estado === "Pendiente";
        if (aActivo && !bActivo) return -1;
        if (!aActivo && bActivo) return 1;
        return a.diasRestantes - b.diasRestantes;
      });

    return jsonResponse({ success: true, compras });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
//  registrarCompra
// ════════════════════════════════════════════════════════════════
function registrarCompra(data) {
  if (!data.proveedor)   throw new Error("El proveedor es obligatorio.");
  if (!data.monto)       throw new Error("El monto es obligatorio.");
  if (!data.fechaCompra) throw new Error("La fecha de compra es obligatoria.");
  if (!data.tipoFactura) throw new Error("El tipo de proveedor es obligatorio.");
  if (!data.diasPago)    throw new Error("El plazo de pago es obligatorio.");
  if (!data.modalidad)   throw new Error("La modalidad de pago es obligatoria.");

  const monto    = parseFloat(data.monto);
  if (isNaN(monto) || monto <= 0) throw new Error("El monto debe ser un número positivo.");

  const diasPago = parseInt(data.diasPago);
  if (isNaN(diasPago) || diasPago < 0) throw new Error("El plazo debe ser un número mayor o igual a cero.");

  const tipoFactura = data.tipoFactura;
  const modalidad   = data.modalidad.toString().trim();

  // Validar campos de cheque (solo Con Factura + modalidad que incluya "Cheque")
  const esCheque = modalidad.toLowerCase().includes("cheque");
  if (esCheque) {
    if (!data.tipoCheque) throw new Error("El tipo de cheque es obligatorio.");
    if (!data.numCheque)  throw new Error("El número de cheque es obligatorio.");
    if (!data.banco)      throw new Error("El banco es obligatorio.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Validar que el proveedor exista
  const provSheet = ss.getSheetByName("Proveedores");
  if (!provSheet) throw new Error("Hoja 'Proveedores' no encontrada.");
  const provRows = provSheet.getDataRange().getValues();
  const provRow  = provRows.slice(1).find(r => r[0].toString().trim() === data.proveedor.trim());
  if (!provRow) throw new Error("Proveedor no encontrado.");

  const partes      = data.fechaCompra.split("-");
  const fechaCompra = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
  const fechaVenc   = new Date(fechaCompra);
  fechaVenc.setDate(fechaVenc.getDate() + diasPago);

  const comprasSheet = ss.getSheetByName("Compras");
  if (!comprasSheet) throw new Error("Hoja 'Compras' no encontrada.");

  const id          = new Date().getTime();
  const descripcion = data.descripcion ? data.descripcion.toString().trim() : "";
  const estadoInicial = tipoFactura === "Con Factura" ? "Automático" : "Pendiente";

  const tipoCheque = esCheque ? data.tipoCheque.toString().trim() : "";
  const numCheque  = esCheque ? data.numCheque.toString().trim()  : "";
  const banco      = esCheque ? data.banco.toString().trim()      : "";

  comprasSheet.appendRow([
    id, new Date(), data.proveedor.trim(), monto, descripcion,
    fechaCompra, fechaVenc, diasPago, estadoInicial, "", tipoFactura, modalidad,
    tipoCheque, numCheque, banco
  ]);

  const lastRow = comprasSheet.getLastRow();
  if (tipoFactura === "Con Factura") {
    comprasSheet.getRange(lastRow, 1, 1, 15).setBackground("#e8f0fe");
  }

  const tipoMsg = tipoFactura === "Con Factura"
    ? "Se registró como pago automático."
    : `Vence el ${formatFecha(fechaVenc)}.`;

  return { success: true, message: `Compra registrada. ${tipoMsg}` };
}

// ════════════════════════════════════════════════════════════════
//  marcarEstado — cambia estado: "Preparado" o "Pagado"
// ════════════════════════════════════════════════════════════════
function marcarEstado(data, nuevoEstado) {
  if (!data.id) throw new Error("ID de compra no proporcionado.");

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Compras");
  if (!sheet) throw new Error("Hoja 'Compras' no encontrada.");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL_ID].toString() !== data.id.toString()) continue;
    if (rows[i][COL_TIPO] === "Con Factura") throw new Error("Las compras con factura no requieren gestión manual.");

    sheet.getRange(i + 1, COL_ESTADO + 1).setValue(nuevoEstado);

    if (nuevoEstado === "Pagado") {
      sheet.getRange(i + 1, COL_FECHA_PAGO + 1).setValue(new Date());
      sheet.getRange(i + 1, 1, 1, 15).setBackground("#e8f5e9"); // verde
    } else if (nuevoEstado === "Preparado") {
      sheet.getRange(i + 1, 1, 1, 15).setBackground("#fff8e1"); // amarillo suave
    }

    return { success: true, message: `Estado actualizado a "${nuevoEstado}".` };
  }
  throw new Error("No se encontró la compra con ese ID.");
}

// ════════════════════════════════════════════════════════════════
//  enviarResumenSemanal — trigger: miércoles 8am
//  Siempre envía. Solo Sin Factura. 4 secciones.
// ════════════════════════════════════════════════════════════════
function enviarResumenSemanal() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Compras");
  if (!sheet) return;

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return;

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);

  const vencidos    = [];  // diff < 0
  const prox3dias   = [];  // 0 <= diff <= 3
  const prox7dias   = [];  // 4 <= diff <= 7
  const prox14dias  = [];  // 8 <= diff <= 14

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[COL_ID])                          continue;
    if (row[COL_TIPO]   !== "Sin Factura")     continue;
    if (row[COL_ESTADO] === "Pagado")          continue;

    const fechaVenc = new Date(row[COL_FECHA_VENC]);
    fechaVenc.setHours(0, 0, 0, 0);
    const diff = Math.round((fechaVenc - hoy) / 86400000);

    const item = {
      proveedor:    row[COL_PROVEEDOR],
      monto:        Number(row[COL_MONTO]),
      descripcion:  row[COL_DESCRIPCION] || "-",
      estado:       row[COL_ESTADO] || "Pendiente",
      fechaVenc:    fechaVenc,
      diasRestantes: diff
    };

    if (diff < 0)                  vencidos.push(item);
    if (diff >= 0 && diff <= 3)    prox3dias.push(item);
    if (diff >= 4 && diff <= 7)    prox7dias.push(item);
    if (diff >= 8 && diff <= 14)   prox14dias.push(item);
  }

  enviarMailResumen(vencidos, prox3dias, prox7dias, prox14dias);
}

// ════════════════════════════════════════════════════════════════
//  enviarMailResumen — reporte consolidado semanal (4 secciones)
// ════════════════════════════════════════════════════════════════
function enviarMailResumen(vencidos, prox3, prox7, prox14) {
  const totalVencidos = vencidos.reduce((s, i) => s + i.monto, 0);
  const total3        = prox3.reduce((s, i) => s + i.monto, 0);
  const total7        = prox7.reduce((s, i) => s + i.monto, 0);
  const total14       = prox14.reduce((s, i) => s + i.monto, 0);
  const totalGeneral  = totalVencidos + total3 + total7 + total14;

  // ── Helper: construir tabla de items ──
  function tablaItems(items, colorHeader, colorBorder, colorBg, colorTexto) {
    if (items.length === 0) return `<p style="margin:8px 0 0;font-size:12px;color:#b0aca6;font-style:italic;">Sin pagos en este período.</p>`;
    return `
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;border:1.5px solid ${colorBorder};border-radius:8px;overflow:hidden;">
        <thead><tr style="background:${colorBg};">
          <th style="padding:8px 12px;text-align:left;font-size:10px;color:${colorTexto};text-transform:uppercase;letter-spacing:1px;">Proveedor</th>
          <th style="padding:8px 12px;text-align:left;font-size:10px;color:${colorTexto};text-transform:uppercase;letter-spacing:1px;">Descripción</th>
          <th style="padding:8px 12px;text-align:left;font-size:10px;color:${colorTexto};text-transform:uppercase;letter-spacing:1px;">Estado</th>
          <th style="padding:8px 12px;text-align:left;font-size:10px;color:${colorTexto};text-transform:uppercase;letter-spacing:1px;">Fecha</th>
          <th style="padding:8px 12px;text-align:right;font-size:10px;color:${colorTexto};text-transform:uppercase;letter-spacing:1px;">Monto</th>
        </tr></thead>
        <tbody>
          ${items.map(item => `
            <tr>
              <td style="padding:9px 12px;border-bottom:1px solid ${colorBorder};font-weight:600;color:#2c2a27;">${item.proveedor}</td>
              <td style="padding:9px 12px;border-bottom:1px solid ${colorBorder};color:#5a5651;font-size:12px;">${item.descripcion}</td>
              <td style="padding:9px 12px;border-bottom:1px solid ${colorBorder};font-size:11px;">${item.estado === 'Preparado' ? '🟡 Preparado' : '⏳ Pendiente'}</td>
              <td style="padding:9px 12px;border-bottom:1px solid ${colorBorder};color:#5a5651;font-size:12px;">${formatFecha(item.fechaVenc)}</td>
              <td style="padding:9px 12px;border-bottom:1px solid ${colorBorder};font-weight:700;color:${colorHeader};text-align:right;">${formatMonto(item.monto)}</td>
            </tr>`).join("")}
        </tbody>
        <tfoot><tr style="background:${colorBg};">
          <td colspan="4" style="padding:9px 12px;font-weight:700;color:${colorTexto};font-size:12px;">TOTAL</td>
          <td style="padding:9px 12px;font-weight:800;color:${colorHeader};font-size:15px;text-align:right;">${formatMonto(items.reduce((s,i)=>s+i.monto,0))}</td>
        </tr></tfoot>
      </table>`;
  }

  // ── Asunto dinámico ──
  let asunto = `📋 Resumen semanal de pagos — ${formatMonto(totalGeneral)}`;
  if (vencidos.length > 0) asunto = `🚨 ${vencidos.length} vencido${vencidos.length>1?'s':''} · ` + asunto;

  const html = `<!DOCTYPE html><html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f5f3ef;font-family:'Helvetica Neue',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3ef;padding:32px 0;">
<tr><td align="center">
<table width="620" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 24px rgba(0,0,0,0.08);">

  <!-- HEADER -->
  <tr><td style="background:#0f0e0d;padding:28px 36px;">
    <p style="margin:0;color:#b8924a;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;">📋 Resumen Semanal · Solo Sin Factura</p>
    <h1 style="margin:8px 0 0;color:#fff;font-size:22px;font-weight:700;">${NOMBRE_SISTEMA}</h1>
    <p style="margin:6px 0 0;color:#8a7a62;font-size:12px;">${formatFechaHora(new Date())}</p>
  </td></tr>

  <!-- TOTALIZADOR GENERAL -->
  <tr><td style="padding:24px 36px 0;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td width="48%" style="background:#fff0f0;border:2px solid #ffcdd2;border-radius:8px;padding:14px 18px;${vencidos.length===0?'opacity:.5':''}">
          <p style="margin:0;font-size:10px;font-weight:700;color:#b71c1c;text-transform:uppercase;letter-spacing:1px;">🚨 Vencidos</p>
          <p style="margin:5px 0 0;font-size:22px;font-weight:800;color:#b71c1c;">${formatMonto(totalVencidos)}</p>
          <p style="margin:2px 0 0;font-size:11px;color:#c62828;">${vencidos.length} pago${vencidos.length!==1?'s':''}</p>
        </td>
        <td width="4%"></td>
        <td width="48%" style="background:var(--gold-faint,#f5edd8);border:1.5px solid #e8d5aa;border-radius:8px;padding:14px 18px;">
          <p style="margin:0;font-size:10px;font-weight:700;color:#7a5c1e;text-transform:uppercase;letter-spacing:1px;">📅 Próximos 14 días</p>
          <p style="margin:5px 0 0;font-size:22px;font-weight:800;color:#2c2a27;">${formatMonto(total3+total7+total14)}</p>
          <p style="margin:2px 0 0;font-size:11px;color:#8a7a62;">${prox3.length+prox7.length+prox14.length} pago${(prox3.length+prox7.length+prox14.length)!==1?'s':''}</p>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- SECCIÓN 1: VENCIDOS -->
  <tr><td style="padding:22px 36px 0;">
    <div style="border-left:4px solid #e53935;padding-left:14px;">
      <p style="margin:0;font-size:13px;font-weight:700;color:#b71c1c;text-transform:uppercase;letter-spacing:1px;">1 · Pagos Vencidos</p>
      <p style="margin:3px 0 0;font-size:12px;color:#c62828;">Fecha anterior a hoy — requieren atención inmediata</p>
    </div>
    ${tablaItems(vencidos, '#b71c1c', '#ffcdd2', '#ffebee', '#b71c1c')}
  </td></tr>

  <!-- SECCIÓN 2: PRÓXIMOS 3 DÍAS -->
  <tr><td style="padding:22px 36px 0;">
    <div style="border-left:4px solid #e65100;padding-left:14px;">
      <p style="margin:0;font-size:13px;font-weight:700;color:#e65100;text-transform:uppercase;letter-spacing:1px;">2 · Próximos 3 días</p>
      <p style="margin:3px 0 0;font-size:12px;color:#bf360c;">Urgentes — vencen entre hoy y el ${formatFecha(diasDesdeHoy(3))}</p>
    </div>
    ${tablaItems(prox3, '#e65100', '#ffe0b2', '#fff3e0', '#e65100')}
  </td></tr>

  <!-- SECCIÓN 3: PRÓXIMA SEMANA -->
  <tr><td style="padding:22px 36px 0;">
    <div style="border-left:4px solid #b8924a;padding-left:14px;">
      <p style="margin:0;font-size:13px;font-weight:700;color:#7a5c1e;text-transform:uppercase;letter-spacing:1px;">3 · Próxima semana</p>
      <p style="margin:3px 0 0;font-size:12px;color:#8a7a62;">Planificación corta — vencen entre el ${formatFecha(diasDesdeHoy(4))} y el ${formatFecha(diasDesdeHoy(7))}</p>
    </div>
    ${tablaItems(prox7, '#b8924a', '#e8d5aa', '#fff8ed', '#7a5c1e')}
  </td></tr>

  <!-- SECCIÓN 4: PRÓXIMA QUINCENA -->
  <tr><td style="padding:22px 36px 0;">
    <div style="border-left:4px solid #1a7a4a;padding-left:14px;">
      <p style="margin:0;font-size:13px;font-weight:700;color:#1a7a4a;text-transform:uppercase;letter-spacing:1px;">4 · Próxima quincena</p>
      <p style="margin:3px 0 0;font-size:12px;color:#2e7d52;">Planificación media — vencen entre el ${formatFecha(diasDesdeHoy(8))} y el ${formatFecha(diasDesdeHoy(14))}</p>
    </div>
    ${tablaItems(prox14, '#1a7a4a', '#b8dfc9', '#f0faf4', '#2e7d52')}
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="padding:28px 36px;border-top:1px solid #f0ece4;margin-top:24px;">
    <p style="margin:0;font-size:11px;color:#b0aca6;text-align:center;">Resumen automático de <strong>${NOMBRE_SISTEMA}</strong> · Todos los miércoles a las 8:00 AM</p>
  </td></tr>

</table>
</td></tr></table>
</body></html>`;

  GmailApp.sendEmail(EMAIL_NOTIFICACION, asunto, "", { htmlBody: html });
}

// ════════════════════════════════════════════════════════════════
//  configurarTrigger — ejecutá UNA SOLA VEZ
// ════════════════════════════════════════════════════════════════
function configurarTrigger() {
  // Elimina triggers anteriores de ambos nombres (por si venías del sistema viejo)
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "verificarVencimientos" || fn === "enviarResumenSemanal") {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Miércoles a las 8am
  ScriptApp.newTrigger("enviarResumenSemanal")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(8)
    .create();
  Logger.log("✅ Trigger configurado: enviarResumenSemanal todos los miércoles a las 8am.");
}

// ════════════════════════════════════════════════════════════════
//  validarLogin — verifica credenciales contra Config_Usuarios
// ════════════════════════════════════════════════════════════════
function validarLogin(data) {
  if (!data.usuario || !data.password) throw new Error("Credenciales incompletas.");

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config_Usuarios");
  if (!sheet) throw new Error("Hoja 'Config_Usuarios' no encontrada. Ejecutá inicializarHojas.");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const usuario  = rows[i][0] ? rows[i][0].toString().trim() : "";
    const password = rows[i][1] ? rows[i][1].toString().trim() : "";
    const permisos = rows[i][2] ? rows[i][2].toString().trim() : "Operador";

    if (usuario.toLowerCase() === data.usuario.toLowerCase().trim() &&
        password === data.password.trim()) {
      return { success: true, usuario, permisos };
    }
  }
  return { success: false, error: "Usuario o contraseña incorrectos." };
}

// ════════════════════════════════════════════════════════════════
//  eliminarCompra — solo Admin puede eliminar
// ════════════════════════════════════════════════════════════════
function eliminarCompra(data) {
  if (!data.id)       throw new Error("ID de compra no proporcionado.");
  if (!data.usuario)  throw new Error("Usuario no proporcionado.");

  // Verificar que el usuario sea Admin
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const usrSheet  = ss.getSheetByName("Config_Usuarios");
  if (!usrSheet) throw new Error("Hoja 'Config_Usuarios' no encontrada.");

  const usrRows = usrSheet.getDataRange().getValues();
  const usrRow  = usrRows.slice(1).find(r =>
    r[0].toString().trim().toLowerCase() === data.usuario.toLowerCase()
  );
  if (!usrRow || usrRow[2].toString().trim() !== "Admin") {
    throw new Error("Sin permisos. Solo el Administrador puede eliminar registros.");
  }

  const sheet = ss.getSheetByName("Compras");
  if (!sheet) throw new Error("Hoja 'Compras' no encontrada.");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL_ID].toString() === data.id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Registro eliminado correctamente." };
    }
  }
  throw new Error("No se encontró la compra con ese ID.");
}

// ════════════════════════════════════════════════════════════════
//  inicializarHojas — ejecutá UNA SOLA VEZ
// ════════════════════════════════════════════════════════════════
function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let provSheet = ss.getSheetByName("Proveedores") || ss.insertSheet("Proveedores");
  provSheet.clearContents();
  provSheet.getRange("A1").setValues([["Proveedor"]]);
  provSheet.getRange("A1").setFontWeight("bold").setBackground("#0f0e0d").setFontColor("#b8924a");
  provSheet.getRange("A2:A4").setValues([["Proveedor 1"],["Proveedor 2"],["Proveedor 3"]]);
  provSheet.setColumnWidth(1, 240);
  provSheet.getRange("A1").setNote("El plazo de pago ahora se ingresa manualmente en cada compra.");

  let comprasSheet = ss.getSheetByName("Compras") || ss.insertSheet("Compras");
  comprasSheet.clearContents();
  const headers = ["ID","Registrado","Proveedor","Monto","Descripción","Fecha Compra","Fecha Vencimiento","Plazo (días)","Estado","Fecha Pago","Tipo","Modalidad"];
  comprasSheet.getRange(1,1,1,headers.length).setValues([headers]);
  comprasSheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#0f0e0d").setFontColor("#b8924a");
  [180,160,180,100,220,120,150,100,110,120,110,130].forEach((w,i) => comprasSheet.setColumnWidth(i+1,w));

  // — Config_Tablas —
  let tabSheet = ss.getSheetByName("Config_Tablas") || ss.insertSheet("Config_Tablas");
  tabSheet.clearContents();
  tabSheet.getRange("A1").setValues([["Formas de Pago"]]);
  tabSheet.getRange("A1").setFontWeight("bold").setBackground("#0f0e0d").setFontColor("#b8924a");
  tabSheet.getRange("A2:A5").setValues([["Efectivo"],["Cheque"],["Efectivo / Cheque"],["Transferencia"]]);
  tabSheet.setColumnWidth(1, 200);

  // — Config_Usuarios —
  let usrSheet = ss.getSheetByName("Config_Usuarios") || ss.insertSheet("Config_Usuarios");
  usrSheet.clearContents();
  usrSheet.getRange("A1:C1").setValues([["Usuario","Password","Permisos"]]);
  usrSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#0f0e0d").setFontColor("#b8924a");
  // Usuarios de ejemplo — CAMBIÁ LAS CONTRASEÑAS antes de usar en producción
  usrSheet.getRange("A2:C3").setValues([
    ["admin",    "admin123",  "Admin"],
    ["operador", "op2024",    "Operador"]
  ]);
  usrSheet.setColumnWidth(1,160); usrSheet.setColumnWidth(2,160); usrSheet.setColumnWidth(3,120);
  // Proteger visualmente (no es seguridad real, solo orientación)
  usrSheet.getRange("A2:C10").setBackground("#fff8e1");

  Logger.log("✅ Hojas inicializadas. IMPORTANTE: cambiá las contraseñas en Config_Usuarios antes de usar en producción.");
}

// ════════════════════════════════════════════════════════════════
//  eliminarTrigger — ejecutá para desactivar el resumen semanal
// ════════════════════════════════════════════════════════════════
function eliminarTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let eliminados = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "enviarResumenSemanal") {
      ScriptApp.deleteTrigger(t);
      eliminados++;
    }
  });
  Logger.log(eliminados > 0
    ? `✅ Trigger eliminado. Ya no se enviarán resúmenes semanales.`
    : `⚠️ No se encontró ningún trigger activo para enviarResumenSemanal.`
  );
}

// ════════════════════════════════════════════════════════════════
//  testMail — para probar el mail con datos de ejemplo
// ════════════════════════════════════════════════════════════════
function testMail() {
  const hoy  = new Date();
  const vencP = new Date(hoy); vencP.setDate(vencP.getDate() - 5);
  const venc2 = new Date(hoy); venc2.setDate(venc2.getDate() + 2);
  const venc6 = new Date(hoy); venc6.setDate(venc6.getDate() + 6);
  const venc12= new Date(hoy); venc12.setDate(venc12.getDate() + 12);
  enviarMailResumen(
    [{ proveedor:"Proveedor A", monto:8500, descripcion:"Materia prima", estado:"Pendiente", fechaVenc:vencP, diasRestantes:-5 }],
    [{ proveedor:"Proveedor B", monto:3200, descripcion:"Servicio",      estado:"Preparado", fechaVenc:venc2, diasRestantes:2  }],
    [{ proveedor:"Proveedor C", monto:5100, descripcion:"Insumos",       estado:"Pendiente", fechaVenc:venc6, diasRestantes:6  }],
    [{ proveedor:"Proveedor D", monto:1800, descripcion:"Flete",         estado:"Pendiente", fechaVenc:venc12,diasRestantes:12 }]
  );
  Logger.log("✅ Mail de prueba enviado.");
}

// ════════════════════════════════════════════════════════════════
//  Helpers
// ════════════════════════════════════════════════════════════════
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function diasDesdeHoy(n) {
  const d = new Date(); d.setHours(0,0,0,0); d.setDate(d.getDate() + n); return d;
}
function formatFecha(date) {
  return new Date(date).toLocaleDateString("es-AR",{day:"2-digit",month:"2-digit",year:"numeric"});
}
function formatFechaHora(date) {
  return new Date(date).toLocaleString("es-AR");
}
function formatMonto(n) {
  return "$\u202F"+Number(n).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2});
}
