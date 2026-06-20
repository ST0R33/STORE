// ════════════════════════════════════════════════════════════════
// OrahBuy — Automações GAS
// Colar em: Apps Script da planilha da Roberta → novo arquivo .gs
// Email admin: roberta.rtg@gmail.com
// ════════════════════════════════════════════════════════════════

const EMAIL_ADMIN   = "roberta.rtg@gmail.com";
const WPP_LOJA     = "5521976264793";
const NOME_LOJA    = "OrahBuy";

// ────────────────────────────────────────────────────────────────
// UTIL
// ────────────────────────────────────────────────────────────────

function getCalendar() {
  return CalendarApp.getDefaultCalendar();
}

function fmtBRL(v) {
  return "R$ " + Number(v).toFixed(2).replace(".", ",");
}

function parseDateBR(str) {
  // "DD/MM/YYYY" ou "YYYY-MM-DD"
  if (!str) return null;
  if (str.includes("/")) {
    const [d, m, y] = str.split("/");
    return new Date(Number(y), Number(m) - 1, Number(d), 9, 0, 0);
  }
  const [y, m, d] = str.split("-");
  return new Date(Number(y), Number(m) - 1, Number(d), 9, 0, 0);
}

// ────────────────────────────────────────────────────────────────
// 1. PEDIDO CRIADO → Calendar + Email admin
// Chamar de dentro do doPost() quando action === "criarPedido"
// ────────────────────────────────────────────────────────────────

function onNovoPedido(pedido) {
  try {
    const id      = pedido["ID Pedido"] || pedido.id || "?";
    const nome    = pedido["Nome Cliente"] || pedido.nome_cliente || "Cliente";
    const total   = pedido["Total"] || pedido.total || 0;
    const status  = pedido["Status"] || "Pendente";
    const dataVenc = pedido["Data_Vencimento"] || null;
    const itens   = pedido["Itens"] || "";

    // ── Google Calendar ──
    const cal = getCalendar();
    const inicio = dataVenc ? parseDateBR(dataVenc) : new Date();
    if (!dataVenc) inicio.setHours(9, 0, 0, 0);
    const fim = new Date(inicio.getTime() + 60 * 60 * 1000);

    const evento = cal.createEvent(
      `🛍️ Pedido #${id} — ${nome} (${fmtBRL(total)})`,
      inicio,
      fim,
      {
        description:
          `Cliente: ${nome}\nTotal: ${fmtBRL(total)}\nStatus: ${status}\n\nItens:\n${itens}`,
      }
    );
    // Lembrar D-1 e no dia às 9h
    evento.addEmailReminder(1440); // 24h antes
    evento.addEmailReminder(0);    // no momento

    // ── Email admin ──
    const assunto = `🛍️ Novo Pedido #${id} — ${nome}`;
    const corpo = `
      <div style="font-family:sans-serif;max-width:520px">
        <div style="background:#FF5500;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">
          <b style="font-size:18px">🛍️ Novo Pedido Recebido!</b>
        </div>
        <div style="background:#f9f9f9;padding:16px 20px;border:1px solid #eee;border-radius:0 0 10px 10px">
          <p><b>Pedido:</b> #${id}</p>
          <p><b>Cliente:</b> ${nome}</p>
          <p><b>Total:</b> ${fmtBRL(total)}</p>
          <p><b>Status:</b> ${status}</p>
          ${dataVenc ? `<p><b>Vencimento:</b> ${dataVenc}</p>` : ""}
          <p><b>Itens:</b><br>${itens}</p>
        </div>
      </div>`;
    GmailApp.sendEmail(EMAIL_ADMIN, assunto, "", { htmlBody: corpo });

    return { ok: true, eventoId: evento.getId() };
  } catch (e) {
    Logger.log("onNovoPedido erro: " + e);
    return { ok: false, erro: String(e) };
  }
}

// ────────────────────────────────────────────────────────────────
// 2. LEMBRETES RECORRENTES → N eventos no Calendar
// Chamar quando admin salva parcelas no detalhe do pedido
// action: "criarLembretesRecorrentes"
// params: { pedidoId, nomeCliente, dataBase (YYYY-MM-DD), nParcelas }
// ────────────────────────────────────────────────────────────────

function criarLembretesRecorrentes(pedidoId, nomeCliente, dataBase, nParcelas) {
  try {
    const cal = getCalendar();
    const [y0, m0, d0] = dataBase.split("-").map(Number);
    const criados = [];

    for (let i = 0; i < nParcelas; i++) {
      let m = m0 + i - 1;
      const y = y0 + Math.floor(m / 12);
      m = m % 12;
      const maxDay = new Date(y, m + 1, 0).getDate();
      const dia = Math.min(d0, maxDay);
      const inicio = new Date(y, m, dia, 9, 0, 0);
      const fim = new Date(inicio.getTime() + 30 * 60 * 1000);

      const ev = cal.createEvent(
        `🔔 Parcela ${i + 1}/${nParcelas} — ${nomeCliente} — Pedido #${pedidoId}`,
        inicio,
        fim,
        {
          description: `Lembrete de cobrança da parcela ${i + 1} de ${nParcelas}.\nCliente: ${nomeCliente}\nPedido: #${pedidoId}`,
        }
      );
      ev.addEmailReminder(1440); // D-1
      ev.addEmailReminder(60);   // 1h antes
      criados.push(ev.getId());
    }

    return { ok: true, criados: criados.length };
  } catch (e) {
    Logger.log("criarLembretesRecorrentes erro: " + e);
    return { ok: false, erro: String(e) };
  }
}

// ────────────────────────────────────────────────────────────────
// 3. STATUS MUDOU → Email admin + (opcional) WA cliente
// Chamar de dentro do handler que atualiza status na planilha
// ────────────────────────────────────────────────────────────────

function onStatusChange(pedido, novoStatus) {
  try {
    const id   = pedido["ID Pedido"] || pedido.id || "?";
    const nome = pedido["Nome Cliente"] || "Cliente";
    const total = pedido["Total"] || 0;

    const icons = {
      "Em Andamento": "🔄", "Enviado": "🚚", "Finalizado": "✅",
      "Cancelado": "❌", "Pendente": "⏳"
    };
    const ico = icons[novoStatus] || "📌";

    // Email admin
    GmailApp.sendEmail(
      EMAIL_ADMIN,
      `${ico} Pedido #${id} → ${novoStatus}`,
      "",
      {
        htmlBody: `
          <div style="font-family:sans-serif;max-width:480px">
            <div style="background:#FF5500;color:#fff;padding:14px 18px;border-radius:8px 8px 0 0">
              <b>${ico} Status Atualizado</b>
            </div>
            <div style="padding:14px 18px;background:#f9f9f9;border:1px solid #eee;border-radius:0 0 8px 8px">
              <p><b>Pedido:</b> #${id}</p>
              <p><b>Cliente:</b> ${nome}</p>
              <p><b>Total:</b> ${fmtBRL(total)}</p>
              <p><b>Novo status:</b> <span style="color:#FF5500;font-weight:700">${novoStatus}</span></p>
            </div>
          </div>`
      }
    );
    return { ok: true };
  } catch (e) {
    Logger.log("onStatusChange erro: " + e);
    return { ok: false, erro: String(e) };
  }
}

// ────────────────────────────────────────────────────────────────
// 4. PAGAMENTO REGISTRADO → Email comprovante admin
// ────────────────────────────────────────────────────────────────

function onPagamentoRegistrado(pedido, valorPago, dataPag) {
  try {
    const id    = pedido["ID Pedido"] || "?";
    const nome  = pedido["Nome Cliente"] || "Cliente";
    const total = pedido["Total"] || 0;

    GmailApp.sendEmail(
      EMAIL_ADMIN,
      `💰 Pagamento Registrado — Pedido #${id}`,
      "",
      {
        htmlBody: `
          <div style="font-family:sans-serif;max-width:480px">
            <div style="background:#00a854;color:#fff;padding:14px 18px;border-radius:8px 8px 0 0">
              <b>💰 Pagamento Confirmado!</b>
            </div>
            <div style="padding:14px 18px;background:#f9f9f9;border:1px solid #eee;border-radius:0 0 8px 8px">
              <p><b>Pedido:</b> #${id}</p>
              <p><b>Cliente:</b> ${nome}</p>
              <p><b>Total do pedido:</b> ${fmtBRL(total)}</p>
              <p><b>Valor recebido:</b> ${fmtBRL(valorPago)}</p>
              <p><b>Data:</b> ${dataPag || new Date().toLocaleDateString("pt-BR")}</p>
            </div>
          </div>`
      }
    );
    return { ok: true };
  } catch (e) {
    Logger.log("onPagamentoRegistrado erro: " + e);
    return { ok: false, erro: String(e) };
  }
}

// ────────────────────────────────────────────────────────────────
// 5. RESUMO DIÁRIO — trigger: todo dia às 8h
// Configurar em: Gatilhos → onTriggerDiario → baseado em tempo → dia → 8h
// ────────────────────────────────────────────────────────────────

function onTriggerDiario() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName("Pedidos") || ss.getSheetByName("pedidos");
    const prSheet = ss.getSheetByName("Produtos") || ss.getSheetByName("produtos");
    if (!pSheet) return;

    const pedRows = pSheet.getDataRange().getValues();
    const headers = pedRows[0];
    const iStatus = headers.indexOf("Status");
    const iNome   = headers.indexOf("Nome Cliente");
    const iTotal  = headers.indexOf("Total");
    const iVenc   = headers.indexOf("Data_Vencimento");
    const iData   = headers.indexOf("Data/Hora");

    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const ontem = new Date(hoje); ontem.setDate(ontem.getDate() - 1);

    const novosHoje = [], vencHoje = [], pendAntigos = [];

    pedRows.slice(1).forEach(row => {
      const status = row[iStatus] || "";
      const nome   = row[iNome]   || "";
      const total  = row[iTotal]  || 0;
      const venc   = iVenc >= 0 ? parseDateBR(String(row[iVenc] || "")) : null;
      const criado = iData >= 0 ? new Date(row[iData]) : null;

      // Novos (criados ontem/hoje)
      if (criado && criado >= ontem && !["Finalizado","Cancelado"].includes(status))
        novosHoje.push({ nome, total, status });

      // Vencendo hoje
      if (venc && venc.getTime() === hoje.getTime() && !["Finalizado","Cancelado"].includes(status))
        vencHoje.push({ nome, total });

      // Pendentes antigos (>48h)
      if (status === "Pendente" && criado && (hoje - criado) > 172800000)
        pendAntigos.push({ nome, total });
    });

    // Estoque crítico
    const estoqueAlertas = [];
    if (prSheet) {
      const prRows = prSheet.getDataRange().getValues();
      const pH = prRows[0];
      const iNomePr = pH.indexOf("Nome do Produto");
      const iEst    = pH.indexOf("Estoque");
      const iSt     = pH.indexOf("Status");
      prRows.slice(1).forEach(row => {
        const est = Number(row[iEst]);
        const st  = row[iSt] || "Ativo";
        if (!isNaN(est) && est <= 3 && st !== "Inativo")
          estoqueAlertas.push({ nome: row[iNomePr], est });
      });
    }

    // Montar email
    const sec = (titulo, items, fn) => items.length
      ? `<h3 style="color:#FF5500;margin:16px 0 6px">${titulo}</h3>
         <ul style="margin:0;padding-left:18px">${items.map(fn).join("")}</ul>`
      : "";

    const html = `
      <div style="font-family:sans-serif;max-width:540px">
        <div style="background:#FF5500;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">
          <b style="font-size:17px">☀️ Resumo Diário ${NOME_LOJA} — ${hoje.toLocaleDateString("pt-BR")}</b>
        </div>
        <div style="padding:16px 20px;background:#fff;border:1px solid #eee;border-radius:0 0 10px 10px">
          ${novosHoje.length === 0 && vencHoje.length === 0 && pendAntigos.length === 0 && estoqueAlertas.length === 0
            ? "<p style='color:#888'>✅ Tudo tranquilo, nenhum alerta hoje.</p>"
            : ""}
          ${sec("🛍️ Pedidos novos / em aberto", novosHoje,
            p => `<li>${p.nome} — ${fmtBRL(p.total)} (${p.status})</li>`)}
          ${sec("⏰ Vencimentos HOJE", vencHoje,
            p => `<li style="color:#FF5500;font-weight:600">${p.nome} — ${fmtBRL(p.total)}</li>`)}
          ${sec("🔴 Pendentes há mais de 48h", pendAntigos,
            p => `<li>${p.nome} — ${fmtBRL(p.total)}</li>`)}
          ${sec("📦 Estoque crítico (≤ 3)", estoqueAlertas,
            p => `<li>${p.nome} — ${p.est} unid.</li>`)}
        </div>
      </div>`;

    GmailApp.sendEmail(EMAIL_ADMIN, `☀️ Resumo Diário ${NOME_LOJA} — ${hoje.toLocaleDateString("pt-BR")}`, "", { htmlBody: html });
  } catch (e) {
    Logger.log("onTriggerDiario erro: " + e);
  }
}

// ────────────────────────────────────────────────────────────────
// COMO ATIVAR O RESUMO DIÁRIO:
//
// No Apps Script → Gatilhos (ícone relógio) → + Adicionar gatilho:
//   Função: onTriggerDiario
//   Origem: baseado em tempo
//   Tipo: temporizador por dia
//   Hora: entre 8h e 9h
// ────────────────────────────────────────────────────────────────


// ────────────────────────────────────────────────────────────────
// 6. RASTREAR PEDIDO — chamado via doGet ?action=rastrear&id=ORB-...
// Retorna: { ok, pedido: { id, status, nome, data, total, itens } }
// ────────────────────────────────────────────────────────────────

function rastrearPedido(idPedido) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Pedidos") || ss.getSheetByName("pedidos");
    if (!sheet) return { ok: false, erro: "Aba Pedidos não encontrada" };

    const rows = sheet.getDataRange().getValues();
    const h = rows[0];
    const iId     = h.indexOf("ID Pedido");
    const iStatus = h.indexOf("Status");
    const iNome   = h.indexOf("Nome Cliente");
    const iTotal  = h.indexOf("Total");
    const iItens  = h.indexOf("Itens");
    const iData   = h.indexOf("Data/Hora");

    const row = rows.slice(1).find(r => String(r[iId]).trim().toUpperCase() === idPedido.trim().toUpperCase());
    if (!row) return { ok: false };

    return {
      ok: true,
      pedido: {
        id:     row[iId],
        status: row[iStatus] || "Pendente",
        nome:   iNome >= 0 ? row[iNome] : "",
        total:  iTotal >= 0 ? row[iTotal] : 0,
        itens:  iItens >= 0 ? row[iItens] : "",
        data:   iData >= 0 ? new Date(row[iData]).toLocaleDateString("pt-BR") : "",
      }
    };
  } catch (e) {
    Logger.log("rastrearPedido erro: " + e);
    return { ok: false, erro: String(e) };
  }
}

// ⚠️ INTEGRAR AO doGet DO SEU Code.gs — adicionar ANTES do default/else:
//
//   case "rastrear":
//     result = rastrearPedido(p.id || "");
//     break;
//
// (se seu doGet usa if/else em vez de switch, adicionar antes do return final:)
//   if (params.action === "rastrear") {
//     return jsonResponse(rastrearPedido(params.id || ""));
//   }
//
// Após adicionar: Implantar → Gerenciar → Nova versão → Implantar

// ────────────────────────────────────────────────────────────────
// INTEGRAR AO doPost() EXISTENTE
// No seu doPost, após salvar o pedido na planilha, chamar:
//
//   onNovoPedido(pedidoObj);
//
// Após atualizar status:
//   onStatusChange(pedidoObj, novoStatus);
//
// Após registrar pagamento:
//   onPagamentoRegistrado(pedidoObj, valorPago, dataPag);
//
// Quando admin salva lembretes recorrentes (nova action "criarLembretes"):
//   criarLembretesRecorrentes(pedidoId, nomeCliente, dataBase, nParcelas);
// ────────────────────────────────────────────────────────────────
