// ═══════════════════════════════════════════════════════════════════
//  TAREFAS — Google Apps Script Backend
//  Cole este código em: script.google.com → Novo projeto
//  Depois: Implantar → Nova implantação → Aplicativo da Web
//  Executar como: Eu  |  Acesso: Qualquer pessoa
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Tarefas';
const SHEET_COMMENTS = 'Comentarios';

// ── Cabeçalhos das abas ──────────────────────────────────────────
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sTarefas = ss.getSheetByName(SHEET_NAME);
  if (!sTarefas) {
    sTarefas = ss.insertSheet(SHEET_NAME);
    sTarefas.appendRow(['id', 'text', 'priority', 'status', 'createdAt']);
    sTarefas.getRange(1,1,1,5).setFontWeight('bold');
  }

  let sComents = ss.getSheetByName(SHEET_COMMENTS);
  if (!sComents) {
    sComents = ss.insertSheet(SHEET_COMMENTS);
    sComents.appendRow(['taskId', 'author', 'text', 'time']);
    sComents.getRange(1,1,1,4).setFontWeight('bold');
  }

  return { sTarefas, sComents };
}

// ── Roteador principal ───────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getTasks') return respond(getTasks());
  return respond({ error: 'Ação inválida' });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}

  const action = body.action;

  if (action === 'addTask')      return respond(addTask(body));
  if (action === 'updateStatus') return respond(updateStatus(body));
  if (action === 'deleteTask')   return respond(deleteTask(body));
  if (action === 'addComment')   return respond(addComment(body));

  return respond({ error: 'Ação inválida' });
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET TASKS ────────────────────────────────────────────────────
function getTasks() {
  const { sTarefas, sComents } = ensureSheets();

  const tRows = sTarefas.getDataRange().getValues();
  const cRows = sComents.getDataRange().getValues();

  // Montar mapa de comentários por taskId
  const commentsMap = {};
  for (let i = 1; i < cRows.length; i++) {
    const [taskId, author, text, time] = cRows[i];
    if (!commentsMap[taskId]) commentsMap[taskId] = [];
    commentsMap[taskId].push({ author, text, time });
  }

  const tasks = [];
  for (let i = 1; i < tRows.length; i++) {
    const [id, text, priority, status, createdAt] = tRows[i];
    if (!id) continue;
    tasks.push({
      id: String(id),
      text,
      priority,
      status,
      createdAt,
      comments: commentsMap[String(id)] || []
    });
  }

  // Mais recentes primeiro
  tasks.reverse();

  return { tasks };
}

// ── ADD TASK ─────────────────────────────────────────────────────
function addTask({ text, priority, createdAt }) {
  const { sTarefas } = ensureSheets();
  const id = Date.now().toString();
  sTarefas.appendRow([id, text, priority || 'media', 'pendente', createdAt]);
  return { ok: true, id };
}

// ── UPDATE STATUS ────────────────────────────────────────────────
function updateStatus({ id, status }) {
  const { sTarefas } = ensureSheets();
  const rows = sTarefas.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sTarefas.getRange(i + 1, 4).setValue(status); // coluna D = status
      return { ok: true };
    }
  }
  return { error: 'Tarefa não encontrada' };
}

// ── DELETE TASK ──────────────────────────────────────────────────
function deleteTask({ id }) {
  const { sTarefas, sComents } = ensureSheets();

  // Remove da aba Tarefas
  const tRows = sTarefas.getDataRange().getValues();
  for (let i = 1; i < tRows.length; i++) {
    if (String(tRows[i][0]) === String(id)) {
      sTarefas.deleteRow(i + 1);
      break;
    }
  }

  // Remove comentários relacionados
  const cRows = sComents.getDataRange().getValues();
  for (let i = cRows.length - 1; i >= 1; i--) {
    if (String(cRows[i][0]) === String(id)) {
      sComents.deleteRow(i + 1);
    }
  }

  return { ok: true };
}

// ── ADD COMMENT ──────────────────────────────────────────────────
function addComment({ taskId, author, text, time }) {
  const { sComents } = ensureSheets();
  sComents.appendRow([taskId, author, text, time]);
  return { ok: true };
}
