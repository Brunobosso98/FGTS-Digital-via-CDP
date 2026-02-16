const { chromium } = require('playwright-core');
const fs = require('fs/promises');
const path = require('path');
const xlsx = require('xlsx');

const CDP_URL = process.env.CDP_URL || 'http://127.0.0.1:9222';
const EXCEL_KEYWORD = process.env.EXCEL_KEYWORD || 'Relação folha de pagamento unificada';
const ACTION_TIMEOUT_MS = Number(process.env.ACTION_TIMEOUT_MS || 10000);
const OUTPUT_BASE_DIR =
  process.env.FGTS_OUTPUT_DIR ||
  'C:\\Dpto. Pessoal\\Trabalhista & Previdenciario\\TEMP\\FGTS';
const RUNTIME_BASE_DIR = process.pkg ? path.dirname(process.execPath) : __dirname;

const MONTH_NAMES = {
  '01': 'janeiro',
  '02': 'fevereiro',
  '03': 'marco',
  '04': 'abril',
  '05': 'maio',
  '06': 'junho',
  '07': 'julho',
  '08': 'agosto',
  '09': 'setembro',
  '10': 'outubro',
  '11': 'novembro',
  '12': 'dezembro',
};

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function sanitizeFilePart(value) {
  return String(value || 'sem_nome')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 140);
}

function onlyDigits(value) {
  return String(value || '').replace(/\D/g, '');
}

function monthInfoFromPreviousMonth() {
  const now = new Date();
  const previousMonthDate = new Date(now.getFullYear(), now.getMonth(), 0);
  const year = String(previousMonthDate.getFullYear());
  const month = String(previousMonthDate.getMonth() + 1).padStart(2, '0');

  return {
    year,
    month,
    monthName: MONTH_NAMES[month],
    competencia: `${month}/${year}`,
    sheetName: `${month}.${year}`,
    logDay: `${year}${month}${String(previousMonthDate.getDate()).padStart(2, '0')}`,
  };
}

function pickColumn(columns, predicates) {
  return (
    columns.find((col) => {
      const normalized = normalizeText(col);
      return predicates.every((predicate) => normalized.includes(predicate));
    }) || null
  );
}

async function findExcelFile() {
  const entries = await fs.readdir(RUNTIME_BASE_DIR, { withFileTypes: true });
  const excelName = entries
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .find((name) => name.toLowerCase().endsWith('.xlsx') && name.includes(EXCEL_KEYWORD));

  if (!excelName) {
    throw new Error(`Arquivo Excel com '${EXCEL_KEYWORD}' nao encontrado em ${RUNTIME_BASE_DIR}`);
  }

  return path.join(RUNTIME_BASE_DIR, excelName);
}

async function readCnpjDataFromExcel(monthInfo) {
  const excelPath = await findExcelFile();
  console.log(`Excel selecionado: ${excelPath}`);

  const workbook = xlsx.readFile(excelPath, { cellDates: false });
  const sheet = workbook.Sheets[monthInfo.sheetName];

  if (!sheet) {
    throw new Error(
      `Aba '${monthInfo.sheetName}' nao encontrada no Excel. Abas disponiveis: ${workbook.SheetNames.join(', ')}`
    );
  }

  const rows = xlsx.utils.sheet_to_json(sheet, {
    defval: '',
    raw: false,
    range: 3,
  });

  if (!rows.length) {
    throw new Error(`Aba '${monthInfo.sheetName}' esta vazia (apos cabecalho na linha 4).`);
  }

  const columns = Object.keys(rows[0]);
  const tipoFolhaCol = pickColumn(columns, ['tipo', 'folha']);
  const cnpjCol = pickColumn(columns, ['cnpj']);
  const empresaCol = pickColumn(columns, ['empresa']);
  const analistaCol = pickColumn(columns, ['analista']);

  if (!tipoFolhaCol || !cnpjCol) {
    throw new Error(
      `Colunas obrigatorias nao encontradas. tipoFolhaCol=${tipoFolhaCol}, cnpjCol=${cnpjCol}. Colunas: ${columns.join(', ')}`
    );
  }

  console.log(`Coluna Tipo Folha: ${tipoFolhaCol}`);
  console.log(`Coluna CNPJ: ${cnpjCol}`);
  console.log(`Coluna Empresa: ${empresaCol || 'nao encontrada (fallback CNPJ)'}`);
  console.log(`Coluna Analista: ${analistaCol || 'nao encontrada (fallback Desconhecido)'}`);

  const data = [];
  const seen = new Set();

  for (const row of rows) {
    const tipoFolhaValue = normalizeText(row[tipoFolhaCol]);
    if (!tipoFolhaValue.includes('com funcion')) {
      continue;
    }

    const cnpjRaw = String(row[cnpjCol] || '').trim();
    if (!cnpjRaw || normalizeText(cnpjRaw).includes('procv')) {
      continue;
    }

    const cnpjDigits = onlyDigits(cnpjRaw);
    if (cnpjDigits.length !== 14) {
      continue;
    }

    const analista = String(row[analistaCol] || 'Desconhecido').trim() || 'Desconhecido';
    const empresa = String(row[empresaCol] || cnpjDigits).trim() || cnpjDigits;

    const dedupeKey = `${cnpjDigits}::${analista}`;
    if (seen.has(dedupeKey)) {
      continue;
    }

    seen.add(dedupeKey);
    data.push({ cnpj: cnpjDigits, analista, empresa });
  }

  if (!data.length) {
    throw new Error('Nenhum CNPJ valido encontrado com tipo folha "com funcionario".');
  }

  console.log(`Total de CNPJs para processar: ${data.length}`);
  return data;
}

async function ensureDir(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

async function buildOutputDirs(analista, monthInfo) {
  const analistaSafe = sanitizeFilePart(analista || 'Desconhecido') || 'Desconhecido';
  const analystDir = path.join(OUTPUT_BASE_DIR, monthInfo.year, monthInfo.monthName, analistaSafe);
  const reportDir = path.join(OUTPUT_BASE_DIR, monthInfo.year, monthInfo.monthName, 'RE');

  await ensureDir(analystDir);
  await ensureDir(reportDir);

  return { analystDir, reportDir, analistaSafe };
}

async function appendErrorLog(analystDir, monthInfo, companyData, reason) {
  const now = new Date();
  const logFileName = `erros_download_${now
    .toISOString()
    .slice(0, 10)
    .replace(/-/g, '')}.txt`;
  const logPath = path.join(analystDir, logFileName);
  const timestamp = now.toISOString().replace('T', ' ').slice(0, 19);
  const line = `[${timestamp}] CNPJ: ${companyData.cnpj} | Empresa: ${companyData.empresa} | Analista: ${companyData.analista} | Competencia: ${monthInfo.competencia} | ERRO: ${reason}\n`;
  await fs.appendFile(logPath, line, 'utf8');
}

async function ensureTrocarPerfilVisible(page) {
  const trocarPerfilButton = page.getByRole('button', { name: /Trocar Perfil/i }).first();
  if (await trocarPerfilButton.isVisible().catch(() => false)) {
    return trocarPerfilButton;
  }

  await page.goto('https://fgtsdigital.sistema.gov.br/portal/servicos', {
    waitUntil: 'domcontentloaded',
  });

  await trocarPerfilButton.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  return trocarPerfilButton;
}

async function validateTrocarPerfilResult(page, cnpj) {
  const modal = page.locator('lib-fgtsd-modal-alterar-perfil').first();
  const invalidMessage = modal
    .locator('.invalid-feedback .message')
    .filter({ hasText: /CPF\/CNPJ inv[aá]lido/i })
    .first();

  const deadline = Date.now() + ACTION_TIMEOUT_MS;
  while (Date.now() < deadline) {
    if (await invalidMessage.isVisible().catch(() => false)) {
      throw new Error(`CPF/CNPJ invalido no modal para ${cnpj}`);
    }

    const modalVisible = await modal.isVisible().catch(() => false);
    if (!modalVisible) {
      return;
    }

    await page.waitForTimeout(200);
  }

  throw new Error(`Modal Trocar Perfil nao fechou para ${cnpj} em ${ACTION_TIMEOUT_MS}ms`);
}

async function setProfileAndCnpj(page, cnpj) {
  const trocarPerfilButton = await ensureTrocarPerfilVisible(page);
  await trocarPerfilButton.click();

  const perfilInput = page.locator('ng-select input[role="combobox"]').first();
  await perfilInput.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  await perfilInput.click();
  await perfilInput.fill('Procurador');
  await perfilInput.press('Enter');

  const cnpjCpfInput = page.locator('input[placeholder="Informe CNPJ ou CPF"]').first();
  await cnpjCpfInput.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  await cnpjCpfInput.click();
  await cnpjCpfInput.fill(cnpj);

  const selecionarButton = page.getByRole('button', { name: /^Selecionar$/i }).first();
  await selecionarButton.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  await selecionarButton.click();
  await validateTrocarPerfilResult(page, cnpj);
}

async function openEmissaoRapida(page) {
  const emissaoUrl = 'https://fgtsdigital.sistema.gov.br/cobranca/#/gestao-guias/emissao-guia-rapida';
  const emissaoUrlMatcher = '**/cobranca/#/gestao-guias/emissao-guia-rapida';

  await page.waitForTimeout(200);

  for (let attempt = 1; attempt <= 2; attempt += 1) {
    try {
      await page.goto(emissaoUrl, { waitUntil: 'domcontentloaded' });
      await page.waitForURL(emissaoUrlMatcher, {
        timeout: ACTION_TIMEOUT_MS,
      });
      return;
    } catch (error) {
      if (attempt === 2) {
        break;
      }
    }
  }

  // Fallback para SPA quando o goto nao estabiliza na primeira empresa.
  await page.goto(emissaoUrl, { waitUntil: 'networkidle' });
  await page.waitForURL(emissaoUrlMatcher, {
    timeout: ACTION_TIMEOUT_MS,
  });
}

async function setCompetencia(page, competencia) {
  const competenciaSelect = page.locator('#selectCompetencia ng-select').first();
  await competenciaSelect.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });

  const competenciaAtual = (
    await competenciaSelect.locator('.ng-value-label').first().textContent().catch(() => '')
  )?.trim();

  if (competenciaAtual === competencia) {
    return;
  }

  await competenciaSelect.click();
  const competenciaInput = competenciaSelect.locator('input[role="combobox"]').first();
  await competenciaInput.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  await competenciaInput.fill(competencia);

  const opcaoCompetencia = page.locator('.ng-dropdown-panel .ng-option', { hasText: competencia }).first();
  if (await opcaoCompetencia.isVisible().catch(() => false)) {
    await opcaoCompetencia.click();
  } else {
    await competenciaInput.press('Enter');
  }
}

async function uncheckRescisorio(page) {
  const rescisorioCheckbox = page.locator('input[name="carregarDebitoRescisorio"]').first();
  await rescisorioCheckbox.waitFor({ state: 'attached', timeout: ACTION_TIMEOUT_MS });
  const visible = await rescisorioCheckbox.isVisible().catch(() => false);
  if (!visible) {
    return;
  }

  const disabled = await rescisorioCheckbox.isDisabled().catch(() => true);
  const checked = await rescisorioCheckbox.isChecked().catch(() => false);
  if (!disabled && checked) {
    await rescisorioCheckbox.uncheck({ force: true });
  }
}

async function makeUniquePath(filePath) {
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);

  let attempt = 0;
  let candidate = filePath;

  while (true) {
    try {
      await fs.access(candidate);
      attempt += 1;
      candidate = path.join(dir, `${base}_${attempt}${ext}`);
    } catch {
      return candidate;
    }
  }
}

async function captureDownload(page, clickAction, targetDir, preferredNameWithoutExt) {
  await ensureDir(targetDir);

  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: ACTION_TIMEOUT_MS }),
    clickAction(),
  ]);

  const suggested = download.suggestedFilename() || 'arquivo.pdf';
  const suggestedExt = path.extname(suggested).toLowerCase() || '.pdf';
  const baseName = sanitizeFilePart(preferredNameWithoutExt) || 'arquivo';
  const initialPath = path.join(targetDir, `${baseName}${suggestedExt === '.pdf' ? '.pdf' : suggestedExt}`);
  const finalPath = await makeUniquePath(initialPath);

  await download.saveAs(finalPath);
  return finalPath;
}

async function processCompany(page, companyData, monthInfo) {
  const dirs = await buildOutputDirs(companyData.analista, monthInfo);
  const companySafe = sanitizeFilePart(companyData.empresa || companyData.cnpj);

  console.log(`\\nProcessando CNPJ ${companyData.cnpj} | Analista: ${companyData.analista} | Empresa: ${companyData.empresa}`);

  await setProfileAndCnpj(page, companyData.cnpj);
  await openEmissaoRapida(page);
  await setCompetencia(page, monthInfo.competencia);
  await uncheckRescisorio(page);

  const pesquisarButton = page.getByRole('button', { name: /^Pesquisar$/i }).first();
  await pesquisarButton.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  await pesquisarButton.click();

  const emitirGuiaButton = page.getByRole('button', { name: /Emitir guia/i }).first();
  await emitirGuiaButton.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
  const guiaPath = await captureDownload(
    page,
    () => emitirGuiaButton.click(),
    dirs.analystDir,
    companySafe
  );

  const imprimirPdfButton = page
    .locator('button[aria-label="Imprimir relatório em PDF"], button[aria-label="Imprimir relatorio em PDF"]')
    .first();

  let clickPrint;
  if (await imprimirPdfButton.isVisible().catch(() => false)) {
    clickPrint = () => imprimirPdfButton.click();
  } else {
    const imprimirPdfXpath = page.locator(
      'xpath=/html/body/app-root/fgtsd-main-layout/div/br-main-layout/div/div/div/main/div[2]/app-emissao-guia-rapida-consignado/div/app-agrupamento-guia-rapida-consignado/div/div/div/div/div[3]/div/button[1]'
    );
    await imprimirPdfXpath.waitFor({ state: 'visible', timeout: ACTION_TIMEOUT_MS });
    clickPrint = () => imprimirPdfXpath.click();
  }

  const reportPath = await captureDownload(page, clickPrint, dirs.reportDir, companySafe);

  console.log(`Guia salva: ${guiaPath}`);
  console.log(`Relatorio salvo: ${reportPath}`);
}

async function main() {
  const monthInfo = monthInfoFromPreviousMonth();
  console.log(`Competencia alvo: ${monthInfo.competencia} | Aba Excel: ${monthInfo.sheetName}`);
  console.log(`Diretorio base de saida: ${OUTPUT_BASE_DIR}`);

  const companies = await readCnpjDataFromExcel(monthInfo);

  const browser = await chromium.connectOverCDP(CDP_URL);
  const context = browser.contexts()[0];
  if (!context) {
    throw new Error('Nenhum contexto encontrado via CDP. Verifique se o Chrome foi aberto com --remote-debugging-port=9222');
  }

  const page = context.pages()[0] || (await context.newPage());
  await page.bringToFront();
  await page.waitForLoadState('domcontentloaded');

  let successCount = 0;
  let failCount = 0;

  for (let i = 0; i < companies.length; i += 1) {
    const company = companies[i];
    console.log(`\\n[${i + 1}/${companies.length}] Iniciando...`);

    let processed = false;
    let lastError = null;

    for (let attempt = 1; attempt <= 2; attempt += 1) {
      try {
        console.log(`Tentativa ${attempt}/2 para CNPJ ${company.cnpj}`);
        await processCompany(page, company, monthInfo);
        successCount += 1;
        processed = true;
        break;
      } catch (error) {
        lastError = error;
        console.error(`Tentativa ${attempt}/2 falhou para ${company.cnpj}: ${error.message}`);
        await page.goto('https://fgtsdigital.sistema.gov.br/portal/servicos', {
          waitUntil: 'domcontentloaded',
        }).catch(() => {});
      }
    }

    if (!processed) {
      failCount += 1;
      try {
        const dirs = await buildOutputDirs(company.analista, monthInfo);
        await appendErrorLog(
          dirs.analystDir,
          monthInfo,
          company,
          lastError?.message || 'Timeout/Falha em duas tentativas'
        );
      } catch (logError) {
        console.error(`Falha ao registrar log de erro para ${company.cnpj}: ${logError.message}`);
      }
      console.log(`CNPJ ${company.cnpj} pulado apos 2 tentativas.`);
    }
  }

  console.log('\\nResumo da execucao:');
  console.log(`Total: ${companies.length}`);
  console.log(`Sucesso: ${successCount}`);
  console.log(`Falhas: ${failCount}`);
}

main().catch((error) => {
  console.error(`Erro fatal: ${error.message}`);
  process.exitCode = 1;
});

