const APP_VERSION = '5.24';

    // Auto version check & Cache Busting
    (function checkAppVersion() {
      const storedVersion = localStorage.getItem('gerador_sql_version');
      if (storedVersion !== APP_VERSION) {
        localStorage.setItem('gerador_sql_version', APP_VERSION);
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get('v') !== APP_VERSION) {
          urlParams.set('v', APP_VERSION);
          window.location.search = urlParams.toString();
        }
      }
    })();

    // ============================================
    // Controlador da interface do Gerador de UPDATE SQL - v5.24
    // ============================================

    let excelData = [];
    let headers = [];
    let currentTableName = 'produtos';
    let generatedSQLs = [];
    let isGenerating = false;
    let originalFileName = '';
    let currentWorkbook = null;
    let currentFile = null;

    let planilhaStats = {
      totalRegistros: 0,
      camposEncontrados: [],
      temClasstrib: false,
      temIbscbs: false,
      temClasstribIbscbs: false,
      camposAusentes: []
    };

    // ============================================
    // MAPEAMENTO DE ALÍQUOTAS
    // ============================================
    // Regra de transição de 2026 da LC 214/2025, arts. 343 e 346.
    // A alíquota municipal do IBS só passa a ser destacada em 2027.
    const ALIQUOTAS_TRANSICAO_2026 = Object.freeze({
      ALIQ_CBS: '0.9',
      ALIQ_IBS_UF: '0.1',
      ALIQ_IBS_MUN: '0'
    });

    const ALIQUOTAS_MAP = {
      '000|000001': {
        ALIQ_CBS: '0.9',
        ALIQ_EFETIVA_CBS: '0.9',
        REDUCAO_ALIQ_CBS: '0',
        ALIQ_IBS_UF: '0.1',
        ALIQ_EFETIVA_IBS_UF: '0.1',
        REDUCAO_ALIQ_IBS_UF: '0',
        ALIQ_IBS_MUN: '0',
        ALIQ_EFETIVA_IBS_MUN: '0',
        REDUCAO_ALIQ_IBS_MUN: '0'
      },
      '200|200047': {
        ALIQ_CBS: '0.9',
        ALIQ_EFETIVA_CBS: '0.54',
        REDUCAO_ALIQ_CBS: '40',
        ALIQ_IBS_UF: '0.1',
        ALIQ_EFETIVA_IBS_UF: '0.06',
        REDUCAO_ALIQ_IBS_UF: '40',
        ALIQ_IBS_MUN: '0',
        ALIQ_EFETIVA_IBS_MUN: '0',
        REDUCAO_ALIQ_IBS_MUN: '40'
      },
      '410|410019': {
        ALIQ_CBS: '0',
        ALIQ_EFETIVA_CBS: '0',
        REDUCAO_ALIQ_CBS: '0',
        ALIQ_IBS_UF: '0',
        ALIQ_EFETIVA_IBS_UF: '0',
        REDUCAO_ALIQ_IBS_UF: '0',
        ALIQ_IBS_MUN: '0',
        ALIQ_EFETIVA_IBS_MUN: '0',
        REDUCAO_ALIQ_IBS_MUN: '0',
        CODNATOPERACAOSAI: '5102'
      },
      '410|410020': {
        ALIQ_CBS: '0',
        ALIQ_EFETIVA_CBS: '0',
        REDUCAO_ALIQ_CBS: '0',
        ALIQ_IBS_UF: '0.1',
        ALIQ_EFETIVA_IBS_UF: '0.1',
        REDUCAO_ALIQ_IBS_UF: '0',
        ALIQ_IBS_MUN: '0',
        ALIQ_EFETIVA_IBS_MUN: '0',
        REDUCAO_ALIQ_IBS_MUN: '0'
      },
      '410|410999': {
        ALIQ_CBS: '0',
        ALIQ_EFETIVA_CBS: '0',
        REDUCAO_ALIQ_CBS: '0',
        ALIQ_IBS_UF: '0',
        ALIQ_EFETIVA_IBS_UF: '0',
        REDUCAO_ALIQ_IBS_UF: '0',
        ALIQ_IBS_MUN: '0',
        ALIQ_EFETIVA_IBS_MUN: '0',
        REDUCAO_ALIQ_IBS_MUN: '0'
      }
    };

    // Dados oficiais da SEFAZ: a tabela pode ser carregada ao lado deste arquivo
    // (classTrib.json) ou manualmente pela interface quando o gerador é aberto localmente.
    const classTribService = new ClassTribService(normalizarValorCodcst);
    let classTribInfo = null;

    function atualizarStatusClassTrib(mensagem, carregada = false) {
      const el = document.getElementById('classTribStatus');
      if (!el) return;
      el.textContent = mensagem;
      el.style.color = carregada ? 'var(--success)' : 'var(--text-muted)';
    }

    function carregarTabelaClassTrib(dados) {
      const info = classTribService.load(dados);
      classTribInfo = { total: info.total, publicacao: info.publication };
      const dataPublicacao = info.publication
        ? new Date(info.publication).toLocaleDateString('pt-BR')
        : 'não informada';
      atualizarStatusClassTrib(`✅ ${info.total} classificações SEFAZ carregadas (publicação ${dataPublicacao})`, true);
      if (excelData.length) analisarPlanilha();
    }

    function obterClassificacaoOficial(ibscbs, classtrib) {
      return classTribService.find(ibscbs, classtrib);
    }

    function classificacaoEstaVigente(classificacao, hoje = new Date()) {
      return classTribService.isCurrent(classificacao, hoje);
    }

    function obterReducaoOficial(classificacao) {
      return classTribService.getReductions(classificacao);
    }

    function classificacaoSemAliquota(classificacao) {
      if (!classificacao) return false;
      if (String(classificacao.TipoAliquota || '').includes('Sem Alíquota')) return true;
      const cst = classificacao.CST;
      // O catálogo embarcado é compilado de classTrib.json e não inclui TipoAliquota.
      // Estes grupos e exceções são as classificações "3 - Sem Alíquota" dessa fonte.
      if (['410', '510', '550', '800', '810', '811', '820'].includes(cst)) return true;
      return ['400|400002', '620|620007'].includes(`${cst}|${classificacao.cClassTrib}`);
    }

    const CAMPOS_ALIQUOTA = [
      'ALIQ_CBS',
      'ALIQ_EFETIVA_CBS',
      'REDUCAO_ALIQ_CBS',
      'ALIQ_IBS_UF',
      'ALIQ_EFETIVA_IBS_UF',
      'REDUCAO_ALIQ_IBS_UF',
      'ALIQ_IBS_MUN',
      'ALIQ_IBS_MU',
      'ALIQ_EFETIVA_IBS_MUN',
      'ALIQ_EFETIVA_IBS_MU',
      'REDUCAO_ALIQ_IBS_MUN',
      'REDUCAO_ALIQ_IBS_MU'
    ];

    function ehCampoAliquota(nomeCampo) {
      if (!nomeCampo) return false;
      const upper = nomeCampo.toUpperCase();
      return CAMPOS_ALIQUOTA.includes(upper);
    }

    function ehCampoSyncIbscbs(nomeCampo) {
      if (!nomeCampo) return false;
      const upper = nomeCampo.toUpperCase();
      return ehAliasIbscbs(upper) || ehAliasClasstrib(upper) || ehCampoAliquota(upper);
    }

    function obterValorAliquota(aliquotasObj, nomeCampo) {
      if (!aliquotasObj || !nomeCampo) return undefined;
      const upper = nomeCampo.toUpperCase();
      if (aliquotasObj[upper] !== undefined) return aliquotasObj[upper];

      if (upper === 'REDUCAO_ALIQ_IBS_MU' && aliquotasObj['REDUCAO_ALIQ_IBS_MUN'] !== undefined) {
        return aliquotasObj['REDUCAO_ALIQ_IBS_MUN'];
      }
      if (upper === 'ALIQ_IBS_MU' && aliquotasObj['ALIQ_IBS_MUN'] !== undefined) {
        return aliquotasObj['ALIQ_IBS_MUN'];
      }
      if (upper === 'ALIQ_EFETIVA_IBS_MU' && aliquotasObj['ALIQ_EFETIVA_IBS_MUN'] !== undefined) {
        return aliquotasObj['ALIQ_EFETIVA_IBS_MUN'];
      }
      return undefined;
    }

    // ============================================
    // FUNÇÕES AUXILIARES E SANITIZAÇÃO
    // ============================================
    function cleanString(val) {
      if (val === undefined || val === null) return '';
      return String(val).replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, ' ').trim();
    }

    function formatarCST(codigo) {
      if (codigo === undefined || codigo === null || codigo === '') return '';
      let strCodigo = cleanString(codigo).replace(/^['"]|['"]$/g, '');
      if (strCodigo === '') return '';
      const mapeamentoCST = {
        '0': '000', '00': '000', '000': '000',
        '20': '020', '020': '020',
        '40': '040', '040': '040',
        '41': '041', '041': '041',
        '60': '060', '060': '060'
      };
      if (mapeamentoCST[strCodigo]) return mapeamentoCST[strCodigo];
      if (/^\d+$/.test(strCodigo)) {
        if (strCodigo.length === 1) return '00' + strCodigo;
        if (strCodigo.length === 2) return '0' + strCodigo;
        if (strCodigo.length === 3) return strCodigo;
        if (strCodigo.length > 3) return strCodigo.substring(0, 3);
      }
      return strCodigo;
    }

    function deveFormatarCST(nomeCampo) {
      return nomeCampo === 'CST' || nomeCampo.toUpperCase() === 'CST';
    }

    const ALIASES_CODCST_IBSCBS = [
      'CODCST_IBSCBS',
      'CODCST_IBS_CBS',
      'CODCSTIBSCBS',
      'CODCST_IBS',
      'CST_IBSCBS',
      'CST_IBS_CBS',
      'CST_IBS',
      'CSTIBSCBS',
      'IBSCBS',
      'IBS_CBS'
    ];

    const ALIASES_CODCST_CLASSTRIB = [
      'CODCST_CLASSTRIB',
      'CODCST_CLASS_TRIB',
      'CODCST_CLASSTRIB_IBS',
      'COD_CLASSTRIB',
      'CLASSTRIB_IBSCBS',
      'CLASSTRIB_IBS_CBS',
      'CLASSTRIB',
      'CODCLASSTRIB',
      'CLASS_TRIB',
      'CLASSTRIB_IBS'
    ];

    // This field is distinct from CODCST_CLASSTRIB. Keep its aliases isolated
    // so a spreadsheet column for one tax field can never populate the other.
    const ALIASES_CODCST_CLASSTRIB_IBSCBS = [
      'CODCST_CLASSTRIB_IBSCBS',
      'CODCST_CLASSTRIB_IBS_CBS'
    ];

    function ehAliasIbscbs(nomeCampo) {
      if (!nomeCampo) return false;
      return ALIASES_CODCST_IBSCBS.includes(nomeCampo.toUpperCase());
    }

    function ehAliasClasstrib(nomeCampo) {
      if (!nomeCampo) return false;
      return ALIASES_CODCST_CLASSTRIB.includes(nomeCampo.toUpperCase());
    }

    function ehAliasClasstribIbscbs(nomeCampo) {
      if (!nomeCampo) return false;
      return ALIASES_CODCST_CLASSTRIB_IBSCBS.includes(nomeCampo.toUpperCase());
    }

    function campoExisteNaPlanilha(nomeCampo) {
      if (!nomeCampo) return false;
      const upper = nomeCampo.toUpperCase();
      if (headers.some(h => h.toUpperCase() === upper)) return true;
      if (ehAliasIbscbs(upper)) return headers.some(h => ehAliasIbscbs(h));
      if (ehAliasClasstribIbscbs(upper)) return headers.some(h => ehAliasClasstribIbscbs(h));
      if (ehAliasClasstrib(upper)) return headers.some(h => ehAliasClasstrib(h));
      return false;
    }

    function normalizarValorCodcst(nomeCampo, valor) {
      if (valor === undefined || valor === null || valor === '') return valor;
      const campo = nomeCampo.toUpperCase();
      let tamanho = 0;
      if (ehAliasIbscbs(campo)) tamanho = 3;
      else if (ehAliasClasstribIbscbs(campo) || ehAliasClasstrib(campo)) tamanho = 6;
      if (tamanho === 0) return valor;
      let str = cleanString(valor).replace(/^['"]|['"]$/g, '');
      if (/^\d+\.0+$/.test(str)) str = str.replace(/\.0+$/, '');
      if (/^\d+$/.test(str)) return str.padStart(tamanho, '0');
      return str;
    }

    function getValorCampo(row, campoNome) {
      if (!campoNome) return undefined;
      const upper = campoNome.toUpperCase();

      let colIndex = headers.findIndex(h => h.toUpperCase() === upper);

      if (colIndex === -1) {
        if (ehAliasIbscbs(upper)) {
          colIndex = headers.findIndex(h => ehAliasIbscbs(h));
        } else if (ehAliasClasstribIbscbs(upper)) {
          colIndex = headers.findIndex(h => ehAliasClasstribIbscbs(h));
        } else if (ehAliasClasstrib(upper)) {
          colIndex = headers.findIndex(h => ehAliasClasstrib(h));
        }
      }

      if (colIndex === -1) return undefined;
      let value = row[colIndex];
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return normalizarValorCodcst(campoNome, String(value).trim());
      }
      return undefined;
    }

    function getAliquotasPorCombinacao(ibscbs, classtrib) {
      const ibs = normalizarValorCodcst('CODCST_IBSCBS', ibscbs) || '';
      const cls = normalizarValorCodcst('CODCST_CLASSTRIB_IBSCBS', classtrib) || '';
      const key = `${ibs}|${cls}`;
      if (ALIQUOTAS_MAP[key]) {
        return ALIQUOTAS_MAP[key];
      }
      return null;
    }

    function formatValue(value) {
      if (value === undefined || value === null || value === '') return 'NULL';
      let strValue = cleanString(value);
      if (strValue === '') return 'NULL';
      return `'${strValue.replace(/'/g, "''")}'`;
    }

    function escapeIdentifier(identifier) {
      if (!identifier) return '';
      if (/[^a-zA-Z0-9_]/.test(identifier)) {
        return `"${identifier.replace(/"/g, '""')}"`;
      }
      return identifier;
    }

    function escapeHtml(str) {
      if (!str) return '';
      return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    // ============================================
    // ANÁLISE DA PLANILHA
    // ============================================
    function analisarPlanilha() {
      planilhaStats.totalRegistros = excelData.length;
      planilhaStats.camposEncontrados = [...headers];
      planilhaStats.temIbscbs = headers.some(h => ehAliasIbscbs(h));
      planilhaStats.temClasstrib = headers.some(h => ehAliasClasstrib(h));
      planilhaStats.temClasstribIbscbs = headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB_IBSCBS' || h.toUpperCase() === 'CODCST_CLASSTRIB_IBS_CBS');
      planilhaStats.camposAusentes = CAMPOS_ALIQUOTA.filter(campo => !campoExisteNaPlanilha(campo));

      const nomeIbscbs = headers.find(h => ehAliasIbscbs(h)) || 'CODCST_IBSCBS';
      const nomeClasstrib = headers.find(h => ehAliasClasstrib(h)) || 'CODCST_CLASSTRIB';
      let resumoOficial = '• Tabela SEFAZ: <span style="color: #64748b;">não carregada — não haverá validação oficial.</span>';
      if (classTribInfo && planilhaStats.temIbscbs && (planilhaStats.temClasstrib || planilhaStats.temClasstribIbscbs)) {
        let validas = 0;
        let invalidas = 0;
        let foraVigencia = 0;
        excelData.forEach(row => {
          const cst = getValorCampo(row, 'CODCST_IBSCBS');
          const classTrib = getValorCampo(row, 'CODCST_CLASSTRIB_IBSCBS') || getValorCampo(row, 'CODCST_CLASSTRIB');
          if (!cst && !classTrib) return;
          const classificacao = obterClassificacaoOficial(
            cst,
            classTrib
          );
          if (!classificacao) invalidas++;
          else if (!classificacaoEstaVigente(classificacao)) foraVigencia++;
          else validas++;
        });
        resumoOficial = `• Validação SEFAZ: <span>✅ ${validas} vigentes</span> · <span style="color: #b45309;">⚠️ ${foraVigencia} fora da vigência</span> · <span style="color: #dc2626;">❌ ${invalidas} não encontradas</span>`;
      } else if (classTribInfo) {
        resumoOficial = `• Tabela SEFAZ: <span>✅ ${classTribInfo.total} classificações carregadas</span> (inclua CST e CLASSTRIB para validar).`;
      }

      const statsPanel = document.getElementById('statsPanel');
      statsPanel.style.display = 'block';
      statsPanel.innerHTML = `
      📊 <strong>Análise da planilha:</strong><br>
      • Total de registros: <span>${planilhaStats.totalRegistros}</span><br>
      • Campos encontrados: <span>${planilhaStats.camposEncontrados.length}</span><br>
      • CST IBSCBS presente: <span>${planilhaStats.temIbscbs ? `✅ Sim (${escapeHtml(nomeIbscbs)})` : '❌ Não'}</span><br>
      • CLASSTRIB presente: <span>${planilhaStats.temClasstrib ? `✅ Sim (${escapeHtml(nomeClasstrib)})` : '❌ Não'}</span><br>
      ${resumoOficial}<br>
      ${planilhaStats.camposAusentes.length > 0 ? `• Campos de alíquota não presentes na planilha (não serão incluídos): <span style="color: #64748b;">${planilhaStats.camposAusentes.length}</span>` : '• Todos os campos de alíquota estão presentes ✅'}
    `;

      const syncToggle = document.getElementById('enableAutoFieldsToggle');
      if (planilhaStats.temIbscbs) {
        syncToggle.disabled = false;
      } else {
        syncToggle.disabled = true;
        syncToggle.checked = false;
      }
    }

    // ============================================
    // RENDERIZAÇÃO DOS CAMPOS
    // ============================================
    function renderFieldSelectors() {
      const syncEnabled = document.getElementById('enableAutoFieldsToggle').checked;

      setFieldsDiv.innerHTML = '';
      if (headers.length === 0) {
        setFieldsDiv.innerHTML = `
        <div class="empty-fields-message">
          <span>📂</span>
          Nenhum campo disponível<br>
          <small style="color: #94a3b8;">Carregue a planilha para ver os campos</small>
        </div>
      `;
        return;
      }

      headers.forEach(header => {
        const div = document.createElement('div');
        div.className = 'field-checkbox';
        let labelHtml = escapeHtml(header);

        if (syncEnabled && ehCampoAliquota(header)) {
          labelHtml += ` <span class="auto-badge">auto</span>`;
        }

        div.innerHTML = `
        <input type="checkbox" class="set-checkbox" value="${escapeHtml(header)}" id="set_${escapeHtml(header)}">
        <label for="set_${escapeHtml(header)}">${labelHtml}</label>
      `;
        setFieldsDiv.appendChild(div);
      });

      whereFieldsDiv.innerHTML = '';
      headers.forEach(header => {
        const div = document.createElement('div');
        div.className = 'field-checkbox';
        div.innerHTML = `
        <input type="checkbox" class="where-checkbox" value="${escapeHtml(header)}" id="where_${escapeHtml(header)}">
        <label for="where_${escapeHtml(header)}">${escapeHtml(header)}</label>
      `;
        whereFieldsDiv.appendChild(div);
      });

      const EXCLUDED_SET_FIELDS = ['CODEMP', 'CAT_COD', 'CAT_DESCR', 'PROD_COD', 'PROD_DESCR'];
      const defaultWhereFields = ['CODEMP', 'CAT_COD', 'PROD_COD'];

      document.querySelectorAll('.where-checkbox').forEach(cb => {
        cb.checked = defaultWhereFields.includes(cb.value.toUpperCase());
      });

      if (syncEnabled) {
        document.querySelectorAll('.set-checkbox').forEach(cb => {
          cb.checked = ehCampoSyncIbscbs(cb.value);
        });
      } else {
        document.querySelectorAll('.set-checkbox').forEach(cb => {
          cb.checked = !EXCLUDED_SET_FIELDS.includes(cb.value.toUpperCase());
        });
      }

      updateSelectAllStates();
      filterFields('setFields', 'searchSetInput');
      filterFields('whereFields', 'searchWhereInput');
    }

    function updateSelectAllStates() {
      const EXCLUDED_SET_FIELDS = ['CODEMP', 'CAT_COD', 'CAT_DESCR', 'PROD_COD', 'PROD_DESCR'];
      const setCheckboxes = Array.from(document.querySelectorAll('.set-checkbox'));
      const whereCheckboxes = Array.from(document.querySelectorAll('.where-checkbox'));

      const eligibleSetCheckboxes = setCheckboxes.filter(cb => !EXCLUDED_SET_FIELDS.includes(cb.value.toUpperCase()));
      if (eligibleSetCheckboxes.length) {
        selectAllSet.checked = eligibleSetCheckboxes.every(cb => cb.checked);
      }
      if (whereCheckboxes.length) {
        selectAllWhere.checked = whereCheckboxes.every(cb => cb.checked);
      }
    }

    function filterFields(gridId, searchInputId) {
      const input = document.getElementById(searchInputId);
      const grid = document.getElementById(gridId);
      if (!input || !grid) return;
      const term = input.value.trim().toLowerCase();
      const items = grid.querySelectorAll('.field-checkbox');
      items.forEach(item => {
        const text = item.textContent.toLowerCase();
        if (!term || text.includes(term)) {
          item.style.display = 'flex';
        } else {
          item.style.display = 'none';
        }
      });
    }

    function getSelectedSetFields() {
      const selected = [];
      document.querySelectorAll('.set-checkbox:checked').forEach(cb => selected.push(cb.value));
      return selected;
    }

    function getSelectedWhereFields() {
      const selected = [];
      document.querySelectorAll('.where-checkbox:checked').forEach(cb => selected.push(cb.value));
      return selected;
    }

    // ============================================
    // PROCESSAMENTO DO ARQUIVO
    // ============================================
    function processFile(file) {
      console.log('Processando arquivo:', file.name);
      showStatus('info', '📂 Processando arquivo...');
      currentFile = file;

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          currentWorkbook = XLSX.read(data, { type: 'array' });
          const sheetNames = currentWorkbook.SheetNames;

          if (sheetNames.length === 0) {
            showStatus('error', '❌ Nenhuma aba encontrada no arquivo.');
            return;
          }

          if (sheetNames.length === 1) {
            processSheet(sheetNames[0]);
            return;
          }

          showSheetSelector(sheetNames);

        } catch (error) {
          console.error('Erro:', error);
          showStatus('error', '❌ Erro ao processar o arquivo.');
        }
      };

      reader.onerror = function () {
        showStatus('error', '❌ Erro ao ler o arquivo');
      };

      reader.readAsArrayBuffer(file);
    }

    function showSheetSelector(sheetNames) {
      const sheetList = document.getElementById('sheetList');
      sheetList.innerHTML = '';

      sheetNames.forEach((name, index) => {
        const sheet = currentWorkbook.Sheets[name];
        const jsonPreview = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
        const rowCount = Math.max(0, jsonPreview.length - 1);

        const item = document.createElement('div');
        item.className = `sheet-item${index === 0 ? ' selected' : ''}`;
        item.innerHTML = `
          <input type="radio" name="sheetSelect" value="${escapeHtml(name)}" id="sheet_${index}" ${index === 0 ? 'checked' : ''}>
          <label for="sheet_${index}">📄 ${escapeHtml(name)}</label>
          <span class="sheet-rows-badge">${rowCount} registro${rowCount !== 1 ? 's' : ''}</span>
        `;

        item.addEventListener('click', () => {
          sheetList.querySelectorAll('.sheet-item').forEach(el => el.classList.remove('selected'));
          item.classList.add('selected');
          item.querySelector('input[type="radio"]').checked = true;
        });

        sheetList.appendChild(item);
      });

      document.getElementById('sheetModalOverlay').classList.add('active');
      showStatus('info', `📑 Arquivo possui ${sheetNames.length} abas. Selecione qual deseja usar.`);
    }

    function processSheet(sheetName) {
      try {
        document.getElementById('sheetModalOverlay').classList.remove('active');

        const sheet = currentWorkbook.Sheets[sheetName];
        if (!sheet) {
          showStatus('error', `❌ Aba "${sheetName}" não encontrada.`);
          return;
        }

        const jsonData = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
          raw: false
        });

        if (jsonData.length < 2) {
          showStatus('error', 'A planilha precisa ter pelo menos cabeçalho e uma linha de dados');
          return;
        }

        let rawHeaders = jsonData[0].map(h => String(h || '').trim());
        const dataRows = jsonData.slice(1);

        const primeiroCabecalhoVazio = rawHeaders.length > 0 && rawHeaders[0] === '';
        const temDadosPrimeiraColuna = dataRows.some(r => r[0] !== undefined && r[0] !== null && String(r[0]).trim() !== '');
        const segundoEhCampoConhecido = rawHeaders.slice(1).some(h => {
          const u = h.toUpperCase();
          return ['CODEMP', 'CAT_COD', 'PROD_COD', 'CAT_DESCR', 'PROD_DESCR', 'ALIQUOTA', 'CST', 'NCM'].includes(u);
        });

        let realinhouCabecalho = false;
        if (primeiroCabecalhoVazio && temDadosPrimeiraColuna && segundoEhCampoConhecido) {
          console.warn('⚠️ Detectado A1 em branco e cabeçalhos deslocados. Realinhando automaticamente...');
          rawHeaders.shift();
          realinhouCabecalho = true;
        }

        const validRawHeaders = rawHeaders.filter(h => h !== '');

        const headerMap = {};
        headers = [];
        validRawHeaders.forEach(h => {
          if (!headerMap[h.toUpperCase()]) {
            headerMap[h.toUpperCase()] = true;
            headers.push(h);
          }
        });

        excelData = [];
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          const hasData = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '');
          if (hasData) {
            const processedRow = [];
            for (let j = 0; j < headers.length; j++) {
              let colIndex = rawHeaders.findIndex(h => h.toUpperCase() === headers[j].toUpperCase());
              if (colIndex === -1) {
                processedRow.push('');
              } else {
                let value = row[colIndex];
                processedRow.push(value !== undefined && value !== null ? cleanString(value) : '');
              }
            }
            excelData.push(processedRow);
          }
        }

        if (headers.length === 0) {
          showStatus('error', 'Nenhum cabeçalho encontrado na planilha');
          return;
        }

        if (realinhouCabecalho) {
          showStatus('warning', '⚡ A célula A1 da planilha estava em branco. Os cabeçalhos foram realinhados automaticamente com a 1ª coluna de dados!');
        }

        if (originalFileName) {
          const suggestedName = originalFileName.replace(/[^a-zA-Z0-9_]/g, '_').toLowerCase();
          if (suggestedName) {
            tableNameInput.value = suggestedName;
            currentTableName = suggestedName;
          }
        }

        const sheetLabel = currentWorkbook.SheetNames.length > 1 ? ` | Aba: <strong>${escapeHtml(sheetName)}</strong>` : '';
        tableInfoDiv.innerHTML = `📊 Arquivo: <strong>${escapeHtml(currentFile.name)}</strong>${sheetLabel} | ${excelData.length} registros | ${headers.length} campos<br><span style="font-size:0.75rem;">✅ Todas as células foram lidas como TEXTO</span>`;
        tableInfoDiv.style.display = 'block';

        analisarPlanilha();
        renderFieldSelectors();

        showStatus('success', `✅ Arquivo carregado! Aba "${sheetName}" — ${excelData.length} registros encontrados.`);
        generatedSQLs = [];
        sqlOutput.innerHTML = '-- Nenhum SQL gerado ainda --';
        sqlCounter.textContent = '0 comandos gerados';

      } catch (error) {
        console.error('Erro:', error);
        showStatus('error', '❌ Erro ao processar a aba.');
      }
    }

    // ============================================
    // GERAÇÃO DE SQL (ASSÍNCRONA EM LOTES)
    // ============================================
    async function generateSQL() {
      if (isGenerating) {
        showStatus('warning', '⏳ Aguarde, já está gerando...');
        return;
      }

      const nomeTabela = tableNameInput.value.trim();
      if (nomeTabela === '') {
        showStatus('error', '❌ O nome da tabela não pode estar vazio.');
        tableNameInput.classList.add('error-input');
        tableNameInput.focus();
        return;
      }

      currentTableName = nomeTabela;
      tableNameInput.classList.remove('error-input');

      const whereFields = getSelectedWhereFields();
      let setFields = getSelectedSetFields();

      if (whereFields.length === 0) {
        showStatus('error', '❌ Selecione pelo menos um campo para o WHERE');
        return;
      }

      if (excelData.length === 0) {
        showStatus('error', '❌ Nenhum dado encontrado.');
        return;
      }

      const camposInexistentes = [];
      for (const field of [...setFields, ...whereFields]) {
        if (!campoExisteNaPlanilha(field)) {
          camposInexistentes.push(field);
        }
      }

      if (camposInexistentes.length > 0) {
        showStatus('error', `❌ Campos NÃO existem na planilha: ${camposInexistentes.join(', ')}`);
        return;
      }

      const syncEnabled = document.getElementById('enableAutoFieldsToggle').checked;
      const temIbscbs = planilhaStats.temIbscbs;
      let classificacoesInvalidas = 0;
      let classificacoesForaVigencia = 0;

      if (classTribInfo && temIbscbs && (planilhaStats.temClasstrib || planilhaStats.temClasstribIbscbs)) {
        excelData.forEach(row => {
          const cst = getValorCampo(row, 'CODCST_IBSCBS');
          const classTrib = getValorCampo(row, 'CODCST_CLASSTRIB_IBSCBS') || getValorCampo(row, 'CODCST_CLASSTRIB');
          if (!cst && !classTrib) return;
          const classificacao = obterClassificacaoOficial(
            cst,
            classTrib
          );
          if (!classificacao) classificacoesInvalidas++;
          else if (!classificacaoEstaVigente(classificacao)) classificacoesForaVigencia++;
        });
        if (classificacoesInvalidas > 0 || classificacoesForaVigencia > 0) {
          const aviso = `A tabela SEFAZ identificou ${classificacoesInvalidas} combinação(ões) inexistente(s) e ${classificacoesForaVigencia} fora da vigência. Deseja gerar os SQLs mesmo assim?`;
          if (!window.confirm(aviso)) {
            showStatus('warning', '⚠️ Geração cancelada para correção das classificações tributárias.');
            return;
          }
        }
      }

      isGenerating = true;
      generateBtn.disabled = true;
      generateBtn.innerHTML = '⏳ Gerando SQLs...';

      const progressContainer = document.getElementById('progressContainer');
      const progressBar = document.getElementById('progressBar');
      const progressText = document.getElementById('progressText');
      const progressPercent = document.getElementById('progressPercent');

      progressContainer.style.display = 'block';
      progressBar.style.width = '0%';
      progressText.textContent = `Gerando SQLs... (0/${excelData.length})`;
      progressPercent.textContent = '0%';

      try {
        const sqlCommands = [];
        let cstFormatCount = 0;
        let syncCount = 0;
        let autoFillCount = 0;
        let registrosIgnorados = 0;

        const totalRows = excelData.length;
        const chunkSize = 500;

        for (let i = 0; i < totalRows; i += chunkSize) {
          const limit = Math.min(i + chunkSize, totalRows);

          for (let j = i; j < limit; j++) {
            const row = excelData[j];
            const statementBuilder = new SqlStatementBuilder(currentTableName, escapeIdentifier, formatValue);
            const setClauses = statementBuilder.setClauses;
            const whereClauses = statementBuilder.whereClauses;
            const processedFields = new Set();
            let hasAutoFill = false;

            const ibscbsSync = getValorCampo(row, 'CODCST_IBSCBS');
            const valorClasstribIbscbs = getValorCampo(row, 'CODCST_CLASSTRIB_IBSCBS');
            const classtribSync = syncEnabled
              ? (valorClasstribIbscbs || getValorCampo(row, 'CODCST_CLASSTRIB'))
              : (getValorCampo(row, 'CODCST_CLASSTRIB') || valorClasstribIbscbs);
            const classificacaoOficial = obterClassificacaoOficial(ibscbsSync, classtribSync);
            // A tabela SEFAZ só completa percentuais de redução. Alíquotas nominais
            // continuam dependendo da configuração fiscal vigente do contribuinte.
            const aliquotasMapeadas = getAliquotasPorCombinacao(ibscbsSync, classtribSync) || {};
            const reducaoOficial = obterReducaoOficial(classificacaoOficial) || {};
            const aliquotasSync = syncEnabled ? { ...aliquotasMapeadas, ...reducaoOficial } : null;

            for (const field of setFields) {
              const fieldUpper = field.toUpperCase();
              if (processedFields.has(fieldUpper)) continue;

              let valorFinal = campoExisteNaPlanilha(field) ? getValorCampo(row, field) : undefined;

              if (fieldUpper === 'CODCST_CLASSTRIB') {
                if (valorFinal === undefined || valorFinal === '') {
                  if (valorClasstribIbscbs) valorFinal = valorClasstribIbscbs;
                }
              }
              if (fieldUpper === 'CODCST_CLASSTRIB_IBSCBS') {
                if (valorFinal === undefined || valorFinal === '') {
                  valorFinal = getValorCampo(row, 'CODCST_CLASSTRIB') || valorClasstribIbscbs;
                }
              }

              // A planilha é a fonte de verdade. O mapa interno só completa
              // campos vazios; nunca pode substituir um valor importado.
              if ((valorFinal === undefined || valorFinal === '') && aliquotasSync) {
                const valMap = obterValorAliquota(aliquotasSync, fieldUpper);
                if (valMap !== undefined) {
                  valorFinal = valMap;
                  hasAutoFill = true;
                }
              }

              if (valorFinal !== undefined && valorFinal !== null && valorFinal !== '') {
                if (deveFormatarCST(field)) {
                  const originalValue = valorFinal;
                  valorFinal = formatarCST(valorFinal);
                  if (originalValue !== valorFinal && originalValue !== '') {
                    cstFormatCount++;
                  }
                }
                setClauses.push(`${escapeIdentifier(field)} = ${formatValue(valorFinal)}`);
                processedFields.add(fieldUpper);
              }
            }

            if (!syncEnabled) {
              const rawClasstrib = getValorCampo(row, 'CODCST_CLASSTRIB');
              const rawAlias = valorClasstribIbscbs;
              const valorCanonical = rawClasstrib || rawAlias || null;

              const temClasstrib = headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB');
              const temAlias = headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB_IBSCBS');
              const classtribSet = processedFields.has('CODCST_CLASSTRIB');
              const aliasSet = processedFields.has('CODCST_CLASSTRIB_IBSCBS');

              if (temAlias && classtribSet && !aliasSet && valorCanonical) {
                setClauses.push(`${escapeIdentifier('CODCST_CLASSTRIB_IBSCBS')} = ${formatValue(valorCanonical)}`);
                processedFields.add('CODCST_CLASSTRIB_IBSCBS');
              }

              if (temClasstrib && aliasSet && !classtribSet && valorCanonical) {
                setClauses.push(`${escapeIdentifier('CODCST_CLASSTRIB')} = ${formatValue(valorCanonical)}`);
                processedFields.add('CODCST_CLASSTRIB');
                syncCount++;
              }
            }

            // Campos presentes na planilha não são sincronizados entre si.
            // Isso impede que uma classificação preencha ou substitua outra.
            if (syncEnabled && !headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB') && !headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB_IBSCBS')) {
              const valorSincronizado = valorClasstribIbscbs || getValorCampo(row, 'CODCST_CLASSTRIB');
              if (valorSincronizado) {
                if (headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB')) {
                  const idCT = escapeIdentifier('CODCST_CLASSTRIB');
                  if (!processedFields.has('CODCST_CLASSTRIB')) {
                    setClauses.push(`${idCT} = ${formatValue(valorSincronizado)}`);
                    processedFields.add('CODCST_CLASSTRIB');
                    syncCount++;
                  } else {
                    const idx = setClauses.findIndex(c => c.startsWith(idCT + ' =') || c.startsWith('CODCST_CLASSTRIB ='));
                    if (idx !== -1) {
                      setClauses[idx] = `${idCT} = ${formatValue(valorSincronizado)}`;
                      syncCount++;
                    }
                  }
                }

                if (headers.some(h => h.toUpperCase() === 'CODCST_CLASSTRIB_IBSCBS')) {
                  const idAlias = escapeIdentifier('CODCST_CLASSTRIB_IBSCBS');
                  if (!processedFields.has('CODCST_CLASSTRIB_IBSCBS')) {
                    setClauses.push(`${idAlias} = ${formatValue(valorSincronizado)}`);
                    processedFields.add('CODCST_CLASSTRIB_IBSCBS');
                  } else {
                    const idx = setClauses.findIndex(c => c.startsWith(idAlias + ' =') || c.startsWith('CODCST_CLASSTRIB_IBSCBS ='));
                    if (idx !== -1) {
                      setClauses[idx] = `${idAlias} = ${formatValue(valorSincronizado)}`;
                    }
                  }
                }
              }
            }

            if (syncEnabled && aliquotasSync) {
              headers.forEach(header => {
                if (!campoExisteNaPlanilha(header)) return;
                const headerUpper = header.toUpperCase();
                if (ehCampoAliquota(headerUpper)) {
                  // Não alterar alíquota existente na planilha durante a sincronização.
                  const valorImportado = getValorCampo(row, header);
                  if (valorImportado !== undefined && valorImportado !== '') return;
                  const valMap = obterValorAliquota(aliquotasSync, headerUpper);
                  if (valMap !== undefined) {
                    if (processedFields.has(headerUpper)) {
                      const id = escapeIdentifier(header);
                      const idx = setClauses.findIndex(c => c.startsWith(id + ' =') || c.startsWith(header + ' ='));
                      if (idx !== -1) {
                        setClauses[idx] = `${id} = ${formatValue(valMap)}`;
                      }
                    } else {
                      setClauses.push(`${escapeIdentifier(header)} = ${formatValue(valMap)}`);
                      processedFields.add(headerUpper);
                      autoFillCount++;
                    }
                  }
                }
              });
            }

            // Validação final: campos existentes na planilha sempre prevalecem
            // sobre qualquer regra aplicada durante a montagem do UPDATE.
            for (const field of setFields) {
              const valorPlanilha = getValorCampo(row, field);
              if (valorPlanilha === undefined || valorPlanilha === null || valorPlanilha === '') continue;

              const valorFinalPlanilha = deveFormatarCST(field)
                ? formatarCST(valorPlanilha)
                : valorPlanilha;
              const identificador = escapeIdentifier(field);
              const clausula = `${identificador} = ${formatValue(valorFinalPlanilha)}`;
              const indice = setClauses.findIndex(c => c.startsWith(identificador + ' ='));

              if (indice === -1) {
                setClauses.push(clausula);
              } else {
                setClauses[indice] = clausula;
              }
            }

            for (const field of whereFields) {
              if (!campoExisteNaPlanilha(field)) continue;
              let valorFinal = getValorCampo(row, field);
              if (valorFinal !== undefined && valorFinal !== null && valorFinal !== '') {
                if (deveFormatarCST(field)) {
                  valorFinal = formatarCST(valorFinal);
                }
                whereClauses.push(`${escapeIdentifier(field)} = ${formatValue(valorFinal)}`);
              }
            }

            if (statementBuilder.canBuild) {
              sqlCommands.push(statementBuilder.build());
              if (hasAutoFill) autoFillCount++;
            } else {
              registrosIgnorados++;
            }
          }

          const pct = Math.round((limit / totalRows) * 100);
          progressBar.style.width = pct + '%';
          progressText.textContent = `Gerando SQLs... (${limit}/${totalRows})`;
          progressPercent.textContent = pct + '%';

          await new Promise(resolve => setTimeout(resolve, 0));
        }

        if (sqlCommands.length === 0) {
          showStatus('error', '❌ Nenhum SQL foi gerado.');
          return;
        }

        generatedSQLs = sqlCommands;

        // Tela, cópia e exportação devem partir da mesma lista completa de SQLs.
        sqlOutput.textContent = generatedSQLs.join('\n\n');

        sqlCounter.textContent = `${sqlCommands.length} comandos gerados`;

        let formatMsg = '';
        if (cstFormatCount > 0) formatMsg += ` 🔧 ${cstFormatCount} CST(s) formatados.`;
        if (syncCount > 0) formatMsg += ` 📌 ${syncCount} registro(s) sincronizados.`;
        if (autoFillCount > 0) formatMsg += ` 📌 ${autoFillCount} registro(s) com auto-fill.`;
        if (classificacoesInvalidas > 0) formatMsg += ` ⚠️ ${classificacoesInvalidas} classificação(ões) não encontrada(s) na SEFAZ.`;
        if (classificacoesForaVigencia > 0) formatMsg += ` ⚠️ ${classificacoesForaVigencia} classificação(ões) fora da vigência.`;
        if (registrosIgnorados > 0) formatMsg += ` ⚠️ ${registrosIgnorados} ignorados.`;

        showStatus('success', `✅ ${sqlCommands.length} UPDATE gerados!${formatMsg}`);

      } catch (error) {
        console.error('Erro:', error);
        showStatus('error', '❌ Erro ao gerar SQLs: ' + error.message);
      } finally {
        isGenerating = false;
        generateBtn.disabled = false;
        generateBtn.innerHTML = '🚀 Gerar UPDATE SQL';
        setTimeout(() => {
          progressContainer.style.display = 'none';
        }, 1200);
      }
    }

    // ============================================
    // FUNÇÕES DE UTILIDADE E EVENTOS
    // ============================================
    function copySQLs() {
      if (generatedSQLs.length === 0) {
        showStatus('error', '❌ Nenhum SQL gerado para copiar');
        return;
      }
      const conteudo = sqlOutput.textContent;
      const copiarComFallback = () => {
        const areaTemporaria = document.createElement('textarea');
        areaTemporaria.value = conteudo;
        areaTemporaria.style.position = 'fixed';
        areaTemporaria.style.opacity = '0';
        document.body.appendChild(areaTemporaria);
        areaTemporaria.select();
        const copiou = document.execCommand('copy');
        areaTemporaria.remove();
        if (!copiou) throw new Error('Não foi possível acessar a área de transferência');
      };
      // Em arquivo local, execute o fallback no próprio clique do botão.
      // Isso preserva a permissão de área de transferência do navegador.
      if (!navigator.clipboard?.writeText || !window.isSecureContext) {
        try {
          copiarComFallback();
          showStatus('success', '📋 SQLs copiados!');
        } catch {
          showStatus('error', '❌ Erro ao copiar SQLs');
        }
        return;
      }

      navigator.clipboard.writeText(conteudo).then(() => {
        showStatus('success', '📋 SQLs copiados!');
      }).catch(() => {
        try {
          copiarComFallback();
          showStatus('success', '📋 SQLs copiados!');
        } catch {
          showStatus('error', '❌ Erro ao copiar SQLs');
        }
      });
    }

    function exportSQLs() {
      if (generatedSQLs.length === 0) {
        showStatus('error', '❌ Nenhum SQL gerado para exportar');
        return;
      }
      let baseFileName = originalFileName || currentTableName;
      baseFileName = baseFileName.replace(/[^a-zA-Z0-9_\-]/g, '_');
      const blob = new Blob([sqlOutput.textContent], { type: 'text/sql;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${baseFileName}_update.sql`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(url);
        a.remove();
      }, 1000);
      showStatus('success', '💾 SQLs exportados!');
    }

    function clearAll() {
      generatedSQLs = [];
      sqlOutput.innerHTML = '-- Nenhum SQL gerado ainda --';
      sqlCounter.textContent = '0 comandos gerados';
      enableAutoFieldsToggle.checked = false;
      document.getElementById('searchSetInput').value = '';
      document.getElementById('searchWhereInput').value = '';
      const autoFieldsInfoEl = document.getElementById('autoFieldsInfo');
      if (autoFieldsInfoEl) autoFieldsInfoEl.style.display = 'none';
      document.getElementById('sheetModalOverlay').classList.remove('active');
      if (originalFileName) {
        tableNameInput.value = originalFileName.replace(/[^a-zA-Z0-9_]/g, '_').toLowerCase();
      } else {
        tableNameInput.value = 'produtos';
      }
      currentTableName = tableNameInput.value;
      tableNameInput.classList.remove('error-input');
      if (headers.length > 0) {
        renderFieldSelectors();
      }
      showStatus('info', '🔄 Configurações resetadas.');
    }

    function showStatus(type, message) {
      statusDiv.className = `status ${type}`;
      statusDiv.innerHTML = message;
    }

    function formatarDataSEFAZ(valor) {
      return valor ? new Date(valor).toLocaleDateString('pt-BR') : 'sem término';
    }

    function renderizarConsultaClassTrib(consulta = '') {
      const resumo = document.getElementById('classTribSearchSummary');
      const resultadosEl = document.getElementById('classTribSearchResults');
      if (!classTribInfo) {
        resumo.textContent = 'Catálogo oficial indisponível.';
        resultadosEl.innerHTML = '';
        return;
      }
      const resultados = classTribService.search(consulta, 100);
      resumo.textContent = resultados.length === 100
        ? 'Exibindo os primeiros 100 resultados. Refine a busca para ver menos itens.'
        : `${resultados.length} classificação(ões) encontrada(s).`;
      resultadosEl.innerHTML = resultados.length === 0
        ? '<div class="class-trib-result">Nenhuma classificação encontrada.</div>'
        : resultados.map(item => `
          <button type="button" class="class-trib-result class-trib-result-button" data-cst="${escapeHtml(item.CST)}" data-class-trib="${escapeHtml(item.cClassTrib)}">
            <div class="class-trib-code">CST ${escapeHtml(item.CST)} · ClassTrib ${escapeHtml(item.cClassTrib)}</div>
            <div class="class-trib-description">${escapeHtml(item.DescricaoClassTrib || 'Descrição não informada.')}</div>
            <div class="class-trib-meta">Redução IBS: ${escapeHtml(String(item.pRedIBS ?? 0))}% · CBS: ${escapeHtml(String(item.pRedCBS ?? 0))}% · Vigência: ${escapeHtml(formatarDataSEFAZ(item.InicioVigencia))} a ${escapeHtml(formatarDataSEFAZ(item.FimVigencia))}</div>
          </button>`).join('');
    }

    function lerNumeroTributario(valor) {
      const bruto = String(valor ?? '').trim();
      const texto = bruto.includes(',') ? bruto.replace(/\./g, '').replace(',', '.') : bruto;
      return Number.isFinite(Number(texto)) ? Number(texto) : 0;
    }

    function formatarNumeroTributario(valor) {
      return Number(valor || 0).toLocaleString('pt-BR', { maximumFractionDigits: 4 });
    }

    function atualizarSimuladorTributario() {
      const simulador = document.getElementById('classTribCalculator');
      const base = lerNumeroTributario(simulador.querySelector('[data-tax="base"]').value);
      ['uf', 'mun', 'cbs'].forEach(tipo => {
        const aliquota = lerNumeroTributario(simulador.querySelector(`[data-tax="aliquota-${tipo}"]`).value);
        const reducao = lerNumeroTributario(simulador.querySelector(`[data-tax="reducao-${tipo}"]`).value);
        const efetiva = aliquota * Math.max(0, 1 - reducao / 100);
        simulador.querySelector(`[data-output="efetiva-${tipo}"]`).value = formatarNumeroTributario(efetiva);
        simulador.querySelector(`[data-output="valor-${tipo}"]`).value = formatarNumeroTributario(base * efetiva / 100);
      });
    }

    function gerarUpdateCardapClassTributaria(cst, codigo) {
      const classificacao = obterClassificacaoOficial(cst, codigo);
      if (!classificacao) return;
      const cstOficial = classificacao.CST;
      const codigoOficial = classificacao.cClassTrib;
      const reducoes = obterReducaoOficial(classificacao);
      const aliquotasMapeadas = getAliquotasPorCombinacao(cstOficial, codigoOficial);
      const semAliquota = classificacaoSemAliquota(classificacao);
      const aliquotasBase = aliquotasMapeadas || (semAliquota
        ? { ALIQ_IBS_UF: '0', ALIQ_IBS_MUN: '0', ALIQ_CBS: '0' }
        : ALIQUOTAS_TRANSICAO_2026);
      const efetiva = (aliquota, reducao) => {
        const resultado = lerNumeroTributario(aliquota) * Math.max(0, 1 - lerNumeroTributario(reducao) / 100);
        return resultado.toFixed(3).replace(/\.?0+$/, '');
      };
      const aliquotas = [
        ['ALIQ_IBS_UF', aliquotasBase.ALIQ_IBS_UF],
        ['ALIQ_EFETIVA_IBS_UF', efetiva(aliquotasBase.ALIQ_IBS_UF, reducoes.REDUCAO_ALIQ_IBS_UF)],
        ['REDUCAO_ALIQ_IBS_UF', reducoes.REDUCAO_ALIQ_IBS_UF],
        ['ALIQ_IBS_MUN', aliquotasBase.ALIQ_IBS_MUN],
        ['ALIQ_EFETIVA_IBS_MUN', efetiva(aliquotasBase.ALIQ_IBS_MUN, reducoes.REDUCAO_ALIQ_IBS_MUN)],
        ['REDUCAO_ALIQ_IBS_MUN', reducoes.REDUCAO_ALIQ_IBS_MUN],
        ['ALIQ_CBS', aliquotasBase.ALIQ_CBS],
        ['ALIQ_EFETIVA_CBS', efetiva(aliquotasBase.ALIQ_CBS, reducoes.REDUCAO_ALIQ_CBS)],
        ['REDUCAO_ALIQ_CBS', reducoes.REDUCAO_ALIQ_CBS]
      ];
      const set = aliquotas.map(([campo, valorAliquota]) => `  ${campo} = ${formatValue(valorAliquota)}`).join(',\n');
      const sql = `UPDATE CARDAP\nSET\n${set}\nWHERE\n  CODCST_IBSCBS = ${formatValue(cstOficial)} AND\n  CODCST_CLASSTRIB = ${formatValue(codigoOficial)} AND\n  CODEMP = '1';`;
      const output = document.getElementById('classTribSqlOutput');
      output.textContent = sql;
      output.hidden = false;
      document.getElementById('copyClassTribSqlBtn').hidden = false;
    }

    function gerarUpdateMovNotaItem(cst, codigo) {
      const classificacao = obterClassificacaoOficial(cst, codigo);
      if (!classificacao) return;
      const codigoOficial = classificacao.cClassTrib;
      const reducoes = obterReducaoOficial(classificacao);
      const aliquotasMapeadas = getAliquotasPorCombinacao(classificacao.CST, codigoOficial);
      const semAliquota = classificacaoSemAliquota(classificacao);
      const aliquotasBase = aliquotasMapeadas || (semAliquota
        ? { ALIQ_IBS_UF: '0', ALIQ_IBS_MUN: '0', ALIQ_CBS: '0' }
        : ALIQUOTAS_TRANSICAO_2026);
      const efetiva = (aliquota, reducao) => {
        const resultado = lerNumeroTributario(aliquota) * Math.max(0, 1 - lerNumeroTributario(reducao) / 100);
        return resultado.toFixed(3).replace(/\.?0+$/, '');
      };
      // Conforme o exemplo informado para a classificação tributária 000001.
      const zeraEfetivas = codigoOficial === '000001';
      const ibsUfEfetiva = zeraEfetivas ? '0' : efetiva(aliquotasBase.ALIQ_IBS_UF, reducoes.REDUCAO_ALIQ_IBS_UF);
      const ibsMunEfetiva = zeraEfetivas ? '0' : efetiva(aliquotasBase.ALIQ_IBS_MUN, reducoes.REDUCAO_ALIQ_IBS_MUN);
      const cbsEfetiva = zeraEfetivas ? '0' : efetiva(aliquotasBase.ALIQ_CBS, reducoes.REDUCAO_ALIQ_CBS);
      const sql = `UPDATE MOV_NOTA_ITEM i\nSET\n` +
        `  i.RT_VLR_BASEIBS_UF_ITEM       = i.MI_VLR_TOTAL_ITEM,\n` +
        `  i.RT_ALIQ_IBS_UF_ITEM          = ${aliquotasBase.ALIQ_IBS_UF},\n` +
        `  i.RT_REDUCAO_ALIQ_IBS_UF_ITEM  = ${reducoes.REDUCAO_ALIQ_IBS_UF},\n` +
        `  i.RT_ALIQ_EFETIVA_IBS_UF_ITEM  = ${ibsUfEfetiva},\n` +
        `  i.RT_VAL_IBS_UF_ITEM           = ((${ibsUfEfetiva} / 100) * i.MI_VLR_TOTAL_ITEM),\n` +
        `  i.RT_VLR_BASEIBS_MUN_ITEM      = i.MI_VLR_TOTAL_ITEM,\n` +
        `  i.RT_ALIQ_IBS_MUN_ITEM         = ${aliquotasBase.ALIQ_IBS_MUN},\n` +
        `  i.RT_REDUCAO_ALIQ_IBS_MUN_ITEM = ${reducoes.REDUCAO_ALIQ_IBS_MUN},\n` +
        `  i.RT_ALIQ_EFETIVA_IBS_MUN_ITEM = ${ibsMunEfetiva},\n` +
        `  i.RT_VAL_IBS_MUN_ITEM          = ${ibsMunEfetiva === '0' ? '0' : `((${ibsMunEfetiva} / 100) * i.MI_VLR_TOTAL_ITEM)`},\n` +
        `  i.RT_VLR_BASECBS_ITEM          = i.MI_VLR_TOTAL_ITEM,\n` +
        `  i.RT_ALIQ_CBS_ITEM             = ${aliquotasBase.ALIQ_CBS},\n` +
        `  i.RT_REDUCAO_ALIQ_CBS_ITEM     = ${reducoes.REDUCAO_ALIQ_CBS},\n` +
        `  i.RT_ALIQ_EFETIVA_CBS_ITEM     = ${cbsEfetiva},\n` +
        `  i.RT_VLR_CBS_ITEM              = ((${cbsEfetiva} / 100) * i.MI_VLR_TOTAL_ITEM)\n` +
        `WHERE\n` +
        `  COALESCE(NULLIF(i.RT_CODCST_CLASSTRIB_IBSCBS_ITEM, ''), '') = ${formatValue(codigoOficial)}\n` +
        `  AND i.NOTAFISCAL IN (SELECT NOTAFISCAL FROM MOV_NOTA WHERE NFE_TRANSMITIDA='N' AND DATAEMISSAO>='01.07.2026');`;
      const output = document.getElementById('classTribSqlOutput');
      output.textContent = sql;
      output.hidden = false;
      document.getElementById('copyClassTribSqlBtn').hidden = false;
    }

    async function copiarUpdateClassTributaria() {
      const sql = document.getElementById('classTribSqlOutput').textContent;
      if (!sql) return;
      try {
        await navigator.clipboard.writeText(sql);
      } catch (error) {
        const area = document.createElement('textarea');
        area.value = sql;
        area.style.position = 'fixed';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        document.execCommand('copy');
        area.remove();
      }
      showStatus('success', '✅ UPDATE da classificação copiado para a área de transferência.');
    }

    function abrirSimuladorTributario(cst, codigo) {
      const classificacao = obterClassificacaoOficial(cst, codigo);
      if (!classificacao) return;
      const aliquotas = getAliquotasPorCombinacao(cst, codigo) || ALIQUOTAS_TRANSICAO_2026;
      const campo = (rotulo, atributo, valor, somenteLeitura = false) => `
        <label class="tax-field">${rotulo}
          <input type="text" inputmode="decimal" ${atributo} value="${escapeHtml(String(valor ?? 0))}" ${somenteLeitura ? 'readonly' : ''}>
        </label>`;
      const cartao = (titulo, tipo, aliquota, reducao) => `
        <section class="tax-card">
          <h5>${titulo}</h5>
          ${campo('Alíquota (%)', `data-tax="aliquota-${tipo}"`, aliquota)}
          <div class="tax-reduction">
            <strong>● Redução de alíquota</strong>
            ${campo('Redução oficial (%)', `data-tax="reducao-${tipo}"`, reducao)}
            ${campo('Alíquota efetiva (%)', `data-output="efetiva-${tipo}"`, 0, true)}
          </div>
          ${campo(`Valor do ${titulo}`, `data-output="valor-${tipo}"`, 0, true)}
        </section>`;

      document.getElementById('classTribCalculator').innerHTML = `
        <h4>Tributação — CST ${escapeHtml(cst)} · ClassTrib ${escapeHtml(codigo)}</h4>
        <p>${escapeHtml(classificacao.DescricaoClassTrib || '')}</p>
        <label class="tax-base-field">Valor da base de cálculo
          <input type="text" inputmode="decimal" data-tax="base" value="100">
        </label>
        <div class="tax-columns">
          ${cartao('IBS Estadual', 'uf', aliquotas.ALIQ_IBS_UF, classificacao.pRedIBS)}
          ${cartao('IBS Municipal', 'mun', aliquotas.ALIQ_IBS_MUN, classificacao.pRedIBS)}
          ${cartao('CBS', 'cbs', aliquotas.ALIQ_CBS, classificacao.pRedCBS)}
        </div>
        <div class="class-trib-sql-actions">
          <button type="button" class="btn-primary" id="generateClassTribSqlBtn" data-cst="${escapeHtml(cst)}" data-class-trib="${escapeHtml(codigo)}">⚡ Gerar UPDATE CARDAP</button>
          <button type="button" class="btn-primary" id="generateMovNotaItemSqlBtn" data-cst="${escapeHtml(cst)}" data-class-trib="${escapeHtml(codigo)}">⚡ Recalcula Mov_Nota_Item</button>
          <button type="button" class="btn-secondary" id="copyClassTribSqlBtn" hidden>📋 Copiar SQL</button>
        </div>
        <pre id="classTribSqlOutput" class="class-trib-sql-output" hidden aria-live="polite"></pre>`;
      document.getElementById('classTribSearchSummary').hidden = true;
      document.getElementById('classTribSearchResults').hidden = true;
      document.getElementById('classTribCalculator').hidden = false;
      document.getElementById('backToClassTribSearchBtn').hidden = false;
      atualizarSimuladorTributario();
    }

    function voltarConsultaClassTrib() {
      document.getElementById('classTribCalculator').hidden = true;
      document.getElementById('classTribSearchSummary').hidden = false;
      document.getElementById('classTribSearchResults').hidden = false;
      document.getElementById('backToClassTribSearchBtn').hidden = true;
      renderizarConsultaClassTrib(classTribSearchInput.value);
      classTribSearchInput.focus();
    }

    // ============================================
    // REGISTRO DE EVENTOS
    // ============================================
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const tableInfoDiv = document.getElementById('tableInfo');
    const enableAutoFieldsToggle = document.getElementById('enableAutoFieldsToggle');
    const setFieldsDiv = document.getElementById('setFields');
    const whereFieldsDiv = document.getElementById('whereFields');
    const generateBtn = document.getElementById('generateBtn');
    const copyBtn = document.getElementById('copyBtn');
    const exportBtn = document.getElementById('exportBtn');
    const clearBtn = document.getElementById('clearBtn');
    const sqlOutput = document.getElementById('sqlOutput');
    const sqlCounter = document.getElementById('sqlCounter');
    const statusDiv = document.getElementById('statusMsg');
    const selectAllSet = document.getElementById('selectAllSet');
    const selectAllWhere = document.getElementById('selectAllWhere');
    const tableNameInput = document.getElementById('tableNameInput');
    const classTribModalOverlay = document.getElementById('classTribModalOverlay');
    const classTribSearchInput = document.getElementById('classTribSearchInput');
    const classTribSearchResults = document.getElementById('classTribSearchResults');

    function carregarCatalogoSEFAZInterno() {
      try {
        if (!Array.isArray(window.CLASS_TRIB_DATA)) {
          throw new Error('Catálogo interno indisponível.');
        }
        carregarTabelaClassTrib(window.CLASS_TRIB_DATA);
      } catch (error) {
        console.error('Erro ao carregar catálogo SEFAZ:', error);
        atualizarStatusClassTrib('Catálogo oficial indisponível.');
      }
    }

    carregarCatalogoSEFAZInterno();

    document.getElementById('openClassTribSearchBtn').addEventListener('click', () => {
      document.getElementById('classTribCalculator').hidden = true;
      document.getElementById('classTribSearchSummary').hidden = false;
      document.getElementById('classTribSearchResults').hidden = false;
      document.getElementById('backToClassTribSearchBtn').hidden = true;
      renderizarConsultaClassTrib(classTribSearchInput.value);
      classTribModalOverlay.classList.add('active');
      classTribSearchInput.focus();
    });
    document.getElementById('closeClassTribModalBtn').addEventListener('click', () => {
      classTribModalOverlay.classList.remove('active');
    });
    classTribSearchInput.addEventListener('input', event => {
      const simulador = document.getElementById('classTribCalculator');
      if (!simulador.hidden) {
        simulador.hidden = true;
        document.getElementById('classTribSearchSummary').hidden = false;
        document.getElementById('classTribSearchResults').hidden = false;
        document.getElementById('backToClassTribSearchBtn').hidden = true;
      }
      renderizarConsultaClassTrib(event.target.value);
    });
    classTribSearchResults.addEventListener('click', event => {
      const resultado = event.target.closest('[data-cst][data-class-trib]');
      if (resultado) abrirSimuladorTributario(resultado.dataset.cst, resultado.dataset.classTrib);
    });
    document.getElementById('classTribCalculator').addEventListener('input', atualizarSimuladorTributario);
    document.getElementById('classTribCalculator').addEventListener('click', event => {
      const generateButton = event.target.closest('#generateClassTribSqlBtn');
      if (generateButton) {
        gerarUpdateCardapClassTributaria(generateButton.dataset.cst, generateButton.dataset.classTrib);
      }
      const movNotaItemButton = event.target.closest('#generateMovNotaItemSqlBtn');
      if (movNotaItemButton) {
        gerarUpdateMovNotaItem(movNotaItemButton.dataset.cst, movNotaItemButton.dataset.classTrib);
      }
      if (event.target.closest('#copyClassTribSqlBtn')) copiarUpdateClassTributaria();
    });
    document.getElementById('backToClassTribSearchBtn').addEventListener('click', voltarConsultaClassTrib);
    classTribModalOverlay.addEventListener('click', event => {
      if (event.target === event.currentTarget) classTribModalOverlay.classList.remove('active');
    });

    uploadArea.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        const file = e.target.files[0];
        originalFileName = file.name.replace(/\.(xlsx|xls|csv)$/i, '');
        processFile(file);
        fileInput.value = '';
      }
    });

    uploadArea.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
      uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadArea.classList.remove('dragover');
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        const file = files[0];
        if (file.name.match(/\.(xlsx|xls|csv)$/i)) {
          originalFileName = file.name.replace(/\.(xlsx|xls|csv)$/i, '');
          processFile(file);
        } else {
          showStatus('error', '❌ Selecione um arquivo Excel ou CSV');
        }
      }
    });

    document.body.addEventListener('dragover', (e) => e.preventDefault());
    document.body.addEventListener('drop', (e) => e.preventDefault());

    selectAllSet.addEventListener('change', (e) => {
      const EXCLUDED_SET_FIELDS = ['CODEMP', 'CAT_COD', 'CAT_DESCR', 'PROD_COD', 'PROD_DESCR'];
      document.querySelectorAll('.set-checkbox').forEach(cb => {
        if (EXCLUDED_SET_FIELDS.includes(cb.value.toUpperCase())) {
          cb.checked = false;
        } else {
          cb.checked = e.target.checked;
        }
      });
    });

    selectAllWhere.addEventListener('change', (e) => {
      document.querySelectorAll('.where-checkbox').forEach(cb => cb.checked = e.target.checked);
    });

    document.getElementById('searchSetInput').addEventListener('input', () => filterFields('setFields', 'searchSetInput'));
    document.getElementById('searchWhereInput').addEventListener('input', () => filterFields('whereFields', 'searchWhereInput'));

    document.getElementById('deselectAllSetBtn').addEventListener('click', () => {
      document.querySelectorAll('.set-checkbox').forEach(cb => cb.checked = false);
      selectAllSet.checked = false;
    });

    document.getElementById('invertSetBtn').addEventListener('click', () => {
      document.querySelectorAll('.set-checkbox').forEach(cb => cb.checked = !cb.checked);
      updateSelectAllStates();
    });

    document.getElementById('deselectAllWhereBtn').addEventListener('click', () => {
      document.querySelectorAll('.where-checkbox').forEach(cb => cb.checked = false);
      selectAllWhere.checked = false;
    });

    document.getElementById('invertWhereBtn').addEventListener('click', () => {
      document.querySelectorAll('.where-checkbox').forEach(cb => cb.checked = !cb.checked);
      updateSelectAllStates();
    });

    enableAutoFieldsToggle.addEventListener('change', (e) => {
      if (e.target.checked && !planilhaStats.temIbscbs) {
        showStatus('warning', '⚠️ CODCST_IBSCBS não encontrado. Sincronização desativada.');
        e.target.checked = false;
        return;
      }
      renderFieldSelectors();
      if (e.target.checked) {
        showStatus('success', '✅ Sincronização ativada. Alíquotas e campos mapeados serão preenchidos automaticamente.');
      } else {
        showStatus('info', '⚠️ Sincronização desativada.');
      }
    });

    document.addEventListener('change', (e) => {
      if (e.target.classList && (e.target.classList.contains('set-checkbox') || e.target.classList.contains('where-checkbox'))) {
        updateSelectAllStates();
      }
    });

    generateBtn.addEventListener('click', generateSQL);
    copyBtn.addEventListener('click', copySQLs);
    exportBtn.addEventListener('click', exportSQLs);
    clearBtn.addEventListener('click', clearAll);

    document.getElementById('sheetConfirmBtn').addEventListener('click', () => {
      const selected = document.querySelector('input[name="sheetSelect"]:checked');
      if (selected) {
        processSheet(selected.value);
      }
    });

    document.getElementById('sheetCancelBtn').addEventListener('click', () => {
      document.getElementById('sheetModalOverlay').classList.remove('active');
      showStatus('info', '⚡ Seleção de aba cancelada.');
    });

    document.getElementById('sheetModalOverlay').addEventListener('click', (e) => {
      if (e.target === e.currentTarget) {
        document.getElementById('sheetModalOverlay').classList.remove('active');
        showStatus('info', '⚡ Seleção de aba cancelada.');
      }
    });

    window.addEventListener('DOMContentLoaded', () => {
      const urlParams = new URLSearchParams(window.location.search);
      if (urlParams.get('v') === APP_VERSION) {
        showStatus('success', `⚡ Aplicação atualizada automaticamente para a versão v${APP_VERSION}! Cache renovado.`);
      }
    });

    console.log('✅ v5.21 - Interface inicializada com serviços fiscais e de SQL separados');
