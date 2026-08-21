/**
 * Repositório em memória da tabela oficial de Classificação Tributária (SEFAZ).
 * Não toma decisões fiscais: apenas disponibiliza validação, vigência e reduções
 * informadas na fonte oficial carregada pelo usuário.
 */
class ClassTribService {
  constructor(normalizeCode) {
    this.normalizeCode = normalizeCode;
    this.index = new Map();
    this.info = null;
  }

  load(data) {
    if (!Array.isArray(data)) {
      throw new Error('O arquivo não contém a lista de CSTs esperada.');
    }

    if (data.length > 0 && data[0].c && data[0].t) {
      return this.loadEmbeddedCatalog(data);
    }

    const index = new Map();
    data.forEach(group => {
      const cst = this.normalizeCode('CODCST_IBSCBS', group.CST);
      (group.classificacoesTributarias || []).forEach(item => {
        const code = this.normalizeCode('CODCST_CLASSTRIB_IBSCBS', item.cClassTrib);
        if (!cst || !code) return;
        index.set(`${cst}|${code}`, { ...item, CST: cst, DescricaoCST: group.DescricaoCST });
      });
    });

    if (index.size === 0) {
      throw new Error('Nenhuma classificação tributária válida foi encontrada no arquivo.');
    }

    this.index = index;
    const publications = [...index.values()].map(item => item.Publicacao).filter(Boolean).sort();
    this.info = { total: index.size, publication: publications.at(-1) || null };
    return this.info;
  }

  loadEmbeddedCatalog(data) {
    const index = new Map();
    data.forEach(item => {
      const cst = this.normalizeCode('CODCST_IBSCBS', item.c);
      const code = this.normalizeCode('CODCST_CLASSTRIB_IBSCBS', item.t);
      if (!cst || !code) return;
      index.set(`${cst}|${code}`, {
        CST: cst,
        cClassTrib: code,
        DescricaoClassTrib: item.n || '',
        pRedIBS: item.ri,
        pRedCBS: item.rc,
        InicioVigencia: item.i,
        FimVigencia: item.f,
        Publicacao: item.p
      });
    });
    if (index.size === 0) throw new Error('O catálogo interno de classificações está vazio.');

    this.index = index;
    const publications = [...index.values()].map(item => item.Publicacao).filter(Boolean).sort();
    this.info = { total: index.size, publication: publications.at(-1) || null };
    return this.info;
  }

  find(cstValue, classTribValue) {
    const cst = this.normalizeCode('CODCST_IBSCBS', cstValue) || '';
    const classTrib = this.normalizeCode('CODCST_CLASSTRIB_IBSCBS', classTribValue) || '';
    return this.index.get(`${cst}|${classTrib}`) || null;
  }

  isCurrent(classification, today = new Date()) {
    if (!classification) return false;
    const start = classification.InicioVigencia ? new Date(classification.InicioVigencia) : null;
    const end = classification.FimVigencia ? new Date(classification.FimVigencia) : null;
    return (!start || start <= today) && (!end || end >= today);
  }

  getReductions(classification) {
    if (!classification) return null;
    const cbs = String(classification.pRedCBS ?? 0);
    const ibs = String(classification.pRedIBS ?? 0);
    return {
      REDUCAO_ALIQ_CBS: cbs,
      REDUCAO_ALIQ_IBS_UF: ibs,
      REDUCAO_ALIQ_IBS_MUN: ibs,
      REDUCAO_ALIQ_IBS_MU: ibs
    };
  }

  search(query, limit = 100) {
    const term = String(query || '').trim().toLocaleLowerCase('pt-BR');
    const records = [...this.index.values()];
    if (!term) return records.slice(0, limit);
    return records.filter(item => [item.CST, item.cClassTrib, item.DescricaoCST, item.DescricaoClassTrib]
      .filter(Boolean)
      .some(value => String(value).toLocaleLowerCase('pt-BR').includes(term)))
      .slice(0, limit);
  }
}

window.ClassTribService = ClassTribService;
