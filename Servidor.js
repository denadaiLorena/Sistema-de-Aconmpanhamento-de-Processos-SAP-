/**
 * Backend (Google Apps Script) do SAP.
 *
 * Responsabilidades principais:
 * - Servir a interface HTML (`Cliente.html`) via Web App.
 * - Ler/escrever dados na planilha (Processos, Clientes, Produtos, Unidades).
 * - Aplicar validações e regras de negócio no servidor.
 */


function doGet(e) {
return HtmlService.createTemplateFromFile('Cliente')
.evaluate()
.setTitle('Sistema de Processos')
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Remove acentos para normalização de texto (busca/validação). */
function removerAcentos_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/** Normaliza o nome do cliente para maiúsculas (padrão de persistência) */
function normalizarNomeCliente_(valor) {
  return removerAcentos_(valor).toUpperCase();
}

/** Extrai apenas dígitos (útil para CPF/CNPJ e IDs numéricos) */
function somenteDigitos_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).replace(/\D/g, '');
}

/**
 * Valida documento (CPF/CNPJ) pelo tamanho (11 ou 14 dígitos)
 * Retorna apenas os dígitos (sem máscara)
 */
function validarDocumentoCpfCnpj_(documento) {
  const d = somenteDigitos_(documento);
  if (!d) throw new Error('Documento (CPF/CNPJ) não informado.');
  if (d.length !== 11 && d.length !== 14) {
    throw new Error('Documento inválido: informe CPF (11 dígitos) ou CNPJ (14 dígitos).');
  }
  return d;
}

/**
 * Validação de “campos obrigatórios” para criação/edição de processo.
 * Mantém a regra de negócio e consistência dos dados antes de gravar na planilha.
 */
function validarCamposObrigatoriosProcesso_(dados) {
  if (!dados) throw new Error('Dados não informados.');

  const contrato = normalizeKey_(dados.contratoCodigo);
  if (!contrato) throw new Error('Contrato não informado.');

  const status = normalizeKey_(dados.status);
  if (!status) throw new Error('Status não informado.');

  const garantia = normalizeKey_(dados.garantia);
  if (!garantia) throw new Error('Garantia não informada.');

  const valor = dados.valor;
  if (valor === null || valor === undefined || String(valor).trim() === '') {
    throw new Error('Valor não informado.');
  }

  const clienteNome = normalizeKey_(normalizarNomeCliente_(dados.clienteNome));
  if (!clienteNome) throw new Error('Nome do cliente não informado.');

  validarDocumentoCpfCnpj_(dados.clienteDocumento);

  const produtoIdSelecionado = normalizeKey_(dados.produtoId);
  if (!produtoIdSelecionado) throw new Error('Produto não informado.');

  const unidadeIdSelecionada = normalizeKey_(dados.unidadeId);
  if (!unidadeIdSelecionada) throw new Error('Unidade não informada.');
}

/**
 * Converte valor para texto seguro no Google Sheets
 */
function comoTextoSheets_(valor) {
  const v = normalizeKey_(valor);
  if (!v) return '';
  // Se já estiver forçado como texto, não duplica.
  if (v.startsWith("'")) return v;
  // Prefixo com apóstrofo força o Google Sheets a manter como texto (ex: 00123).
  return "'" + v;
}

/**
 * Busca um cliente na aba Clientes pelo CPF/CNPJ (apenas dígitos)
 * Usado no frontend para autocompletar o nome ao digitar o documento
 */
function buscarClientePorDocumentoServidor(documento) {
  const alvo = somenteDigitos_(documento);
  if (!alvo) return { encontrado: false };

  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");
  const abaClientes = getAbaClientes_(ss);

  const valores = abaClientes.getDataRange().getValues();
  // [0]=ID, [1]=Nome, [2]=Documento
  for (let i = 1; i < valores.length; i++) {
    const clienteId = normalizeKey_(valores[i][0]);
    const nome = normalizeKey_(valores[i][1]);
    const doc = somenteDigitos_(valores[i][2]);

    if (!clienteId || !doc) continue;
    if (doc === alvo) {
      return {
        encontrado: true,
        clienteId,
        nome,
      };
    }
  }

  return { encontrado: false };
}

/**
 * Retorna a lista de processos já enriquecida (join) com cliente/produto/unidade.
 */
function buscarDadosCompletos() {
  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");

  const abaProcessos = getAbaProcessos_(ss);
  const abaClientes = getAbaClientes_(ss);
  const abaProdutos = getAbaProdutos_(ss);
  const abaUnidades = getAbaUnidades_(ss);

  const processos = abaProcessos.getDataRange().getValues();
  const clientes = abaClientes.getDataRange().getValues();
  const produtos = abaProdutos.getDataRange().getValues();
  const unidades = abaUnidades.getDataRange().getValues();

  // Remove cabeçalhos
  processos.shift();
  clientes.shift();
  produtos.shift();
  unidades.shift();

  //MAPAS (JOIN)
  const clientesMap = {};
  clientes.forEach(l => {
    const key = normalizeKey_(l[0]);
    if (!key) return;
    const nome = l[1];
    const documento = l[2];

    clientesMap[key] = { nome, documento };
  });

  const produtosMap = {};
  produtos.forEach(l => {
    const key = normalizeKey_(l[0]);
    if (!key) return;
    produtosMap[key] = {
      segmentoEsubsegmento: l[1],
      nome: l[2],
    };
  });

  const unidadesMap = {};
  unidades.forEach(l => {
    const key = normalizeKey_(l[0]);
    if (!key) return;
    unidadesMap[key] = {
      cid: l[1],
      nome: l[2]
    };
  });

  // PROCESSAMENTO FINAL
  // Regras de negócio enviadas prontas para o frontend:
  // - `prioridade`: true quando há garantia (diferente de NENHUMA)
  // - `isViavel`: true quando valor >= 15000
  const resultado = processos.map(p => {
    const processoId = normalizeKey_(p[0]);
    const clienteId = normalizeKey_(p[1]);
    const produtoId = normalizeKey_(p[2]);
    const unidadeId = normalizeKey_(p[3]);

    let cliente = clientesMap[clienteId] || {};
    let clienteIdResolvido = clienteId;
    const produto = produtosMap[produtoId] || {};
    const unidade = unidadesMap[unidadeId] || {};

    if (!cliente.nome && clienteId) {
      Logger.log('[JOIN] Cliente não encontrado | processo=' + processoId + ' | valorEmProcessos=' + clienteId);
    }

    const valorContrato = Number(p[5]) || 0;
    const tipoGarantia = p[6];

    return {
      processo: p[0],

      clienteId: clienteIdResolvido,
      produtoId: produtoId,
      unidadeId: unidadeId,

      clienteNome: cliente.nome || 'Não encontrado',
      clienteDocumento: cliente.documento || 'N/A',

      produtoNome: produto.nome || 'N/A',
      segmentoEsubseguimento: produto.segmentoEsubsegmento || 'N/A',

      unidadeNome: unidade.nome || 'N/A',
      unidadeCID: unidade.cid || 'N/A',

      contratoCodigo: p[4],
      contratoValor: valorContrato,

      garantiaTipo: tipoGarantia,
      andamento: p[7],

      // Envia como timestamp (ms) para o frontend formatar no fuso do computador do usuário
      dataCriacao: p[8] instanceof Date ? p[8].getTime() : (p[8] ? new Date(p[8]).getTime() : null),
      dataAtualizacao: p[9] instanceof Date ? p[9].getTime() : (p[9] ? new Date(p[9]).getTime() : null),

      prioridade: tipoGarantia && tipoGarantia.trim().toUpperCase() !== "NENHUMA",
      isViavel: Number(valorContrato) >= 15000
    };
  });
  return resultado;
}

/** Verifica se um ID existe na primeira coluna de uma aba (pós-cabeçalho). */
function idExisteNaAba_(aba, id) {
  const chave = normalizeKey_(id);
  if (!chave) return false;

  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (normalizeKey_(valores[i][0]) === chave) return true;
  }
  return false;
}

/**
 * Lista catálogos (Produtos e Unidades) para popular selects no frontend
 * Retorna objetos simples com id/nome e campos auxiliares (segSub/cid)
 */
function listarCatalogosServidor() {
  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");
  const abaProdutos = getAbaProdutos_(ss);
  const abaUnidades = getAbaUnidades_(ss);

  const produtos = abaProdutos.getDataRange().getValues();
  const unidades = abaUnidades.getDataRange().getValues();
  produtos.shift();
  unidades.shift();

  return {
    produtos: produtos
      .map((l) => ({ id: normalizeKey_(l[0]), segSub: normalizeKey_(l[1]), nome: normalizeKey_(l[2]) }))
      .filter((p) => !!p.id),
    unidades: unidades
      .map((l) => ({ id: normalizeKey_(l[0]), cid: normalizeKey_(l[1]), nome: normalizeKey_(l[2]) }))
      .filter((u) => !!u.id),
  };
}

/** Acessa a aba Processos (falha com erro amigável se não existir) */
function getAbaProcessos_(ss) {
  const aba = ss.getSheetByName('Processos');
  if (!aba) throw new Error('Aba Processos não encontrada');
  return aba;
}

/** Acessa a aba Clientes (falha com erro amigável se não existir) */
function getAbaClientes_(ss) {
  const aba = ss.getSheetByName('Clientes');
  if (!aba) throw new Error('Aba Clientes não encontrada');
  return aba;
}

/** Acessa a aba Produtos (falha com erro amigável se não existir) */
function getAbaProdutos_(ss) {
  const aba = ss.getSheetByName('Produtos');
  if (!aba) throw new Error('Aba Produtos não encontrada');
  return aba;
}

/** Acessa a aba Unidades (falha com erro amigável se não existir) */
function getAbaUnidades_(ss) {
  const aba = ss.getSheetByName('Unidades');
  if (!aba) throw new Error('Aba Unidades não encontrada');
  return aba;
}

/** Normaliza um valor para chave (trim + string); usado para comparar IDs. */
function normalizeKey_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/** Encontra o índice do processo pelo ID. */
function encontrarLinhaProcessoPorId_(linhas, processoId) {
  const id = normalizeKey_(processoId);
  if (!id) return -1;

  for (let i = 1; i < linhas.length; i++) {
    if (normalizeKey_(linhas[i][0]) === id) return i;
  }
  return -1;
}

/** Protege string para uso seguro em RegExp. */
function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Gera o próximo ID com prefixo (ex.: PROCESSO_001 -> PROCESSO_002).
 */
function proximoIdPrefixado_(linhas, prefixoPadrao) {
  // Padrão: PREFIXO_<n>
  // Mantém o prefixo (incluindo caixa) e preserva padding de zeros se existir no maior ID encontrado.
  let maxNumero = 0;
  let prefixo = prefixoPadrao;
  let padWidth = 0;

  const re = new RegExp('^(' + escapeRegExp_(prefixoPadrao) + ')(\\d+)$', 'i');

  for (let i = 1; i < linhas.length; i++) {
    const raw = normalizeKey_(linhas[i][0]);
    if (!raw) continue;

    const m = raw.match(re);
    if (!m) continue;

    const numStr = m[2];
    const n = Number(numStr);
    if (!Number.isFinite(n)) continue;

    if (n > maxNumero) {
      maxNumero = n;
      prefixo = m[1];
      padWidth = numStr.length;
    }
  }

  const proximo = maxNumero + 1;
  const proximoStr = padWidth > 0 ? String(proximo).padStart(padWidth, '0') : String(proximo);
  return prefixo + proximoStr;
}

/** Endpoint chamado pelo frontend para obter um novo ID de processo. */
function gerarNovoProcessoId() {
  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");
  const aba = getAbaProcessos_(ss);
  const valores = aba.getDataRange().getValues();
  return proximoIdPrefixado_(valores, 'PROCESSO_');
}

/**
 * Cria um novo processo.
 * Fluxo:
 * - valida campos
 * - cria um novo cliente na aba Clientes
 * - grava o processo na aba Processos usando o clienteId gerado
 */
function criarProcessoServidor(dados) {
  if (!dados || normalizeKey_(dados.id) === '') {
    throw new Error('ID do processo não informado.');
  }

  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");
  const abaProcessos = getAbaProcessos_(ss);
  const abaClientes = getAbaClientes_(ss);
  const abaProdutos = getAbaProdutos_(ss);
  const abaUnidades = getAbaUnidades_(ss);

  const processoId = normalizeKey_(dados.id);
  const nowMs = Number(dados.clientNowMs);
  const now = Number.isFinite(nowMs) ? new Date(nowMs) : new Date();

  validarCamposObrigatoriosProcesso_(dados);

  const clienteNome = normalizeKey_(normalizarNomeCliente_(dados.clienteNome));
  const produtoIdSelecionado = normalizeKey_(dados.produtoId);
  const unidadeIdSelecionada = normalizeKey_(dados.unidadeId);

  // Gera ID do cliente e cria registro na aba Clientes
  const clientesValores = abaClientes.getDataRange().getValues();
  const clienteId = proximoIdPrefixado_(clientesValores, 'CLIENTE_');
  abaClientes.appendRow([clienteId, clienteNome, comoTextoSheets_(dados.clienteDocumento)]);
  const linhaClienteGravada = abaClientes.getLastRow();

  // Valida IDs selecionados existem nas abas
  if (!idExisteNaAba_(abaProdutos, produtoIdSelecionado)) {
    throw new Error('Produto não encontrado pelo ID: ' + produtoIdSelecionado);
  }
  if (!idExisteNaAba_(abaUnidades, unidadeIdSelecionada)) {
    throw new Error('Unidade não encontrada pelo ID: ' + unidadeIdSelecionada);
  }

  // Cria processo
  const linhaProcesso = [
    processoId,
    clienteId,
    produtoIdSelecionado,
    unidadeIdSelecionada,
    comoTextoSheets_(dados.contratoCodigo),
    Number(dados.valor) || 0,
    normalizeKey_(dados.garantia),
    normalizeKey_(dados.status),
    now,
    now,
  ];

  // Evita duplicidade pelo ID
  const existentes = abaProcessos.getDataRange().getValues();
  if (encontrarLinhaProcessoPorId_(existentes, processoId) !== -1) {
    throw new Error('Já existe um processo com o ID: ' + processoId);
  }

  abaProcessos.appendRow(linhaProcesso);
  const linhaProcessoGravada = abaProcessos.getLastRow();
  const processoSalvo = abaProcessos.getRange(linhaProcessoGravada, 1, 1, 10).getValues()[0];
  if (normalizeKey_(processoSalvo[0]) !== processoId) {
    throw new Error('Falha ao confirmar salvamento do processo na planilha.');
  }

  const processoSalvoOut = processoSalvo.map((v) => (v instanceof Date ? v.getTime() : v));

  return {
    ok: true,
    acao: 'criado',
    processoId,
    linhaProcessos: linhaProcessoGravada,
    processo: processoSalvoOut,
    clienteId,
    linhaClientes: linhaClienteGravada,
  };
}

/** Atualiza o nome do cliente na aba Clientes, dado um clienteId existente. */
function atualizarNomeClientePorId_(ss, clienteId, novoNome) {
  const id = normalizeKey_(clienteId);
  const nome = normalizeKey_(normalizarNomeCliente_(novoNome)).replace(/\s+/g, ' ');
  if (!id) throw new Error('ID do cliente vazio.');
  if (!nome) throw new Error('Nome do cliente vazio.');

  const abaClientes = ss.getSheetByName('Clientes');
  if (!abaClientes) throw new Error('Aba Clientes não encontrada');

  const valores = abaClientes.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (normalizeKey_(valores[i][0]) === id) {
      // Coluna B (Nome)
      abaClientes.getRange(i + 1, 2).setValue(nome);
      return;
    }
  }

  throw new Error('Cliente não encontrado pelo ID: ' + id);
}

/** Atualiza CPF/CNPJ do cliente na aba Clientes, mantendo-o como texto no Sheets. */
function atualizarDocumentoClientePorId_(ss, clienteId, novoDocumento) {
  const id = normalizeKey_(clienteId);
  const documento = normalizeKey_(novoDocumento);
  if (!id) throw new Error('ID do cliente vazio.');

  validarDocumentoCpfCnpj_(documento);

  const abaClientes = ss.getSheetByName('Clientes');
  if (!abaClientes) throw new Error('Aba Clientes não encontrada');

  const valores = abaClientes.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (normalizeKey_(valores[i][0]) === id) {
      // Coluna C (Documento)
      abaClientes.getRange(i + 1, 3).setValue(comoTextoSheets_(documento));
      return;
    }
  }

  throw new Error('Cliente não encontrado pelo ID: ' + id);
}

/**
 * Edita um processo existente.
 * Regra importante: o clienteId é derivado da linha do processo (não do frontend),
 * garantindo consistência e evitando que alguém edite o processo “apontando” para outro cliente.
 */
function editarProcessoServidor(dados) {
  if (!dados || dados.id === null || dados.id === undefined || String(dados.id).trim() === '') {
    throw new Error('ID do processo não informado. Não é possível salvar a edição.');
  }

  const clienteNomeInformado = (dados.clienteNome !== undefined) ? dados.clienteNome : dados.cliente;
  if (clienteNomeInformado === null || clienteNomeInformado === undefined || String(clienteNomeInformado).trim() === '') {
    throw new Error('Nome do cliente não informado. Não é possível salvar a edição.');
  }

  validarCamposObrigatoriosProcesso_({
    contratoCodigo: dados.contratoCodigo,
    status: dados.status,
    garantia: dados.garantia,
    valor: dados.valor,
    clienteNome: clienteNomeInformado,
    clienteDocumento: dados.clienteDocumento,
    produtoId: dados.produtoId,
    unidadeId: dados.unidadeId,
  });

  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");
  const aba = getAbaProcessos_(ss);
  const abaProdutos = getAbaProdutos_(ss);
  const abaUnidades = getAbaUnidades_(ss);

  const linhas = aba.getDataRange().getValues();

  const idx = encontrarLinhaProcessoPorId_(linhas, dados.id);
  if (idx === -1) throw new Error('Processo não encontrado: ' + dados.id);

  const row = idx + 1; // idx=1 corresponde à linha 2 (após cabeçalho)

  // Sempre deriva o clienteId a partir do ID do processo
  const clienteIdAtual = normalizeKey_(linhas[idx][1]);
  if (!clienteIdAtual) {
    throw new Error('Processo ' + dados.id + ' está sem ID de cliente. Não é possível atualizar o nome.');
  }

  const produtoIdAtual = normalizeKey_(linhas[idx][2]);
  const unidadeIdAtual = normalizeKey_(linhas[idx][3]);

  // Atualizações nas tabelas de referência sempre via IDs atuais do processo
  atualizarNomeClientePorId_(ss, clienteIdAtual, clienteNomeInformado);

  if (dados.clienteDocumento !== undefined) {
    atualizarDocumentoClientePorId_(ss, clienteIdAtual, dados.clienteDocumento);
  }

  // Permite trocar as chaves de Produto/Unidade selecionadas no catálogo
  const novoProdutoId = normalizeKey_(dados.produtoId);
  if (!novoProdutoId) throw new Error('Produto (ID) não informado.');
  if (!idExisteNaAba_(abaProdutos, novoProdutoId)) {
    throw new Error('Produto não encontrado pelo ID: ' + novoProdutoId);
  }

  const novaUnidadeId = normalizeKey_(dados.unidadeId);
  if (!novaUnidadeId) throw new Error('Unidade (ID) não informada.');
  if (!idExisteNaAba_(abaUnidades, novaUnidadeId)) {
    throw new Error('Unidade não encontrada pelo ID: ' + novaUnidadeId);
  }

  const contratoCodigo = comoTextoSheets_(dados.contratoCodigo);
  const valor = Number(dados.valor) || 0;
  const garantia = normalizeKey_(dados.garantia);
  const status = normalizeKey_(dados.status);
  const nowMs = Number(dados.clientNowMs);
  const now = Number.isFinite(nowMs) ? new Date(nowMs) : new Date();

  // Atualiza somente a linha editada (bem mais leve que setValues no DataRange inteiro)
  aba.getRange(row, 3).setValue(novoProdutoId);   // C: produtoId
  aba.getRange(row, 4).setValue(novaUnidadeId);   // D: unidadeId
  aba.getRange(row, 5).setValue(contratoCodigo);  // E: contrato
  aba.getRange(row, 6).setValue(valor);           // F: valor
  aba.getRange(row, 7).setValue(garantia);        // G: garantia
  aba.getRange(row, 8).setValue(status);          // H: status
  aba.getRange(row, 10).setValue(now);            // J: data atualização

  if (produtoIdAtual && !idExisteNaAba_(abaProdutos, produtoIdAtual)) {
    Logger.log('Aviso: produtoId atual não existe mais: ' + produtoIdAtual);
  }
  if (unidadeIdAtual && !idExisteNaAba_(abaUnidades, unidadeIdAtual)) {
    Logger.log('Aviso: unidadeId atual não existe mais: ' + unidadeIdAtual);
  }

  const processoSalvo = aba.getRange(row, 1, 1, 10).getValues()[0];
  if (normalizeKey_(processoSalvo[0]) !== normalizeKey_(dados.id)) {
    throw new Error('Falha ao confirmar salvamento da edição na planilha.');
  }

  const processoSalvoOut = processoSalvo.map((v) => (v instanceof Date ? v.getTime() : v));

  return {
    ok: true,
    acao: 'editado',
    processoId: normalizeKey_(dados.id),
    linhaProcessos: row,
    processo: processoSalvoOut,
  };
}

/** Exclui um processo pelo ID (apenas remove a linha na aba Processos). */
function excluirProcessoServidor(processoId) {
  const id = normalizeKey_(processoId);
  if (!id) throw new Error('ID do processo não informado. Não é possível excluir.');

  const ss = SpreadsheetApp.openById("17W-FbFkoJ4Msr0tYawse8uneYUS0qo9nMi3sjsWts58");
  const aba = getAbaProcessos_(ss);
  const linhas = aba.getDataRange().getValues();

  const idx = encontrarLinhaProcessoPorId_(linhas, id);
  if (idx === -1) throw new Error('Processo não encontrado: ' + id);

  aba.deleteRow(idx + 1);
  return 'Excluído com sucesso';
}