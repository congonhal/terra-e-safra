// ============================================================
// TERRA & SAFRA — Backup diário automático por e-mail
// Roda sozinho via GitHub Actions, não precisa do navegador aberto.
// Atualizado em 22/07/2026 pra bater com a lógica financeira atual
// do app (fertilizantes, consultoria rural, boletos, divisão 50/50
// até fim de 2026 com Ronaldo/Lucas, custo de fiado proporcional).
// ============================================================

const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

const SUPABASE_URL = 'https://zpqvyexitrdllujoehvp.supabase.co';
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_APP_PASSWORD = process.env.EMAIL_APP_PASSWORD;
const EMAIL_DESTINO = 'biancamarisant@gmail.com';

const sb = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);

// ---------- Helpers gerais (espelham a lógica do app) ----------
const TIPO_LABEL = { remedio: 'Medicamentos', racao: 'Rações', proteinado: 'Nutrição Animal' };
function categoriaLabel(tipo) { return TIPO_LABEL[tipo] || tipo; }
function isDescricaoTipo(tipo) { return tipo === 'remedio' || tipo === 'racao' || tipo === 'proteinado'; }
function tipoLabel(tipo) { return tipo === 'remedio' ? 'Medicamento' : tipo === 'racao' ? 'Ração' : 'Nutrição Animal'; }
function ehFertilizante(tipo) { return tipo === 'Fertilizantes'; }

function hojeStr() {
  const now = new Date();
  const brt = new Date(now.toLocaleString('en-US', { timeZone: 'America/Sao_Paulo' }));
  const y = brt.getFullYear(), m = String(brt.getMonth() + 1).padStart(2, '0'), d = String(brt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}
function formatDate(d) { if (!d) return ''; const [y, m, dd] = d.split('-'); return `${dd}/${m}/${y}`; }
function diasParaVencer(d) {
  if (!d) return null;
  const hoje = new Date(hojeStr() + 'T00:00:00');
  const val = new Date(d + 'T00:00:00');
  return Math.round((val - hoje) / 86400000);
}
function diasDesde(d) { if (!d) return null; return -diasParaVencer(d); }
function statusLabel(d) {
  const dias = diasParaVencer(d);
  if (dias === null) return 'sem validade definida';
  if (dias < 0) return `vencido há ${Math.abs(dias)} dia(s)`;
  if (dias === 0) return 'vence hoje';
  if (dias <= 30) return `vence em ${dias} dia(s)`;
  return `válido (${dias} dias)`;
}
function brl(n) { return Number(n || 0).toFixed(2); }

// Divisão societária: 50% empresa até fim de 2026, 100% a partir de 2027
function fatorEmpresaFertilizante(dataVenda) {
  if (!dataVenda) return 1;
  const ano = Number(String(dataVenda).slice(0, 4));
  return ano <= 2026 ? 0.5 : 1;
}
function fatorEmpresaConsultoria(dataProjeto) {
  if (!dataProjeto) return 1;
  const ano = Number(String(dataProjeto).slice(0, 4));
  return ano <= 2026 ? 0.5 : 1;
}

function bookAppend(wb, rows, sheetName) {
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Aviso: 'Nenhum dado encontrado' }]);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}
function bufferOf(wb) { return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); }

async function main() {
  console.log('Buscando dados do Supabase...');
  const [
    { data: produtos, error: e1 },
    { data: lotes, error: e2 },
    { data: vendasFiado, error: e3 },
    { data: vendasItens, error: e4 },
    { data: vendasPagamentos, error: e5 },
    { data: despesas, error: e6 },
    { data: aportes, error: e7 },
    { data: retiradas, error: e8 },
    { data: entradasExtra, error: e9 },
    { data: movimentacoes, error: e10 },
    { data: vendasAvista, error: e11 },
    { data: vendasAvistaItens, error: e12 },
    { data: precosHistorico, error: e13 },
    { data: fertilizanteVendas, error: e14 },
    { data: fertBoletos, error: e15 },
    { data: fertBoletosPagamentos, error: e16 },
    { data: consultoriaProjetos, error: e17 },
  ] = await Promise.all([
    sb.from('produtos').select('*'),
    sb.from('lotes').select('*'),
    sb.from('vendas_fiado').select('*'),
    sb.from('vendas_fiado_itens').select('*'),
    sb.from('vendas_fiado_pagamentos').select('*'),
    sb.from('despesas').select('*'),
    sb.from('socios_aportes').select('*'),
    sb.from('socios_retiradas').select('*'),
    sb.from('entradas_extra').select('*'),
    sb.from('lotes_movimentacoes').select('*'),
    sb.from('vendas_avista').select('*'),
    sb.from('vendas_avista_itens').select('*'),
    sb.from('precos_historico').select('*'),
    sb.from('fertilizante_vendas').select('*'),
    sb.from('fertilizante_boletos').select('*'),
    sb.from('fertilizante_boletos_pagamentos').select('*'),
    sb.from('consultoria_projetos').select('*'),
  ]);
  const erro = e1 || e2 || e3 || e4 || e5 || e6 || e7 || e8 || e9 || e10 || e11 || e12 || e13 || e14 || e15 || e16 || e17;
  if (erro) { throw new Error('Erro ao buscar dados: ' + erro.message); }

  const vendas = vendasFiado.map(v => ({
    ...v,
    itens: vendasItens.filter(it => it.venda_id === v.id),
    pagamentos: vendasPagamentos.filter(p => p.venda_id === v.id),
  }));
  const avista = vendasAvista.map(v => ({
    ...v,
    itens: vendasAvistaItens.filter(it => it.venda_id === v.id),
  }));

  function todosPagamentosFiado() {
    const lista = [];
    vendas.forEach(v => (v.pagamentos || []).forEach(p => lista.push({ ...p, cliente_nome: v.cliente_nome })));
    return lista;
  }

  // ---------- Estoque / produtos ----------
  function saldo(l) { return Number(l.quantidade_inicial || 0) - Number(l.qtd_vendida || 0); }
  function totalVendido(l) { return Number(l.qtd_vendida || 0) * Number(l.preco_venda_final || 0); }
  function totalInvestido(l) { return (Number(l.quantidade_inicial || 0) - Number(l.qtd_brinde || 0)) * Number(l.valor_pago || 0); }
  function impostoPago(l) { return Number(l.preco_venda_final || 0) * (Number(l.imposto_pct || 0) / 100) * Number(l.qtd_vendida || 0); }
  function lucroEstimado(l) { return (Number(l.preco_venda_final || 0) - Number(l.valor_pago || 0)) * Number(l.qtd_vendida || 0); }
  function precoSugerido(l) {
    const perc = (Number(l.imposto_pct || 0) + Number(l.margem_pct || 0)) / 100;
    if (perc >= 1) return null;
    return Number(l.valor_pago || 0) / (1 - perc);
  }

  // ---------- Fiado (usa valor_final com desconto quando existe) ----------
  function totalVendaFiado(v) { return (v.valor_final !== null && v.valor_final !== undefined) ? Number(v.valor_final) : v.itens.reduce((s, it) => s + Number(it.quantidade) * Number(it.valor_unitario), 0); }
  function totalPagoFiado(v) {
    if (v.pagamentos && v.pagamentos.length > 0) return v.pagamentos.reduce((s, p) => s + Number(p.valor), 0);
    if (v.status === 'pago') return totalVendaFiado(v);
    return 0;
  }
  function saldoDevedorFiado(v) { return Math.max(0, Math.round((totalVendaFiado(v) - totalPagoFiado(v)) * 100) / 100); }
  function diasSemPagamentoCliente(nome) {
    const vc = vendas.filter(v => v.cliente_nome === nome);
    const abertas = vc.filter(v => saldoDevedorFiado(v) > 0);
    if (abertas.length === 0) return null;
    let ultima = null;
    vc.forEach(v => (v.pagamentos || []).forEach(p => { if (!ultima || p.data_pagamento > ultima) ultima = p.data_pagamento; }));
    if (!ultima) ultima = abertas.reduce((min, v) => (!min || v.data_venda < min) ? v.data_venda : min, null);
    return diasDesde(ultima);
  }
  function statusClienteCor(nome) {
    const dias = diasSemPagamentoCliente(nome);
    if (dias === null) return 'verde';
    if (dias > 90) return 'vermelho';
    if (dias > 60) return 'amarelo';
    return 'verde';
  }
  // custo do fiado só na proporção do que já foi pago (usado no "lucro sem o fiado em aberto")
  function custoFiadoProporcionalPago(movsPeriodo) {
    let custo = 0;
    movsPeriodo.forEach(m => {
      if (!m.venda_fiado_id) return;
      const l = lotes.find(x => x.id === m.lote_id);
      if (!l) return;
      const venda = vendas.find(v => v.id === m.venda_fiado_id);
      if (!venda) return;
      const totalVenda = totalVendaFiado(venda);
      const totalPago = totalPagoFiado(venda);
      if (totalVenda <= 0) return;
      const fracaoPaga = Math.min(1, totalPago / totalVenda);
      custo += Number(m.quantidade) * Number(l.valor_pago || 0) * fracaoPaga;
    });
    return custo;
  }

  function totalDespesasPagasPor(nome) { return despesas.filter(d => d.responsavel === nome).reduce((s, d) => s + Number(d.valor), 0); }

  // ---------- Fertilizantes (módulo separado) ----------
  function fertValorRelevante(v) { return v.origem === 'fabrica' ? Number(v.valor_comissao || 0) : Number(v.valor_final || 0); }
  function fertImpostoDaVenda(v) {
    if (v.origem === 'fabrica') {
      return Number(v.valor_comissao || 0) * (Number(v.imposto_pct || 0) / 100);
    }
    const valorFinal = Number(v.valor_final || 0);
    const impostoPct = Number(v.imposto_pct || 0);
    if (v.forma_pagamento !== 'boleto') {
      const naoTributa = v.forma_pagamento === 'dinheiro';
      return naoTributa ? 0 : valorFinal * (impostoPct / 100);
    }
    const boleto = fertBoletos.find(b => b.fertilizante_venda_id === v.id);
    if (!boleto || Number(boleto.valor_total || 0) <= 0) return 0;
    const baixas = fertBoletosPagamentos.filter(p => p.boleto_id === boleto.id);
    return baixas.reduce((s, p) => {
      const fracao = Number(p.valor || 0) / Number(boleto.valor_total);
      const naoTributaBaixa = (p.forma_pagamento || 'dinheiro') === 'dinheiro';
      return s + (naoTributaBaixa ? 0 : valorFinal * fracao * (impostoPct / 100));
    }, 0);
  }
  function fertLucroDaVenda(v) {
    if (v.origem === 'fabrica') return Number(v.valor_comissao || 0) - fertImpostoDaVenda(v);
    const custoTotal = Number(v.quantidade_toneladas || 0) * Number(v.custo_tonelada || 0);
    return Number(v.valor_final || 0) - custoTotal - fertImpostoDaVenda(v);
  }

  // ---------- Cálculos gerais: Entradas / Lucro (fatia da empresa) ----------
  const movsAvista = movimentacoes.filter(m => m.tipo === 'venda' && m.quantidade > 0 && m.venda_avista_id);
  const movsFiadoMov = movimentacoes.filter(m => m.tipo === 'venda' && m.quantidade > 0 && m.venda_fiado_id);

  function custoEImpostoAvista() {
    let custoProdutos = 0, impostoEstimado = 0;
    movsAvista.forEach(m => {
      const l = lotes.find(x => x.id === m.lote_id);
      if (!l) return;
      const p = produtos.find(pr => pr.id === l.produto_id);
      const fatorEmpresa = (p && ehFertilizante(p.tipo)) ? fatorEmpresaFertilizante(m.data) : 1;
      custoProdutos += Number(m.quantidade) * Number(l.valor_pago || 0) * fatorEmpresa;
      const venda = avista.find(v => v.id === m.venda_avista_id);
      if (venda && venda.forma_pagamento) {
        const naoTributa = venda.forma_pagamento === 'dinheiro';
        if (!naoTributa) {
          const nome = p ? p.name : '';
          const itemCorrespondente = venda.itens.find(it => nome && it.produto_nome.startsWith(nome));
          const precoUnitario = itemCorrespondente ? Number(itemCorrespondente.valor_unitario) : Number(l.preco_venda_final || 0);
          const bruto = venda.itens.reduce((s, it) => s + Number(it.quantidade) * Number(it.valor_unitario), 0);
          const fator = bruto > 0 ? Number(venda.valor_recebido || 0) / bruto : 1;
          impostoEstimado += Number(m.quantidade) * precoUnitario * fator * (Number(l.imposto_pct || 0) / 100);
        }
      }
    });
    return { custoProdutos, impostoEstimado };
  }

  const receitaAvista = avista.reduce((s, v) => s + Number(v.valor_recebido || 0), 0);
  const pagamentosFiadoTodos = todosPagamentosFiado().filter(p => p.forma_pagamento);
  const receitaFiadoPago = pagamentosFiadoTodos.reduce((s, p) => s + Number(p.valor || 0), 0);
  const { custoProdutos: custoAvista, impostoEstimado: impostoAvista } = custoEImpostoAvista();
  const custoFiadoPago = custoFiadoProporcionalPago(movsFiadoMov);

  function impostoPagamentosFiado(pagamentosDoPeriodo) {
    let total = 0;
    pagamentosDoPeriodo.forEach(pg => {
      const venda = vendas.find(v => (v.pagamentos || []).some(x => x.id === pg.id));
      if (!venda) return;
      const naoTributa = pg.forma_pagamento === 'dinheiro';
      if (naoTributa) return;
      const totalVenda = totalVendaFiado(venda);
      if (totalVenda <= 0) return;
      const itens = venda.itens || [];
      const bruto = itens.reduce((s, it) => s + Number(it.quantidade) * Number(it.valor_unitario), 0);
      const fatorDesconto = bruto > 0 ? totalVenda / bruto : 1;
      const fracaoDessePagamento = Number(pg.valor) / totalVenda;
      itens.forEach(it => {
        const mv = movsFiadoMov.find(m => m.venda_fiado_id === venda.id);
        const l = mv ? lotes.find(x => x.id === mv.lote_id) : null;
        const impostoPct = l ? Number(l.imposto_pct || 0) : 0;
        total += Number(it.quantidade) * Number(it.valor_unitario) * fatorDesconto * fracaoDessePagamento * (impostoPct / 100);
      });
    });
    return total;
  }
  const impostoFiadoPago = impostoPagamentosFiado(pagamentosFiadoTodos);

  function ajusteReceitaSocietaria() {
    let ajuste = 0;
    movsAvista.forEach(m => {
      const l = lotes.find(x => x.id === m.lote_id);
      if (!l) return;
      const p = produtos.find(pr => pr.id === l.produto_id);
      if (!p || !ehFertilizante(p.tipo)) return;
      const fatorEmpresa = fatorEmpresaFertilizante(m.data);
      if (fatorEmpresa >= 1) return;
      const venda = avista.find(v => v.id === m.venda_avista_id);
      if (!venda) return;
      const bruto = venda.itens.reduce((s, it) => s + Number(it.quantidade) * Number(it.valor_unitario), 0);
      const fator = bruto > 0 ? Number(venda.valor_recebido || 0) / bruto : 1;
      const item = venda.itens.find(it => it.produto_nome.startsWith(p.name));
      const precoUnit = item ? Number(item.valor_unitario) : Number(l.preco_venda_final || 0);
      ajuste += Number(m.quantidade) * precoUnit * fator * (1 - fatorEmpresa);
    });
    pagamentosFiadoTodos.forEach(pg => {
      const venda = vendas.find(v => (v.pagamentos || []).some(x => x.id === pg.id));
      if (!venda) return;
      const fatorEmpresa = fatorEmpresaFertilizante(venda.data_venda);
      if (fatorEmpresa >= 1) return;
      const totalVenda = totalVendaFiado(venda);
      if (totalVenda <= 0) return;
      const fracao = Number(pg.valor) / totalVenda;
      const movsDaVenda = movimentacoes.filter(m => m.venda_fiado_id === venda.id && m.tipo === 'venda' && m.quantidade > 0);
      movsDaVenda.forEach(m => {
        const l = lotes.find(x => x.id === m.lote_id);
        if (!l) return;
        const p = produtos.find(pr => pr.id === l.produto_id);
        if (!p || !ehFertilizante(p.tipo)) return;
        const item = venda.itens.find(it => it.produto_nome.startsWith(p.name));
        const precoUnit = item ? Number(item.valor_unitario) : Number(l.preco_venda_final || 0);
        ajuste += Number(m.quantidade) * precoUnit * fracao * (1 - fatorEmpresa);
      });
    });
    return ajuste;
  }
  const ajusteSocietario = ajusteReceitaSocietaria();
  const totalDespesasManual = despesas.filter(d => !d.eh_imposto).reduce((s, d) => s + Number(d.valor), 0);
  const totalImpostosManual = despesas.filter(d => d.eh_imposto).reduce((s, d) => s + Number(d.valor), 0);

  const lucroVendasProdutos = (receitaAvista + receitaFiadoPago - ajusteSocietario) - (custoAvista + custoFiadoPago) - (impostoAvista + impostoFiadoPago) - totalDespesasManual;

  const fertPagas = fertilizanteVendas.filter(v => v.status_pagamento === 'pago');
  const totalEntradasFertilizantes = fertPagas.reduce((s, v) => s + fertValorRelevante(v) * fatorEmpresaFertilizante(v.data), 0);
  const totalImpostoEstimadoFertilizantes = fertPagas.reduce((s, v) => s + fertImpostoDaVenda(v), 0);
  const lucroVendasFertilizantes = fertPagas.reduce((s, v) => s + fertLucroDaVenda(v) * fatorEmpresaFertilizante(v.data), 0);

  const consultPagos = consultoriaProjetos.filter(p => p.status_servico === 'aprovado' && p.status_pagamento === 'pago');
  const totalEntradasConsultoria = consultPagos.reduce((s, p) => s + Number(p.valor_final || 0) * fatorEmpresaConsultoria(p.data), 0);
  const lucroVendasConsultoria = totalEntradasConsultoria;

  const totalEntradasVendas = receitaAvista + receitaFiadoPago - ajusteSocietario;
  const totalEntradasExtra = entradasExtra.reduce((s, e) => s + Number(e.valor), 0);
  const totalEntradasGeral = totalEntradasVendas + totalEntradasFertilizantes + totalEntradasConsultoria + totalEntradasExtra;

  const totalGastoProduto = lotes.reduce((s, l) => s + totalInvestido(l), 0);
  const impostosEstimadosGeral = impostoAvista + impostoFiadoPago + totalImpostoEstimadoFertilizantes;
  const saidasGeral = totalGastoProduto + totalDespesasManual + totalImpostosManual;
  const lucroLiquido = totalEntradasGeral - saidasGeral;
  const lucroVendas = lucroVendasProdutos + lucroVendasFertilizantes + lucroVendasConsultoria;

  console.log('Montando planilha de Estoque...');
  const wbEstoque = XLSX.utils.book_new();
  const linhasEstoque = lotes.map(l => {
    const p = produtos.find(pr => pr.id === l.produto_id) || {};
    const sug = precoSugerido(l);
    return {
      'Produto': p.name || '', 'Categoria': categoriaLabel(p.tipo), 'Variante/Tamanho': l.variante || '',
      'Lote': l.nome, 'Nota Fiscal': l.com_nf === false ? 'Sem NF' : 'Com NF',
      'Fornecedor': l.fornecedor || '', 'Contato Fornecedor': l.fornecedor_contato || '',
      'Unidade': l.observacao || 'un', 'Qtd Comprada': l.quantidade_inicial, 'Qtd Vendida': l.qtd_vendida || 0, 'Qtd Brinde (grátis)': l.qtd_brinde || 0,
      'Saldo': saldo(l), 'Valor Pago (unitário)': Number(l.valor_pago || 0), 'Imposto %': Number(l.imposto_pct || 0),
      'Margem %': Number(l.margem_pct || 0), 'Preço Sugerido': sug !== null ? Number(sug.toFixed(2)) : '',
      'Preço de Venda Final': Number(l.preco_venda_final || 0), 'Total Investido': Number(totalInvestido(l).toFixed(2)),
      'Total Vendido (receita)': Number(totalVendido(l).toFixed(2)), 'Total Imposto Pago': Number(impostoPago(l).toFixed(2)),
      'Lucro Produtos': Number(lucroEstimado(l).toFixed(2)), 'Data de Entrada': l.data_entrada ? formatDate(l.data_entrada) : '',
      'Validade': l.data_validade ? formatDate(l.data_validade) : '',
      'Qtd Vendida Atualizada Em': l.qtd_vendida_atualizada_em ? formatDate(l.qtd_vendida_atualizada_em) : '',
      'Última Conferência Física': l.ultima_atualizacao ? formatDate(l.ultima_atualizacao) : '',
      'Qtd Contada na Última Conferência': l.qtd_conferida ?? '',
    };
  });
  bookAppend(wbEstoque, linhasEstoque, 'Estoque');

  console.log('Montando planilha de Vencimentos...');
  const wbVenc = XLSX.utils.book_new();
  let rowsVenc = lotes.map(l => {
    const p = produtos.find(pr => pr.id === l.produto_id);
    return (p && isDescricaoTipo(p.tipo)) ? { ...l, produtoNome: p.name, produtoTipo: p.tipo } : null;
  }).filter(Boolean).filter(l => saldo(l) > 0);
  rowsVenc.sort((a, b) => { if (!a.data_validade) return 1; if (!b.data_validade) return -1; return a.data_validade.localeCompare(b.data_validade); });
  const linhasVenc = rowsVenc.map(l => ({
    'Produto': l.produtoNome, 'Variante/Tamanho': l.variante || '', 'Lote': l.nome, 'Tipo': tipoLabel(l.produtoTipo),
    'Quantidade Atual': saldo(l), 'Validade': l.data_validade ? formatDate(l.data_validade) : '', 'Situação': statusLabel(l.data_validade),
  }));
  bookAppend(wbVenc, linhasVenc, 'Vencimentos');

  console.log('Montando planilha Financeira (Geral)...');
  const wbFin = XLSX.utils.book_new();
  bookAppend(wbFin, [
    { Item: '💰 LUCRO VENDAS (produtos sem fiado em aberto + fertilizantes/consultoria já pagos, só a fatia da empresa)', 'Valor (R$)': Number(lucroVendas.toFixed(2)) },
    { Item: '', 'Valor (R$)': '' },
    { Item: 'Entrada — Vendas Produtos', 'Valor (R$)': Number(totalEntradasVendas.toFixed(2)) },
    { Item: 'Entrada — Vendas Fertilizantes (fatia da empresa)', 'Valor (R$)': Number(totalEntradasFertilizantes.toFixed(2)) },
    { Item: 'Entrada — Consultoria Rural (fatia da empresa)', 'Valor (R$)': Number(totalEntradasConsultoria.toFixed(2)) },
    { Item: 'Entrada — Atípicas', 'Valor (R$)': Number(totalEntradasExtra.toFixed(2)) },
    { Item: 'TOTAL DE ENTRADAS', 'Valor (R$)': Number(totalEntradasGeral.toFixed(2)) },
    { Item: '', 'Valor (R$)': '' },
    { Item: 'Impostos estimados (produtos + fertilizantes)', 'Valor (R$)': Number(impostosEstimadosGeral.toFixed(2)) },
    { Item: 'Impostos pagos (DARF)', 'Valor (R$)': Number(totalImpostosManual.toFixed(2)) },
    { Item: 'Despesa Produtos (todo o estoque, desde sempre)', 'Valor (R$)': Number(totalGastoProduto.toFixed(2)) },
    { Item: 'Despesas operacionais', 'Valor (R$)': Number(totalDespesasManual.toFixed(2)) },
    { Item: '', 'Valor (R$)': '' },
    { Item: 'Lucro líquido até o momento', 'Valor (R$)': Number(lucroLiquido.toFixed(2)) },
  ], 'Resumo Geral');

  const categoriasPresentes = Array.from(new Set(produtos.map(p => categoriaLabel(p.tipo))));
  const entradasLinhas = [
    ...categoriasPresentes.map(cat => {
      const totalCat = lotes.filter(l => { const p = produtos.find(pr => pr.id === l.produto_id); return p && categoriaLabel(p.tipo) === cat; }).reduce((s, l) => s + totalVendido(l), 0);
      return { Tipo: 'Venda de produtos', 'Categoria/Descrição': cat, Data: '', 'Valor (R$)': Number(totalCat.toFixed(2)) };
    }),
    ...entradasExtra.map(e => ({ Tipo: 'Atípica', 'Categoria/Descrição': e.descricao, Data: formatDate(e.data), 'Valor (R$)': Number(Number(e.valor).toFixed(2)) })),
  ];
  bookAppend(wbFin, entradasLinhas, 'Entradas');

  const saidasLinhas = [
    { Tipo: 'Produto (Estoque)', Categoria: 'Compra de estoque', Responsável: 'Empresa', Data: '', Descrição: '', 'Valor (R$)': Number(totalGastoProduto.toFixed(2)) },
    ...despesas.map(d => ({ Tipo: d.eh_imposto ? 'Imposto' : 'Despesa', Categoria: d.categoria, Responsável: d.responsavel || 'Empresa', Data: formatDate(d.data), Descrição: d.descricao || '', 'Valor (R$)': Number(Number(d.valor).toFixed(2)) })),
  ];
  bookAppend(wbFin, saidasLinhas, 'Saídas');

  const SOCIOS = ['Lucas', 'Gabriel', 'Ronaldo'];
  const sociosResumoLinhas = SOCIOS.map(nome => {
    const aportado = aportes.filter(a => a.socio === nome).reduce((s, a) => s + Number(a.valor), 0);
    const despesasPagas = totalDespesasPagasPor(nome);
    const totalDevido = aportado + despesasPagas;
    const retiradoCapital = retiradas.filter(r => r.socio === nome && r.tipo === 'reposicao_capital').reduce((s, r) => s + Number(r.valor), 0);
    const retiradoLucro = retiradas.filter(r => r.socio === nome && r.tipo === 'lucro').reduce((s, r) => s + Number(r.valor), 0);
    return {
      Sócio: nome, 'Aportado diretamente (R$)': Number(aportado.toFixed(2)), 'Despesas pagas do bolso (R$)': Number(despesasPagas.toFixed(2)),
      'Total devido a ele (R$)': Number(totalDevido.toFixed(2)), 'Retirado como capital (R$)': Number(retiradoCapital.toFixed(2)),
      'Saldo de capital a repor (R$)': Number(Math.max(0, totalDevido - retiradoCapital).toFixed(2)), 'Retirado como lucro (R$)': Number(retiradoLucro.toFixed(2)),
    };
  });
  bookAppend(wbFin, sociosResumoLinhas, 'Sócios');

  console.log('Montando planilha de Vendas Fiado...');
  const wbFiado = XLSX.utils.book_new();
  const vendasLinhas = [];
  vendas.forEach(v => {
    const total = totalVendaFiado(v), pago = totalPagoFiado(v), sd = saldoDevedorFiado(v);
    v.itens.forEach(it => {
      vendasLinhas.push({
        Cliente: v.cliente_nome, Telefone: v.cliente_telefone || '', 'Data da venda': formatDate(v.data_venda),
        Produto: it.produto_nome, Quantidade: it.quantidade, 'Valor unit. (R$)': Number(it.valor_unitario),
        'Subtotal (R$)': Number((it.quantidade * it.valor_unitario).toFixed(2)), 'Total da venda (R$)': Number(total.toFixed(2)),
        'Pago até agora (R$)': Number(pago.toFixed(2)), 'Saldo devedor (R$)': Number(sd.toFixed(2)),
      });
    });
  });
  bookAppend(wbFiado, vendasLinhas, 'Vendas');

  const pagamentosLinhas = [];
  vendas.forEach(v => (v.pagamentos || []).forEach(p => pagamentosLinhas.push({ Cliente: v.cliente_nome, 'Data do pagamento': formatDate(p.data_pagamento), 'Valor (R$)': Number(Number(p.valor).toFixed(2)), 'Forma de pagamento': p.forma_pagamento || 'não classificado' })));
  bookAppend(wbFiado, pagamentosLinhas, 'Pagamentos');

  const nomesClientes = Array.from(new Set(vendas.map(v => v.cliente_nome))).sort();
  const clientesLinhas = nomesClientes.map(nome => {
    const dias = diasSemPagamentoCliente(nome);
    const cor = statusClienteCor(nome);
    return { Cliente: nome, Situação: cor === 'vermelho' ? 'Vermelho (evitar fiado)' : cor === 'amarelo' ? 'Amarelo (atenção)' : 'Verde (em dia)', 'Dias sem pagamento': dias === null ? '' : dias };
  });
  bookAppend(wbFiado, clientesLinhas, 'Reputação Clientes');

  console.log('Montando planilha de Vendas à Vista...');
  const wbAvista = XLSX.utils.book_new();
  const formaLabel = { dinheiro: 'Dinheiro', pix: 'PIX', cartao_credito: 'Cartão Crédito', cartao_debito: 'Cartão Débito', boleto: 'Boleto' };
  const linhasAvista = [];
  avista.forEach(v => {
    v.itens.forEach(it => {
      linhasAvista.push({
        Data: formatDate(v.data), Produto: it.produto_nome, Quantidade: it.quantidade, 'Valor unit. (R$)': Number(it.valor_unitario),
        'Subtotal (R$)': Number((it.quantidade * it.valor_unitario).toFixed(2)), 'Valor bruto da venda (R$)': Number(Number(v.valor_bruto).toFixed(2)),
        'Valor recebido (R$)': Number(Number(v.valor_recebido).toFixed(2)),
        'Forma de pagamento': v.forma_pagamento ? formaLabel[v.forma_pagamento] : 'pendente de classificar',
        'Taxa maquininha (R$)': v.taxa_maquina_valor ? Number(Number(v.taxa_maquina_valor).toFixed(2)) : 0,
      });
    });
  });
  bookAppend(wbAvista, linhasAvista, 'Vendas à Vista');

  const pagamentosFiadoLinhas = todosPagamentosFiado().map(p => ({
    Cliente: p.cliente_nome, 'Data do pagamento': formatDate(p.data_pagamento), 'Valor (R$)': Number(Number(p.valor).toFixed(2)),
    'Forma de pagamento': p.forma_pagamento ? formaLabel[p.forma_pagamento] : 'pendente de classificar',
    'Taxa maquininha (R$)': p.taxa_maquina_valor ? Number(Number(p.taxa_maquina_valor).toFixed(2)) : 0,
  }));
  bookAppend(wbAvista, pagamentosFiadoLinhas, 'Pagamentos Fiado Classificados');

  console.log('Montando planilha de Fertilizantes...');
  const wbFert = XLSX.utils.book_new();
  const linhasFertVendas = fertilizanteVendas.map(v => {
    const lucro = fertLucroDaVenda(v);
    const fatorEmpresa = fatorEmpresaFertilizante(v.data);
    return {
      'Nº Pedido': v.numero_pedido || '', Data: formatDate(v.data), Cliente: v.cliente,
      Origem: v.origem === 'fabrica' ? 'Fábrica' : 'Estoque', Produto: v.produto_nome || '',
      'Qtd (t)': v.quantidade_toneladas, 'Valor/ton (R$)': Number(v.valor_tonelada || 0), 'Valor final (R$)': Number(v.valor_final || 0),
      '% Comissão': v.origem === 'fabrica' ? Number(v.pct_comissao || 0) : '', 'Valor comissão (R$)': v.origem === 'fabrica' ? Number(v.valor_comissao || 0) : '',
      'Custo tonelada (R$)': v.origem === 'fabrica' ? '' : Number(v.custo_tonelada || 0), 'Forma pagamento': v.forma_pagamento || '',
      'Imposto %': Number(v.imposto_pct || 0), 'Imposto (R$)': Number(fertImpostoDaVenda(v).toFixed(2)),
      'Lucro (R$)': Number(lucro.toFixed(2)), 'Lucro Empresa (R$)': Number((lucro * fatorEmpresa).toFixed(2)), 'Lucro Sócio (R$)': Number((lucro * (1 - fatorEmpresa)).toFixed(2)),
      Status: v.status_venda === 'cancelado' ? 'Cancelado' : 'Realizado', Pagamento: v.status_pagamento === 'pago' ? 'Pago' : 'Pendente',
      Parceria: v.valor_parceria != null ? Number(v.valor_parceria) : '', 'Sobra Frete': v.valor_frete_sobra != null ? Number(v.valor_frete_sobra) : '',
      Observações: v.observacoes || '',
    };
  });
  bookAppend(wbFert, linhasFertVendas, 'Vendas Fertilizantes');

  const linhasFertBoletos = fertBoletos.map(b => {
    const pago = fertBoletosPagamentos.filter(p => p.boleto_id === b.id).reduce((s, p) => s + Number(p.valor || 0), 0);
    return {
      Tipo: b.tipo === 'pagar' ? 'A Pagar' : 'A Receber', Descrição: b.descricao, 'Cliente/Fornecedor': b.cliente_fornecedor || '',
      'Valor total (R$)': Number(b.valor_total || 0), Parcelas: b.numero_parcelas, 'Data criação': formatDate(b.data_criacao),
      'Pago até agora (R$)': Number(pago.toFixed(2)), 'Saldo (R$)': Number((Number(b.valor_total || 0) - pago).toFixed(2)), Observações: b.observacoes || '',
    };
  });
  bookAppend(wbFert, linhasFertBoletos, 'Boletos');

  console.log('Montando planilha de Consultoria Rural...');
  const wbConsult = XLSX.utils.book_new();
  const linhasConsult = consultoriaProjetos.map(p => {
    const fatorEmpresa = fatorEmpresaConsultoria(p.data);
    const valorEmpresa = Number(p.valor_final || 0) * fatorEmpresa;
    return {
      Data: formatDate(p.data), Cliente: p.cliente, Cidade: p.cidade || '', Telefone: p.telefone || '',
      'Tipo de serviço': p.tipo_servico || '', 'Valor do serviço (R$)': Number(p.valor_servico || 0), '% Cobrada': Number(p.pct_cobrada || 0),
      'Valor final (R$)': Number(p.valor_final || 0), 'Status serviço': p.status_servico, 'Status pagamento': p.status_pagamento === 'pago' ? 'Pago' : 'Pendente',
      'Valor Empresa (R$)': Number(valorEmpresa.toFixed(2)), 'Valor Lucas (R$)': Number((Number(p.valor_final || 0) - valorEmpresa).toFixed(2)), Observações: p.observacoes || '',
    };
  });
  bookAppend(wbConsult, linhasConsult, 'Projetos');

  console.log('Montando planilha de Histórico de Preços...');
  const wbPrecos = XLSX.utils.book_new();
  const linhasPrecos = precosHistorico
    .map(ph => {
      const p = produtos.find(pr => pr.id === ph.produto_id);
      return { Produto: p ? p.name : '—', Categoria: p ? categoriaLabel(p.tipo) : '', Data: formatDate(ph.data), 'Preço (R$)': Number(Number(ph.preco).toFixed(2)) };
    })
    .sort((a, b) => a.Produto.localeCompare(b.Produto) || a.Data.localeCompare(b.Data));
  bookAppend(wbPrecos, linhasPrecos, 'Histórico de Preços');

  console.log('Enviando e-mail...');
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: EMAIL_USER, pass: EMAIL_APP_PASSWORD },
  });
  const dataHoje = formatDate(hojeStr());
  await transporter.sendMail({
    from: `Terra & Safra <${EMAIL_USER}>`,
    to: EMAIL_DESTINO,
    subject: `📦 Backup diário Terra & Safra — ${dataHoje}`,
    text: `Segue em anexo o backup automático de hoje (${dataHoje}):\n\n- Estoque.xlsx\n- Vencimentos.xlsx\n- Financeiro.xlsx\n- VendasFiado.xlsx\n- VendasAvista.xlsx\n- Fertilizantes.xlsx\n- ConsultoriaRural.xlsx\n- HistoricoPrecos.xlsx\n\nEsse e-mail é gerado e enviado sozinho, todo dia às 21h, direto do sistema.`,
    attachments: [
      { filename: `Estoque-${hojeStr()}.xlsx`, content: bufferOf(wbEstoque) },
      { filename: `Vencimentos-${hojeStr()}.xlsx`, content: bufferOf(wbVenc) },
      { filename: `Financeiro-${hojeStr()}.xlsx`, content: bufferOf(wbFin) },
      { filename: `VendasFiado-${hojeStr()}.xlsx`, content: bufferOf(wbFiado) },
      { filename: `VendasAvista-${hojeStr()}.xlsx`, content: bufferOf(wbAvista) },
      { filename: `Fertilizantes-${hojeStr()}.xlsx`, content: bufferOf(wbFert) },
      { filename: `ConsultoriaRural-${hojeStr()}.xlsx`, content: bufferOf(wbConsult) },
      { filename: `HistoricoPrecos-${hojeStr()}.xlsx`, content: bufferOf(wbPrecos) },
    ],
  });
  console.log('Backup enviado com sucesso para', EMAIL_DESTINO);
}

main().catch(err => { console.error('ERRO NO BACKUP:', err); process.exit(1); });
