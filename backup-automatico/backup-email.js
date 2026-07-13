// ============================================================
// TERRA & SAFRA — Backup diário automático por e-mail
// Roda sozinho via GitHub Actions, não precisa do navegador aberto.
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

// ---------- Helpers (espelham a lógica do app) ----------
const TIPO_LABEL = { remedio: 'Medicamentos', racao: 'Rações', proteinado: 'Nutrição Animal' };
function categoriaLabel(tipo) { return TIPO_LABEL[tipo] || tipo; }
function isDescricaoTipo(tipo) { return tipo === 'remedio' || tipo === 'racao' || tipo === 'proteinado'; }
function tipoLabel(tipo) { return tipo === 'remedio' ? 'Medicamento' : tipo === 'racao' ? 'Ração' : 'Nutrição Animal'; }

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
  ]);
  const erro = e1 || e2 || e3 || e4 || e5 || e6 || e7 || e8 || e9 || e10 || e11 || e12 || e13;
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
  function totalVendaFiado(v) { return v.itens.reduce((s, it) => s + Number(it.quantidade) * Number(it.valor_unitario), 0); }
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
  function totalDespesasPagasPor(nome) { return despesas.filter(d => d.responsavel === nome).reduce((s, d) => s + Number(d.valor), 0); }

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

  console.log('Montando planilha Financeira...');
  const wbFin = XLSX.utils.book_new();
  const totalEntradasVendas = lotes.reduce((s, l) => s + totalVendido(l), 0);
  const totalEntradasExtra = entradasExtra.reduce((s, e) => s + Number(e.valor), 0);
  const totalEntradasGeral = totalEntradasVendas + totalEntradasExtra;
  const totalGastoProduto = lotes.reduce((s, l) => s + totalInvestido(l), 0);
  const totalDespesasManual = despesas.filter(d => !d.eh_imposto).reduce((s, d) => s + Number(d.valor), 0);
  const totalImpostosManual = despesas.filter(d => d.eh_imposto).reduce((s, d) => s + Number(d.valor), 0);
  const totalDespesasGeral = totalGastoProduto + totalDespesasManual;
  const totalSaidasGeral = totalDespesasGeral + totalImpostosManual;
  const lucroLiquido = totalEntradasGeral - totalSaidasGeral;

  bookAppend(wbFin, [
    { Item: 'Total de entradas', 'Valor (R$)': Number(totalEntradasGeral.toFixed(2)) },
    { Item: '  · Venda de produtos (Estoque)', 'Valor (R$)': Number(totalEntradasVendas.toFixed(2)) },
    { Item: '  · Entradas atípicas', 'Valor (R$)': Number(totalEntradasExtra.toFixed(2)) },
    { Item: 'Despesas (produto + operacionais)', 'Valor (R$)': Number(totalDespesasGeral.toFixed(2)) },
    { Item: '  · Gasto com produto (Estoque)', 'Valor (R$)': Number(totalGastoProduto.toFixed(2)) },
    { Item: '  · Despesas manuais', 'Valor (R$)': Number(totalDespesasManual.toFixed(2)) },
    { Item: 'Impostos (DARF)', 'Valor (R$)': Number(totalImpostosManual.toFixed(2)) },
    { Item: 'Total de saídas', 'Valor (R$)': Number(totalSaidasGeral.toFixed(2)) },
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
  const formaLabel = { dinheiro: 'Dinheiro', pix: 'PIX', cartao_credito: 'Cartão Crédito', cartao_debito: 'Cartão Débito' };
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
    text: `Segue em anexo o backup automático de hoje (${dataHoje}):\n\n- Estoque.xlsx\n- Vencimentos.xlsx\n- Financeiro.xlsx\n- VendasFiado.xlsx\n- VendasAvista.xlsx\n- HistoricoPrecos.xlsx\n\nEsse e-mail é gerado e enviado sozinho, todo dia às 21h, direto do sistema.`,
    attachments: [
      { filename: `Estoque-${hojeStr()}.xlsx`, content: bufferOf(wbEstoque) },
      { filename: `Vencimentos-${hojeStr()}.xlsx`, content: bufferOf(wbVenc) },
      { filename: `Financeiro-${hojeStr()}.xlsx`, content: bufferOf(wbFin) },
      { filename: `VendasFiado-${hojeStr()}.xlsx`, content: bufferOf(wbFiado) },
      { filename: `VendasAvista-${hojeStr()}.xlsx`, content: bufferOf(wbAvista) },
      { filename: `HistoricoPrecos-${hojeStr()}.xlsx`, content: bufferOf(wbPrecos) },
    ],
  });
  console.log('Backup enviado com sucesso para', EMAIL_DESTINO);
}

main().catch(err => { console.error('ERRO NO BACKUP:', err); process.exit(1); });
