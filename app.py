import React, { useState, useMemo } from 'react';
import { Upload, FileSpreadsheet, TrendingUp, Users, Download, Search, LogIn, Filter, X } from 'lucide-react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

const DashboardBI = () => {
  const [autenticado, setAutenticado] = useState(False);
  const [senha, setSenha] = useState('');
  const [dadosBrutos, setDadosBrutos] = useState([]);
  const [carregando, setCarregando] = useState(False);
  
  const [dataInicial, setDataInicial] = useState('');
  const [dataFinal, setDataFinal] = useState('');
  const [vendedorFiltro, setVendedorFiltro] = useState('');
  const [estadoFiltro, setEstadoFiltro] = useState('');
  const [mesFiltro, setMesFiltro] = useState('');
  const [anoFiltro, setAnoFiltro] = useState('');
  const [buscaCliente, setBuscaCliente] = useState('');
  const [clienteSelecionado, setClienteSelecionado] = useState(null);
  const [mostrarSugestoes, setMostrarSugestoes] = useState(False);
  const [telaAtiva, setTelaAtiva] = useState('dashboard');
  const [modoDebug, setModoDebug] = useState(False);
  const [vendedorExpandido, setVendedorExpandido] = useState(null);
  const [topClientesQtd, setTopClientesQtd] = useState(10);

  const handleLogin = () => {
    if (senha === 'admin123') {
      setAutenticado(True);
    } else {
      alert('Senha incorreta!');
    }
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setCarregando(True);
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);
        
        setDadosBrutos(data);
        setCarregando(False);
      } catch (error) {
        alert('Erro ao processar arquivo Excel');
        setCarregando(False);
      }
    };

    reader.readAsBinaryString(file);
  };

  const dadosProcessados = useMemo(() => {
    return dadosBrutos.map(row => {
      // Converter data do Excel (pode vir como número serial)
      let dataConvertida = null;
      if (row.DataEmissao) {
        # Se for número serial do Excel (dias desde 1900-01-01)
        if (typeof row.DataEmissao === 'number') {
          dataConvertida = new Date((row.DataEmissao - 25569) * 86400 * 1000);
        } else {
          // Se for string, tentar converter
          dataConvertida = new Date(row.DataEmissao);
        }
      }

      // Garantir que TotalProduto seja número
      const totalProduto = parseFloat(row.TotalProduto) || 0;
      
      // Calcular Valor_Real
      const valorReal = row.TipoMov === 'NF Venda' ? totalProduto : -totalProduto;

      return {
        ...row,
        TotalProduto: totalProduto,
        Valor_Real: valorReal,
        DataEmissaoFormatada: dataConvertida,
        Mes: dataConvertida ? dataConvertida.getMonth() + 1 : null,
        Ano: dataConvertida ? dataConvertida.getFullYear() : null
      };
    });
  }, [dadosBrutos]);

  const dadosFiltrados = useMemo(() => {
    let dados = dadosProcessados;

    if (dataInicial) {
      const dtInicial = new Date(dataInicial);
      dados = dados.filter(d => d.DataEmissaoFormatada >= dtInicial);
    }

    if (dataFinal) {
      const dtFinal = new Date(dataFinal);
      dados = dados.filter(d => d.DataEmissaoFormatada <= dtFinal);
    }

    if (vendedorFiltro) {
      dados = dados.filter(d => d.Vendedor === vendedorFiltro);
    }

    if (estadoFiltro) {
      dados = dados.filter(d => d.Estado === estadoFiltro);
    }

    if (mesFiltro) {
      dados = dados.filter(d => d.Mes === parseInt(mesFiltro));
    }

    if (anoFiltro) {
      dados = dados.filter(d => d.Ano === parseInt(anoFiltro));
    }

    return dados;
  }, [dadosProcessados, dataInicial, dataFinal, vendedorFiltro, estadoFiltro, mesFiltro, anoFiltro]);

  const notasUnicas = useMemo(() => {
    const notas = new Map();
    dadosFiltrados.forEach(row => {
      if (!notas.has(row.Numero_NF)) {
        notas.set(row.Numero_NF, row);
      }
    });
    return Array.from(notas.values());
  }, [dadosFiltrados]);

  const faturamentoTotal = useMemo(() => {
    return notasUnicas.reduce((acc, row) => acc + (row.Valor_Real || 0), 0);
  }, [notasUnicas]);

  const clientesUnicos = useMemo(() => {
    const clientes = new Set(dadosFiltrados.map(d => d.CPF_CNPJ));
    return clientes.size;
  }, [dadosFiltrados]);

  const vendedores = useMemo(() => {
    return [...new Set(dadosProcessados.map(d => d.Vendedor))].filter(Boolean).sort();
  }, [dadosProcessados]);

  const estados = useMemo(() => {
    return [...new Set(dadosProcessados.map(d => d.Estado))].filter(Boolean).sort();
  }, [dadosProcessados]);

  const dadosTemporais = useMemo(() => {
    const grupos = {};
    notasUnicas.forEach(row => {
      if (row.DataEmissaoFormatada) {
        const mes = `${row.Ano}-${String(row.Mes).padStart(2, '0')}`;
        if (!grupos[mes]) grupos[mes] = 0;
        grupos[mes] += row.Valor_Real;
      }
    });
    return Object.entries(grupos)
      .map(([mes, valor]) => ({ mes, valor }))
      .sort((a, b) => a.mes.localeCompare(b.mes));
  }, [notasUnicas]);

  const vendasPorEstado = useMemo(() => {
    const grupos = {};
    notasUnicas.forEach(row => {
      const estado = row.Estado || 'Sem Estado';
      if (!grupos[estado]) {
        grupos[estado] = { 
          estado, 
          valor: 0,
          count: 0 
        };
      }
      grupos[estado].valor += (row.Valor_Real || 0);
      grupos[estado].count += 1;
    });
    return Object.values(grupos)
      .sort((a, b) => b.valor - a.valor)
      .slice(0, 10);
  }, [notasUnicas]);

  const relatorioPositivacao = useMemo(() => {
    const todosClientes = new Set(dadosProcessados.map(d => d.CPF_CNPJ));
    const grupos = {};

    dadosFiltrados.forEach(row => {
      if (row.TipoMov === 'NF Venda') {
        if (!grupos[row.Vendedor]) {
          grupos[row.Vendedor] = {
            vendedor: row.Vendedor,
            clientes: new Set(),
            clientesLista: []
          };
        }
        grupos[row.Vendedor].clientes.add(row.CPF_CNPJ);
        if (!grupos[row.Vendedor].clientesLista.find(c => c.cpf === row.CPF_CNPJ)) {
          grupos[row.Vendedor].clientesLista.push({
            cpf: row.CPF_CNPJ,
            razao: row.RazaoSocial,
            cidade: row.Cidade,
            estado: row.Estado
          });
        }
      }
    });

    return Object.values(grupos).map(g => ({
      vendedor: g.vendedor,
      qtdAtendidos: g.clientes.size,
      totalBase: todosClientes.size,
      percentual: ((g.clientes.size / todosClientes.size) * 100).toFixed(1),
      clientes: g.clientesLista
    }));
  }, [dadosProcessados, dadosFiltrados]);

  const clientesSemCompra = useMemo(() => {
    const clientesComVenda = new Set(
      dadosFiltrados.filter(d => d.TipoMov === 'NF Venda').map(d => d.CPF_CNPJ)
    );
    
    const todosClientes = {};
    dadosProcessados.forEach(row => {
      if (!todosClientes[row.CPF_CNPJ]) {
        todosClientes[row.CPF_CNPJ] = {
          cpf: row.CPF_CNPJ,
          razao: row.RazaoSocial,
          cidade: row.Cidade,
          estado: row.Estado,
          valorHistorico: 0
        };
      }
      if (row.TipoMov === 'NF Venda') {
        todosClientes[row.CPF_CNPJ].valorHistorico += row.TotalProduto;
      }
    });

    return Object.values(todosClientes)
      .filter(c => !clientesComVenda.has(c.cpf))
      .sort((a, b) => b.valorHistorico - a.valorHistorico);
  }, [dadosProcessados, dadosFiltrados]);

  const historicoCliente = useMemo(() => {
    if (!buscaCliente) return [];
    return dadosProcessados.filter(d => 
      d.RazaoSocial?.toLowerCase().includes(buscaCliente.toLowerCase())
    ).sort((a, b) => (b.DataEmissaoFormatada || 0) - (a.DataEmissaoFormatada || 0));
  }, [dadosProcessados, buscaCliente]);

  const downloadCSV = (data, filename) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dados');
    XLSX.writeFile(wb, `${filename}.xlsx`);
  };

  if (!autenticado) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-500 to-blue-700 flex items-center justify-center p-4">
        <div className="bg-white rounded-lg shadow-2xl p-8 w-full max-w-md">
          <div className="flex justify-center mb-6">
            <LogIn className="w-16 h-16 text-blue-500" />
          </div>
          <h1 className="text-2xl font-bold text-center mb-6">Dashboard BI</h1>
          <input
            type="password"
            placeholder="Digite a senha"
            value={senha}
            onChange={(e) => setSenha(e.target.value)}
            onKeyPress={(e) => e.key === 'Enter' && handleLogin()}
            className="w-full px-4 py-2 border rounded-lg mb-4"
          />
          <button
            onClick={handleLogin}
            className="w-full bg-blue-500 text-white py-2 rounded-lg hover:bg-blue-600"
          >
            Entrar
          </button>
          <p className="text-xs text-gray-500 text-center mt-4">Senha: admin123</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-blue-600 text-white p-4 shadow-lg">
        <div className="max-w-7xl mx-auto flex justify-between items-center">
          <h1 className="text-2xl font-bold">Dashboard BI - Vendas</h1>
          <button
            onClick={() => setAutenticado(False)}
            className="px-4 py-2 bg-blue-700 rounded hover:bg-blue-800"
          >
            Sair
          </button>
        </div>
      </div>

      <div className="max-w-7xl mx-auto p-6">
        {dadosBrutos.length === 0 && (
          <div className="bg-white rounded-lg shadow p-6 mb-6">
            <h2 className="text-xl font-semibold mb-4 flex items-center gap-2">
              <Upload className="w-5 h-5" />
              Upload da Planilha
            </h2>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center">
              <FileSpreadsheet className="w-12 h-12 mx-auto mb-4 text-gray-400" />
              <label className="cursor-pointer">
                <span className="bg-blue-500 text-white px-6 py-2 rounded-lg hover:bg-blue-600 inline-block">
                  Selecionar Arquivo Excel
                </span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                />
              </label>
            </div>
          </div>
        )}

        {dadosBrutos.length > 0 && (
          <>
            <div className="bg-white rounded-lg shadow mb-6 p-4">
              <div className="flex gap-2 flex-wrap items-center justify-between">
                <div className="flex gap-2 flex-wrap">
                  <button
                    onClick={() => setTelaAtiva('dashboard')}
                    className={`px-4 py-2 rounded ${telaAtiva === 'dashboard' ? 'bg-blue-500 text-white' : 'bg-gray-200'}`}
                  >
                    Dashboard
                  </button>
                  <button
                    onClick={() => setTelaAtiva('positivacao')}
                    className={`px-4 py-2 rounded ${telaAtiva === 'positivacao' ? 'bg-blue-500 text-white' : 'bg-gray-200'}`}
                  >
                    Positivação
                  </button>
                  <button
                    onClick={() => setTelaAtiva('churn')}
                    className={`px-4 py-2 rounded ${telaAtiva === 'churn' ? 'bg-blue-500 text-white' : 'bg-gray-200'}`}
                  >
                    Clientes sem Compra
                  </button>
                  <button
                    onClick={() => setTelaAtiva('historico')}
                    className={`px-4 py-2 rounded ${telaAtiva === 'historico' ? 'bg-blue-500 text-white' : 'bg-gray-200'}`}
                  >
                    Histórico
                  </button>
                  <button
                    onClick={() => setTelaAtiva('rankings')}
                    className={`px-4 py-2 rounded ${telaAtiva === 'rankings' ? 'bg-blue-500 text-white' : 'bg-gray-200'}`}
                  >
                    Rankings
                  </button>
                </div>
                <button
                  onClick={() => setModoDebug(!modoDebug)}
                  className="px-3 py-1 text-xs bg-yellow-500 text-white rounded hover:bg-yellow-600"
                >
                  {modoDebug ? '🐛 Debug ON' : 'Debug OFF'}
                </button>
              </div>
            </div>

            {/* PAINEL DEBUG */}
            {modoDebug && (
              <div className="bg-yellow-50 border-2 border-yellow-400 rounded-lg p-4 mb-6">
                <h3 className="font-bold text-yellow-800 mb-3">🐛 Modo Debug - Diagnóstico</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
                  <div>
                    <p className="font-semibold text-gray-700">Dados Carregados:</p>
                    <p>Total de linhas: {dadosBrutos.length}</p>
                    <p>Após processamento: {dadosProcessados.length}</p>
                    <p>Após filtros: {dadosFiltrados.length}</p>
                    <p>Notas únicas: {notasUnicas.length}</p>
                  </div>
                  <div>
                    <p className="font-semibold text-gray-700">Primeira Linha:</p>
                    <p className="text-xs bg-white p-2 rounded overflow-x-auto">
                      {JSON.stringify(dadosProcessados[0], null, 2).slice(0, 300)}...
                    </p>
                  </div>
                  <div>
                    <p className="font-semibold text-gray-700">Valores Detectados:</p>
                    <p>TipoMov únicos: {[...new Set(dadosBrutos.map(d => d.TipoMov))].join(', ')}</p>
                    <p>Anos encontrados: {[...new Set(dadosProcessados.map(d => d.Ano))].filter(Boolean).sort().join(', ')}</p>
                  </div>
                  <div>
                    <p className="font-semibold text-gray-700">Exemplo de Conversão:</p>
                    {dadosProcessados[0] && (
                      <>
                        <p>DataEmissao original: {dadosBrutos[0]?.DataEmissao}</p>
                        <p>Data convertida: {dadosProcessados[0]?.DataEmissaoFormatada?.toLocaleDateString()}</p>
                        <p>Ano: {dadosProcessados[0]?.Ano}</p>
                        <p>TotalProduto: {dadosProcessados[0]?.TotalProduto}</p>
                        <p>Valor_Real: {dadosProcessados[0]?.Valor_Real}</p>
                      </>
                    )}
                  </div>
                </div>
              </div>
            )}

            <div className="bg-white rounded-lg shadow p-4 mb-6">
              <h3 className="font-semibold mb-3 flex items-center gap-2">
                <Filter className="w-4 h-4" />
                Filtros Globais
              </h3>
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
                <input
                  type="date"
                  value={dataInicial}
                  onChange={(e) => setDataInicial(e.target.value)}
                  className="px-3 py-2 border rounded text-sm"
                  placeholder="Data Inicial"
                />
                <input
                  type="date"
                  value={dataFinal}
                  onChange={(e) => setDataFinal(e.target.value)}
                  className="px-3 py-2 border rounded text-sm"
                  placeholder="Data Final"
                />
                <select
                  value={vendedorFiltro}
                  onChange={(e) => setVendedorFiltro(e.target.value)}
                  className="px-3 py-2 border rounded text-sm"
                >
                  <option value="">Todos Vendedores</option>
                  {vendedores.map(v => <option key={v} value={v}>{v}</option>)}
                </select>
                <select
                  value={estadoFiltro}
                  onChange={(e) => setEstadoFiltro(e.target.value)}
                  className="px-3 py-2 border rounded text-sm"
                >
                  <option value="">Todos Estados</option>
                  {estados.map(e => <option key={e} value={e}>{e}</option>)}
                </select>
                <select
                  value={mesFiltro}
                  onChange={(e) => setMesFiltro(e.target.value)}
                  className="px-3 py-2 border rounded text-sm"
                >
                  <option value="">Todos Meses</option>
                  {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => (
                    <option key={m} value={m}>{m}</option>
                  ))}
                </select>
                <select
                  value={anoFiltro}
                  onChange={(e) => setAnoFiltro(e.target.value)}
                  className="px-3 py-2 border rounded text-sm"
                >
                  <option value="">Todos Anos</option>
                  {[2023, 2024, 2025, 2026].map(a => (
                    <option key={a} value={a}>{a}</option>
                  ))}
                </select>
              </div>
              <button
                onClick={() => {
                  setDataInicial('');
                  setDataFinal('');
                  setVendedorFiltro('');
                  setEstadoFiltro('');
                  setMesFiltro('');
                  setAnoFiltro('');
                }}
                className="mt-3 text-sm text-blue-600 hover:text-blue-800 flex items-center gap-1"
              >
                <X className="w-4 h-4" />
                Limpar Filtros
              </button>
            </div>

            {telaAtiva === 'dashboard' && (
              <>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
                  <div className="bg-white rounded-lg shadow p-6">
                    <div className="flex items-center justify-between">
                      <div>
                        <p className="text-gray-600 text-sm">Faturamento Total Líquido</p>
                        <p className="text-3xl font-bold text-green-600">
                          R$ {faturamentoTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                        </p>
                      </div>
                      <TrendingUp className="w-12 h-12 text-green-500" />
                    </div>
                  </div>
                  <div className="bg-white rounded-lg shadow p-6">
                    <div className="flex items-center justify-between">
                      <div>
                        <p className="text-gray-600 text-sm">Clientes Únicos</p>
                        <p className="text-3xl font-bold text-blue-600">{clientesUnicos}</p>
                      </div>
                      <Users className="w-12 h-12 text-blue-500" />
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
                  <div className="bg-white rounded-lg shadow p-6">
                    <h3 className="font-semibold mb-4">Evolução de Vendas</h3>
                    <ResponsiveContainer width="100%" height={250}>
                      <LineChart data={dadosTemporais}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis dataKey="mes" />
                        <YAxis />
                        <Tooltip />
                        <Line type="monotone" dataKey="valor" stroke="#3B82F6" />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="bg-white rounded-lg shadow p-6">
                    <h3 className="font-semibold mb-4">Top 10 Estados</h3>
                    <ResponsiveContainer width="100%" height={250}>
                      <BarChart data={vendasPorEstado}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis dataKey="estado" />
                        <YAxis />
                        <Tooltip 
                          formatter={(value) => `R$ ${value.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`}
                        />
                        <Bar dataKey="valor" fill="#10B981" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </>
            )}

            {telaAtiva === 'positivacao' && (
              <div className="bg-white rounded-lg shadow p-6">
                <div className="flex justify-between items-center mb-4">
                  <h3 className="font-semibold text-lg">Relatório de Positivação por Vendedor</h3>
                  <button
                    onClick={() => {
                      const exportData = relatorioPositivacao.flatMap(item => 
                        item.clientes.map(c => ({
                          Vendedor: item.vendedor,
                          QtdAtendidos: item.qtdAtendidos,
                          TotalBase: item.totalBase,
                          Percentual: item.percentual + '%',
                          Cliente_RazaoSocial: c.razao,
                          Cliente_CPFCNPJ: c.cpf,
                          Cliente_Cidade: c.cidade,
                          Cliente_Estado: c.estado,
                          Cliente_ValorTotal: c.valorTotal
                        }))
                      );
                      downloadCSV(exportData, 'positivacao_detalhada');
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
                  >
                    <Download className="w-4 h-4" />
                    Exportar Completo
                  </button>
                </div>

                {/* TABELA RESUMIDA */}
                <div className="overflow-x-auto mb-6">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-100">
                      <tr>
                        <th className="px-4 py-2 text-left">Vendedor</th>
                        <th className="px-4 py-2 text-left">Clientes Atendidos</th>
                        <th className="px-4 py-2 text-left">Total Base do Vendedor</th>
                        <th className="px-4 py-2 text-left">% Positivação</th>
                        <th className="px-4 py-2 text-left">Ações</th>
                      </tr>
                    </thead>
                    <tbody>
                      {relatorioPositivacao.map((item, idx) => (
                        <tr key={idx} className="border-t hover:bg-gray-50">
                          <td className="px-4 py-2 font-semibold">{item.vendedor}</td>
                          <td className="px-4 py-2">{item.qtdAtendidos}</td>
                          <td className="px-4 py-2">{item.totalBase}</td>
                          <td className="px-4 py-2">
                            <span className="px-2 py-1 bg-blue-100 text-blue-800 rounded">
                              {item.percentual}%
                            </span>
                          </td>
                          <td className="px-4 py-2">
                            <button
                              onClick={() => setVendedorExpandido(vendedorExpandido === item.vendedor ? null : item.vendedor)}
                              className="px-3 py-1 bg-blue-500 text-white rounded text-xs hover:bg-blue-600"
                            >
                              {vendedorExpandido === item.vendedor ? 'Ocultar Detalhes' : 'Ver Detalhes'}
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* TABELA DETALHADA POR VENDEDOR */}
                {vendedorExpandido && (
                  <div className="border-t-4 border-blue-500 pt-6">
                    <div className="flex justify-between items-center mb-4">
                      <h4 className="font-semibold text-lg">
                        Detalhamento - {vendedorExpandido}
                      </h4>
                      <button
                        onClick={() => {
                          const vendedor = relatorioPositivacao.find(v => v.vendedor === vendedorExpandido);
                          if (vendedor) {
                            downloadCSV(vendedor.clientes.map(c => ({
                              RazaoSocial: c.razao,
                              CPFCNPJ: c.cpf,
                              Cidade: c.cidade,
                              Estado: c.estado,
                              ValorTotal: c.valorTotal
                            })), `clientes_${vendedorExpandido}`);
                          }
                        }}
                        className="flex items-center gap-2 px-3 py-1 bg-green-500 text-white rounded text-sm hover:bg-green-600"
                      >
                        <Download className="w-3 h-3" />
                        Exportar
                      </button>
                    </div>

                    <div className="bg-blue-50 p-4 rounded-lg mb-4">
                      <p className="text-sm">
                        <strong>Total de Clientes Atendidos:</strong> {relatorioPositivacao.find(v => v.vendedor === vendedorExpandido)?.clientes.length || 0}
                      </p>
                      <p className="text-sm">
                        <strong>Valor Total Vendido:</strong> R$ {
                          ((relatorioPositivacao.find(v => v.vendedor === vendedorExpandido)?.clientes || [])
                            .reduce((acc, c) => acc + (c.valorTotal || 0), 0))
                            .toLocaleString('pt-BR', { minimumFractionDigits: 2 })
                        }
                      </p>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead className="bg-gray-200">
                          <tr>
                            <th className="px-4 py-2 text-left">Razão Social</th>
                            <th className="px-4 py-2 text-left">CPF/CNPJ</th>
                            <th className="px-4 py-2 text-left">Cidade</th>
                            <th className="px-4 py-2 text-left">Estado</th>
                            <th className="px-4 py-2 text-left">Valor Total</th>
                          </tr>
                        </thead>
                        <tbody>
                          {relatorioPositivacao
                            .find(v => v.vendedor === vendedorExpandido)
                            ?.clientes.map((cliente, idx) => (
                              <tr key={idx} className="border-t hover:bg-gray-50">
                                <td className="px-4 py-2">{cliente.razao}</td>
                                <td className="px-4 py-2">{cliente.cpf}</td>
                                <td className="px-4 py-2">{cliente.cidade}</td>
                                <td className="px-4 py-2">{cliente.estado}</td>
                                <td className="px-4 py-2 font-semibold text-green-700">
                                  R$ {(cliente.valorTotal || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                </td>
                              </tr>
                            ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}

            {telaAtiva === 'churn' && (
              <div className="bg-white rounded-lg shadow p-6">
                <div className="flex justify-between items-center mb-4">
                  <h3 className="font-semibold text-lg">Clientes sem Compra no Período</h3>
                  <button
                    onClick={() => downloadCSV(clientesSemCompra, 'clientes_sem_compra')}
                    className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
                  >
                    <Download className="w-4 h-4" />
                    Exportar
                  </button>
                </div>
                <p className="text-sm text-gray-600 mb-4">
                  Total: {clientesSemCompra.length} clientes | Valor Potencial Perdido: R$ {clientesSemCompra.reduce((acc, c) => acc + c.valorHistorico, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-100">
                      <tr>
                        <th className="px-4 py-2 text-left">Razão Social</th>
                        <th className="px-4 py-2 text-left">CPF/CNPJ</th>
                        <th className="px-4 py-2 text-left">Vendedor</th>
                        <th className="px-4 py-2 text-left">Cidade</th>
                        <th className="px-4 py-2 text-left">Estado</th>
                        <th className="px-4 py-2 text-left">Valor Histórico</th>
                      </tr>
                    </thead>
                    <tbody>
                      {clientesSemCompra.slice(0, 50).map((cliente, idx) => (
                        <tr key={idx} className="border-t hover:bg-gray-50">
                          <td className="px-4 py-2">{cliente.razao}</td>
                          <td className="px-4 py-2">{cliente.cpf}</td>
                          <td className="px-4 py-2 font-semibold text-blue-700">{cliente.vendedor}</td>
                          <td className="px-4 py-2">{cliente.cidade}</td>
                          <td className="px-4 py-2">{cliente.estado}</td>
                          <td className="px-4 py-2">R$ {cliente.valorHistorico.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {telaAtiva === 'rankings' && (
              <>
                {/* RANKING DE VENDEDORES */}
                <div className="bg-white rounded-lg shadow p-6 mb-6">
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="font-semibold text-lg">Ranking de Vendedores por Valor</h3>
                    <button
                      onClick={() => downloadCSV(rankingVendedores, 'ranking_vendedores')}
                      className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
                    >
                      <Download className="w-4 h-4" />
                      Exportar
                    </button>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead className="bg-gray-100">
                        <tr>
                          <th className="px-4 py-2 text-left">#</th>
                          <th className="px-4 py-2 text-left">Vendedor</th>
                          <th className="px-4 py-2 text-left">Valor Total</th>
                          <th className="px-4 py-2 text-left">Qtd Notas</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rankingVendedores.map((item, idx) => (
                          <tr key={idx} className="border-t hover:bg-gray-50">
                            <td className="px-4 py-2 font-bold text-gray-500">#{idx + 1}</td>
                            <td className="px-4 py-2 font-semibold">{item.vendedor}</td>
                            <td className="px-4 py-2 text-green-700 font-bold">
                              R$ {item.valor.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </td>
                            <td className="px-4 py-2">{item.qtdNotas}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* RANKING DE CLIENTES */}
                <div className="bg-white rounded-lg shadow p-6">
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="font-semibold text-lg">Ranking de Clientes por Valor</h3>
                    <div className="flex gap-2 items-center">
                      <select
                        value={topClientesQtd}
                        onChange={(e) => setTopClientesQtd(parseInt(e.target.value))}
                        className="px-3 py-2 border rounded text-sm"
                      >
                        <option value={10}>Top 10</option>
                        <option value={20}>Top 20</option>
                        <option value={50}>Top 50</option>
                      </select>
                      <button
                        onClick={() => downloadCSV(rankingClientes.slice(0, topClientesQtd), 'ranking_clientes')}
                        className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
                      >
                        <Download className="w-4 h-4" />
                        Exportar
                      </button>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead className="bg-gray-100">
                        <tr>
                          <th className="px-4 py-2 text-left">#</th>
                          <th className="px-4 py-2 text-left">Razão Social</th>
                          <th className="px-4 py-2 text-left">CPF/CNPJ</th>
                          <th className="px-4 py-2 text-left">Cidade</th>
                          <th className="px-4 py-2 text-left">Estado</th>
                          <th className="px-4 py-2 text-left">Valor Total</th>
                          <th className="px-4 py-2 text-left">Qtd Notas</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rankingClientes.slice(0, topClientesQtd).map((item, idx) => (
                          <tr key={idx} className="border-t hover:bg-gray-50">
                            <td className="px-4 py-2 font-bold text-gray-500">#{idx + 1}</td>
                            <td className="px-4 py-2 font-semibold">{item.razao}</td>
                            <td className="px-4 py-2">{item.cpf}</td>
                            <td className="px-4 py-2">{item.cidade}</td>
                            <td className="px-4 py-2">{item.estado}</td>
                            <td className="px-4 py-2 text-green-700 font-bold">
                              R$ {item.valor.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </td>
                            <td className="px-4 py-2">{item.qtdNotas}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}

            {telaAtiva === 'historico' && (
              <div className="bg-white rounded-lg shadow p-6">
                <h3 className="font-semibold text-lg mb-4">Histórico de Vendas por Cliente</h3>
                
                {/* BUSCA COM AUTOCOMPLETE */}
                <div className="mb-4">
                  <div className="flex gap-2">
                    <div className="relative flex-1">
                      <Search className="absolute left-3 top-3 w-4 h-4 text-gray-400" />
                      <input
                        type="text"
                        placeholder="Buscar por Razão Social ou CPF/CNPJ..."
                        value={buscaCliente}
                        onChange={(e) => {
                          setBuscaCliente(e.target.value);
                          setClienteSelecionado(null);
                          setMostrarSugestoes(True);
                        }}
                        onFocus={() => setMostrarSugestoes(True)}
                        className="w-full pl-10 pr-4 py-2 border rounded"
                      />
                      
                      {/* LISTA DE SUGESTÕES */}
                      {mostrarSugestoes && sugestoesClientes.length > 0 && (
                        <div className="absolute z-10 w-full mt-1 bg-white border rounded-lg shadow-lg max-h-60 overflow-y-auto">
                          {sugestoesClientes.map((cliente, idx) => (
                            <div
                              key={idx}
                              onClick={() => {
                                setClienteSelecionado(cliente.razao);
                                setBuscaCliente(cliente.razao);
                                setMostrarSugestoes(False);
                              }}
                              className="px-4 py-3 hover:bg-blue-50 cursor-pointer border-b last:border-b-0"
                            >
                              <p className="font-semibold text-sm">{cliente.razao}</p>
                              <p className="text-xs text-gray-600">
                                CPF/CNPJ: {cliente.cpf} | {cliente.cidade} - {cliente.estado}
                              </p>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    
                    {(buscaCliente || clienteSelecionado) && (
                      <button
                        onClick={() => {
                          setBuscaCliente('');
                          setClienteSelecionado(null);
                          setMostrarSugestoes(False);
                        }}
                        className="px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600"
                      >
                        Limpar
                      </button>
                    )}
                    
                    {historicoCliente.length > 0 && (
                      <button
                        onClick={() => downloadCSV(historicoCliente, 'historico_cliente')}
                        className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
                      >
                        <Download className="w-4 h-4" />
                        Exportar
                      </button>
                    )}
                  </div>
                  
                  {/* INFO DO CLIENTE SELECIONADO */}
                  {clienteSelecionado && historicoCliente.length > 0 && (
                    <div className="mt-3 p-3 bg-blue-50 border border-blue-200 rounded">
                      <p className="text-sm">
                        <strong>Cliente:</strong> {historicoCliente[0].RazaoSocial}
                      </p>
                      <p className="text-sm">
                        <strong>CPF/CNPJ:</strong> {historicoCliente[0].CPF_CNPJ} | 
                        <strong> Cidade:</strong> {historicoCliente[0].Cidade} - {historicoCliente[0].Estado}
                      </p>
                      <p className="text-sm">
                        <strong>Total de Registros:</strong> {historicoCliente.length}
                      </p>
                    </div>
                  )}
                </div>

                {/* TABELA DE HISTÓRICO */}
                {historicoCliente.length > 0 ? (
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead className="bg-gray-100">
                        <tr>
                          <th className="px-4 py-2 text-left">Data</th>
                          <th className="px-4 py-2 text-left">Tipo</th>
                          <th className="px-4 py-2 text-left">NF</th>
                          <th className="px-4 py-2 text-left">Código</th>
                          <th className="px-4 py-2 text-left">Produto</th>
                          <th className="px-4 py-2 text-left">Qtd</th>
                          <th className="px-4 py-2 text-left">Valor</th>
                        </tr>
                      </thead>
                      <tbody>
                        {historicoCliente.map((item, idx) => (
                          <tr key={idx} className="border-t hover:bg-gray-50">
                            <td className="px-4 py-2">{item.DataEmissaoFormatada?.toLocaleDateString('pt-BR') || item.DataEmissao}</td>
                            <td className="px-4 py-2">
                              <span className={`px-2 py-1 rounded text-xs ${item.TipoMov === 'NF Venda' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
                                {item.TipoMov}
                              </span>
                            </td>
                            <td className="px-4 py-2">{item.Numero_NF}</td>
                            <td className="px-4 py-2">{item.CodigoProduto}</td>
                            <td className="px-4 py-2">{item.NomeProduto}</td>
                            <td className="px-4 py-2">{item.Quantidade}</td>
                            <td className="px-4 py-2">R$ {(item.TotalProduto || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : buscaCliente ? (
                  <div className="text-center py-8 text-gray-500">
                    <Search className="w-12 h-12 mx-auto mb-2 text-gray-300" />
                    <p>Nenhum resultado encontrado para "{buscaCliente}"</p>
                    <p className="text-sm">Tente digitar o nome completo ou parte do CPF/CNPJ</p>
                  </div>
                ) : (
                  <div className="text-center py-8 text-gray-400">
                    <Search className="w-12 h-12 mx-auto mb-2 text-gray-300" />
                    <p>Digite o nome do cliente para ver o histórico</p>
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};


export default DashboardBI;
