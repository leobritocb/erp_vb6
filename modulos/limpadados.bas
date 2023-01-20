Attribute VB_Name = "limpadados"
Sub limpaCadCliente()
cadCliente.txbCod = ultimoRegistroN("TB_Clientes")
cadCliente.txbData = Date

cadCliente.txbNome = ""
cadCliente.txbApelido = ""
cadCliente.txbDtNascimento = ""
cadCliente.txbCelular = ""
cadCliente.txbEmail = ""
cadCliente.txbCpf = ""
cadCliente.txbRg = ""
cadCliente.txbTelefone = ""
cadCliente.txbDtNascimento = ""
cadCliente.cbxEstCivil = ""
cadCliente.cbxSexo = ""

cadCliente.txbRua = ""
cadCliente.txbNumeroEnd = ""
cadCliente.txbComplemento = ""
cadCliente.txbBairro = ""
cadCliente.txbCidade = ""
cadCliente.txbCep = ""
cadCliente.cbxUf = ""

cadCliente.txbLimiteCredito = "0"

cadCliente.txbObs = ""

cadCliente.optBtnFisico.Value = True

End Sub

Sub limpaListClientes()
    clistClientes.txbNome = ""
    clistClientes.txbApelido = ""
    clistClientes.cbxSexo = ""
    clistClientes.cbxEstCivil = ""
    clistClientes.txbRg = ""
    clistClientes.txbCpf = ""
    clistClientes.txbDtNascimento = ""
    clistClientes.txbRua = ""
    clistClientes.txbNumeroEnd = ""
    clistClientes.txbComplemento = ""
    clistClientes.txbCidade = ""
    clistClientes.txbBairro = ""
    clistClientes.txbCep = ""
    'clistClientes.lblCodUf = ""
    clistClientes.txbTelefone = ""
    clistClientes.txbCelular = ""
    clistClientes.txbEmail = ""
    clistClientes.txbData = ""
    clistClientes.txbLimiteCredito = ""
    clistClientes.txbUltimaCompra = ""
    clistClientes.txbObs = ""
End Sub

Sub limpaCadEntregador()
cadEntregador.txbCod = ""
cadEntregador.txbData = Date

cadEntregador.txbNome = ""
cadEntregador.txbCelular = ""
cadEntregador.txbCpf = ""
cadEntregador.txbTelefone = ""

cadEntregador.txbRua = ""
cadEntregador.txbNumeroEnd = ""
cadEntregador.txbComplemento = ""
cadEntregador.txbBairro = ""
cadEntregador.txbCidade = ""
cadEntregador.txbCep = ""
cadEntregador.cbxUf = ""

cadEntregador.txbVeiculo = ""
cadEntregador.txbPlaca = ""
cadEntregador.txbModelo = ""
cadEntregador.txbMarca = ""

cadEntregador.txbObs = ""


cadEntregador.txbCod = 0
End Sub

Sub limpaCadFornecedor()
cadFornecedor.txbCod = ultimoRegistroN("TB_Fornecedor")
cadFornecedor.txbData = Date

cadFornecedor.txbNome = ""
cadFornecedor.txbApelido = ""
cadFornecedor.txbCelular = ""
cadFornecedor.txbEmail = ""
cadFornecedor.txbCpf = ""
cadFornecedor.txbTelefone = ""

cadFornecedor.txbRua = ""
cadFornecedor.txbNumeroEnd = ""
cadFornecedor.txbComplemento = ""
cadFornecedor.txbBairro = ""
cadFornecedor.txbCidade = ""
cadFornecedor.txbCep = ""
cadFornecedor.cbxUf = ""

cadFornecedor.txbObs = ""

End Sub

Sub limpaCadProdutos()
cadProduto.txbCodProduto = ultimoRegistroN("TB_Produtos")

cadProduto.txbCodBarras = ""
cadProduto.txbDescricao = ""
cadProduto.cbxUn = ""
cadProduto.cbxMarca = ""
cadProduto.cbxCategoria = ""

cadProduto.txbPrecoCusto = 0
cadProduto.txbPrecoVenda = 0
cadProduto.txbEstoque = ""
cadProduto.txbLucroPorc = 0
cadProduto.txbPeso = ""
End Sub

Sub limpaPDV()
pdv.listVenda.ListItems.Clear
pdv.codProduto = ""
pdv.totalVenda = "0,00"
pdv.lblStatuscaixa.Caption = "CAIXA ABERTO"
pdv.lblStatuscaixa.BackColor = &HC000&
pdv.lblOperador.Caption = "PADRÃO"
pdv.lblCliente.Caption = "CLIENTE NÃO INFORMADO"
pdv.lblEntregador.Caption = "SEM ENTREGA"
pdv.lblVendedor.Caption = "VENDEDOR NÃO INFORMADO"
End Sub

Sub iniciaVendaPDV()
    pdv.listVenda.ListItems.Clear
    pdv.listVenda.ColumnHeaders.Clear
    pdv.listVenda.View = lvwReport
    pdv.listVenda.GridLines = False
    pdv.listVenda.ColumnHeaders.Add , , "#"
    pdv.listVenda.ColumnHeaders.Add , , "COD"
    pdv.listVenda.ColumnHeaders.Add , , "Produto"
    pdv.listVenda.ColumnHeaders.Add , , "Un"
    pdv.listVenda.ColumnHeaders.Add , , "Qt"
    pdv.listVenda.ColumnHeaders.Add , , "Unt"
    pdv.listVenda.ColumnHeaders.Add , , "Desc"
    pdv.listVenda.ColumnHeaders.Add , , "Total"
    pdv.listVenda.ColumnHeaders(1).Width = 17
    pdv.listVenda.ColumnHeaders(2).Width = 56
    pdv.listVenda.ColumnHeaders(3).Width = 188
    pdv.listVenda.ColumnHeaders(4).Width = 30
    pdv.listVenda.ColumnHeaders(5).Width = 33
    pdv.listVenda.ColumnHeaders(6).Width = 50
    pdv.listVenda.ColumnHeaders(7).Width = 50
    pdv.listVenda.ColumnHeaders(8).Width = 43
Total = 0
pdv.totalVenda = Format(0, "0.00")
pdv.vlrunProduto = "0,00"
pdv.vlrTotalProduto = "0,00"
pdv.quantProduto = 1
pdv.lblStatuscaixa.Caption = "CAIXA ABERTO"
pdv.lblStatuscaixa.BackColor = &H8000&
End Sub


Sub limpaAvista()
    pagamento.cbxFormaPagamentoAvista = "DINHEIRO"
    pagamento.txbRecebidoAvista = pagamento.lblAreceberGeral.Caption

End Sub

Sub limpaAprazo()
    'pagamento.cbxFormaPagamentoPrazo = "CARNÊ"
   ' pagamento.cbxFormaEntrada = "DINHEIRO"
    pagamento.txbEntrada = "0,00"
    pagamento.txbAcrescimoPrazo = "0,00"
    pagamento.lblValorParceladoPrazo.Caption = "0,00"
    pagamento.txbNumParcelasPrazo = 1
    pagamento.txbPrazoDias = 30
    pagamento.optValorPrazo.Value = True
    pagamento.lblTotalPrazo.Caption = pagamento.lblAreceberGeral.Caption
    pagamento.lblValorParceladoPrazo.Caption = pagamento.lblTotalPrazo.Caption
    pagamento.txbDT1Vencimento = Date

    colunas = Array("ID", "Pagamento", "Parcela", "Vencimento", "Valor")
    Call limpaLV(pagamento.lvPrazo, colunas)

    pagamento.btnAprazo.Enabled = True
End Sub

Sub limpaCartao()
    'pagamento.cbxTipoCartao.Value = "CRÉDITO"
    pagamento.cbxBandeira = ""
    pagamento.cbxFinanceira = ""
    pagamento.txbAcrescimoCartao = "0,00"
    pagamento.cbxParcelasCartao.Clear
    pagamento.lblTotalCartao.Caption = pagamento.lblAreceberGeral.Caption
    pagamento.optValorCartao.Value = True
    calculaParcelas
    pagamento.btnCartao.Enabled = True
End Sub

Sub limpaAdOutros()
    'pagamento.cbxFormaPagamentoOutros = "BOLETO"
    pagamento.txbDataOutros = Date
    pagamento.cbxMetodo.ListIndex = 0
    pagamento.txbValorOutros = pagamento.lblAreceberGeral.Caption
    pagamento.cbxFormaPagamentoOutros.SetFocus
End Sub

Sub limpaOutros()
    'pagamento.cbxFormaPagamentoOutros = "BOLETO"
    pagamento.txbDataOutros = Date
    pagamento.cbxMetodo = "A VISTA"
    pagamento.txbValorOutros = pagamento.lblAreceberGeral.Caption

    colunas = Array("ID", "Pagamento", "Método", "Data", "Valor")
    Call limpaLV(pagamento.lvOutros, colunas)

    pagamento.btnOutros.Enabled = True
End Sub

Sub limpaPagamento()
    pagamento.lblSubtotalGeral.Caption = pdv.totalVenda.Caption
    pagamento.lblAcrescimoGeral.Caption = "0,00"
    pagamento.lblDescontoGeral.Caption = "0,00"
    pagamento.lblAreceberGeral.Caption = "0,00"
    pagamento.lblRecebidoGeral.Caption = "0,00"
    pagamento.lblTroco.Caption = "0,00"
End Sub
