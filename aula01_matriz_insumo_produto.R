rm(list = ls())
setwd("C:/Users/raphael.vieira/OneDrive - Valec/Área de Trabalho/UNB/00 - GRADUAÇÃO ECONOMIA/2023-2 - ANÁLISE INSUMO PRODUTO/Aulas R")

# ====== Área de funções ========
LeTabela = function(sNomeArquivo, sSheet, nLinhaInicial,
                    nLinhaFinal, nColunaInicial, nColunaFinal)
{
  dSheet = read.xlsx(sNomeArquivo, sSheet, colNames = FALSE, rowNames = FALSE)
  mMatriz = dSheet[nLinhaInicial: nLinhaFinal, nColunaInicial:nColunaFinal]
  mMatriz = apply(as.matrix.noquote(mMatriz), 2, as.numeric)
  return(nMatriz)
}


# ====== Área de biblioteca ========
install.packages("openxlsx")
library(openxlsx)
library(dplyr)

# ====== Área do código principal ======
# Variável de Parâmetros
aAno = 2019
nSetores = 68
nProdutos = 128
nColunaEstoque = 6
nColunaII = 4
nColunaIPI = 5
nColunaICMS = 6
nColunaOI = 7
nColunaMargemComercio = 2
nColunaMargemTransporte = 3
nColunaExportacao = 1

# Lendo oferta
# mOferta = LeTabela(sNomeArquivo, "oferta", 4,131,3,9)

dSheet = read.xlsx('68_tab1_2019.xlsx',"oferta", colNames = FALSE, rowNames = FALSE)
mMatriz = dSheet[4:131, 3:9]
mOferta = apply(as.matrix.noquote(mMatriz), 2, as.numeric)

# Lendo Produção
# mProducao = LeTabela(sNomeArquivo, "producao", 4,131,3,70)

dSheet = read.xlsx('68_tab1_2019.xlsx',"producao", colNames = FALSE, rowNames = FALSE)
mMatriz = dSheet[4:131, 3:70]
mProducao = apply(as.matrix.noquote(mMatriz), 2, as.numeric)

# Lendo Importação
dSheet = read.xlsx('68_tab1_2019.xlsx',"importacao", colNames = FALSE, rowNames = FALSE)
mMatriz = dSheet[4:131, 3]
vImportacao = apply(data.matrix(mMatriz), 2, as.numeric) # transformando vetor em matriz com duas dimensões


# Lendo Consumo Intermediário (essa é a matriz de usos a preços de consumidor)
dSheet = read.xlsx('68_tab2_2019.xlsx',"CI", colNames = FALSE, rowNames = FALSE)
mMatriz = dSheet[4:131, 3:70]
mConsumoIntermediario = apply(as.matrix.noquote(mMatriz), 2, as.numeric)

# Lendo Demanda Final
dSheet = read.xlsx('68_tab2_2019.xlsx',"demanda", colNames = FALSE, rowNames = FALSE)
mMatriz = dSheet[4:131, 3:8]
mDemandaFinal = apply(as.matrix.noquote(mMatriz), 2, as.numeric)

# Lendo Valor Adicionado
dSheet = read.xlsx('68_tab2_2019.xlsx',"VA", colNames = FALSE, rowNames = FALSE)
mMatriz = dSheet[4:17, 2:69]
mVA = apply(as.matrix.noquote(mMatriz), 2, as.numeric)

mDemandaFinalSemEstoque = mDemandaFinal
mDemandaFinalSemEstoque[, nColunaEstoque] <- 0.0
mConsumoTotalSemEstoque <- cbind(mConsumoIntermediario, mDemandaFinalSemEstoque) # função que concatena vetores (junta matrizes pelas colunas)

vTotalProduto = rowSums(mConsumoTotalSemEstoque)
mDistribuicao = mConsumoTotalSemEstoque / vTotalProduto
mDistribuicao[is.na(mDistribuicao)] <- 0

mValorIPI = mOferta[, nColunaIPI] * mDistribuicao
mValorICMS = mOferta[, nColunaICMS] * mDistribuicao
mValorOI = mOferta[, nColunaOI] * mDistribuicao

vVetorEntrada <- mOferta[, nColunaMargemComercio]
vMargem <- c(93, 94)

vPropMargem = vVetorEntrada[vMargem[1]: vMargem[2]] / sum(vVetorEntrada[vMargem[1]:vMargem[2]])
mMargemComercio = vVetorEntrada * mDistribuicao
mMargemDistribuida = colSums(mMargemComercio[1:(vMargem[1]-1), ]) + colSums(mMargemComercio[(vMargem[2]+1):nrow(mMargemComercio), ])
mMargemComercio[vMargem[1]:vMargem[2], ] = t(t(vPropMargem)) %*% mMargemDistribuida * (-1)

vVetorEntrada = mOferta[, nColunaMargemTransporte]
vMargem = c(95, 98)

vPropMargem = vVetorEntrada[vMargem[1]:vMargem[2]] / sum(vVetorEntrada[vMargem[1]:vMargem[2]])
mMargemTransporte = vVetorEntrada * mDistribuicao
mMargemDistribuida = colSums(mMargemTransporte[1:(vMargem[1]-1), ]) + colSums(mMargemTransporte[(vMargem[2]+1):nrow(mMargemTransporte), ])
mMargemTransporte[vMargem[1]:vMargem[2], ] = t(t(vPropMargem)) %*% mMargemDistribuida * (-1)


mDemandaFinalSemExportacao = mDemandaFinalSemEstoque
mDemandaFinalSemExportacao[, nColunaExportacao] = 0.0
mConsumoTotalSemExportacao = cbind(mConsumoIntermediario, mDemandaFinalSemExportacao)

vTotalProduto = rowSums(mConsumoTotalSemExportacao)
mDistribuicaoImportacao = mConsumoTotalSemExportacao / vTotalProduto
mDistribuicaoImportacao[is.na(mDistribuicaoImportacao)] <- 0

mValorII = mOferta[, nColunaII] * mDistribuicaoImportacao
mImportacao = vImportacao[,] * mDistribuicaoImportacao


mConsumoTotalPrecoMercado = cbind(mConsumoIntermediario, mDemandaFinal)
mConsumoTotalPrecoBase = mConsumoTotalPrecoMercado - mMargemComercio - mMargemTransporte - mValorIPI - mValorICMS - mValorOI - mImportacao - mValorII

