# Projeto InsperData Imangai


# MÓDULO 1 - Processo de Limpeza de Boletos

## 1. Regras de negócio adotadas

**Identificação dos boletos**
- Cada linha válida da planilha original representa um registro de boleto.
- Registros com o mesmo cliente e a mesma data de vencimento são agrupados em um único boleto.
- O mesmo cliente com vencimentos diferentes terá boletos separados.

**Tratamento dos dados**
- **Cliente:** remove textos entre parênteses e espaços desnecessários.
- **Data vecto:** mantém a data de vencimento.
- **Documento:** mantém o primeiro documento encontrado no grupo.
- **Unid. princ:** mantém a primeira unidade encontrada no grupo.
- **Parc:** reúne os números de parcelas distintos.
- **TC:** reúne os códigos distintos, organizados por prioridade: PM, PM VAGA, P2, demais P, 02 e 01.
- **Indicador V:** é convertido para VAGA. Exemplo: P4V → P4 VAGA.
- **Total:** soma os valores de todos os registros agrupados, inclusive linhas idênticas.
- **Sequência:** numera os boletos de cada cliente em ordem de vencimento, começando em 1.
- **Emitido:** permanece vazio para preenchimento posterior pelo financeiro.

**Ordenação**
- Os boletos são organizados alfabeticamente pelo nome do cliente.
- Para o mesmo cliente, os vencimentos são organizados do mais antigo para o mais recente.

**Conferência financeira**
- A soma dos valores da planilha limpa deve ser igual à soma dos registros válidos da planilha original.
- Caso exista divergência, o processamento deve sinalizar o problema.

## 2. Fluxo do código

1. **Leitura:** abre o relatório original em Excel.
2. **Seleção:** identifica as colunas necessárias e descarta cabeçalhos, subtotais e informações adicionais.
3. **Limpeza:** padroniza nomes, datas, valores e identificadores.
4. **Agrupamento:** reúne os registros por cliente e data de vencimento.
5. **Consolidação:** soma os valores, reúne as parcelas e aplica as regras dos demais campos.
6. **Organização:** ordena os boletos e atribui a sequência por cliente.
7. **Validação:** confere se o valor total foi preservado.
8. **Exportação:** gera uma nova planilha Excel com as nove colunas definidas.

## 3. Resultado esperado

Uma planilha simplificada, contendo uma linha por boleto consolidado, pronta para a próxima etapa de emissão.

O arquivo original não é alterado e a formatação visual será realizada posteriormente.

**Pontos principais para validação do financeiro:** 
1. critério de agrupamento
2. prioridade dos códigos TC 
3. manutenção do primeiro Documento/Unidade
4. soma de registros repetidos 
5. numeração sequencial por cliente.

# MÓDULO 2 - Processo de Limpeza de Cobrança Escritural

- Recebe `planilha_limpa.xlsx` (Módulo 1) e `base_geral.xlsm`
- Identifica os nomes únicos dos proprietários na planilha limpa
- Busca o histórico de boletos de cada proprietário na base geral
- Ordena os boletos por `DATA VENC.`, do mais recente para o mais antigo
- Analisa os três boletos mais recentes de cada proprietário:
  - **EMITIR:** pelo menos um dos três está com status `PAGO`
  - **NÃO EMITIR:** nenhum dos três está com status `PAGO`
  - **Sem histórico:** classifica como `EMITIR`
  - **Menos de três boletos:** analisa os registros existentes
  - Somente o status `PAGO` é considerado pagamento
- Gera `planilha_inadimplencia.xlsx` com duas abas:
  1. **Análise de emissão:** relação de proprietários, últimos vencimentos, status e decisão de emissão
  2. **Histórico completo:** cópia de todos os boletos dos proprietários analisados, organizados por data e com cores para facilitar a revisão humana
- Não altera os arquivos originais
- A emissão depende de aprovação humana posterior

