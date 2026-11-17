# RH — base de pessoal da CEPLAC

Conferência automática da base de servidores com o SIAPE.

## Como atualizar (todo mês)
1. No GitHub, abra a pasta `rh/entrada/`.
2. **Add file › Upload files** e envie os dois arquivos novos:
   - a planilha de servidores por unidade (`.xls` ou `.xlsx`, as 10 abas);
   - o PDF do SIAPE "Total de servidores por situação".
3. Clique em **Commit changes**.
4. Aba **Actions**: a execução "RH — conferir base de pessoal" roda sozinha (cerca de 1 minuto).
5. Resultado:
   - `rh/RELATORIO.md`: relatório (SIAPE × planilha, totais por unidade, mudanças);
   - `rh/dados/mudancas.csv`: lista das mudanças;
   - `rh/dados/pessoal.json` e `siape.json`: base tratada, usada pelo painel;
   - o mesmo relatório aparece no resumo da execução em Actions.

Vale o arquivo **mais recente de cada tipo**: não precisa apagar os antigos (o histórico do GitHub guarda tudo).

## Primeira vez
Envie primeiro a planilha do mês anterior (base de referência) e depois a nova, em dois envios separados. Assim a conferência mostra o que mudou.

## O que o relatório confere
- Ativos permanentes do SIAPE × servidores da planilha (sem inativos, vagos e duplicados, CLT e CCE).
- Total de cada aba × rodapé "TOTAL DE SERVIDORES" da própria aba (acusa rodapé desatualizado).
- Nomes repetidos entre abas: fica a linha da unidade; a da Diretoria é descartada e herda a função.
- Entradas, saídas, mudanças de unidade e de função em relação à base anterior.

## Observações
- O repositório deve ser **privado**: os arquivos têm nomes de servidores.
- Rodar à mão: Actions › "RH — conferir base de pessoal" › Run workflow.
- Teste local: `python rh/scripts/atualizar.py` (precisa de pandas, xlrd, openpyxl, pypdf).
