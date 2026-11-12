# Gastos da Obra - B4-04 Fazenda Reall

App web (celular/PC) para lancar gastos da obra direto na planilha Google Sheets, com comprovante (foto ou PDF) salvo no Google Drive.

- **Gasto**: grava na aba `Lancamentos` (ID automatico, Data, Item, Categoria, Tipo, Valor, Observacao).
- **Conta mensal**: grava/atualiza o mes na aba `Recorrentes` (Condominio, Agua, Luz, Pago?, Observacao).
- **Comprovantes**: salvos na pasta `Comprovantes - B4-04 Fazenda Reall` (criada ao lado da planilha). O link vai para a coluna `Comprovante`, criada automaticamente nas duas abas.
- **Sem sinal na obra**: o registro fica guardado no aparelho e e enviado quando a rede voltar.

```
docs/            -> app (publicado pelo GitHub Pages)
apps-script/     -> backend (colar no Apps Script da planilha)
```

## 1. Planilha no Google Sheets

1. Suba `B4-04 - Fazenda Reall.xlsx` para o Google Drive.
2. Abra e va em **Arquivo > Salvar como Planilhas Google** (o Apps Script nao grava em .xlsx).
3. Use daqui pra frente somente a versao Google Sheets.

## 2. Apps Script

1. Na planilha: **Extensoes > Apps Script**.
2. Apague o conteudo de `Codigo.gs` e cole `apps-script/Code.gs`.
3. (Opcional) Em **Configuracoes do projeto**, marque "Mostrar appsscript.json" e cole `apps-script/appsscript.json` (fuso America/Bahia).
4. Salve. Recarregue a planilha: aparece o menu **App de Gastos**.
5. **App de Gastos > Configurar PIN de acesso** - defina um PIN de 6+ caracteres. Autorize o script quando pedir.
6. **App de Gastos > Testar estrutura da planilha** - confira categorias, itens recorrentes e linhas de total.

## 3. Publicar o backend

1. No Apps Script: **Implantar > Nova implantacao > Tipo: App da Web**.
2. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**.
3. Copie a URL que termina em `/exec`.

> "Qualquer pessoa" e necessario para o app no GitHub Pages conseguir chamar o script. A protecao e o PIN: quem nao tem o PIN nao grava nem le nada. Nao compartilhe a URL + PIN.
>
> Ao alterar o `Code.gs`, use **Implantar > Gerenciar implantacoes > Editar > Nova versao** para manter a mesma URL.

## 4. Publicar o app no GitHub Pages

1. Crie um repositorio no GitHub (ex.: `gastos-obra`) e envie estes arquivos.
2. **Settings > Pages > Source: Deploy from a branch > Branch: main / pasta `/docs`**.
3. O app fica em `https://SEU-USUARIO.github.io/gastos-obra/`.

GitHub Pages gratuito exige repositorio publico. O repositorio **nao contem** URL, PIN nem dados - esses ficam apenas no aparelho de quem configurou.

## 5. Usar

1. Abra o link no celular > aba **Config** > cole a URL `/exec` e o PIN > Salvar.
2. No celular: menu do navegador > **Adicionar a tela inicial** (abre como app).
3. Lance os gastos. A aba **Resumo** mostra o total da aba Lancamentos e os ultimos 15 registros.

## Regras do backend

- Lancamento entra na primeira linha vazia antes de `TOTAL DOS LANCAMENTOS`; se nao houver, insere linha.
- ID = maior ID existente + 1.
- Categorias e Tipos sao lidos da validacao de dados da propria planilha (coluna D e E).
- Recorrentes: se o mes/ano ja existe, so os campos preenchidos sao atualizados; senao, cria na primeira linha vazia antes de `TOTAIS`. Os itens sao as colunas entre `Ano` e `Total do mes` - se voce criar uma coluna `Mensalidade` nesse intervalo, ela aparece no app automaticamente.
- `CORRIGIR_TOTAL_LANC = true`: a cada lancamento o script reescreve a formula do total para `=SUBTOTAL(109,F3:F<linha anterior>)`.
- Gravacoes simultaneas sao serializadas (LockService).

Os dados sao sensiveis e exigem conferencia humana na planilha.
