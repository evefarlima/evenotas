# Notas

App de notas pessoal (celular e computador), com pastas, tags, busca, modo escuro e funcionamento offline.
As notas ficam numa planilha do Google, acessada por um script (Apps Script).

## Configurar a planilha (uma vez)

1. Crie uma planilha nova no Google Sheets com o nome **Notas**.
2. Menu **Extensões > Apps Script**. Apague o que estiver escrito e cole o conteúdo de `apps-script/Codigo.gs`. Clique no disquete para salvar.
3. No topo, escolha a função **configurar** e clique em **Executar**. Autorize com a sua conta Google
   (se aparecer "O Google não verificou este app", clique em **Avançado > Acessar Notas (não seguro)** — o app é seu).
4. Volte à planilha: a aba **Config** mostra a sua chave na célula **B1**.
5. No Apps Script: **Implantar > Nova implantação**, tipo **App da Web**.
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
   Clique em **Implantar** e copie o **URL do app da Web** (termina em `/exec`).

## Usar o app

1. Abra o endereço do app (GitHub Pages).
2. Cole o URL do app da Web e digite a chave. Clique em **Conectar**.
3. No celular, use **Adicionar à tela inicial** para ter o ícone.

Para trocar a chave, mude a célula B1 da aba Config. Os aparelhos com a chave antiga pedem a nova.

Se alterar o código do Apps Script no futuro, publique de novo em
**Implantar > Gerenciar implantações > lápis > Versão: Nova versão > Implantar** (o endereço continua o mesmo).
