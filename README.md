# Ocorrências do Projeto Centro

Painel interativo com os boletins de ocorrência (BOs) semanais do Projeto Centro: furtos e roubos com circunscrição no 3º DP (Campos Elíseos) e no 77º DP (Santa Cecília), centro de São Paulo.

## O que tem nesta pasta

| Arquivo | Para que serve |
|---|---|
| `index.html` | O painel. É a página que o GitHub Pages publica. |
| `dados/base_2024.js` | A base que o painel lê (mesmo conteúdo do CSV). |
| `dados/base_2024.csv` | A base consolidada, um BO por linha, para abrir no Excel, R ou Stata. |
| `consolidar_boletins.py` | Script que transforma as planilhas semanais (.xlsx) de um ano num CSV. |

## Para ver no seu computador

Descompacte o zip e abra o `index.html` com dois cliques. A pasta `dados` precisa ficar ao lado dele.

## Como publicar no GitHub Pages (sem usar linha de comando)

1. Crie uma conta em https://github.com (se ainda não tiver).
2. Clique em **New repository**. Dê um nome, por exemplo `painel-projeto-centro`, marque **Public** e clique em **Create repository**.
3. Na página do repositório, clique em **uploading an existing file**. Arraste o `index.html`, o `README.md`, o `consolidar_boletins.py` e a pasta `dados` inteira. Depois do envio, confira se a pasta `dados` aparece na lista de arquivos do repositório. Se não aparecer, clique em **Add file → Upload files** e arraste a pasta de novo. Clique em **Commit changes**.
4. Vá em **Settings → Pages**. Em **Branch**, escolha `main` e a pasta `/ (root)`. Clique em **Save**.
5. Em 1 ou 2 minutos, o painel fica no ar em `https://SEU-USUARIO.github.io/painel-projeto-centro/`.

## Como fazer o Google encontrar o painel

O Google costuma achar sozinho as páginas do GitHub Pages, mas isso pode levar semanas. Para acelerar:

1. Entre em https://search.google.com/search-console e adicione o endereço do painel como **prefixo do URL**.
2. Para verificar que o site é seu, escolha o método **Arquivo HTML**. Baixe o arquivo que o Google oferece e suba esse arquivo no repositório, na mesma pasta do `index.html`.
3. Depois de verificado, use **Inspeção de URL → Solicitar indexação**.
4. Ajuda muito colocar o link em lugares que o Google já visita: seu Lattes, LinkedIn, a página do grupo de pesquisa ou o README do seu perfil no GitHub.

## Como incluir outros anos

1. Separe as planilhas semanais de cada ano numa pasta (`boletins/2023`, `boletins/2025`...). Mantenha os nomes originais dos arquivos (`01__SEMANA_-_PROJETO_CENTRO_CRIMES_01JAN_A_07JAN.xlsx`), porque o script lê deles o número da semana e as datas.
2. Instale as dependências uma vez: `pip install pandas openpyxl`.
3. Rode o script para cada ano:
   ```
   python consolidar_boletins.py boletins/2023 2023 base_2023.csv
   ```
   O script gera dois arquivos, `base_2023.csv` e `base_2023.js`, e avisa quando uma planilha tem colunas deslocadas e as corrige (como aconteceu na semana 6 de 2024).
4. Suba os dois arquivos novos (.csv e .js) na pasta `dados` do repositório.
5. No `index.html`, procure a linha
   ```
   const ARQUIVOS = ['dados/base_2024.js'];
   ```
   e acrescente o arquivo novo:
   ```
   const ARQUIVOS = ['dados/base_2023.js', 'dados/base_2024.js'];
   ```
   Pelo GitHub, dá para editar direto no site: abra o arquivo e clique no lápis. O filtro **Ano** passa a mostrar os anos novos automaticamente.

## Antes de publicar

O site fica aberto para qualquer pessoa, e o CSV em `dados/` pode ser baixado por qualquer um. A base tem o número de cada BO e o endereço da ocorrência. Confirme que esses dados podem ser divulgados nesse nível de detalhe.

## Notas sobre os dados

- **Semana** é a do boletim semanal em que o BO foi listado. Dia e hora são os da ocorrência.
- **Natureza** segue a classificação final da planilha. "Outros" reúne BOs marcados como "não se trata de roubo ou furto", saques ou movimentações não autorizadas e crimes tentados.
- Grafias diferentes da mesma categoria foram unificadas, e campos vazios aparecem como "não informado".
- O histórico (texto) de cada BO não entra na base.
