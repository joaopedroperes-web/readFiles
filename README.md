# Busca de informações em PDFs

Programa de terminal em Python que consulta cada PDF pela API do OpenRouter e salva uma planilha em `ArquivosProcessados`.

## Instalação e execução

Instale o Python 3.10 ou superior, abra um terminal na pasta deste projeto e instale as dependências:

```powershell
python -m pip install -r requirements.txt
```

Crie sua chave no [OpenRouter](https://openrouter.ai/settings/keys). No PowerShell, configure a chave para a sessão atual:

```powershell
$env:OPENROUTER_API_KEY = "SUA_CHAVE_AQUI"
python buscar_pdfs.py
```

Use somente sua chave real no lugar do exemplo e não a compartilhe. A variável precisa ser configurada novamente ao abrir outro terminal. O programa lê a variável de ambiente; ele não lê arquivos `.env`.

Para diagnosticar a autenticação usando essa mesma variável, execute:

```powershell
python buscar_pdfs.py --verificar-chave
```

Essa opção faz uma única chamada GET a `https://openrouter.ai/api/v1/key`, sem ler ou enviar PDFs e sem gerar Excel. Não exibe a chave nem o corpo completo da resposta do serviço. `HTTP 200` confirma a autenticação; `HTTP 401` indica que o OpenRouter a recusou. Variável ausente, falha de conexão e resposta inválida têm mensagens próprias. A verificação retorna código de saída 0 em caso de sucesso e 1 em caso de erro. A autenticação aceita não garante disponibilidade ou acesso ao modelo. Depois da verificação, use `python buscar_pdfs.py` para realizar a pesquisa normalmente.

A verificação também mostra apenas os contadores de requisições gratuitas usadas, limite diário e restantes, quando `free_model_daily_requests` estiver presente na resposta. Esses contadores consideram o dia UTC. Se não forem enviados pelo serviço, o programa informa que a cota não está disponível; não interpreta isso como saldo zero.

## Quando aparecer HTTP 429

O erro pode ser um limite da conta (diário ou por minuto) ou uma limitação do provedor que serve o modelo. O programa classifica mensagens reconhecidas do serviço e mantém a causa como indeterminada quando faltam detalhes. Após cada falha ao processar um documento, mostra o nome do PDF, o motivo e pergunta:

```text
[1] Tentar processar este documento novamente
[2] Prosseguir sem processar este documento
```

A nova tentativa só ocorre se você escolher `1`. Em erros temporários, o programa informa a espera antes da escolha: ela cresce de 15 para 30, 60, 120, 240 e no máximo 300 segundos; se `Retry-After` indicar um tempo maior (em segundos ou data HTTP), esse tempo é respeitado. A opção `2` pula imediatamente o documento sem aguardar esse tempo. Uma entrada inválida repete a pergunta. Ctrl+C interrompe a execução. O menu também aparece para falhas de leitura, permissão, autenticação e limite de tokens. Cada repetição relê o PDF, permitindo corrigir o arquivo antes de tentar novamente.

Execute `python buscar_pdfs.py --verificar-chave` para consultar a cota. Se estiver esgotada, aguarde a renovação; se houver cota e o provedor estiver limitado, tente a busca mais tarde. A cota diária não mede a capacidade disponível do provedor. O intervalo de 3,1 segundos do programa não elimina limites compartilhados com outros aplicativos nem indisponibilidade do modelo. Consulte os [limites oficiais do OpenRouter](https://openrouter.ai/docs/api-reference/limits).

Na primeira execução sem PDFs, o programa cria `Arquivos` e mostra onde colocar os documentos. Coloque os arquivos nessa pasta e execute novamente. Digite o que deseja encontrar, por exemplo: `Qual é o número do contrato?`.

Todos os caminhos são relativos à localização de `buscar_pdfs.py`, mesmo quando você o executa a partir de outro diretório:

```text
buscar_pdfs.py
requirements.txt
Arquivos/
    contrato.pdf
    documento.pdf
ArquivosProcessados/
    resultados_DATA_HORA_IDENTIFICADOR.xlsx
```

O programa cria `ArquivosProcessados` ao terminar o lote. Os PDFs originais permanecem em `Arquivos`. A planilha usa duas colunas: a pergunta digitada e `arquivo`. Cada PDF analisado com sucesso gera uma linha, com a resposta ou exatamente `vazio` quando a informação não for encontrada. Documentos pulados pela opção `2` não entram na planilha e não são marcados como `vazio`. Se nenhum documento for processado, nenhum Excel novo é gerado. Guarde a saída do terminal se precisar consultar os arquivos pulados. Problemas de configuração anteriores ao processamento, como variável de chave ausente, são informados antes de iniciar o lote.

Depois de salvar o Excel novo, o programa exclui os anteriores criados por ele nesta pasta, mesmo que pertençam a outra pergunta. Somente arquivos com o padrão completo de nome gerado pelo programa são removidos; planilhas de outros nomes são preservadas. Se a gravação falhar ou a execução for cancelada antes de gerar o resultado, o Excel anterior permanece. Se um Excel antigo estiver aberto e não puder ser removido, o programa avisa e mantém o novo salvo.

## Funcionamento e limites

- Modelo: `xingchen-agi/xing4.0-29b-a4b:free`, via `https://openrouter.ai/api/v1/chat/completions`. A [página do modelo](https://openrouter.ai/xingchen-agi/xing4.0-29b-a4b:free) e a [documentação da API](https://openrouter.ai/docs/api-reference/overview) descrevem disponibilidade e parâmetros.
- O texto integral extraído de cada PDF é enviado ao serviço externo. Não há combinação de PDFs, resumos prévios ou cortes do documento.
- Cada chamada analisa somente um PDF. Após uma falha, a decisão de repetir ou pular é do usuário. O programa espera 3,1 segundos entre PDFs e usa timeout de 180 segundos por chamada. Falhas não são convertidas em `vazio`.
- O prompt separa instruções, pergunta e documento. A API deve retornar JSON com `resposta`; respostas vazias, inválidas ou encerradas por falta de tokens de saída são rejeitadas e o usuário escolhe repetir ou pular. O modelo precisa aceitar o formato estruturado e os parâmetros enviados. Não há troca automática de modelo.
- Para limitar tokens, as instruções são curtas, o raciocínio é desativado e a saída tem teto de 1024 tokens. `vazio` é a conclusão do modelo, cuja precisão depende do conteúdo e da pergunta.
- A estimativa da entrada usa o tamanho em bytes UTF-8, mais 2048 tokens de estrutura. Entrada e saída devem caber em 90% da referência de 262144 tokens. É uma estimativa conservadora, não uma contagem pelo tokenizador exato; pode recusar PDFs que caberiam no limite. O provedor também pode aplicar limites menores. Documentos grandes demais são sinalizados, nunca cortados ou divididos.
- Apenas PDFs diretamente dentro de `Arquivos` são lidos; a extensão aceita maiúsculas e minúsculas. Não há busca em subpastas.
- Não há OCR, senha informada pelo usuário nem interpretação de imagens, tabelas como imagem ou manuscritos. PDFs mistos têm apenas a camada de texto analisada. A ordem de leitura em PDFs com colunas ou tabelas depende da extração feita pelo pypdf.
- O Excel novo tem nome único; após a gravação, os resultados anteriores gerados pelo programa são excluídos. Células são gravadas como texto para evitar fórmulas vindas do documento; caracteres de controle incompatíveis com Excel são removidos apenas na exportação.

## Verificação local

```powershell
python -m unittest -v
```

Os testes criam PDFs temporários, verificam extração e Excel, simulam a resposta do serviço e verificam falhas. Não usam chave real e não enviam documentos ao OpenRouter.
