# Imposto de Renda — operações em bolsa

Programa para apurar IR sobre operações em bolsa: você lança as notas de
corretagem, ele diz quanto pagar, em que guia e até quando.

Substitui a planilha Excel que deu origem a ele. Os dados passam a viver num
banco próprio; o Excel entra uma vez só, para trazer o histórico, e depois pode
ser aposentado.

## Como rodar no Windows

**Uma vez só:** se o computador ainda não tem Python, baixe em
[python.org/downloads](https://www.python.org/downloads/). Na primeira tela do
instalador, marque **"Add python.exe to PATH"** — é a caixinha lá embaixo, e sem
ela nada funciona.

Depois:

1. Baixe este repositório: botão verde **Code → Download ZIP**.
2. Descompacte onde preferir (a Área de Trabalho serve).
3. Dê **dois cliques em `Imposto de Renda.bat`**.

Na primeira vez uma janela preta aparece por cerca de um minuto, preparando
tudo. Nas seguintes o programa abre direto — e quando uma versão nova precisar
de alguma biblioteca a mais, a janela preta reaparece por alguns segundos para
instalá-la sozinha.

Se o Windows disser que "protegeu o computador", clique em **Mais informações →
Executar assim mesmo**. É o aviso padrão de arquivo baixado da internet.

## Importar notas de corretagem em PDF

**Arquivo → Importar notas em PDF.** Um arquivo pode conter várias notas, e uma
nota pode ocupar várias folhas — o programa separa sozinho. Lê o leiaute da
B3/Bovespa (BTG e Rico, entre outras) e o da BM&F, para futuros.

Antes de gravar qualquer coisa, quatro travas:

**A nota tem de fechar consigo mesma.** Os totais lidos são comparados com os
que a própria nota declara — compras e vendas à vista, opções, valor das
operações, líquido, e em futuros o ajuste do dia. Diferença de um centavo já
barra a importação e mostra o que não bateu. Um leitor que erra em silêncio
seria pior que digitar à mão.

**Nenhuma linha de negócio pode ficar pelo caminho.** Linha que começa pela
coluna de negociação é negócio; se o leitor não conseguir entendê-la, ela é
mostrada inteira, como veio do PDF, e a nota não entra. A conferência dos
totais já acusaria a falta, mas só em centavos — diz que a nota não fecha, não
diz onde. E nota que não declara total nenhum não teria nem isso.

**Todo ativo precisa de código e tipo.** As corretoras não escrevem o ativo do
mesmo jeito: o BTG põe `LFTB11` na especificação, a Rico põe `FIC INFR BTG CI`
— o nome do fundo, sem o código. E um papel terminado em 11 tanto pode ser FII
quanto ETF ou FIDC. Em vez de adivinhar, o programa pergunta uma vez e guarda
no catálogo; da próxima aquele ativo já entra identificado.

Opções e futuros ele deduz sozinho, do mercado da operação.

Os formatos variam mais do que parece. A Rico numera a nota como `797.915`, com
separador de milhar, e escreve o contrato como `WIN Z25`; o BTG escreve
`339729` e `WING25`. A Rico ainda junta Bovespa e BM&F no mesmo arquivo, e o
BTG os separa. Tudo isso é tratado.

**Os marcadores de rodapé.** A nota sobrescreve símbolos nos negócios para
avisar coisas — `#` é negócio direto, `@` diz que a corretora atuou na
contraparte, `*` é negócio gerado pelo sistema, e cada corretora acrescenta os
seus. Ao extrair o texto do PDF esses símbolos grudam na coluna vizinha, antes
ou depois do dado: `D#`, `@17/06/2026`, `2,53*`, `C#`. O leitor os separa em
qualquer coluna e guarda como observação. Eles nunca sobram na especificação,
que é a chave do catálogo — se sobrassem, a mesma opção com e sem marcador
viraria dois ativos diferentes.

**A coluna D/C em branco.** O `D`/`C` do fim da linha diz se o financeiro é
débito ou crédito, e acompanha o valor: quando o ajuste do dia fecha em
`0,00`, não há o que debitar nem creditar e a corretora deixa o campo vazio.
Uma linha assim é negócio como qualquer outra e entra normalmente. Já o D/C
faltando **sobre um valor** não é a nota, é coluna embaralhada na extração —
essa linha é denunciada, porque adivinhar o sinal ali trocaria lucro por
prejuízo.

**Futuros: o valor do ponto.** Em futuros o preço da nota é a cotação em
pontos, e o resultado em reais depende do multiplicador do contrato: o ponto do
mini-índice vale R$ 0,20, o do índice cheio R$ 1,00, o do mini-dólar R$ 10,00.
Uma nota de mini-índice fechada com 99 contratos girando de 158.895,35 para
158.850,96 dá

```
(158.850,96 − 158.895,35) × 99 × R$ 0,20 = −R$ 879,00
```

que é exatamente o "Ajuste day trade" que a nota declara. O programa grava o
preço em pontos, como a nota imprime, e guarda o valor do ponto junto do
lançamento — quem multiplica é a apuração.

Isso vira mais uma trava: **a conta do multiplicador tem de reproduzir o ajuste
declarado**. Cadastre o ponto do mini-índice como R$ 0,50 e a importação para,
em vez de gravar um resultado duas vezes e meia maior.

Vêm cadastrados `WIN`, `IND`, `WDO` e `DOL`. Contrato de fora — boi, milho,
café — é perguntado uma vez e guardado, como acontece com os papéis. Ele não é
chutado: assumir 1,0 num mini-índice quintuplicaria o resultado.

Reimportar o mesmo arquivo não duplica nada: notas já gravadas são ignoradas.

### Quando a nota não passa: corrigir à mão

Trava não pode virar beco sem saída. Escolha a nota na lista e clique em
**Corrigir à mão...** — ou dê um duplo clique nela. Abre o editor com o que o
leitor conseguiu entender já preenchido: cabeçalho, custos e papéis. No alto,
em vermelho, o que ficou por resolver, com a linha crua da nota tal como veio
do PDF, para você conferir contra o papel e digitar o que falta.

Dentro do editor, **Alterar** traz o papel escolhido de volta para os campos —
que é o que permite acertar uma quantidade sem redigitar o preço médio de uma
nota com dezenas de negócios. Duplo clique na linha faz o mesmo.

O que sai daí entra no banco como você deixou: a conferência da nota não vale
mais como garantia, e é por isso que só se chega aqui de propósito. As demais
notas do arquivo seguem pelo caminho normal.

## Primeiro uso: trazer o histórico

**Arquivo → Importar planilha do Excel.** Escolha a sua planilha antiga e todo o
histórico entra de uma vez — notas, custos e lançamentos.

A importação é segura de repetir: notas que já estão no banco são ignoradas, e
reimportar o mesmo arquivo não duplica nada.

Ela também é fiel. Um teste automatizado importa os 1.699 lançamentos de
2017 a 2025 e confere que a apuração resultante é **idêntica à do Excel, nos 71
meses e nas 31 guias** — inclusive a ordem dos lançamentos dentro de cada dia,
que em 20 datas do histórico vinha intercalada entre notas diferentes e da qual
o preço médio depende.

## O dia a dia

### Lançamentos

Cada nota de corretagem é lançada uma vez, com seus custos e os papéis
negociados nela. **Nova nota**, preencha o cabeçalho, some os papéis um a um, e
salve. Duplo clique abre uma nota para edição, e dentro dela **Alterar** — ou
outro duplo clique — traz um papel já lançado de volta para os campos, para
corrigir no lugar em vez de remover e redigitar.

Isso corrige uma torção da planilha, em que a mesma nota aparecia espalhada por
várias linhas de `Notas` e às vezes por mais de uma linha de `Custos`, sem nada
que amarrasse as duas coisas.

Um papel já lançado antes reaparece com o tipo preenchido sozinho.

**Qualquer tabela do programa ordena por clique no cabeçalho.** Um clique
ordena crescente, outro inverte, o terceiro devolve a ordem de origem — que nos
lançamentos é a cronológica, e é dela que o preço médio depende. Datas ordenam
por data, valores por grandeza (R$ 1.000,00 depois de R$ 9,00, não antes), e
texto ignora acento e maiúscula.

### Catálogo de papéis: ida e volta pelo Excel

O catálogo é o que traduz o ativo como a corretora o escreve — `FIC INFR BTG
CI` — para o código e o tipo — `BDIF11`, `FII`. Ele se preenche sozinho
conforme você lança notas, e a importação de PDF pergunta o que não conhece.

Para quem prefere deixar tudo pronto de antemão, **Arquivo → Exportar catálogo
de papéis** grava um `.xlsx` ou `.csv`, e **Importar catálogo de papéis** o
traz de volta. Dá para montar a lista inteira de fora — a relação de FIIs ou de
ações da B3, por exemplo — e trazer centenas de papéis de uma vez, em vez de
cadastrar um a um pela tela. Depois disso as notas em PDF entram sem perguntar
nada.

O arquivo tem três colunas, mas **duas bastam**: código e tipo. A especificação
só é necessária para corretoras que escrevem o nome do fundo em vez do ticker;
sem ela, o próprio código faz as vezes. O cabeçalho é reconhecido por nome,
então `Código`, `codigo`, `Ativo` ou `Ticker` servem igualmente, e o tipo entra
como está cadastrado — quem digitou `acao` fica com `AÇÃO`.

Na volta, duas opções: **acrescentar e atualizar**, que mexe só no que veio no
arquivo, ou **substituir tudo**, que faz o catálogo passar a ser exatamente o
arquivo — e nesse caso a tela avisa quantos papéis serão apagados antes de
fazê-lo.

Arquivo com defeito não entra pela metade: tipo que não existe no cadastro,
papel sem código ou especificação repetida barram a importação inteira, com a
linha de cada problema apontada. Meio catálogo trocado seria pior que nenhum.

### Apuração

Quatro visões, filtráveis por ano:

| Aba | O que mostra |
|---|---|
| Guias a recolher | competência, valor, código e vencimento de cada DARF |
| Imposto por mês | o imposto de cada bolso, mês a mês |
| Retido na fonte | o que o banco já recolheu, para conferência |
| Posição em aberto | o que restou em carteira, por papel e por corretora |

A aba **Posição em aberto** mostra o mesmo estoque de duas maneiras. Em cima,
consolidado por papel: quantidade, preço médio e custo total. Embaixo, o mesmo
estoque separado por corretora — porque a ficha de bens e direitos pede o CNPJ
de quem guarda o papel, e um papel comprado em duas corretoras vira duas linhas
na declaração.

O preço médio é o mesmo nas duas visões, e é sempre o do papel. O custo de
aquisição é do contribuinte e não se reparte por custódia: transferir papel de
uma corretora para outra não muda o custo de nada. Por isso a soma das linhas
por corretora sempre bate com a linha consolidada.

Quando uma corretora fica com quantidade negativa num papel que no total está
comprado, a apuração avisa: é o rastro de uma transferência de custódia que não
foi lançada — comprou numa, transferiu, vendeu na outra. O imposto continua
certo; o que está errado é só a foto de onde o papel está.

**Exportar para Excel** grava tudo num `.xlsx`, com sete abas. Duas servem para
conferir contra os PDFs da corretora:

*Notas* — **uma linha por nota de corretagem**, com o número impresso nela, a
data, a corretora e os cinco campos de custo discriminados: liquidação,
registro, emolumentos, corretagem e IRRF retido. Mais o custo total, quantos
papéis a nota tem e o volume negociado. É a aba que se põe lado a lado com o
PDF.

*Lançamentos* — cada operação com o número da nota, os mesmos cinco campos
repetidos, a fatia que coube àquela linha no rateio, e o lucro e o imposto
apurados.

O IRRF fica fora do custo total nas duas: é imposto retido, não despesa.

### Tipos e alíquotas

O que era a aba `Variáveis`. Só precisa de atenção se a legislação mudar. Cada
tipo define alíquotas, regime de custo, isenção mensal, bolso de compensação e
grupo de guia.

**Tabela regressiva** é um campo à parte, e vale saber como ele interage com a
alíquota: quando há tabela escolhida, quem manda na operação normal é o prazo
de cada lote, e a alíquota fixa digitada acima fica inerte. A tela avisa isso
na hora. Para voltar a uma alíquota única, escolha *(nenhuma)* na tabela.

## Os dois regimes de custo

A diferença central em relação à planilha. Renda variável apura pelo **preço
médio** do estoque; renda fixa apura por **FIFO**, com cada lote preservado e
consumido na ordem em que foi comprado.

Comprou 100 a R$ 200, depois 500 a R$ 220, e vendeu 200 a R$ 230:

| Regime | Custo aplicado | Lucro | IR a 15% |
|---|---|---|---|
| Preço médio | 200 × R$ 216,67 | R$ 2.666,67 | R$ 400,00 |
| **FIFO** | 100 × R$ 200 + 100 × R$ 220 | **R$ 4.000,00** | **R$ 600,00** |

Em FIFO cada lote consumido carrega a sua data de aquisição, e é isso que
permite aplicar tabela regressiva: uma única venda pode ser tributada em duas
alíquotas, uma por lote.

## Renda fixa é conferência, não guia

O imposto de renda fixa é retido na fonte pelo banco ou pela corretora. O
programa calcula assim mesmo — não para você pagar de novo, mas para conferir o
que já foi recolhido. Esses ativos ficam **fora** de qualquer DARF.

Vêm cadastrados `CDB`, `TESOURO`, `DEBÊNTURE` e `RENDA FIXA` pela tabela
regressiva de 22,5% a 15%, mais `LCI/LCA` e `DEB INCENTIVADA` como isentos.

`ETF RF` é o híbrido, e não usa tabela: são **15% fixos** sobre o ganho,
independente do prazo. O ganho de capital é retido na fonte pelo
administrador, mas o **day trade** é recolhido pelo investidor e gera guia.

## Como a apuração funciona

**Rateio de custos.** As taxas de cada nota são divididas entre a perna day
trade e a perna normal pela participação de cada uma no giro financeiro. Day
trade gira duas vezes, porque compra e vende.

**Bolsos estanques.** Prejuízo de um bolso nunca abate ganho de outro. São três:
fundos (FII, Fiagro, FIDC), demais operações normais, e day trade. Cada um
carrega o próprio saldo negativo para o mês seguinte.

**Isenção mensal.** R$ 20.000 de vendas por mês, só para ações. O teste é sobre
o volume vendido, não sobre o lucro — e sobre o vendido em **operação normal**.
Day trade não goza da isenção, então também não conta para medir o teto: quem
vendeu R$ 15.000 à vista e girou outros R$ 50.000 em day trade no mesmo mês
continua isento na parte comum.

**Guias.** Dois grupos — fundos e o resto —, cada um abatendo o IRRF retido no
mês. Retenção que sobra vira crédito para o mês seguinte. Saldo abaixo de R$ 10
não é recolhido e acumula (art. 68 da Lei 9.430/96). Vencimento no último dia
útil do mês seguinte, recuando de fim de semana e feriado.

**Day trade a 19%.** A corretora já recolhe 1% sobre o ganho, então resta 19%.
Por isso esse 1% **não** deve ser lançado no campo IRRF da nota — aquele campo é
para o dedo-duro de 0,005% das operações normais. Lançar o 1% ali abateria duas
vezes.

## O que mudou em relação à planilha

Divergências encontradas na conversão e corrigidas no comportamento padrão.

| # | Na planilha | Aqui |
|---|---|---|
| 1 | Renda fixa apurada por preço médio | FIFO, com lotes datados |
| 2 | `ETF RF` com day trade a 0% por causa de `IF(0%−1%<0, 0, …)`, contra o que a própria documentação dela mandava — R$ 1.291,41 de lucro passaram sem tributação em 2024 | 19% sobre o ganho |
| 3 | `FIAG` sem perfil: o `XLOOKUP` caía em zero e o Fiagro saía sem imposto, em silêncio | Fiagro tributado como FII, e tipo desconhecido levanta erro |
| 4 | Vencimento aproximado por competência + 45 dias | Último dia útil do mês seguinte |
| 5 | Sem o limite de dispensa de R$ 10 | Saldo abaixo do mínimo acumula |
| 6 | IRRF da nota inteiro no tipo da primeira linha dela | Rateado pelo giro de cada tipo |
| 7 | `FIAG` no grupo de guia das ações | No grupo dos fundos |
| 8 | Prejuízo de mês isento entrava na compensação | Configurável |
| 9 | Ordem cronológica manual, aba protegida, limite de 1.028 linhas | Ordem preservada na gravação, sem limite |
| 10 | Custos da mesma nota espalhados por várias linhas | Uma nota, um conjunto de custos |
| 11 | Teto de isenção medido pela venda inteira, day trade somado junto — bastava girar em day trade para perder a isenção da parte comum | Só a venda em operação normal mede o teto |

Para conferir a migração antes de confiar nos números corrigidos, a linha de
comando tem `--modo-planilha`, que reproduz o Excel exatamente, defeitos
inclusive.

## Linha de comando

Para quem preferir o terminal, ou para rodar em servidor:

```bash
pip install -r requirements.txt
python -m imposto_renda Imposto_de_renda.xlsm --ano 2024 --posicao
python -m imposto_renda Imposto_de_renda.xlsm --modo-planilha
```

`--posicao` imprime as duas visões do estoque: por papel e por corretora.

Como biblioteca:

```python
from imposto_renda import Apurador, gerar_darfs
from imposto_renda.banco import Banco

with Banco() as banco:
    config = banco.configuracao()
    operacoes, custos = banco.para_apuracao()
    apuracao = Apurador(config).apurar(operacoes, custos)

    for darf in gerar_darfs(apuracao, config).a_recolher():
        print(darf.competencia, darf.valor, darf.vencimento)
```

## Onde ficam os dados

Um arquivo SQLite. No Windows fica em:

```
C:\Users\<você>\AppData\Roaming\ImpostoDeRenda\dados.db
```

`AppData` é uma pasta oculta, então **Arquivo → Onde ficam os dados** mostra o
caminho exato, o tamanho e abre a pasta no Explorer. O rodapé do programa
também exibe o caminho.

É um arquivo comum: pode ser copiado, guardado em backup e levado para outro
computador.

### Backup e restauração

**Arquivo → Fazer backup** grava uma cópia íntegra, mesmo com o programa
aberto. **Arquivo → Restaurar backup** substitui o conteúdo atual pelo da
cópia — e recusa arquivos que não sejam backups deste programa, em vez de
deixar tudo quebrado.

### Apagar lançamentos

Três caminhos, conforme o tamanho da faxina:

| Quero apagar | Como |
|---|---|
| Uma nota | selecione e clique em **Excluir selecionadas**, ou tecle Delete |
| Algumas notas | Ctrl+clique para escolher avulsas, Shift+clique para um intervalo |
| Um ano inteiro, ou tudo | **Arquivo → Apagar lançamentos** |

Em **Apagar lançamentos** você escolhe o ano ou "Todos", e a janela oferece
salvar um backup antes — marcado por padrão. Se você cancelar a gravação do
backup, o apagamento também é cancelado.

Apagar lançamentos **não** apaga tipos, alíquotas nem corretoras: recomeçar do
zero não custa reconfigurar tudo. Para zerar inclusive o cadastro, apague o
próprio arquivo `dados.db` — ele é recriado na próxima abertura.

## Testes

```bash
pip install pytest
python -m pytest tests/ -q
```

187 testes. A regressão contra a planilha confere custo, lucro, preço médio,
estoque e imposto nas 1.699 linhas; imposto, compensação e DARF nos 71 meses com
movimento; e as 24 guias que ela emitiu. Diferença máxima observada:
5,7 × 10⁻¹⁴. O restante cobre FIFO, tabela regressiva, renda fixa na fonte,
bolsos estanques, isenção mensal, vencimento, limite de dispensa, migração para
o banco, exclusão em massa, cópias de segurança, o cadastro de tipos, a ida e
volta do catálogo por planilha, a ordenação das tabelas e a leitura de notas em
PDF — esta última sobre o texto real de notas do BTG e da Rico, com os dados
pessoais substituídos.

Futuros têm suíte própria, e a mais convincente delas roda sobre as notas
reais: em **todas as notas de futuros dos fixtures que declaram o ajuste do
dia**, a conta `quantidade × Δpreço × valor do ponto` reproduz o valor
declarado ao centavo — inclusive numa nota com dois vencimentos diferentes de
mini-índice na mesma folha.

## Aviso

Ferramenta de apoio. Não substitui contador nem orientação profissional. Confira
os valores antes de recolher.
