# Simulador de Investimentos em Fundos Imobiliários

Ferramenta em Excel que simula aportes mensais em Fundos Imobiliários (FIIs), projeta o patrimônio acumulado e divide o aporte entre seis tipos de fundo conforme o perfil de investidor escolhido.

## As perguntas que a ferramenta responde

| Pergunta | Onde aparece |
|---|---|
| Quanto investir por mês | Aba **App**, célula D11 |
| Por quantos anos | Aba **App**, célula D12 |
| Qual a taxa de rendimento mensal | Aba **App**, célula D13 |
| Quanto de patrimônio vai acumular | Aba **App**, célula D14 |
| Quanto vai receber de dividendos por mês | Aba **App**, célula D15 |

A tabela de **cenários** (linhas 19–23) repete o mesmo cálculo para 2, 5, 10, 20 e 30 anos. A tabela de **perfil do investidor** (linhas 29–36) divide o valor mensal entre os 6 tipos de fundo, com um gráfico de pizza que atualiza sozinho ao trocar o perfil.

## Como o VF e o PROCV entram nos cálculos

**VF (Valor Futuro)** calcula o patrimônio acumulado a partir de aportes mensais constantes:

```
=FV(taxa_mensal, anos*12, aporte*-1)
```

Nesta fórmula de Valor Futuro, colocamos a taxa de rendimento mensal(%), seguido do número de anos multiplicado por 12 (quantidade de meses em cada ano), e depois colocamos o valor do aporte mensal,multiplicado por -1, para tornar o valor positivo (uma vez que pelo investimento ser um pagamento, o valor fica negativo).

**PROCV com chave composta** busca o percentual de cada tipo de fundo para o perfil selecionado:

```
=VLOOKUP($C$26&"-"&B30, Apoio!$A:$D, 4, FALSE)
```

Um PROCV comum(LOOKUP), busca um valor baseado em uma única coluna-chave. Quando aexiste repetição de termos em uma coluna, o PROCV vai retornar apenas o primeiro termo. Usar chave composta é usar para procurar duas colunas distintas. É, basicamente, criar na mão, o equivalente a uma busca por duas condições ao mesmo tempo — o PROCV nativo só aceita uma coluna de busca, então a chave composta “engana” ele transformando duas colunas em uma só.

## Intervalos nomeados

| Nome | Célula | O que representa |
|---|---|---|
| `salario` | App!D6 | Salário mensal informado |
| `rendimento_carteira` | App!D7 | Rendimento médio da carteira (% a.m.) |
| `aporte` | App!D11 | Valor mensal a investir |
| `anos` | App!D12 | Prazo do investimento em anos |
| `taxa_mensal` | App!D13 | Taxa de rendimento mensal da simulação |
| `patrimonio` | App!D14 | Patrimônio acumulado calculado |

Usar intervalos nomeados,dá a mesma segurança do $, só que também deixa a fórmula legível por si só — você troca “decorar onde cada coisa está” por “ler o nome e já saber o que é”.

## Percentuais de cada perfil

Definidos por mim (não são os da aula), cada perfil soma exatamente 100% — há uma linha de verificação na aba **Apoio** confirmando isso:

| Tipo de fundo | Conservador | Moderado | Arrojado |
|---|---|---|---|
| Papel | 40% | 25% | 15% |
| Tijolo | 40% | 30% | 15% |
| Híbridos | 10% | 15% | 10% |
| FoFs | 10% | 10% | 10% |
| Desenvolvimento | 0% | 10% | 25% |
| Hotelaria | 0% | 10% | 25% |

Os percentuais de cada perfil foram criados por mim, seguindo uma lógica própria de risco crescente: concentrei o perfil Conservador nos fundos de Papel e Tijolo (considerados mais estáveis) e zerei Desenvolvimento e Hotelaria (mais voláteis); o perfil Arrojado faz o caminho inverso. Não são dados de mercado real nem recomendação de investimento — são valores ilustrativos para fins de estudo, como o próprio enunciado do desafio permite.

## O que mudei em relação à ferramenta do Expert

Fórmulas

Nas linhas de cenário (2/5/10/20/30 anos), o Expert usa referência absoluta tradicional ($D$19, $D$17 com cifrão). Na minha versão, todas as fórmulas de cenário usam os intervalos nomeados (taxa_mensal, aporte) — o Expert só usava os nomes na fórmula principal (patrimônio/dividendo), não nos cenários.
Adicionei uma linha de verificação na aba Apoio que soma o percentual de cada perfil e confirma que fecha 100% — isso não existe na planilha do Expert.

Perfis e percentuais

Troquei o nome do perfil de maior risco: o Expert usa “Agressivo”; eu usei “Arrojado”.
Os percentuais de cada perfil são completamente diferentes dos dele (os valores e a lógica por trás eu expliquei na resposta anterior sobre “de onde vieram esses números”).

Formatação

Apliquei formatação de moeda (R$) e percentual explícita em todas as células numéricas.
Usei bordas nas tabelas e uma legenda fixa no topo explicando “azul = editável, cinza = cálculo”.   

## Prints

<img width="723" height="486" alt="Captura de tela 2026-10-15 154249" src="https://github.com/user-attachments/assets/0f0c769a-280f-4354-beae-c1dc45550d8c" />
<img width="723" height="477" alt="Captura de tela 2026-10-15 154227" src="https://github.com/user-attachments/assets/7ef64274-1452-4c49-9057-28b832809e56" />

---

**Aviso:** salário, aporte e taxas usados são valores de exemplo para fins de estudo, não recomendação de investimento nem dados reais.
