# 🏢 FII INVEST — Simulador de Fundos Imobiliários no Excel

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-Planilha-217346?logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Projeto_educacional-087F8C)
![DIO](https://img.shields.io/badge/Projeto-DIO-12263F)

Um simulador de **investimentos em fundos imobiliários (FIIs)** desenvolvido no **Microsoft Excel** para explorar aportes recorrentes, juros compostos, patrimônio futuro e renda mensal hipotética. O projeto faz parte de um desafio prático da **Digital Innovation One (DIO)**.

> **Aviso:** esta ferramenta é exclusivamente educacional. Taxas, perfis e pesos de alocação são **hipóteses**, não recomendações ou previsões de rentabilidade. FIIs têm riscos, e dividendos não são garantidos.

## 🎯 Perguntas que a ferramenta responde

| Pergunta | Onde consultar |
| --- | --- |
| Quanto investir por mês? | **Configurações!C12**; exibido em **Painel!C11** |
| Por quantos anos investir? | **Configurações!C13**; exibido em **Painel!C12** |
| Qual é o rendimento mensal assumido? | **Configurações!C11**; exibido em **Painel!C13** |
| Quanto patrimônio será acumulado? | **Painel!G11**, além dos cenários na aba **Projeções** |
| Quanto será recebido mensalmente em dividendos hipotéticos? | **Painel!G12**, além dos cenários na aba **Projeções** |

## 📊 Demonstração (valores fictícios)

Parâmetros da simulação inicial: **R$ 500/mês**, **10 anos**, taxa constante de **0,80% ao mês** e **perfil Moderado**.

| Indicador | Resultado estimado |
| --- | ---: |
| Total aportado | R$ 60.000,00 |
| Patrimônio futuro | R$ 100.108,73 |
| Dividendos mensais hipotéticos ao final do prazo | R$ 800,87 |
| Ganho acumulado | R$ 40.108,73 |

Os cálculos assumem aportes **no fim de cada mês** e reinvestimento dos rendimentos durante a fase de acumulação, com taxa constante. Não consideram inflação, alterações de cotação, vacância, tributação ou custos operacionais. O valor de dividendos é uma **estimativa de renda ao fim do prazo**, não pagamentos mensais assegurados ao longo da simulação.

## 🧭 Estrutura do arquivo

| Aba | Finalidade |
| --- | --- |
| **Painel** | Resumo de parâmetros, indicadores financeiros e alocação mensal por categoria |
| **Configurações** | Células editáveis para salário fictício, taxa, aporte, prazo e perfil |
| **Perfis** | Tabela de apoio com chaves compostas e percentuais de alocação |
| **Projeções** | Projeções para **2, 5, 10, 20 e 30 anos** |
| **Guia** | Explicação das funções e das referências usadas no projeto |

## 🧮 Fórmulas utilizadas

### Valor futuro — VF

**Excel em português:**

```excel
=SE(C13=0;G10;VF(C13;C12*12;-C11;0;0))
```

Implementação na aba **Painel!G11** (armazenada no arquivo em inglês, como `IF` / `FV`, padrão interno do XLSX):

```excel
=IF(C13=0,G10,FV(C13,C12*12,-C11,0,0))
```

- `C13`: taxa de rendimento mensal;
- `C12*12`: número de meses;
- `-C11`: aporte mensal em sinal negativo (saída de caixa);
- o primeiro `0`: saldo inicial;
- o segundo `0`: aporte no **final** do período;
- `SE` trata o caso de rendimento zero, para o qual o patrimônio equivale à soma dos aportes.

Dividendos mensais hipotéticos: `=G11*C13` (**Painel!G12**).

### Consulta por perfil — PROCV e chave composta

**Excel em português, para a primeira categoria:**

```excel
=SEERRO(PROCV($C$14&"|"&B22;Perfis!$B$11:$E$28;4;FALSO);0)
```

A chave composta junta **perfil** e **categoria**, por exemplo `Moderado|Papel`, e busca o percentual correspondente na tabela da aba Perfis. A referência fixa `$C$14` conserva o perfil selecionado; `B22` muda de linha ao copiar a fórmula para as outras categorias.

O valor em reais é calculado por `=$C$11*C22` (**Painel!D22**). As categorias são consultadas em **Painel!C22:C27**.

**Importante:** `SEERRO` evita exibir um erro visual em uma busca inválida, mas um total de alocação diferente de 100% deve ser investigado.

### Outras funções

- **SOMASE (SUMIF):** confere se cada perfil da tabela de apoio soma 100%.
- **Referências absolutas (`$`):** mantêm fixos os parâmetros ao calcular cenários para prazos diferentes.
- **Validação de dados:** a célula **Configurações!C14** contém a lista de perfis possíveis.

### Intervalos nomeados

**Ainda não implementados nesta versão.** As fórmulas existentes usam referências de células, incluindo referências absolutas. Numa evolução futura, pretende-se nomear `aporte`, `taxa_mensal`, `anos` e `perfil`, e substituir as referências nas fórmulas. Isso está registrado aqui para não apresentar uma funcionalidade ausente como concluída.

## 🗂️ Perfis de investidor e percentuais

Os percentuais abaixo foram **definidos para demonstração didática**, sem origem em uma recomendação de corretora, consultoria ou carteira real.

| Categoria de FII | Conservador | Moderado | Arrojado |
| --- | ---: | ---: | ---: |
| Papel | 35% | 25% | 15% |
| Tijolo | 20% | 20% | 20% |
| Híbridos | 15% | 15% | 15% |
| FOFs | 10% | 10% | 10% |
| Logística | 15% | 15% | 20% |
| Shoppings | 5% | 15% | 20% |
| **Total** | **100%** | **100%** | **100%** |

A alocação é **uma demonstração de fórmulas**, não uma classificação obrigatória de risco dos segmentos. A seleção do perfil atualiza os percentuais e os valores do aporte em cada categoria.

## 🧰 Como usar

1. Baixe [**FII_Invest_4_0_Corrigido_DIO.xlsx**](./FII_Invest_4_0_Corrigido_DIO.xlsx).
2. Abra o arquivo no **Excel desktop** (habilite a edição, caso solicitado).
3. Na aba **Configurações**, altere **apenas as células amarelas**, usando dados fictícios em um repositório público.
4. Escolha entre **Conservador**, **Moderado** e **Arrojado**.
5. Volte ao **Painel** para conferir patrimônio, renda mensal hipotética e divisão do aporte.
6. Na aba **Projeções**, compare os horizontes de 2, 5, 10, 20 e 30 anos.
7. Confira se as porcentagens de cada perfil somam **100%**.

## 🖼️ Evidências de funcionamento

Adicione duas imagens reais da **versão final**, com **os mesmos valores de aporte, taxa e prazo**, alternando somente o perfil:

- `imagens/painel-moderado.png` — Painel com perfil Moderado;
- `imagens/painel-arrojado.png` — Painel com perfil Arrojado.

Depois de adicionar os arquivos, remova os marcadores de comentário das linhas abaixo para exibi-las automaticamente no GitHub:

<!-- ![Painel do FII INVEST com perfil Moderado](imagens/painel-moderado.png) -->
<!-- ![Painel do FII INVEST com perfil Arrojado](imagens/painel-arrojado.png) -->

## ✨ O que foi personalizado em relação ao projeto de referência

- Painel com **indicadores de total aportado, patrimônio futuro, dividendos estimados e ganho acumulado**;
- Indicadores adicionais de **ganho sobre aportes** e **renda mensal hipotética sobre o total aportado**;
- Separação entre **entrada de dados**, **perfis**, **projeções** e **guia explicativo**;
- Verificação de alocação dos perfis e tratamento de **taxa de rendimento zero**;
- Estrutura e identidade visual **FII INVEST**, com valores de demonstração.

## 🚧 Possíveis evoluções

- Adicionar intervalos nomeados às fórmulas;
- Comparar, em uma segunda simulação, o aporte informado com a hipótese de **30% do salário**;
- Incluir uma simulação com inflação, custos, cenários de taxa variável e volatilidade;
- Aperfeiçoar os controles de entrada e validar faixas realistas de prazo e rendimento.

## 🎓 Sobre o desafio

Projeto de prática em Excel realizado a partir da proposta da **DIO (Digital Innovation One)**, no tema de simulação de fundos imobiliários, com referência às aulas de **Felipe Aguiar (Felipão)**.

A construção desta versão foi auxiliada por ferramentas de IA, com revisão e validação pelo estudante.
