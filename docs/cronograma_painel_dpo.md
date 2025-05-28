# Projeto e Cronograma Painel DPO

## Introdução
Este projeto se refere à elaboração do Painel (Dashboard) de Business Intelligence da Diretoria de Planejamento e Orçamento (DPO) da Superintendência de Planejamento e Finanças (SPF) da Subsecretaria de Gestão e Finanças (SUBGEF) da Secretaria de Planejamento e Gestão do Estado de Minas Gerais (SEPLAG-MG).

O principal fator identificado durante a investigação realizada foi a elevada complexidade técnica associada à manutenção e ao uso da planilha atualmente em vigor. Utilizada há vários anos, essa planilha passou por sucessivos ajustes e adaptações, o que comprometeu sua estrutura tanto como base de dados quanto como ferramenta de consulta. Ademais, verificou-se que a forma de organização dos dados dificulta a compreensão por parte das áreas técnicas externas, que enfrentam obstáculos para acessar as informações de maneira clara e visualmente organizada. Isso evidencia uma lacuna entre a forma como os dados estão dispostos e sua efetiva utilidade no apoio às demandas dessas áreas.

Além disso, os dados presentes na planilha são consumidos e utilizados por agentes de diversos níveis organizacionais: tanto pelos técnicos da DPO, quanto áreas demandantes e tomadores de decisão. Portanto, proporcionar uma comunicação clara e acessível para as informações presentes nessa fonte de dados se apresenta como de suma importância.

Dessa forma, foi identificada a necessidade de melhoria na representação e visualização dos dados utilizados no contexto de trabalho da Diretoria por meio do desenvolvimento de um dashboard de Business Intelligence.

## Objetivos
- Subsidiar o desenvolvimento de um dashboard com informações de BI (Business Intelligence), com o intuito de prestar apoio e melhoria nos processos de trabalho e na tomada de decisão, que permita tanto à DPO quanto às áreas externas acessar, interpretar e monitorar os dados de forma mais fluida.
- Documentar o processo de trabalho realizado pelos técnicos desta Assessoria e servir de método base para a elaboração de futuros painéis ou automações de processos de trabalho.
- Subsidiar a elaboração de um cronograma com detalhamento de prazos, metas e entregas previstas.

**Resultado esperado:** Entrega de um painel BI validado pelos usuários e em conformidade com as demandas operacionais da DPO, das áreas externas e dos tomadores de decisão.

## Metodologia
Para a definição do escopo do projeto, foram realizadas diversas conversas e reuniões com a DPO, além de estudo e investigação da planilha Excel utilizada pelos técnicos como fonte de dados.

O painel será estruturado em duas abas:
1. **Informações técnicas**: voltadas para o contexto de trabalho operacional da DPO, facilitando o acompanhamento diário das áreas externas e técnicos.
2. **Informações gerenciais**: reunindo os principais indicadores para avaliação e tomada de decisão dos gestores.

A abordagem será incremental, com entregas evolutivas em fases e ciclos, garantindo:
- Entregas parciais e rápidas, permitindo valor desde fases iniciais.
- Engajamento contínuo dos envolvidos, com validação colaborativa.
- Redução de retrabalho, evitando grandes mudanças no final.

**Ferramentas:** Power BI, Power Query, Python (pandas), GitHub, Excel, SharePoint.

## Fases e Ciclos

### Fase 0 – Alinhamento e Planejamento Inicial
- **Conversas iniciais** com a DPO (técnicos e gestor): mapeamento do processo, entendimento de necessidades e requisitos.
- **Estudo da fonte de dados**: análise da estrutura, qualidade e relacionamento dos dados.
- **Definição das fases** e elaboração do cronograma de entregas.
- **Validação do planejamento** com a DPO.

Ao final desta fase, será produzido o cronograma detalhado.

### Fase 1 – Dashboard Informações Áreas e Técnicos
- **Conversas com DPO**: definição de escopo, prioridades e layout (1 dia por ciclo).
- **Modelagem e ETL**: limpeza, padronização e estruturação dos dados (3 dias por ciclo).
- **Desenvolvimento do Painel**: criação de visuais e interações (3 dias por ciclo).
- **Validação**: testes e ajustes com técnicos da DPO (3 dias por ciclo).

### Fase 2 – Dashboard Informações Gerenciais
- **Conversas com gestores**: escopo de indicadores estratégicos (1 dia por ciclo).
- **Modelagem de dados gerenciais**: preparação de métricas para gestão (3 dias por ciclo).
- **Desenvolvimento dos gráficos gerenciais**: construção de painéis (3 dias por ciclo).
- **Validação**: testes e ajustes com gestores (3 dias por ciclo).

### Fase 3 – Validação e Homologação Final
- **Preparação para piloto**: consolidação do painel e documentação de uso (1 dia).
- **Teste piloto**: execução em ambiente real com usuários-chave (3 dias).
- **Coleta de feedback**: análise de sugestões e correções (3 dias).
- **Ajustes finais**: incorporação de feedbacks (3 dias).
- **Homologação formal**: entrega definitiva (3 dias).

**Cronograma estimado:** 30 dias úteis de desenvolvimento + 15 dias úteis de validação e homologação.

## Gestão de Riscos e Desafios
- **Retrabalho por atrasos na validação**: estabelecer janela de validação clara.
- **Limitações de hardware ou ferramenta**: monitorar capacidade e performance.
- **Complexidade técnica superior ao previsto**: planejamento de buffers.
- **Necessidade de ciclos adicionais**: ajustar cronograma conforme feedback.

## Rotina de Alimentação de Dados
Detalhar o processo de atualização de dados via GitHub, garantindo rastreabilidade, versionamento e automação das fontes de dados.

---
