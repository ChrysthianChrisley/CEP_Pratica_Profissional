# Projeto de Pesquisa e Prática Profissional – Mestrado em Telessaúde e Saúde Digital (UERJ)

Este repositório reúne toda a documentação institucional, portfólio de prática profissional e arquivos regulatórios para submissão ao **Comitê de Ética em Pesquisa do Hospital Universitário Pedro Ernesto (CEP/HUPE – UERJ)** via **Plataforma Brasil**.

---

## 🎯 Tema da Pesquisa
**"Arquitetura de baixo custo para monitoramento preventivo do pé diabético: proposta de sistema com sensores e modelo preditivo baseado em Random Forest"**

* **Autor / Mestrando:** Chrysthian Chrisley Tadeu Santos Silva
* **Orientadora:** Prof.ª Dra. Rosa Maria Esteves Moreira da Costa
* **Instituição:** Universidade do Estado do Rio de Janeiro (UERJ)
* **Programa:** Centro Biomédico – Faculdade de Ciências Médicas – Programa de Pós-Graduação em Telessaúde e Saúde Digital (PPGTSD)
* **Local de Coleta / Campo:** Policlínica Piquet Carneiro (PPC/UERJ) – Setor de Baropodometria / Serviço de Fisioterapia
* **Comitê de Ética:** Comitê de Ética em Pesquisa do Hospital Universitário Pedro Ernesto (CEP/HUPE – UERJ)

---

## ⚙️ Especificações Técnicas do Protótipo (Versão com Sensirion SHT31-D)
* **Microcontrolador:** **Wemos Lolin D32 V1** (ESP32 ESP-WROOM-32, clock de 240MHz, Wi-Fi 802.11 b/g/n, Bluetooth/BLE, **16MB Flash**, **8MB PSRAM**, slot para cartão MicroSD/TF integrado em modo SPI, dimensões de 65 x 25,4 mm e peso ultraleve de apenas **7,5g**).
* **Alimentação:** **Bateria de Polímero de Lítio (LiPo) 3.7V 600mAh** com cabo conector polarizado **JST PH 2.0mm**, acoplada diretamente à interface de gerenciamento de carga da placa (máx 500mA), proporcionando um módulo de tornozelo plano, leve e de alta segurança.
* **Sensores de Pressão Plantar:** 3 sensores piezorresistivos **FSR 402** posicionados nas regiões anatômicas de maior sobrecarga mecânica (antepé/metatarsos e retropé/calcâneo).
* **Sensor de Temperatura e Umidade:** **Sensirion SHT31-D** (Breakout Board I2C com resistores pull-up de 10k e capacitores integrados, compatível com 3.3V, alta precisão de **±0,3°C** e **±2% UR**, endereço I2C configurável 0x44/0x45 e perfil ultraplano que elimina atrito mecânico na sola do pé).
* **Armazenamento Seguro:** Gravação local de redundância no cartão MicroSD em caso de oscilação do link Bluetooth durante a marcha.
* **Software:** Aplicativo móvel **"Monitor do Pé"** (HTML5/JavaScript) com mapa anatômico plantar em escala de cores e alertas em tempo real.
* **Inteligência Artificial:** Modelo preditivo baseado em **Random Forest** integrando pressão plantar, temperatura, umidade relativa e variáveis clínicas para estratificação precoce de risco de ulceração no pé diabético.
* **Custo Unitário:** **R$ 337,65 por pé** (R$ 675,30 o par) — viabilidade comprovada para escalabilidade no SUS.

---

## 📁 Estrutura do Repositório

```text
CEP_Pratica_Profissional/
│
├── 01_Submissao_CEP_Palmilha_PPC/            # Documentos preenchidos para a Plataforma Brasil
│   ├── 04_Termo_de_Anuencia_Institucional_PPC.docx
│   ├── 05_TCLE_Adulto_Palmilha_Inteligente.docx
│   ├── 09_Declaracao_de_Isencao_de_Custos.docx
│   ├── Instrumento_Coleta_de_Dados_Palmilha_Baropodometro.docx
│   └── Projeto_de_Pesquisa_Integra_Palmilha_Inteligente.docx
│
├── 02_Mestrado_Pratica_Profissional/          # Documentos do Mestrado (Estágio de Campo - 20h)
│   ├── Portfolio_Pratica_Profissional_Telessaude_UERJ.docx
│   └── Portfólio das Atividades Obrigatórias (Estágio Docente e Estágio Profissional - Modelo Único Obrigatório) .pdf
│
├── 03_Modelos_e_Manuais_CEP_Oficiais/        # Modelos originais e guias do CEP/HUPE
│   ├── 01. Guia para encaminhamento de projetos com seres humanos  – LEIA-ME.pdf
│   ├── 02. Calendário das reuniões em 2026.pdf
│   ├── 03. Endereço para site da Plataforma Brasil.docx
│   ├── 04. Modelo do TERMO DE ANUÊNCIA INSTITUCIONAL e vinculo do pesquisador principal .docx
│   ├── 05. TCLE - Termo de Consentimento Livre e Esclarecido (Adulto).docx
│   ├── 06. TCLE - Termo de Consentimento Livre e Esclarecido (Responsáveis de menores de 18 anos).docx
│   ├── 07. TALE - Termo de Assentimento (Criança e adolescente).docx
│   ├── 08. TERMO - SOLICITAÇÃO DE DISPENSA DO TCLE, AUTORIZAÇÃO PARA PESQUISA EM PRONTUÁRIO E BANCO DE DADOS.docx
│   ├── 09. Declaração de isenção de custos.docx
│   ├── 10. MODELO DE RESPOSTA ÀS PENDÊNCIAS AO PARECER CONSUBSTANCIADO Nº XXXXX DO CEP_HUPE.docx
│   ├── 11 - Manual Pesquisador para Plataforma Brasil - Versão 38.pdf
│   ├── RELATÓRIO FINAL DE PESQUISA.docx
│   └── RELATÓRIO PARCIAL DE PESQUISA.docx
│
├── [-] - Qualificação/                       # Texto e apresentações do Exame de Qualificação
│   └── 0 - Qualificação - Chrysthian Chrisley FINAL.pdf
│
├── scripts/                                  # Scripts de automação e geração de documentos
│   └── generate_all_docs.py
│
├── .gitignore
└── README.md
```

---

## 📋 Checklist de Submissão na Plataforma Brasil

1. [ ] Cadastrar o protocolo de pesquisa na [Plataforma Brasil](https://plataformabrasil.saude.gov.br/) selecionando a instituição vinculada **Hospital Universitário Pedro Ernesto (CNPJ: 33.540.014/0017-14)**.
2. [ ] Gerar e assinar a **Folha de Rosto** (assinada por Chrysthian Chrisley e pela direção da PPC/HUPE).
3. [ ] Converter todos os arquivos da pasta `01_Submissao_CEP_Palmilha_PPC/` para formato **PDF**, com nomes padronizados sem acentos ou espaços (ex.: `Projeto_Palmilha.pdf`, `TCLE_Adulto.pdf`, `Anuencia_PPC.pdf`, `Declaracao_Custos.pdf`, `Ficha_Coleta.pdf`).
4. [ ] Respeitar o prazo limite de submissão (**05 de outubro de 2026**) para deliberação na reunião do comitê de **15 de outubro de 2026**.

---

## 🏛️ Contato do CEP/HUPE (UERJ)
* **Endereço:** Av. 28 de Setembro, nº 77 – Prédio do CePeM, 2º andar, sala 33 – Vila Isabel, Rio de Janeiro – RJ.
* **Telefone / WhatsApp:** (21) 2868-8253
* **E-mail:** cep@hupe.uerj.br
* **Atendimento:** Segunda a sexta-feira, das 09:00h às 13:00h.
