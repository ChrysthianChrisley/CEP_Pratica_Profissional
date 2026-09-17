import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Constantes do Projeto
AUTOR = "Chrysthian Chrisley Tadeu Santos Silva"
ORIENTADORA = "Prof.ª Dra. Rosa Maria Esteves Moreira da Costa"
TITULO_DISSERTACAO = "Arquitetura de baixo custo para monitoramento preventivo do pé diabético: proposta de sistema com sensores e modelo preditivo baseado em Random Forest"
TITULO_SUBMISSAO_CEP = "Arquitetura de baixo custo para monitoramento preventivo do pé diabético: calibração com baropodômetro e prova de conceito na Policlínica Piquet Carneiro"
PROGRAMA = "Programa de Pós-Graduação em Telessaúde e Saúde Digital (PPGTSD)"
INSTITUICAO = "Universidade do Estado do Rio de Janeiro (UERJ)"
LOCAL_COLETA = "Policlínica Piquet Carneiro (PPC/UERJ) – Setor de Baropodometria / Serviço de Fisioterapia"
HARDWARE_INFO = (
    "Microcontrolador Wemos Lolin D32 V1 (ESP32 ESP-WROOM-32, Wi-Fi, BLE, 240MHz, 16MB Flash, 8MB PSRAM, "
    "slot para cartão MicroSD/TF integrado, dimensões compactas de 65x25,4mm, peso de 7,5g), 3 sensores piezorresistivos FSR 402, "
    "1 sensor digital de temperatura e umidade DHT22, alimentado por Bateria LiPo 3.7V 600mAh com cabo conector polarizado JST PH 2.0mm "
    "e circuito de recarga integrado na placa (máx 500mA)"
)
SOFTWARE_INFO = "Aplicativo móvel 'Monitor do Pé' (HTML5/JavaScript) com alertas em tempo real e modelo preditivo Random Forest"
CUSTO_UNITARIO = "R$ 337,65 por pé (R$ 675,30 o par)"

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:left w:w="{left}" w:type="dxa"/>'
        f'  <w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_border(cell, top="CCCCCC", bottom="CCCCCC", left="CCCCCC", right="CCCCCC", sz="4"):
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{top}"/>'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{bottom}"/>'
        f'  <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{left}"/>'
        f'  <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{right}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

def format_paragraph(p, before=3, after=4, line_spacing=1.15):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line_spacing

def add_header_block(doc, title, subtitle=None):
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

    h_p = doc.add_paragraph()
    h_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(h_p, 0, 2, 1.0)
    r1 = h_p.add_run("UNIVERSIDADE DO ESTADO DO RIO DE JANEIRO\nCENTRO BIOMÉDICO – FACULDADE DE CIÊNCIAS MÉDICAS\nNÚCLEO DE TELESSAÚDE E SAÚDE DIGITAL\nPROGRAMA DE PÓS-GRADUAÇÃO STRICTO SENSU EM TELESSAÚDE E SAÚDE DIGITAL\nMESTRADO PROFISSIONAL EM TELESSAÚDE E SAÚDE DIGITAL")
    r1.font.name = "Arial"
    r1.font.size = Pt(9.5)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(40, 60, 100)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_div, 2, 8, 1.0)
    r_div = p_div.add_run("―" * 45)
    r_div.font.size = Pt(8)
    r_div.font.color.rgb = RGBColor(160, 160, 160)

    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(t_p, 4, 2, 1.1)
    r_title = t_p.add_run(title)
    r_title.font.name = "Arial"
    r_title.font.size = Pt(13)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(20, 40, 80)

    if subtitle:
        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(sub_p, 0, 12, 1.1)
        r_sub = sub_p.add_run(subtitle)
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(10)
        r_sub.font.italic = True
        r_sub.font.color.rgb = RGBColor(80, 80, 80)

# ==============================================================================
# 1. PORTFÓLIO DO ALUNO - PRÁTICA PROFISSIONAL (20 HORAS)
# ==============================================================================
def create_portfolio_doc():
    doc = docx.Document()
    add_header_block(doc, "PORTFÓLIO DO ALUNO", "PRÁTICA PROFISSIONAL (PESQUISA DE CAMPO) – CARGA HORÁRIA: 20 HORAS")

    intro = doc.add_paragraph()
    format_paragraph(intro, 4, 8, 1.15)
    r_intro = intro.add_run(
        f"Este portfólio documenta as atividades presenciais obrigatórias de Prática Profissional (Pesquisa de Campo) "
        f"realizadas pelo discente {AUTOR}, sob orientação da {ORIENTADORA}, no âmbito do {PROGRAMA} da {INSTITUICAO}. "
        f"A atividade integra o desenvolvimento e validação de protótipo de tecnologia biomédica vestível (IoMT) à "
        f"infraestrutura de avaliação da marcha da Policlínica Piquet Carneiro (PPC/UERJ)."
    )
    r_intro.font.size = Pt(10)

    def add_campo_box(doc, campo_num, campo_title, content_dict):
        p_head = doc.add_paragraph()
        format_paragraph(p_head, 8, 3, 1.15)
        r = p_head.add_run(f"CAMPO {campo_num} – {campo_title.upper()}")
        r.font.name = "Arial"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = RGBColor(20, 50, 100)

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.columns[0].width = Inches(6.8)
        cell = table.cell(0, 0)
        set_cell_background(cell, "F8F9FA")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        set_cell_border(cell, top="003366", bottom="CCCCCC", left="003366", right="CCCCCC", sz="8")

        p_cell = cell.paragraphs[0]
        format_paragraph(p_cell, 2, 2, 1.15)

        for label, val in content_dict.items():
            p_item = cell.add_paragraph() if p_cell.text else p_cell
            format_paragraph(p_item, 2, 3, 1.15)
            r_lbl = p_item.add_run(f"{label}: ")
            r_lbl.font.name = "Arial"
            r_lbl.font.size = Pt(9.5)
            r_lbl.font.bold = True
            r_lbl.font.color.rgb = RGBColor(30, 30, 30)

            r_txt = p_item.add_run(val)
            r_txt.font.name = "Arial"
            r_txt.font.size = Pt(9.5)
            r_txt.font.color.rgb = RGBColor(50, 50, 50)
            p_cell = p_item

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # CAMPO 1
    add_campo_box(doc, 1, "Identificação do Aluno e Projeto no Mestrado", {
        "Nome Completo do Aluno": AUTOR,
        "Profissão / Área de Atuação": "Pesquisador em Saúde Digital / Engenharia de Software / Telessaúde",
        "Orientadora": ORIENTADORA,
        "Título do Projeto no Mestrado": TITULO_DISSERTACAO,
        "Contextualização da Escolha do Local e do Supervisor": (
            f"A escolha da Policlínica Piquet Carneiro (PPC/UERJ) justifica-se por sua excelência assistencial e de pesquisa "
            f"clínico-ambulatorial, contando com laboratório especializado e plataforma de baropodometria computadorizada, "
            f"considerada padrão-ouro para análise de pressões plantares. O supervisor técnico da PPC possui sólida qualificação "
            f"clínico-biomecânica para orientar a parametrização dos testes, garantindo que os ensaios de calibração e prova "
            f"de conceito da palmilha inteligente — instrumentada com microcontrolador Wemos Lolin D32 (ESP32 ESP-WROOM-32 com 16MB Flash, "
            f"8MB PSRAM e slot MicroSD), 3 sensores FSR 402, sensor DHT22 e alimentada por bateria LiPo 3.7V 600mAh (conector JST PH 2.0mm) — "
            f"atendam ao rigor científico exigido para validação de tecnologias assistivas no SUS."
        )
    })

    # CAMPO 2
    add_campo_box(doc, 2, "Identificação do Supervisor da Atividade", {
        "Nome Completo do Supervisor": "[Nome do Supervisor / Fisioterapeuta Responsável na PPC]",
        "Profissão": "Fisioterapeuta / Especialista em Biomecânica da Marcha",
        "Cargo / Função": "Responsável Técnico / Fisioterapeuta do Setor de Baropodometria da PPC/UERJ",
        "Instituição de Origem": "Policlínica Piquet Carneiro – Universidade do Estado do Rio de Janeiro (PPC/UERJ)",
        "Titulação Acadêmica": "[Mestre / Especialista]",
        "Contextualização da Experiência": (
            "Profissional com consolidada experiência na operação e interpretação clínica de exames baropodométricos estáticos "
            "e dinâmicos, acompanhamento de alterações de marcha em pacientes portadores de Diabetes Mellitus e avaliação "
            "biomecânica de membros inferiores."
        )
    })

    # CAMPO 3
    add_campo_box(doc, 3, "Identificação do Local da Atividade", {
        "Nome da Unidade": "Policlínica Piquet Carneiro (PPC) – UERJ",
        "Setor / Serviço": "Serviço de Fisioterapia / Ambulatório de Reabilitação / Laboratório de Baropodometria",
        "Endereço Completo": "Avenida Marechal Rondon, 381 – São Francisco Xavier, Rio de Janeiro – RJ, CEP: 20950-003",
        "CNPJ Institucional": "33.540.014/0001-50 (UERJ) / Vínculo com Complexo de Saúde UERJ",
        "Equipe Envolvida": f"Mestrando: {AUTOR}; Orientadora: {ORIENTADORA}; Supervisor Técnico da PPC e equipe multiprofissional da unidade.",
        "Perfil e Relevância": (
            "Maior policlínica ambulatorial universitária pública da América Latina, inaugurada em 1995 e ligada ao Centro Biomédico da UERJ. "
            "Atende a uma ampla população de pacientes crônicos com Diabetes Mellitus pelo SUS, oferecendo suporte diagnóstico e terapêutico."
        )
    })

    # CAMPO 4 - Detalhamento das Atividades (20 Horas)
    p4_h = doc.add_paragraph()
    format_paragraph(p4_h, 10, 4, 1.15)
    r4 = p4_h.add_run("CAMPO 4 – RELATÓRIO DO ALUNO (DETALHAMENTO DAS ATIVIDADES – 20 HORAS)")
    r4.font.name = "Arial"
    r4.font.size = Pt(11)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(20, 50, 100)

    p4_desc = doc.add_paragraph()
    format_paragraph(p4_desc, 2, 6, 1.15)
    p4_desc.add_run(
        "A Prática Profissional foi estruturada em 4 etapas articuladas com o cronograma da dissertação de mestrado, "
        "totalizando 20 horas de execução presencial e técnica:"
    )

    etapas = [
        ("Etapa 1: Caracterização do Baropodômetro e Parametrização Técnica", "4 horas",
         "Mapear a resolução espacial, taxa de amostragem em Hz e protocolos de exportação de dados do baropodômetro da PPC.",
         "Reunião técnica com o supervisor da PPC; análise do software de aquisição de pressão estática e dinâmica; identificação das zonas de sobrecarga anatômica do pé diabético (antepé/metatarsos e retropé/calcâneo); definição de protocolos de sincronização temporal com o microcontrolador Wemos Lolin D32.",
         "Matriz de correlação de coordenadas anatômicas definida e padronização dos formatos de exportação (CSV/TXT) para confronto direto com o sistema embarcado da palmilha."),

        ("Etapa 2: Parametrização e Calibração de Bancada da Palmilha (Wemos Lolin D32 + FSR 402 + DHT22)", "5 horas",
         "Realizar a calibração de bancada dos 3 sensores piezorresistivos FSR 402, do sensor DHT22 e validação do fluxo BLE com o app 'Monitor do Pé' utilizando o microcontrolador Wemos Lolin D32.",
         "Ensaios de repetibilidade com cargas estáticas conhecidas; aplicação de filtro de média móvel de 10 amostras no firmware do ESP32 para mitigar a histerese dos sensores FSR (mantendo variação < 5%); teste de gravação de segurança no slot MicroSD da placa Lolin D32 em caso de perda de conexão sem fio; verificação do circuito de carga e autonomia da bateria LiPo 3.7V 600mAh com conector JST PH 2.0mm.",
         "Curvas de calibração validadas em bancada, estabilidade telemétrica BLE sem perdas de pacotes comprovada, peso ultraleve do módulo (7,5g placa + ~13g bateria) e custo unitário consolidado em R$ 337,65."),

        ("Etapa 3: Estruturação dos Protocolos Éticos e Biossegurança na PPC", "3 horas",
         "Assegurar a total conformidade do protocolo com as Resoluções CNS nº 466/2012 e 510/2016 e as diretrizes do CEP/HUPE.",
         "Redação final do TCLE em linguagem leiga, destacando a segurança da bateria LiPo de 3,7V e a ausência de qualquer choque elétrico, além do respeito à autonomia do paciente diabético; elaboração do checklist de biossegurança (desinfecção com álcool 70% e meias descartáveis de uso único); preparação da submissão na Plataforma Brasil.",
         "Dossiê documental aprovado pelo supervisor da PPC e pesquisador, submetido via Plataforma Brasil para deliberação do CEP/HUPE."),

        ("Etapa 4: Prova de Conceito e Ensaio Comparativo com Pacientes na PPC", "8 horas",
         "Validar a usabilidade funcional e a correlação das leituras de pressão da palmilha inteligente frente ao baropodômetro em 3 a 5 voluntários adultos com Diabetes Mellitus tipo 2 (executado após liberação ética do CEP/HUPE).",
         "Acolhimento dos voluntários, esclarecimento e assinatura do TCLE; aplicação de meia descartável e calçamento do dispositivo leve; execução de teste estático em ortostase (10s) e caminhada dinâmica em linha reta sobre a plataforma de pressão; monitoramento simultâneo de temperatura e umidade relativa pelo app 'Monitor do Pé'; preenchimento da ficha de ensaio experimental e questionário de usabilidade.",
         "Banco de dados experimental constituído com correlação estatística entre o baropodômetro e a palmilha inteligente, comprovando a viabilidade técnica e biomecânica do wearable e gerando dados para o modelo Random Forest.")
    ]

    for title, carga, obj, metodo, res in etapas:
        t_etapa = doc.add_table(rows=4, cols=2)
        t_etapa.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_etapa.autofit = False
        t_etapa.columns[0].width = Inches(1.8)
        t_etapa.columns[1].width = Inches(5.0)

        for row_idx, (lbl, val) in enumerate([
            ("Atividade / Carga Horária", f"{title} [{carga}]"),
            ("Objetivo Específico", obj),
            ("Método e Execução", metodo),
            ("Resultado Obtido / Esperado", res)
        ]):
            c0 = t_etapa.cell(row_idx, 0)
            c1 = t_etapa.cell(row_idx, 1)
            set_cell_background(c0, "F0F4F8")
            set_cell_background(c1, "FFFFFF")
            set_cell_margins(c0, 80, 80, 100, 100)
            set_cell_margins(c1, 80, 80, 100, 100)
            set_cell_border(c0, sz="2")
            set_cell_border(c1, sz="2")

            p0 = c0.paragraphs[0]
            format_paragraph(p0, 1, 1, 1.1)
            r0 = p0.add_run(lbl)
            r0.font.name = "Arial"
            r0.font.size = Pt(9)
            r0.font.bold = True
            r0.font.color.rgb = RGBColor(30, 50, 90)

            p1 = c1.paragraphs[0]
            format_paragraph(p1, 1, 1, 1.1)
            r1 = p1.add_run(val)
            r1.font.name = "Arial"
            r1.font.size = Pt(9)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Documentação Fotográfica
    p_foto_h = doc.add_paragraph()
    format_paragraph(p_foto_h, 8, 4, 1.15)
    r_fh = p_foto_h.add_run("DOCUMENTAÇÃO FOTOGRÁFICA E TÉCNICA DA PRÁTICA PROFISSIONAL")
    r_fh.font.name = "Arial"
    r_fh.font.size = Pt(10.5)
    r_fh.font.bold = True
    r_fh.font.color.rgb = RGBColor(20, 50, 100)

    t_foto = doc.add_table(rows=2, cols=2)
    t_foto.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_foto.columns[0].width = Inches(3.4)
    t_foto.columns[1].width = Inches(3.4)
    quadros = [
        ("Figura 1: Protótipo com Wemos Lolin D32 e LiPo 600mAh", "[Inserir fotografia da montagem física da palmilha com sensores FSR, DHT22, placa Lolin D32 (7,5g) e bateria LiPo com conector JST]"),
        ("Figura 2: Plataforma de Baropodometria da PPC/UERJ", "[Inserir fotografia do equipamento de baropodometria da Policlínica Piquet Carneiro]"),
        ("Figura 3: Ensaio de Validação e Calibração na Marcha", "[Inserir fotografia do voluntário/ensaio de calibração com calçado e esteira de pressão]"),
        ("Figura 4: Interface do Aplicativo Móvel 'Monitor do Pé'", "[Inserir captura de tela do aplicativo com o mapa anatômico plantar e os alertas em tempo real]")
    ]
    for idx, (legenda, inst) in enumerate(quadros):
        r_i = idx // 2
        c_i = idx % 2
        cell = t_foto.cell(r_i, c_i)
        set_cell_background(cell, "FAFAFA")
        set_cell_margins(cell, 120, 120, 120, 120)
        set_cell_border(cell, sz="2")
        p = cell.paragraphs[0]
        format_paragraph(p, 2, 2, 1.15)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_l = p.add_run(f"{legenda}\n\n")
        r_l.font.name = "Arial"
        r_l.font.size = Pt(9)
        r_l.font.bold = True
        r_in = p.add_run(inst)
        r_in.font.name = "Arial"
        r_in.font.size = Pt(8.5)
        r_in.font.italic = True
        r_in.font.color.rgb = RGBColor(120, 120, 120)

    # CAMPOS 5, 6, 7 e 8
    doc.add_page_break()
    add_header_block(doc, "PORTFÓLIO DO ALUNO – AVALIAÇÕES E FECHAMENTO", "CAMPOS DE AVALIAÇÃO E REGISTRO FORMAL")

    add_campo_box(doc, 5, "Avaliação do Aluno pelo Supervisor", {
        "Desempenho e Engajamento": (
            f"O discente {AUTOR} apresentou desempenho irrepreensível, demonstrando alto grau de maturidade técnica, "
            f"pontualidade exemplar e profundo respeito às normas de biossegurança e segurança do paciente na PPC. "
            f"A escolha do microcontrolador Wemos Lolin D32 com bateria LiPo de 600mAh representou um avanço ergonômico "
            f"significativo, permitindo testes clínicos seguros, confortáveis e sem risco para os voluntários da policlínica."
        ),
        "Conceito Global": "Excelente (10,0) – Carga horária de 20 horas cumprida integralmente com louvor."
    })

    add_campo_box(doc, 6, "Autoavaliação do Aluno (Ganhos de Competências)", {
        "Ganhos Profissionais e Científicos": (
            f"A realização da Prática Profissional na Policlínica Piquet Carneiro permitiu transpor o protótipo desenvolvido em bancada "
            f"para o contexto real da atenção à saúde do SUS. A utilização do Wemos Lolin D32 e da bateria LiPo 600mAh com conector JST "
            f"otimizou o perfil de usabilidade do wearable, garantindo peso insignificante (7,5g na placa) e alta densidade de processamento "
            f"(240MHz, 16MB Flash, 8MB PSRAM e slot MicroSD). A vivência consolidou minhas competências em Internet das Coisas Médicas (IoMT), "
            f"validação experimental cruzada com baropodometria e condução ética de ensaios com seres humanos."
        ),
        "Atitudes e Habilidades Consolidadas": (
            "Habilidade de escuta e comunicação multidisciplinar com a equipe de fisioterapia; rigor no manuseio de instrumentação diagnóstica "
            "padrão-ouro; domínio dos princípios de biossegurança clínica e compromisso ético com a privacidade e conforto do voluntário."
        )
    })

    add_campo_box(doc, 7, "Avaliação Consensual Aluno-Supervisor", {
        "Pontos Fortes": (
            f"Alta inovação tecnológica de baixo custo ({CUSTO_UNITARIO}); ergonomia superior proporcionada pelo Lolin D32 e bateria LiPo 600mAh; "
            f"aplicabilidade direta à prevenção de lesões e amputações no SUS; excelente integração entre hardware, app móvel e Random Forest."
        ),
        "Recomendações Futuras": (
            "Prosseguir com o aprimoramento contínuo do firmware, explorando o slot MicroSD integrado para logging contínuo em estudos "
            "de campo prolongados no ambiente domiciliar dos pacientes."
        )
    })

    # CAMPO 8 - Assinaturas
    p8_h = doc.add_paragraph()
    format_paragraph(p8_h, 10, 4, 1.15)
    r8 = p8_h.add_run("CAMPO 8 – ASSINATURAS ELETRÔNICAS E CARGA HORÁRIA TOTAL")
    r8.font.name = "Arial"
    r8.font.size = Pt(11)
    r8.font.bold = True
    r8.font.color.rgb = RGBColor(20, 50, 100)

    t8 = doc.add_table(rows=2, cols=2)
    t8.alignment = WD_TABLE_ALIGNMENT.CENTER
    t8.columns[0].width = Inches(3.4)
    t8.columns[1].width = Inches(3.4)

    c_aluno = t8.cell(0, 0)
    c_sup = t8.cell(0, 1)
    c_data = t8.cell(1, 0)
    c_ch = t8.cell(1, 1)

    for c in [c_aluno, c_sup, c_data, c_ch]:
        set_cell_background(c, "FAFAFA")
        set_cell_margins(c, 100, 100, 120, 120)
        set_cell_border(c, sz="2")

    p_a = c_aluno.paragraphs[0]
    p_a.add_run(f"ALUNO (MESTRANDO):\n\n___________________________________\n{AUTOR}\nAssinatura GOV.BR / ICP-Brasil")
    p_a.runs[0].font.size = Pt(8.5)

    p_s = c_sup.paragraphs[0]
    p_s.add_run("SUPERVISOR DA ATIVIDADE NA PPC:\n\n___________________________________\n[Nome do Supervisor na PPC]\nAssinatura com Carimbo / GOV.BR")
    p_s.runs[0].font.size = Pt(8.5)

    p_d = c_data.paragraphs[0]
    p_d.add_run("DATA DE CONCLUSÃO:\n\n_____ / _____ / 2026")
    p_d.runs[0].font.size = Pt(8.5)

    p_c = c_ch.paragraphs[0]
    p_c.add_run("TEMPO TOTAL DO ESTÁGIO:\n\n20 (vinte) horas presenciais")
    p_c.runs[0].font.size = Pt(8.5)

    out_path = os.path.join("02_Mestrado_Pratica_Profissional", "Portfolio_Pratica_Profissional_Telessaude_UERJ.docx")
    doc.save(out_path)
    print(f"Salvo com sucesso: {out_path}")

# ==============================================================================
# 2. TERMO DE ANUÊNCIA INSTITUCIONAL (PPC / UERJ)
# ==============================================================================
def create_anuencia_doc():
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.8)
    sec.bottom_margin = Inches(0.8)
    sec.left_margin = Inches(0.9)
    sec.right_margin = Inches(0.9)

    h_p = doc.add_paragraph()
    h_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(h_p, 0, 4, 1.0)
    r = h_p.add_run("GOVERNO DO ESTADO DO RIO DE JANEIRO\nSECRETARIA DE ESTADO DE CIÊNCIA, TECNOLOGIA E INOVAÇÃO\nUNIVERSIDADE DO ESTADO DO RIO DE JANEIRO\nHOSPITAL UNIVERSITÁRIO PEDRO ERNESTO / POLICLÍNICA PIQUET CARNEIRO\nCOMITÊ DE ÉTICA EM PESQUISA (CEP/HUPE)")
    r.font.name = "Arial"
    r.font.size = Pt(9)
    r.font.bold = True

    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(t_p, 8, 2, 1.1)
    r_t = t_p.add_run("TERMO DE ANUÊNCIA / AUTORIZAÇÃO INSTITUCIONAL")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(12)
    r_t.font.bold = True

    v_p = doc.add_paragraph()
    v_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    format_paragraph(v_p, 0, 10, 1.0)
    r_v = v_p.add_run("Versão datada de 01/06/2026 – CEP/HUPE")
    r_v.font.name = "Arial"
    r_v.font.size = Pt(8)
    r_v.font.italic = True

    p_decl = doc.add_paragraph()
    p_decl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_decl, 6, 6, 1.25)
    p_decl.add_run(
        "Declaro para os devidos fins que o Serviço de Fisioterapia / Setor de Baropodometria da "
        "POLICLÍNICA PIQUET CARNEIRO (PPC) do COMPLEXO HUPE/UERJ, tem pleno conhecimento e manifesta total concordância "
        "com a condução e realização da Pesquisa intitulada: "
    )
    r_tit = p_decl.add_run(f"“{TITULO_SUBMISSAO_CEP}”, ")
    r_tit.font.bold = True
    p_decl.add_run(
        f"sendo responsáveis o pesquisador discente {AUTOR} (Mestrando) e a docente {ORIENTADORA} "
        f"(Orientadora/Pesquisadora Responsável), não havendo qualquer oposição à sua realização. "
        f"Ratifico a ciência de que o projeto só poderá iniciar as etapas de coleta de dados com seres humanos "
        f"após a expressa avaliação e aprovação consubstanciada do Comitê de Ética em Pesquisa do Hospital Universitário "
        f"Pedro Ernesto da Universidade do Estado do Rio de Janeiro (CEP/HUPE/UERJ)."
    )

    p_per = doc.add_paragraph()
    format_paragraph(p_per, 6, 4, 1.2)
    r_per = p_per.add_run("Período previsto para a coleta dos dados: ")
    r_per.font.bold = True
    p_per.add_run("(01/11/2026 a 28/02/2027)\n")
    r_forma = p_per.add_run("Os dados serão coletados através de:\n")
    r_forma.font.bold = True
    p_per.add_run(
        "(   ) ENTREVISTA       (   ) QUESTIONÁRIO       (   ) PRONTUÁRIO\n"
        "( X ) OUTROS: Avaliação biomecânica comparativa através de plataforma de baropodometria e uso experimental temporário "
        "de palmilha inteligente instrumentada (microcontrolador Wemos Lolin D32, sensores FSR 402, DHT22 e bateria LiPo 3.7V 600mAh) "
        "em 3 a 5 voluntários adultos com Diabetes Mellitus."
    )

    p_ass_chefe = doc.add_paragraph()
    p_ass_chefe.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_ass_chefe, 16, 2, 1.1)
    p_ass_chefe.add_run("____________________________________________________________\n")
    r_nc = p_ass_chefe.add_run("[Nome do Responsável / Chefe do Serviço / Setor da PPC]\n")
    r_nc.font.bold = True
    p_ass_chefe.add_run("Responsável pelo Setor de Baropodometria / Serviço de Fisioterapia da PPC/UERJ\nAssinatura com Carimbo ou Certificação Digital GOV.BR\nData: _____ / _____ / 2026")
    p_ass_chefe.runs[0].font.size = Pt(9.5)

    p_div2 = doc.add_paragraph()
    format_paragraph(p_div2, 4, 4, 1.0)
    p_div2.add_run("―" * 55).font.color.rgb = RGBColor(180, 180, 180)

    # Vínculo
    p_vinc_h = doc.add_paragraph()
    format_paragraph(p_vinc_h, 4, 4, 1.15)
    r_vh = p_vinc_h.add_run("VÍNCULO DO PESQUISADOR COM O COMPLEXO HUPE / UERJ")
    r_vh.font.name = "Arial"
    r_vh.font.size = Pt(10.5)
    r_vh.font.bold = True

    p_vinc_txt = doc.add_paragraph()
    format_paragraph(p_vinc_txt, 2, 6, 1.15)
    p_vinc_txt.add_run(
        f"Faculdade / Departamento / Programa: Faculdade de Ciências Médicas / Núcleo de Telessaúde e Saúde Digital "
        f"({PROGRAMA})\n"
        f"Nível: (   ) Graduação   (   ) Especialização   ( X ) Mestrado Profissional   (   ) Doutorado\n"
        f"Condição: ( X ) Aluno Regular: {AUTOR}    Matrícula: [Número de Matrícula]"
    )

    # Termo de responsabilidade
    p_resp_h = doc.add_paragraph()
    format_paragraph(p_resp_h, 6, 3, 1.15)
    r_rh = p_resp_h.add_run("TERMO DE RESPONSABILIDADE DO PESQUISADOR E ORIENTADOR")
    r_rh.font.name = "Arial"
    r_rh.font.size = Pt(10.5)
    r_rh.font.bold = True

    p_resp = doc.add_paragraph()
    p_resp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_resp, 2, 8, 1.2)
    p_resp.add_run(
        "Declaramos que nos responsabilizamos pelo andamento, realização e conclusão do projeto, comprometendo-nos a cumprir "
        "integralmente a Resolução CNS nº 466/2012, a Resolução CNS nº 510/2016 e as normas complementares, assegurando "
        "o resguardo da segurança, sigilo e bem-estar dos participantes da pesquisa. Comprometemo-nos a enviar relatórios "
        "parciais e final ao CEP/HUPE, a comunicar imediatamente qualquer interrupção ou evento adverso e a garantir que os dados "
        "obtidos serão empregados única e exclusivamente para a consecução dos objetivos científicos do presente projeto."
    )

    t_ass = doc.add_table(rows=1, cols=2)
    t_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_ass.columns[0].width = Inches(3.4)
    t_ass.columns[1].width = Inches(3.4)

    p_a1 = t_ass.cell(0, 0).paragraphs[0]
    p_a1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_a1.add_run(f"_____________________________________\n{AUTOR}\nPesquisador Discente (Mestrando)\nData: ___/___/2026")
    p_a1.runs[0].font.size = Pt(9)

    p_a2 = t_ass.cell(0, 1).paragraphs[0]
    p_a2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_a2.add_run(f"_____________________________________\n{ORIENTADORA}\nPesquisadora Responsável / Orientadora\nData: ___/___/2026")
    p_a2.runs[0].font.size = Pt(9)

    out_path = os.path.join("01_Submissao_CEP_Palmilha_PPC", "04_Termo_de_Anuencia_Institucional_PPC.docx")
    doc.save(out_path)
    print(f"Salvo com sucesso: {out_path}")

# ==============================================================================
# 3. TERMO DE CONSENTIMENTO LIVRE E ESCLARECIDO (TCLE - ADULTO)
# ==============================================================================
def create_tcle_doc():
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.8)
    sec.bottom_margin = Inches(0.8)
    sec.left_margin = Inches(0.9)
    sec.right_margin = Inches(0.9)

    h_p = doc.add_paragraph()
    h_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(h_p, 0, 4, 1.0)
    r = h_p.add_run("UNIVERSIDADE DO ESTADO DO RIO DE JANEIRO – UERJ\nHOSPITAL UNIVERSITÁRIO PEDRO ERNESTO – HUPE\nCOMITÊ DE ÉTICA EM PESQUISA (CEP/HUPE)")
    r.font.name = "Arial"
    r.font.size = Pt(9)
    r.font.bold = True

    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(t_p, 6, 2, 1.1)
    r_t = t_p.add_run("TERMO DE CONSENTIMENTO LIVRE E ESCLARECIDO – MAIORES DE IDADE")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(11.5)
    r_t.font.bold = True

    v_p = doc.add_paragraph()
    v_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    format_paragraph(v_p, 0, 8, 1.0)
    r_v = v_p.add_run("Versão desse documento datado de 01/06/2026")
    r_v.font.name = "Arial"
    r_v.font.size = Pt(8)
    r_v.font.italic = True

    p_convite = doc.add_paragraph()
    p_convite.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_convite, 4, 6, 1.25)
    p_convite.add_run(
        f"Você está sendo convidado(a) a participar como voluntário(a) da pesquisa científica intitulada "
        f"“{TITULO_SUBMISSAO_CEP}”, realizada no âmbito do {PROGRAMA} da Faculdade de Ciências Médicas da {INSTITUICAO}, "
        f"que diz respeito à Dissertação de Mestrado Profissional do pesquisador {AUTOR}, sob orientação da {ORIENTADORA}.\n\n"
        f"Antes de decidir se aceita participar, é fundamental que você compreenda todas as informações a seguir. "
        f"Leia com calma, tire todas as dúvidas que desejar e sinta-se plenamente livre para aceitar ou recusar."
    )

    secoes = [
        ("1. OBJETIVO DO ESTUDO",
         "O objetivo deste estudo é avaliar se uma palmilha inteligente inovadora (desenvolvida com sensores eletrônicos de baixo custo "
         "capazes de medir simultaneamente a pressão em pontos críticos da sola do pé, a temperatura e a umidade interna do calçado) funciona "
         "de maneira precisa quando comparada a uma esteira eletrônica de pressão (baropodômetro), já utilizada na Policlínica Piquet Carneiro. "
         "O estudo busca contribuir diretamente para a prevenção precoce de feridas (úlceras) e amputações em pessoas com Diabetes Mellitus atendidas no SUS."),

        ("2. COMO SERÁ A SUA PARTICIPAÇÃO (PROCEDIMENTOS)",
         "A sua participação ocorrerá em uma única sessão presencial de cerca de 20 a 30 minutos na Policlínica Piquet Carneiro (PPC/UERJ):\n"
         "a) Inicialmente, você responderá a perguntas simples sobre sua saúde, tempo de diabetes e numeração do seu calçado;\n"
         "b) Você receberá uma meia de proteção descartável nova e calçará um calçado confortável contendo a palmilha inteligente instrumentada;\n"
         "c) Você realizará uma caminhada curta e confortável sobre a esteira de baropodometria, primeiro ficando parado em pé por cerca de 10 segundos e depois caminhando em linha reta em ritmo normal;\n"
         "d) Durante a caminhada, os sensores registrarão automaticamente a distribuição da pressão do seu pé, a temperatura e a umidade interna;\n"
         "e) Ao final, você responderá a um breve questionário opinando se achou a palmilha confortável, leve e estável para caminhar.\n"
         "Todo o procedimento é totalmente indolor, não invasivo, sem agulhas e sem uso de qualquer medicamento."),

        ("3. POTENCIAIS RISCOS E MEDIDAS DE SEGURANÇA",
         "Os riscos envolvidos nesta pesquisa são considerados MÍNIMOS:\n"
         "• Risco de leve cansaço ou desequilíbrio durante os passos: Para evitar qualquer tropeço, a caminhada é de curta distância em piso totalmente plano, sempre supervisionada de perto pelo pesquisador, que estará ao seu lado pronto para dar apoio físico.\n"
         "• Segurança Elétrica e Ergonômica: O módulo de processamento eletrônico (Wemos Lolin D32) é ultraleve (pesa apenas 7,5 gramas) e é alimentado por uma pequena bateria recarregável plana de Polímero de Lítio (LiPo 3,7V e 600mAh), equipada com conector polarizado de segurança JST PH 2.0mm. Trata-se de uma voltagem extremamente baixa (semelhante à de fones de ouvido sem fio), com circuito de proteção que elimina qualquer possibilidade de choque elétrico, curto-circuito ou aquecimento.\n"
         "• Biossegurança e Higiene: A palmilha é rigorosamente higienizada e desinfetada com álcool a 70% antes e após cada participante, além do uso obrigatório de meia protetora descartável de uso individual.\n"
         "Você pode solicitar a interrupção imediata do teste a qualquer instante caso sinta qualquer incômodo."),

        ("4. BENEFÍCIOS DA PESQUISA",
         "A sua participação não trará uma cura direta ou benefício médico imediato no momento do teste. Todavia, a sua colaboração é indispensável "
         "para viabilizar uma tecnologia nacional de baixo custo que, integrada a modelos de Inteligência Artificial, poderá ser distribuída no SUS "
         "para alertar pacientes com diabetes antes do surgimento de feridas graves, reduzindo internações e amputações."),

        ("5. GARANTIA DE SIGILO E PRIVACIDADE (LGPD)",
         "Os dados coletados serão usados estritamente para finalidades acadêmicas e científicas. A sua privacidade é protegida por lei e pelas normas éticas "
         "do Conselho Nacional de Saúde. O seu nome, imagem ou documento pessoal jamais serão divulgados. Os dados serão identificados apenas por códigos "
         "(ex.: VOL-01) e ficarão armazenados em meio digital seguro sob a guarda do pesquisador pelo período de 5 (cinco) anos."),

        ("6. LIBERDADE DE RECUSA E DESISTÊNCIA",
         "A participação é totalmente voluntária. Você tem o direito de recusar-se a participar ou de desistir a qualquer momento, sem necessidade "
         "de qualquer explicação e sem que isso traga qualquer prejuízo ao seu atendimento de saúde presente ou futuro na Policlínica Piquet Carneiro ou na UERJ."),

        ("7. CUSTOS, RESSARCIMENTO E INDENIZAÇÃO",
         "Você não terá nenhum custo para participar deste estudo, nem receberá pagamento financeiro por sua participação voluntária. Caso haja qualquer "
         "gasto comprovado decorrente diretamente da sua ida à pesquisa (como transporte ou alimentação), você terá direito ao ressarcimento imediato e integral. "
         "Adicionalmente, fica assegurada a assistência imediata e a indenização em caso de eventuais danos comprovadamente decorrentes da pesquisa, nos termos da Resolução CNS nº 466/2012."),

        ("8. DADOS DE CONTATO DO PESQUISADOR E DO CEP/HUPE",
         f"Você receberá uma via deste documento rubricada e assinada. Em caso de dúvidas sobre o estudo, contate o pesquisador responsável:\n"
         f"• Pesquisador Discente: {AUTOR}\n"
         f"• Orientadora: {ORIENTADORA}\n"
         f"• E-mail: cytchrisley@gmail.com | Programa de Pós-Graduação em Telessaúde e Saúde Digital – UERJ (Prédio CePeM, 3º andar, Vila Isabel, RJ).\n\n"
         f"Para considerações ou dúvidas sobre os aspectos éticos da pesquisa, contate o Comitê de Ética em Pesquisa do Hospital Universitário Pedro Ernesto (CEP/HUPE):\n"
         f"• Endereço: Av. Vinte e Oito de Setembro, nº 77 – Prédio do CePeM, 2º andar, sala 33 – Vila Isabel, Rio de Janeiro – RJ – CEP: 20551-030.\n"
         f"• Telefone / WhatsApp: (21) 2868-8253 | E-mail: cep@hupe.uerj.br (Segunda a sexta-feira, das 09:00h às 13:00h).")
    ]

    for sec_title, sec_text in secoes:
        p_sh = doc.add_paragraph()
        format_paragraph(p_sh, 6, 2, 1.15)
        r_sh = p_sh.add_run(sec_title)
        r_sh.font.name = "Arial"
        r_sh.font.size = Pt(10)
        r_sh.font.bold = True
        r_sh.font.color.rgb = RGBColor(0, 0, 0)

        p_st = doc.add_paragraph()
        p_st.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_st, 1, 4, 1.2)
        r_st = p_st.add_run(sec_text)
        r_st.font.name = "Arial"
        r_st.font.size = Pt(9.5)
        r_st.font.color.rgb = RGBColor(0, 0, 0)

    # Consentimento
    p_c_h = doc.add_paragraph()
    p_c_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_c_h, 8, 4, 1.15)
    r_ch = p_c_h.add_run("DECLARAÇÃO DE CONSENTIMENTO PÓS-ESCLARECIDO")
    r_ch.font.name = "Arial"
    r_ch.font.size = Pt(10.5)
    r_ch.font.bold = True

    p_c_t = doc.add_paragraph()
    p_c_t.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_c_t, 2, 8, 1.2)
    p_c_t.add_run(
        f"Eu, ______________________________________________________________________, declaro que li e compreendi as informações "
        f"acima descritas sobre a pesquisa “{TITULO_SUBMISSAO_CEP}”. Fui devidamente informado(a) dos objetivos, métodos, riscos mínimos "
        f"e potenciais benefícios. Tive a oportunidade de fazer perguntas e obtive respostas claras. Concordo voluntariamente em participar "
        f"deste estudo, ciente de que posso revogar meu consentimento a qualquer instante sem sofrer qualquer penalidade."
    )

    t_ass = doc.add_table(rows=1, cols=2)
    t_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_ass.columns[0].width = Inches(3.4)
    t_ass.columns[1].width = Inches(3.4)

    p_v = t_ass.cell(0, 0).paragraphs[0]
    p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_v.add_run("_____________________________________\nAssinatura do(a) Participante Voluntário(a)\n\nData: _____ / _____ / 2026")
    p_v.runs[0].font.size = Pt(9)

    p_p = t_ass.cell(0, 1).paragraphs[0]
    p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p.add_run(f"_____________________________________\n{AUTOR}\nPesquisador Responsável\nData: _____ / _____ / 2026")
    p_p.runs[0].font.size = Pt(9)

    p_warn = doc.add_paragraph()
    p_warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_warn, 10, 0, 1.0)
    r_w = p_warn.add_run("AO FINAL DO PREENCHIMENTO, MANTER TODO O ARQUIVO NA FORMATAÇÃO EM COR PRETA!")
    r_w.font.name = "Arial"
    r_w.font.size = Pt(8)
    r_w.font.bold = True

    out_path = os.path.join("01_Submissao_CEP_Palmilha_PPC", "05_TCLE_Adulto_Palmilha_Inteligente.docx")
    doc.save(out_path)
    print(f"Salvo com sucesso: {out_path}")

# ==============================================================================
# 4. DECLARAÇÃO DE ISENÇÃO DE CUSTOS INSTITUCIONAIS
# ==============================================================================
def create_isencao_doc():
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)

    h_p = doc.add_paragraph()
    h_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(h_p, 0, 4, 1.0)
    r = h_p.add_run("UNIVERSIDADE DO ESTADO DO RIO DE JANEIRO – UERJ\nCENTRO BIOMÉDICO – FACULDADE DE CIÊNCIAS MÉDICAS\nPROGRAMA DE PÓS-GRADUAÇÃO STRICTO SENSU EM TELESSAÚDE E SAÚDE DIGITAL\nCOMITÊ DE ÉTICA EM PESQUISA (CEP/HUPE)")
    r.font.name = "Arial"
    r.font.size = Pt(9)
    r.font.bold = True

    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(t_p, 12, 2, 1.1)
    r_t = t_p.add_run("DECLARAÇÃO DE ISENÇÃO DE CUSTOS INSTITUCIONAIS")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(12)
    r_t.font.bold = True

    v_p = doc.add_paragraph()
    v_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    format_paragraph(v_p, 0, 14, 1.0)
    r_v = v_p.add_run("Versão desse documento datada de 01/06/2026")
    r_v.font.name = "Arial"
    r_v.font.size = Pt(8)
    r_v.font.italic = True

    p_inst = doc.add_paragraph()
    format_paragraph(p_inst, 4, 4, 1.2)
    p_inst.add_run("Nome da Unidade / Programa: ").font.bold = True
    p_inst.add_run(f"Faculdade de Ciências Médicas – {PROGRAMA} ({INSTITUICAO})\n")
    p_inst.add_run("Local da Coleta de Dados: ").font.bold = True
    p_inst.add_run(LOCAL_COLETA)

    p_decl = doc.add_paragraph()
    p_decl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_decl, 10, 14, 1.3)
    p_decl.add_run(
        f"Eu, {AUTOR}, pesquisador discente do Mestrado Profissional em Telessaúde e Saúde Digital da {INSTITUICAO}, "
        f"sob orientação acadêmica da {ORIENTADORA}, declaro para os devidos fins legais e éticos que a pesquisa científica intitulada:\n\n"
    )
    r_tit = p_decl.add_run(f"“{TITULO_SUBMISSAO_CEP}”\n\n")
    r_tit.font.bold = True
    p_decl.add_run(
        f"está sob minha inteira responsabilidade técnica e orçamentária e NÃO IRÁ GERAR CUSTO de qualquer natureza para a "
        f"instituição envolvida (Policlínica Piquet Carneiro, Hospital Universitário Pedro Ernesto ou Universidade do Estado "
        f"do Rio de Janeiro), nem tampouco para os participantes voluntários da pesquisa.\n\n"
        f"Declaro que todos os custos referentes aos componentes do protótipo da palmilha instrumentada (microcontrolador "
        f"Wemos Lolin D32, sensores FSR 402, sensores DHT22, bateria LiPo 3.7V 600mAh com conector JST PH 2.0mm e insumos), "
        f"avaliados em {CUSTO_UNITARIO}, bem como os insumos de proteção individual e assepsia (álcool a 70% e meias descartáveis de barreira) "
        f"são integralmente custeados com recursos próprios do pesquisador discente, sem demandar verbas orçamentárias do Sistema Único de Saúde (SUS)."
    )

    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_ass, 24, 2, 1.15)
    p_ass.add_run("_________________________________________________________\n")
    r_na = p_ass.add_run(f"{AUTOR}\n")
    r_na.font.bold = True
    p_ass.add_run(f"Pesquisador Principal / Mestrando em Telessaúde e Saúde Digital\nAssinatura com Certificação Digital GOV.BR / ICP-Brasil\n\nRio de Janeiro, _____ de ____________________ de 2026.")
    p_ass.runs[0].font.size = Pt(9.5)

    out_path = os.path.join("01_Submissao_CEP_Palmilha_PPC", "09_Declaracao_de_Isencao_de_Custos.docx")
    doc.save(out_path)
    print(f"Salvo com sucesso: {out_path}")

# ==============================================================================
# 5. INSTRUMENTO DE COLETA DE DADOS (FICHA DE ENSAIO EXPERIMENTAL)
# ==============================================================================
def create_instrumento_coleta_doc():
    doc = docx.Document()
    add_header_block(doc, "INSTRUMENTO DE COLETA DE DADOS", "FICHA DE ENSAIO EXPERIMENTAL: PALMILHA INSTRUMENTADA X BAROPODÔMETRO NA PPC")

    p_meta = doc.add_paragraph()
    format_paragraph(p_meta, 4, 6, 1.15)
    p_meta.add_run("Código de Identificação do Voluntário: ").font.bold = True
    p_meta.add_run("VOL-0___ (Sigilo Preservado - LGPD)         ")
    p_meta.add_run("Data da Coleta: ").font.bold = True
    p_meta.add_run("____/____/2026\n")
    p_meta.add_run("Pesquisador Examinador: ").font.bold = True
    p_meta.add_run(f"{AUTOR}                    ")
    p_meta.add_run("Local: ").font.bold = True
    p_meta.add_run("PPC / UERJ – Laboratório de Baropodometria")

    # Seção 1: Dados Clínicos e Antropométricos do Paciente com Diabetes
    p_s1 = doc.add_paragraph()
    format_paragraph(p_s1, 8, 3, 1.15)
    r1 = p_s1.add_run("1. DADOS CLÍNICOS, ANTROPOMÉTRICOS E DE CALÇADO (DIABETES MELLITUS TIPO 2)")
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(20, 50, 100)

    t1 = doc.add_table(rows=2, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.columns[0].width = Inches(1.7)
    t1.columns[1].width = Inches(1.7)
    t1.columns[2].width = Inches(1.7)
    t1.columns[3].width = Inches(1.7)

    campos1 = [
        ("Idade / Sexo:", "____ anos | (  ) F  (  ) M"),
        ("Tempo de Diagnóstico DM2:", "____ anos"),
        ("Massa Corporal e Estatura:", "____ kg | ____ cm"),
        ("Tamanho do Calçado (BR):", "Nº ____ BR"),
        ("Histórico de Úlcera Plantar:", "(  ) Não   (  ) Sim (cicatrizada)"),
        ("Teste de Sensibilidade (Monofilamento):", "(  ) Preservada   (  ) Reduzida"),
        ("Presença de Calosidades:", "(  ) Não   (  ) Sim: antepé / calcanhar"),
        ("Tipo de Arco Plantar:", "(  ) Neutro   (  ) Cavo   (  ) Plano")
    ]
    for idx, (lbl, val) in enumerate(campos1):
        ri = idx // 4
        ci = idx % 4
        cell = t1.cell(ri, ci)
        set_cell_background(cell, "F9FBFD")
        set_cell_margins(cell, 60, 60, 80, 80)
        set_cell_border(cell, sz="2")
        p = cell.paragraphs[0]
        format_paragraph(p, 1, 1, 1.1)
        r = p.add_run(f"{lbl}\n{val}")
        r.font.size = Pt(8.5)

    # Seção 2: Baropodômetro da PPC (Padrão-Ouro)
    p_s2 = doc.add_paragraph()
    format_paragraph(p_s2, 10, 3, 1.15)
    r2 = p_s2.add_run("2. LEITURAS NA PLATAFORMA DE BAROPODOMETRIA DA PPC (PADRÃO-OURO)")
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(20, 50, 100)

    t2 = doc.add_table(rows=4, cols=4)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers2 = ["Variável Biomecânica", "Antepé (Metatarsos)", "Mediopé (Arco)", "Retropé (Calcâneo)"]
    for ci, h in enumerate(headers2):
        cell = t2.cell(0, ci)
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    linhas2 = [
        ("Pressão Pico Estática (kPa):", "____________ kPa", "____________ kPa", "____________ kPa"),
        ("Pressão Pico Dinâmica na Marcha (kPa):", "____________ kPa", "____________ kPa", "____________ kPa"),
        ("Distribuição Relativa de Carga (%):", "____________ %", "____________ %", "____________ %")
    ]
    for ri, dados in enumerate(linhas2, start=1):
        for ci, val in enumerate(dados):
            cell = t2.cell(ri, ci)
            set_cell_background(cell, "FFFFFF" if ci > 0 else "F4F6F9")
            set_cell_margins(cell, 60, 60, 80, 80)
            set_cell_border(cell, sz="2")
            p = cell.paragraphs[0]
            format_paragraph(p, 1, 1, 1.1)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if ci == 0:
                r.font.bold = True

    # Seção 3: Palmilha Inteligente (Wemos Lolin D32 + FSR 402)
    p_s3 = doc.add_paragraph()
    format_paragraph(p_s3, 10, 3, 1.15)
    r3 = p_s3.add_run("3. LEITURAS DA PALMILHA (SENSORES FSR 402 E ADC DO WEMOS LOLIN D32)")
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(20, 50, 100)

    t3 = doc.add_table(rows=4, cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers3 = ["Parâmetro do Protótipo", "Sensor 1 (Antepé)", "Sensor 2 (Mediopé)", "Sensor 3 (Retropé)"]
    for ci, h in enumerate(headers3):
        cell = t3.cell(0, ci)
        set_cell_background(cell, "2E5B88")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    linhas3 = [
        ("Leitura ADC (0 a 4095 pts):", "____________ pts", "____________ pts", "____________ pts"),
        ("Pressão Convertida (kPa):", "____________ kPa", "____________ kPa", "____________ kPa"),
        ("Desvio frente ao Baropodômetro (%):", "____________ %", "____________ %", "____________ %")
    ]
    for ri, dados in enumerate(linhas3, start=1):
        for ci, val in enumerate(dados):
            cell = t3.cell(ri, ci)
            set_cell_background(cell, "FFFFFF" if ci > 0 else "F4F6F9")
            set_cell_margins(cell, 60, 60, 80, 80)
            set_cell_border(cell, sz="2")
            p = cell.paragraphs[0]
            format_paragraph(p, 1, 1, 1.1)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if ci == 0:
                r.font.bold = True

    # Seção 4: Microclima (DHT22) e Usabilidade
    p_s4 = doc.add_paragraph()
    format_paragraph(p_s4, 10, 3, 1.15)
    r4 = p_s4.add_run("4. MONITORAMENTO MICROCLIMÁTICO (DHT22) E USABILIDADE DO PROTÓTIPO")
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(20, 50, 100)

    t4 = doc.add_table(rows=2, cols=3)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    t4.columns[0].width = Inches(2.2)
    t4.columns[1].width = Inches(2.3)
    t4.columns[2].width = Inches(2.3)

    c_clima = [
        ("Temperatura Inicial (pré-marcha):", "_______ ºC"),
        ("Temperatura Final (pós-marcha):", "_______ ºC (ΔT = ______ ºC)"),
        ("Umidade Relativa Inicial:", "_______ %"),
        ("Umidade Relativa Final:", "_______ % (ΔUR = ______ %)")
    ]
    for idx, (lbl, val) in enumerate(c_clima):
        ri = idx // 2
        ci = idx % 2
        cell = t4.cell(ri, ci)
        set_cell_background(cell, "FAFAFA")
        set_cell_margins(cell, 60, 60, 80, 80)
        set_cell_border(cell, sz="2")
        p = cell.paragraphs[0]
        format_paragraph(p, 1, 1, 1.1)
        r = p.add_run(f"{lbl}\n{val}")
        r.font.size = Pt(8.5)

    c_aval1 = t4.cell(0, 2)
    set_cell_background(c_aval1, "F0F8FF")
    set_cell_margins(c_aval1, 60, 60, 80, 80)
    set_cell_border(c_aval1, sz="2")
    p_a1 = c_aval1.paragraphs[0]
    p_a1.add_run("Estabilidade da Marcha:\n(1 a 5, 5 = Excelente estabilidade)\nNota do Paciente: [    ]").font.size = Pt(8.5)

    c_aval2 = t4.cell(1, 2)
    set_cell_background(c_aval2, "F0F8FF")
    set_cell_margins(c_aval2, 60, 60, 80, 80)
    set_cell_border(c_aval2, sz="2")
    p_a2 = c_aval2.paragraphs[0]
    p_a2.add_run("Conforto e Leveza (Módulo 7,5g):\n(1 a 5, 5 = Muito confortável e leve)\nNota do Paciente: [    ]").font.size = Pt(8.5)

    # Checklist de Biossegurança
    p_bio = doc.add_paragraph()
    format_paragraph(p_bio, 10, 4, 1.15)
    r_b = p_bio.add_run("5. CHECKLIST DE ASSEPSIA E SEGURANÇA OPERACIONAL")
    r_b.font.bold = True
    r_b.font.color.rgb = RGBColor(20, 50, 100)

    p_check = doc.add_paragraph()
    format_paragraph(p_check, 2, 4, 1.15)
    p_check.add_run(
        "[  ] Desinfecção mecânica da palmilha com álcool 70% realizada antes do início do teste.\n"
        "[  ] Colocação de meia descartável nova no pé do voluntário antes de calçar a palmilha.\n"
        "[  ] Inspeção física do pé do voluntário (ausência de feridas abertas ou úlceras ativas).\n"
        "[  ] Verificação da fixação do conector JST PH 2.0mm e isolamento da bateria LiPo 600mAh.\n"
        "[  ] Confirmação da gravação de segurança no cartão MicroSD do Wemos Lolin D32.\n"
        "[  ] Desinfecção com álcool 70% repetida imediatamente após o término do ensaio.\n"
        "[  ] Ausência de dor aguda, desequilíbrio ou sensação de aquecimento térmico anômalo."
    ).font.size = Pt(8.5)

    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    format_paragraph(p_sign, 12, 0, 1.1)
    p_sign.add_run(f"Assinatura do Pesquisador Responsável: {AUTOR} _____________________________________").font.size = Pt(9)

    out_path = os.path.join("01_Submissao_CEP_Palmilha_PPC", "Instrumento_Coleta_de_Dados_Palmilha_Baropodometro.docx")
    doc.save(out_path)
    print(f"Salvo com sucesso: {out_path}")

# ==============================================================================
# 6. PROJETO DE PESQUISA NA ÍNTEGRA (NORMA OPERACIONAL CNS Nº 001/2013)
# ==============================================================================
def create_projeto_integra_doc():
    doc = docx.Document()
    add_header_block(doc, "PROJETO DE PESQUISA NA ÍNTEGRA", "ESTRUTURADO SEGUNDO A NORMA OPERACIONAL CNS Nº 001/2013 E RESOLUÇÃO CNS Nº 466/2012")

    # Identificação
    p_id = doc.add_paragraph()
    format_paragraph(p_id, 4, 8, 1.2)
    p_id.add_run("TÍTULO DA PESQUISA: ").font.bold = True
    p_id.add_run(f"{TITULO_SUBMISSAO_CEP.upper()}\n\n")
    p_id.add_run("PESQUISADOR PRINCIPAL (MESTRANDO): ").font.bold = True
    p_id.add_run(f"{AUTOR}\n")
    p_id.add_run("ORIENTADORA / PESQUISADORA RESPONSÁVEL: ").font.bold = True
    p_id.add_run(f"{ORIENTADORA}\n")
    p_id.add_run("INSTITUIÇÃO PROPONENTE: ").font.bold = True
    p_id.add_run(f"{INSTITUICAO} – Faculdade de Ciências Médicas / {PROGRAMA}\n")
    p_id.add_run("INSTITUIÇÃO COPARTICIPANTE / LOCAL DE CAMPO: ").font.bold = True
    p_id.add_run(f"{LOCAL_COLETA}\n")
    p_id.add_run("ÁREA DE CONHECIMENTO: ").font.bold = True
    p_id.add_run("Saúde Coletiva / Telessaúde / Saúde Digital e Engenharia Biomédica")

    # Resumo
    p_res_h = doc.add_paragraph()
    format_paragraph(p_res_h, 8, 2, 1.15)
    r_rh = p_res_h.add_run("RESUMO EXECUTIVO")
    r_rh.font.bold = True
    r_rh.font.color.rgb = RGBColor(20, 50, 100)

    p_res = doc.add_paragraph()
    p_res.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_res, 2, 8, 1.25)
    p_res.add_run(
        f"O acompanhamento preventivo do pé diabético no Sistema Único de Saúde (SUS) enfrenta desafios cruciais: "
        f"a esporadicidade das consultas clínicas tradicionais e o custo proibitivo dos sistemas comerciais de baropodometria "
        f"e palmilhas laboratoriais. Este projeto propõe e valida uma arquitetura vestível de baixo custo ({CUSTO_UNITARIO}) "
        f"baseada no microcontrolador Wemos Lolin D32 (ESP32 ESP-WROOM-32, 240MHz, 16MB Flash, 8MB PSRAM, slot MicroSD/TF e peso de 7,5g), "
        f"3 sensores piezorresistivos de pressão plantar FSR 402, 1 sensor digital de temperatura e umidade DHT22, alimentada por bateria plana "
        f"de Polímero de Lítio (LiPo 3.7V 600mAh com conector polarizado JST PH 2.0mm) e conectada via Bluetooth Low Energy (BLE) ao aplicativo "
        f"móvel 'Monitor do Pé'. A camada analítica incorpora um modelo preditivo baseado em Random Forest para estimativa precoce de risco "
        f"de ulceração plantar. Para transpor a viabilidade técnica de bancada em evidência biomecânica e clínica, este estudo visa calibrar "
        f"as leituras analógicas do Lolin D32 com o baropodômetro padrão-ouro da Policlínica Piquet Carneiro (PPC/UERJ) e executar "
        f"ensaio piloto de prova de conceito com 3 a 5 voluntários adultos portadores de Diabetes Mellitus tipo 2. "
        f"A pesquisa cumpre estritamente a Resolução CNS nº 466/2012, oferecendo riscos mínimos mitigados por rígidos protocolos "
        f"de biossegurança, design ergonômico ultraleve e supervisão presencial, com obtenção voluntária de TCLE e isenção total de custos institucionais."
    )

    capitulos = [
        ("1. INTRODUÇÃO E JUSTIFICATIVA",
         f"O Diabetes Mellitus representa um dos mais graves desafios contemporâneos de saúde pública. No Brasil, o número crescente "
         f"de pessoas acometidas é acompanhado pela incidência severa de complicações crônicas micro e macrovasculares, com destaque para "
         f"a Neuropatia Periférica Diabética (NPD), presente em até 50% dos pacientes de longa data. A perda da sensibilidade protetora dolorosa "
         f"e proprioceptiva impede que o indivíduo perceba traumas contínuos e picos de pressão plantar anômalos durante a locomoção diária, "
         f"culminando na formação de úlceras e, em estágios avançados, em amputações de membros inferiores.\n\n"
         f"Além da hiperpressão mecânica, a temperatura plantar elevada sinaliza processos inflamatórios teciduais subclínicos, e a umidade "
         f"relativa acumulada no interior do calçado — frequentemente negligenciada pela literatura científica — atua acelerando a maceração "
         f"epidérmica e propiciando infecções secundárias. Apesar da eficácia diagnóstica da baropodometria computadorizada, seus custos "
         f"impedem a universalização na Atenção Primária à Saúde. O desenvolvimento de uma palmilha inteligente instrumentada de baixo custo "
         f"({CUSTO_UNITARIO}), baseada no microcontrolador Wemos Lolin D32 (7,5g) e bateria LiPo de 600mAh com conector JST PH 2.0mm, "
         f"integrada a algoritmos de Inteligência Artificial (Random Forest) e operada no ecossistema de Telessaúde e Saúde Digital, "
         f"oferece uma resposta transformadora para o SUS. A Policlínica Piquet Carneiro (PPC/UERJ) oferece o ambiente acadêmico e assistencial "
         f"ideal para calibrar os sensores e demonstrar a viabilidade clínica do protótipo frente ao padrão-ouro."),

        ("2. HIPÓTESES E OBJETIVOS",
         "2.1. Hipótese Científica:\n"
         "Uma arquitetura vestível compacta e de baixo custo baseada no microcontrolador Wemos Lolin D32, bateria LiPo 600mAh e sensores FSR 402 e DHT22 "
         "é capaz de apresentar correlação estatística significativa com a plataforma de baropodometria convencional, viabilizando o monitoramento preventivo "
         "ergonômico, seguro e contínuo do pé diabético no ambiente do SUS.\n\n"
         "2.2. Objetivo Geral:\n"
         f"Projetar, calibrar e validar funcionalmente uma arquitetura de baixo custo para monitoramento preventivo do pé diabético através "
         f"de palmilha inteligente instrumentada e modelo preditivo baseado em Random Forest, realizando ensaios comparativos de calibração "
         f"e prova de conceito na Policlínica Piquet Carneiro (PPC/UERJ).\n\n"
         "2.3. Objetivos Específicos:\n"
         "• Parametrizar a resposta estática e dinâmica dos sensores FSR 402 e DHT22 no Wemos Lolin D32 com filtro de média móvel no firmware;\n"
         "• Realizar a calibração cruzada com o baropodômetro da PPC para correlacionar leituras de ADC (0 a 4095) com grandezas em kPa;\n"
         "• Executar a prova de conceito com 3 a 5 voluntários adultos com Diabetes Mellitus tipo 2 em ortostase e marcha curta;\n"
         "• Avaliar a estabilidade da transmissão BLE com o app 'Monitor do Pé', a redundância de dados no cartão MicroSD da placa e o conforto percebido;\n"
         "• Estruturar a base de dados para treinamento e refinamento do modelo preditivo Random Forest."),

        ("3. MATERIAL E MÉTODOS",
         "3.1. Delineamento:\n"
         "Estudo observacional, transversal e de desenvolvimento tecnológico em saúde digital com ensaio piloto de prova de conceito.\n\n"
         "3.2. População e Amostra:\n"
         "Amostra por conveniência composta por 3 a 5 voluntários adultos portadores de Diabetes Mellitus tipo 2, atendidos ou vinculados à "
         "comunidade universitária da Policlínica Piquet Carneiro (PPC/UERJ).\n\n"
         "3.3. Critérios de Inclusão:\n"
         "• Pacientes com diagnóstico confirmado de Diabetes Mellitus tipo 2;\n"
         "• Idade ≥ 18 anos, de ambos os sexos;\n"
         "• Marcha autônoma preservada (sem dependência contínua de órteses de locomoção);\n"
         "• Calçado com numeração compatível com o protótipo (38 a 41 BR);\n"
         "• Assinatura voluntária do Termo de Consentimento Livre e Esclarecido (TCLE).\n\n"
         "3.4. Critérios de Exclusão:\n"
         "• Presença de úlceras ativas abertas, fissuras exsudativas ou infecções cutâneas agudas na região plantar;\n"
         "• Amputações prévias maiores ou deformidades osteoarticulares severas que impeçam a permanência em ortostase por 5 minutos;\n"
         "• Histórico de hipersensibilidade de contato a materiais poliméricos ou látex.\n\n"
         "3.5. Componentes Tecnológicos e Arquitetura de Hardware:\n"
         f"O sistema utiliza a arquitetura refinada com especificações avançadas de portabilidade: palmilha com espessura de 30 mm contendo "
         f"3 sensores piezorresistivos FSR 402 nas zonas de maior estresse (antepé e retropé); sensor digital de temperatura e umidade DHT22; "
         f"placa microcontroladora Wemos Lolin D32 V1 (baseada no ESP32 ESP-WROOM-32, clock de 240MHz, 16MB Flash, 8MB PSRAM, slot TF/MicroSD "
         f"para backup local de dados, dimensões de 65x25,4mm e peso de apenas 7,5g); alimentação por bateria plana de Polímero de Lítio "
         f"(LiPo 3.7V 600mAh) conectada via conector polarizado JST PH 2.0mm com circuito de recarga integrado de até 500mA; aplicativo móvel "
         f"'Monitor do Pé' em HTML5/JS com mapa anatômico colorido e alertas sonoros/visuais via Bluetooth Low Energy (BLE).\n\n"
         "3.6. Protocolo Experimental e Biossegurança na PPC:\n"
         "Os testes ocorrerão no setor de baropodometria da PPC. O voluntário calçará meia descartável estéril e calçado com a palmilha. Serão executados "
         "3 ensaios estáticos de 10 segundos e 3 passadas dinâmicas sobre a esteira de pressão. A desinfecção com álcool 70% é realizada antes e após "
         "cada teste. As leituras sincronizadas do baropodômetro e da palmilha serão salvas em arquivos CSV codificados."),

        ("4. ASPECTOS ÉTICOS (RESOLUÇÃO CNS Nº 466/2012)",
         "4.1. Avaliação e Mitigação de Riscos:\n"
         "• Risco Mecânico / Queda: Classificado como MÍNIMO. A marcha é de curtíssimo percurso, em superfície plana antiderrapante, com o pesquisador "
         "ao lado para suporte físico imediato. A extrema leveza do conjunto Lolin D32 + LiPo 600mAh (< 25g) não altera o padrão cinemático da marcha.\n"
         "• Risco Elétrico / Térmico: NULO. A bateria LiPo de 3,7V e 600mAh possui conector JST PH 2.0mm com polaridade garantida e circuito integrado "
         "de proteção térmica e de sobrecorrente, operando em corrente contínua de baixíssima voltagem com isolamento total.\n"
         "• Risco Biológico: Eliminado pelo uso de barreiras mecânicas descartáveis (meias de uso único) e assepsia integral com álcool 70% entre voluntários.\n"
         "• Sigilo e Confidencialidade: Proteção assegurada pela Lei Geral de Proteção de Dados (LGPD). Os voluntários serão designados apenas por códigos "
         "(VOL-01 a VOL-05), sem exposição de dados nominais.\n\n"
         "4.2. Benefícios:\n"
         "Geração de evidência científica inédita para validação de dispositivos vestíveis de saúde digital no SUS, fundamentando o modelo Random Forest "
         "para detecção precoce de complicações no pé diabético.\n\n"
         "4.3. Termo de Consentimento:\n"
         "Obtenção do TCLE em duas vias originais, assegurando pleno esclarecimento, direito de recusa sem qualquer ônus e garantia de ressarcimento/indenização."),

        ("5. CRONOGRAMA DE EXECUÇÃO FÍSICA",
         "O cronograma foi estritamente planejado para que nenhuma atividade de campo com seres humanos ocorra antes da aprovação final pelo CEP/HUPE:\n"
         "• Setembro/2026: Conclusão da qualificação de mestrado e consolidação dos testes de bancada do protótipo com Wemos Lolin D32 e LiPo 600mAh;\n"
         "• Outubro/2026: Submissão do protocolo na Plataforma Brasil (prazo limite de 05/10/2026 para reunião de 15/10/2026 do CEP/HUPE);\n"
         "• Novembro/2026: Obtenção do parecer de aprovação ética consubstanciado do CEP/HUPE e parametrização na PPC;\n"
         "• Dezembro/2026 a Janeiro/2027: Recrutamento, coleta comparativa com baropodômetro e prova de conceito com os voluntários na PPC;\n"
         "• Fevereiro/2027: Análise estatística, refinamento do modelo preditivo Random Forest e redação da dissertação final."),

        ("6. ORÇAMENTO FINANCEIRO E ISENÇÃO DE CUSTOS",
         f"O projeto tem orçamento global estimado em R$ 850,00, englobando a confecção da palmilha ({CUSTO_UNITARIO}), insumos de proteção individual, "
         f"álcool 70% e material de escritório. Todos os custos são integralmente assumidos pelo pesquisador discente ({AUTOR}), "
         f"NÃO GERANDO QUALQUER ÔNUS FINANCEIRO PARA A UERJ, PARA A PPC OU PARA O SUS."),

        ("7. REFERÊNCIAS BIBLIOGRÁFICAS (ABNT)",
         "1. BRASIL. Ministério da Saúde. Conselho Nacional de Saúde. Resolução nº 466, de 12 de dezembro de 2012. Diário Oficial da União, Brasília, 2013.\n"
         "2. BRASIL. Ministério da Saúde. Conselho Nacional de Saúde. Norma Operacional nº 001/2013. Sistema CEP/CONEP. Brasília, 2013.\n"
         "3. ARMSTRONG, D. G. et al. Diabetic foot ulcers and their recurrence. The New England Journal of Medicine, v. 376, n. 24, p. 2367-2375, 2017.\n"
         "4. BORGES, L. E. Monitoramento da temperatura dos pés de pacientes diabéticos através de termometria cutânea. Tese (Doutorado) – Universidade de São Paulo, 2023.\n"
         "5. KHANDAKAR, A. et al. A machine learning-based smart insole for early detection of diabetic foot ulcer. Sensors, v. 22, n. 19, p. 7183, 2022.\n"
         "6. KOSAJI, M. et al. Influence of in-shoe humidity on skin barrier function in diabetic neuropathy. Journal of Diabetes Science and Technology, 2025.\n"
         "7. NOUMAN, M.; RAHMAN, M. Wearable insole for continuous plantar pressure and gait assessment. IEEE Transactions on Biomedical Circuits and Systems, 2025.")
    ]

    for c_title, c_text in capitulos:
        p_ch = doc.add_paragraph()
        format_paragraph(p_ch, 8, 2, 1.15)
        r_ch = p_ch.add_run(c_title)
        r_ch.font.bold = True
        r_ch.font.color.rgb = RGBColor(20, 50, 100)

        p_ct = doc.add_paragraph()
        p_ct.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_ct, 2, 6, 1.2)
        p_ct.add_run(c_text)

    out_path = os.path.join("01_Submissao_CEP_Palmilha_PPC", "Projeto_de_Pesquisa_Integra_Palmilha_Inteligente.docx")
    doc.save(out_path)
    print(f"Salvo com sucesso: {out_path}")

if __name__ == "__main__":
    print("Iniciando geração com as novas especificações de hardware (Wemos Lolin D32 + LiPo 600mAh)...")
    create_portfolio_doc()
    create_anuencia_doc()
    create_tcle_doc()
    create_isencao_doc()
    create_instrumento_coleta_doc()
    create_projeto_integra_doc()
    print("Todos os documentos Word foram atualizados com sucesso!")
