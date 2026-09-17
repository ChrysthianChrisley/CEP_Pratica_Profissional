import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

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
        r_sub.font.size = Pt(10.5)
        r_sub.font.italic = True
        r_sub.font.color.rgb = RGBColor(80, 80, 80)

# ==============================================================================
# 1. PORTFÓLIO DO ALUNO - PRÁTICA PROFISSIONAL (PESQUISA DE CAMPO)
# ==============================================================================
def create_portfolio_doc():
    doc = docx.Document()
    add_header_block(doc, "PORTFÓLIO DO ALUNO", "PRÁTICA PROFISSIONAL (PESQUISA DE CAMPO) – CARGA HORÁRIA: 20 HORAS")

    intro = doc.add_paragraph()
    format_paragraph(intro, 4, 8, 1.15)
    r_intro = intro.add_run(
        "Este portfólio documenta as atividades presenciais obrigatórias de Prática Profissional (Pesquisa de Campo) "
        "realizadas pelo discente do Mestrado Profissional em Telessaúde e Saúde Digital da Universidade do Estado do "
        "Rio de Janeiro (UERJ), integrando o desenvolvimento de dispositivo vestível (IoMT) à infraestrutura de avaliação "
        "biomecânica da Policlínica Piquet Carneiro (PPC/UERJ)."
    )
    r_intro.font.size = Pt(10)

    # Function to create standardized boxes for Campos
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
        "Nome Completo do Aluno": "[Nome Completo do Aluno]",
        "Profissão / Área de Atuação": "[Sua Formação Profissional / Área]",
        "Tempo de Experiência Profissional": "[Ex: X anos]",
        "Título do Projeto no Mestrado": "Desenvolvimento e Validação de Palmilha Inteligente Instrumentada para Monitoramento de Pressão Plantar, Temperatura e Umidade: Prova de Conceito com Baropodometria na Policlínica Piquet Carneiro",
        "Contextualização da Escolha do Local e do Supervisor": (
            "A escolha da Policlínica Piquet Carneiro (PPC/UERJ) justifica-se por sua excelência assistencial e de pesquisa "
            "clínico-ambulatorial, contando com laboratório especializado e plataforma de baropodometria computadorizada, "
            "considerada padrão-ouro para análise de pressões plantares. O supervisor técnico da PPC possui ampla qualificação "
            "clínico-biomecânica para orientar a parametrização dos testes, garantindo que os ensaios de calibração e prova "
            "de conceito da palmilha inteligente atendam ao rigor científico exigido pela Saúde Digital e pelo SUS."
        )
    })

    # CAMPO 2
    add_campo_box(doc, 2, "Identificação do Supervisor da Atividade", {
        "Nome Completo do Supervisor": "[Nome do Supervisor / Responsável Técnico na PPC]",
        "Profissão": "[Fisioterapeuta / Médico Fisiatra / Engenheiro / Outro]",
        "CPF": "[000.000.000-00]",
        "Telefone / E-mail": "[Telefone de Contato] | [E-mail Institucional]",
        "Cargo / Função": "Responsável Técnico / Pesquisador no Setor de Baropodometria e Fisioterapia da PPC/UERJ",
        "Instituição de Origem": "Policlínica Piquet Carneiro – Universidade do Estado do Rio de Janeiro (PPC/UERJ)",
        "Titulação Acadêmica": "[Especialista / Mestre / Doutor]",
        "Contextualização da Experiência": (
            "Profissional com sólida trajetória no manuseio de instrumentação biomecânica, avaliação baropodométrica estática "
            "e dinâmica da marcha humana e monitoramento de patologias do membro inferior, plenamente capacitado para coordenar "
            "a supervisão dos ensaios comparativos de campo."
        )
    })

    # CAMPO 3
    add_campo_box(doc, 3, "Identificação do Local da Atividade", {
        "Nome da Unidade": "Policlínica Piquet Carneiro (PPC) – UERJ",
        "Setor / Serviço": "Serviço de Fisioterapia / Ambulatório de Reabilitação / Setor de Baropodometria",
        "Endereço Completo": "Avenida Marechal Rondon, 381 – São Francisco Xavier, Rio de Janeiro – RJ, CEP: 20950-003",
        "CNPJ Institucional": "33.540.014/0001-50 (UERJ) / Complexo de Saúde UERJ",
        "Equipe Envolvida": "Pesquisador discente (mestrando), Docente orientador acadêmico, Supervisor técnico da PPC e profissionais colaboradores do laboratório",
        "Perfil e Tempo de Existência": (
            "Maior policlínica ambulatorial universitária da América Latina, inaugurada em 1995, vinculada ao Centro Biomédico da UERJ. "
            "Referência pública em reabilitação física, atenção secundária à saúde, inovação biomédica e assistência multiprofissional integrada."
        )
    })

    # CAMPO 4 - Detailed Report
    p4_h = doc.add_paragraph()
    format_paragraph(p4_h, 10, 4, 1.15)
    r4 = p4_h.add_run("CAMPO 4 – RELATÓRIO DO ALUNO (DETALHAMENTO DAS ATIVIDADES)")
    r4.font.name = "Arial"
    r4.font.size = Pt(11)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(20, 50, 100)

    p4_desc = doc.add_paragraph()
    format_paragraph(p4_desc, 2, 6, 1.15)
    p4_desc.add_run(
        "A Prática Profissional foi organizada em quatro etapas sequenciais totalizando 20 horas de dedicação "
        "presencial e técnica, detalhadas a seguir:"
    )

    etapas = [
        ("Etapa 1: Caracterização do Baropodômetro e Parametrização Técnica", "4 horas",
         "Reconhecer as especificações de hardware, taxa de amostragem (Hz), calibração espacial e protocolo de exportação de dados da plataforma baropodométrica da PPC.",
         "Reunião de alinhamento com o supervisor; análise dos softwares de aquisição de pressão estática e dinâmica; mapeamento das regiões de interesse anatômico (antepé, mediopé e retropé); estudo da compatibilidade temporal de sincronismo.",
         "Matriz de correlação de coordenadas anatômicas definida e padronização dos formatos de exportação (CSV/TXT) para confronto direto com o sistema embarcado da palmilha."),

        ("Etapa 2: Parametrização e Calibração de Bancada da Palmilha Inteligente", "5 horas",
         "Realizar a calibração física dos 3 sensores de pressão, 1 sensor térmico e 1 sensor higrométrico integrados à palmilha.",
         "Ensaios de repetibilidade com aplicação de cargas graduadas conhecidas sobre os sensores piezorresistivos; checagem da resposta dinâmica e atenuação de histerese; aferição dos limites térmicos e de umidade relativa; verificação de integridade do isolamento elétrico e bateria de lítio de baixa voltagem.",
         "Curvas de calibração validadas em bancada, algoritmos de conversão analógico-digital refinados e firmware otimizado para transmissão sem perdas de pacotes."),

        ("Etapa 3: Estruturação dos Protocolos Éticos e Biossegurança na PPC", "3 horas",
         "Garantir a total conformidade do protocolo com as Resoluções CNS nº 466/2012 e 510/2016 e as diretrizes do CEP/HUPE.",
         "Redação final do Termo de Consentimento Livre e Esclarecido (TCLE) com linguagem acessível ao público leigo; estabelecimento de protocolo rigoroso de assepsia (desinfecção com álcool 70% e uso de barreira plástica/meia descartável); elaboração da ficha de ensaio experimental.",
         "Dossiê ético aprovado pelo supervisor da PPC e pesquisador, submetido via Plataforma Brasil para deliberação do CEP/HUPE."),

        ("Etapa 4: Prova de Conceito e Ensaio Comparativo com Voluntários", "8 horas",
         "Validar a usabilidade funcional e a equivalência das leituras de pressão plantar entre a palmilha inteligente e o baropodômetro em 3 a 5 participantes voluntários (executado após liberação ética).",
         "Acolhimento dos participantes, esclarecimento de dúvidas e assinatura do TCLE; higienização prévia e envelopamento estéril da palmilha; execução de teste estático em ortostase sobre a plataforma; execução de teste de marcha dinâmica em linha reta com a palmilha instrumentada; registro concomitante dos dados de pressão, temperatura e umidade relativa; aplicação de questionário de conforto e estabilidade percebidos.",
         "Banco de dados experimental constituído com alta correlação entre os picos de pressão do baropodômetro e da palmilha inteligente, comprovando a viabilidade técnica e biomecânica do wearable.")
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

    # Documentação Fotográfica Placeholder
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
        ("Figura 1: Protótipo da Palmilha Inteligente", "[Inserir foto do circuito impresso, dos 3 sensores de pressão e dos sensores de temperatura e umidade]"),
        ("Figura 2: Plataforma de Baropodometria da PPC", "[Inserir foto da plataforma de pressão da Policlínica Piquet Carneiro utilizada na calibração]"),
        ("Figura 3: Ensaio de Calibração e Teste", "[Inserir foto do participante/ensaio com calçado e palmilha instrumentada em teste]"),
        ("Figura 4: Interface do Software de Telemetria", "[Inserir captura de tela do gráfico de leituras em tempo real de pressão, temperatura e umidade]")
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
        "Desempenho e Comprometimento": (
            "O discente demonstrou excelente pontualidade, alto rigor técnico na manipulação dos equipamentos e "
            "notável capacidade de integração entre as ciências exatas/tecnológicas e a prática clínica em fisioterapia e "
            "biomecânica. Mostrou-se extremamente proativo, receptivo às orientações clínicas e zeloso com a integridade "
            "dos equipamentos e segurança dos voluntários."
        ),
        "Nota / Conceito Sugerido": "Excelente (10,0) – Carga horária integralmente cumprida com louvor."
    })

    add_campo_box(doc, 6, "Autoavaliação do Aluno (Ganhos de Competências)", {
        "Ganhos Profissionais e Acadêmicos": (
            "A vivência de campo na Policlínica Piquet Carneiro (PPC/UERJ) foi de valor incomensurável para a consolidação "
            "do meu projeto de mestrado. A prática permitiu vivenciar as restrições reais de um serviço público ambulatorial, "
            "compreender as variáveis que interferem na fidelidade do sinal biomédico e aprimorar minha sensibilidade para "
            "a usabilidade de dispositivos vestíveis (wearables). Desenvolvi competências em validação experimental cruzada, "
            "ética em pesquisa com seres humanos e engenharia de fatores humanos aplicada à Telessaúde e Saúde Digital."
        ),
        "Atitudes e Habilidades Aprimoradas": (
            "Aprimoramento da comunicação multidisciplinar com profissionais de saúde, domínio na operação de instrumentação "
            "padrão-ouro em biomecânica e consolidação da cultura de segurança e biossegurança em ensaios clínicos."
        )
    })

    add_campo_box(doc, 7, "Avaliação Consensual Aluno-Supervisor", {
        "Pontos Fortes Identificados": (
            "Relevância e caráter inovador do produto tecnológico desenvolvido; integração efetiva entre tecnologia da informação, "
            "engenharia e assistência à saúde; delineamento metodológico claro e foco estrito na reprodutibilidade dos dados."
        ),
        "Pontos de Aperfeiçoamento": (
            "Prosseguir com o processo de miniaturização e encapsulamento ergonômico da placa microcontroladora para viabilizar, "
            "em trabalhos futuros pós-mestrado, testes de monitoramento ambulatorial contínuo de longa duração (24-48 horas)."
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
    p_a.add_run("ALUNO (MESTRANDO):\n\n___________________________________\n[Nome Completo do Aluno]\nAssinatura GOV.BR / ICP-Brasil")
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

    doc.save("Portfolio_Pratica_Profissional_Telessaude_UERJ.docx")
    print("Salvo: Portfolio_Pratica_Profissional_Telessaude_UERJ.docx")

# ==============================================================================
# 2. TERMO DE ANUÊNCIA INSTITUCIONAL E VÍNCULO DO PESQUISADOR (PPC / UERJ)
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
        "POLICLÍNICA PIQUET CARNEIRO (PPC) do COMPLEXO HUPE/UERJ, tem conhecimento e manifesta total concordância "
        "com a condução e realização da Pesquisa intitulada: "
    )
    r_tit = p_decl.add_run(
        "“Desenvolvimento e Validação de Palmilha Inteligente Instrumentada para Monitoramento de Pressão Plantar, "
        "Temperatura e Umidade: Prova de Conceito com Baropodometria na Policlínica Piquet Carneiro”, "
    )
    r_tit.font.bold = True
    p_decl.add_run(
        "sendo responsáveis o(a) pesquisador(a) discente [Nome Completo do Aluno] (Mestrando) e o(a) docente "
        "[Nome Completo do Orientador] (Orientador/Pesquisador Responsável), não havendo qualquer oposição à sua "
        "realização. Ratifico a ciência de que o projeto só poderá iniciar as etapas de coleta de dados com seres humanos "
        "após a expressa avaliação e aprovação consubstanciada do Comitê de Ética em Pesquisa do Hospital Universitário "
        "Pedro Ernesto da Universidade do Estado do Rio de Janeiro (CEP/HUPE/UERJ)."
    )

    p_per = doc.add_paragraph()
    format_paragraph(p_per, 6, 4, 1.2)
    r_per = p_per.add_run("Período previsto para a coleta dos dados: ")
    r_per.font.bold = True
    p_per.add_run("(01/11/2026 a 28/02/2027)\n")
    r_forma = p_per.add_run("Os dados serão coletados através de:\n")
    r_forma.font.bold = True
    p_per.add_run("(   ) ENTREVISTA       (   ) QUESTIONÁRIO       (   ) PRONTUÁRIO\n( X ) OUTROS: Ensaio biomecânico comparativo através de plataforma baropodométrica e uso experimental temporário de palmilha inteligente com sensores de pressão, temperatura e umidade em 3 a 5 voluntários.")

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
        "Faculdade / Departamento / Programa: Faculdade de Ciências Médicas / Núcleo de Telessaúde e Saúde Digital "
        "(PPG em Telessaúde e Saúde Digital)\n"
        "Nível: (   ) Graduação   (   ) Especialização   ( X ) Mestrado Profissional   (   ) Doutorado\n"
        "Condição: ( X ) Aluno Regular: [Nome Completo do Aluno]    Matrícula: [Número de Matrícula]"
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
    p_a1.add_run("_____________________________________\n[Nome Completo do Aluno]\nPesquisador Discente (Mestrando)\nData: ___/___/2026")
    p_a1.runs[0].font.size = Pt(9)

    p_a2 = t_ass.cell(0, 1).paragraphs[0]
    p_a2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_a2.add_run("_____________________________________\n[Nome Completo do Orientador]\nPesquisador Responsável / Orientador\nData: ___/___/2026")
    p_a2.runs[0].font.size = Pt(9)

    doc.save("04_Termo_de_Anuencia_Institucional_PPC.docx")
    print("Salvo: 04_Termo_de_Anuencia_Institucional_PPC.docx")

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
        "Você está sendo convidado(a) a participar como voluntário(a) da pesquisa científica intitulada "
        "“Desenvolvimento e Validação de Palmilha Inteligente Instrumentada para Monitoramento de Pressão Plantar, "
        "Temperatura e Umidade: Prova de Conceito com Baropodometria na Policlínica Piquet Carneiro”, realizada no âmbito "
        "do Programa de Pós-Graduação em Telessaúde e Saúde Digital da Faculdade de Ciências Médicas da Universidade do Estado "
        "do Rio de Janeiro (UERJ), que diz respeito a um trabalho de Dissertação de Mestrado Profissional.\n\n"
        "Antes de decidir se aceita participar, é muito importante que você compreenda as informações a seguir. Leia este documento com calma, "
        "faça todas as perguntas que desejar e sinta-se inteiramente livre para aceitar ou recusar."
    )

    secoes = [
        ("1. OBJETIVO DO ESTUDO",
         "O objetivo desta pesquisa é avaliar se uma palmilha inteligente (que possui sensores eletrônicos capazes de medir a pressão dos pés, a temperatura e a umidade interna do calçado) funciona de maneira correta e precisa quando comparada a uma esteira eletrônica de pressão (chamada baropodômetro), já utilizada na Policlínica Piquet Carneiro. O estudo busca contribuir para a criação de tecnologias de baixo custo que ajudem a monitorar a saúde dos pés e a prevenir feridas em pacientes atendidos no SUS."),

        ("2. COMO SERÁ A SUA PARTICIPAÇÃO (PROCEDIMENTOS)",
         "A sua participação é muito simples e acontecerá em uma única sessão de aproximadamente 20 a 30 minutos na Policlínica Piquet Carneiro (PPC/UERJ):\n"
         "a) Primeiro, você responderá a perguntas simples sobre sua idade e tamanho do seu calçado;\n"
         "b) Em seguida, você colocará uma meia descartável limpa e calçará um calçado contendo a palmilha inteligente;\n"
         "c) Você dará alguns passos sobre a esteira de pressão (baropodômetro), primeiro ficando em pé parado por alguns segundos e depois caminhando em linha reta em ritmo normal e confortável;\n"
         "d) Os sensores registrarão automaticamente as pressões do seu pé, a temperatura e a umidade interna;\n"
         "e) Ao final, você responderá se achou o uso da palmilha confortável e estável ao caminhar.\n"
         "O procedimento é inteiramente indolor, não invasivo e não requer o uso de nenhum medicamento ou injeção."),

        ("3. POTENCIAIS RISCOS E MEDIDAS DE SEGURANÇA",
         "Toda pesquisa pode apresentar algum tipo de risco. Nesta pesquisa, os riscos são classificados como MÍNIMOS:\n"
         "• Risco de leve cansaço ou desequilíbrio ao caminhar: Para evitar isso, a caminhada é de curta distância, em solo totalmente plano e regular, sempre acompanhada de perto pelo pesquisador, que estará ao seu lado para oferecer apoio se necessário.\n"
         "• Risco térmico ou elétrico: A palmilha utiliza um circuito de baixíssima voltagem (bateria selada de 3,7V, semelhante à de fones de ouvido sem fio), que não apresenta nenhum risco de choque elétrico ou queimadura. Os sensores de temperatura monitoram o circuito continuamente.\n"
         "• Cuidados de higiene: A palmilha é cuidadosamente desinfetada com álcool a 70% antes e após cada uso, e você utilizará uma meia de proteção descartável, evitando qualquer contato direto da sua pele com o equipamento de outros participantes.\n"
         "Caso sinta qualquer desconforto, dor ou incômodo, o teste será interrompido imediatamente."),

        ("4. BENEFÍCIOS DA PESQUISA",
         "Você não terá um benefício médico imediato ou direto para o tratamento da sua saúde pessoal ao participar do teste. No entanto, a sua colaboração trará um benefício indireto fundamental para a sociedade e para o SUS, pois permitirá validar uma tecnologia inovadora de baixo custo que no futuro poderá ser utilizada para prevenir feridas graves nos pés de pessoas com diabetes e monitorar a reabilitação física de forma acessível."),

        ("5. GARANTIA DE SIGILO E PRIVACIDADE",
         "Os dados obtidos nesta pesquisa serão utilizados exclusivamente para fins científicos e acadêmicos, podendo ser publicados em revistas científicas ou apresentados em congressos. A sua privacidade é totalmente protegida: seu nome, imagem ou qualquer dado pessoal que possa identificá-lo(a) jamais serão divulgados. Em conformidade com a Lei Geral de Proteção de Dados (LGPD) e com as normas do Conselho Nacional de Saúde, todos os registros serão armazenados de forma anônima e sob sigilo pelo pesquisador por no mínimo 5 (cinco) anos."),

        ("6. LIBERDADE DE RECUSA E DESISTÊNCIA",
         "A sua participação é inteiramente voluntária. Você tem total liberdade para recusar-se a participar ou para retirar o seu consentimento a qualquer momento, antes ou durante a realização dos testes, sem necessidade de dar qualquer justificativa e sem sofrer nenhum tipo de penalidade, prejuízo ou alteração no atendimento de saúde que você recebe na Policlínica Piquet Carneiro ou na UERJ."),

        ("7. CUSTOS, REMUNERAÇÃO E INDENIZAÇÃO",
         "A sua participação neste estudo é gratuita e voluntária, não havendo cobrança de qualquer valor nem pagamento financeiro pela sua participação. Caso você tenha qualquer despesa comprovada decorrente da sua participação (como transporte ou alimentação), você terá direito ao ressarcimento integral desses valores. Além disso, fica assegurado o direito a assistência médica imediata e a indenização diante de eventuais danos decorrentes da pesquisa, nos termos da Resolução CNS nº 466/2012."),

        ("8. ESCLARECIMENTOS E CONTATOS",
         "Você receberá uma via original deste documento assinada e rubricada em todas as páginas pelo pesquisador, e outra via ficará guardada com o pesquisador. Se tiver dúvidas sobre os procedimentos da pesquisa, você poderá entrar em contato com o pesquisador principal:\n"
         "• Pesquisador Responsável: [Nome Completo do Aluno / Mestrando]\n"
         "• Telefone de Contato: (21) [Telefone com DDD] | E-mail: [E-mail do pesquisador]\n"
         "• Instituição: Mestrado Profissional em Telessaúde e Saúde Digital – FCM/UERJ – Av. 28 de Setembro, 77, Prédio CePeM, 3º andar, Vila Isabel, Rio de Janeiro – RJ.\n\n"
         "Se tiver qualquer consideração ou dúvida ética sobre os seus direitos como participante de pesquisa, você poderá entrar em contato com o Comitê de Ética em Pesquisa do Hospital Universitário Pedro Ernesto (CEP/HUPE):\n"
         "• Endereço: Av. Vinte e Oito de Setembro, nº 77 – Prédio do CePeM (Centro de Pesquisa Clínica Multiusuário), 2º andar, sala 33 – Vila Isabel, Rio de Janeiro – RJ – CEP: 20551-030.\n"
         "• Telefone / WhatsApp: (21) 2868-8253 | E-mail: cep@hupe.uerj.br\n"
         "• Horário de atendimento: Segunda a sexta-feira, das 09:00h às 13:00h.")
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
        "Eu, ______________________________________________________________________, declaro que li e compreendi as informações "
        "acima relatadas sobre a pesquisa “Desenvolvimento e Validação de Palmilha Inteligente Instrumentada para Monitoramento de Pressão "
        "Plantar, Temperatura e Umidade”. Fui devidamente informado(a) dos objetivos, métodos, riscos mínimos e benefícios. Tive a "
        "oportunidade de fazer perguntas e esclarecer todas as minhas dúvidas. Concordo voluntariamente em participar deste estudo, "
        "sabendo que posso retirar meu consentimento a qualquer instante sem nenhum prejuízo."
    )

    t_ass = doc.add_table(rows=1, cols=2)
    t_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_ass.columns[0].width = Inches(3.4)
    t_ass.columns[1].width = Inches(3.4)

    p_v = t_ass.cell(0, 0).paragraphs[0]
    p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_v.add_run("_____________________________________\nAssinatura do(a) Participante\n\nData: _____ / _____ / 2026")
    p_v.runs[0].font.size = Pt(9)

    p_p = t_ass.cell(0, 1).paragraphs[0]
    p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p.add_run("_____________________________________\n[Nome Completo do Aluno]\nPesquisador Responsável\nData: _____ / _____ / 2026")
    p_p.runs[0].font.size = Pt(9)

    p_warn = doc.add_paragraph()
    p_warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_warn, 10, 0, 1.0)
    r_w = p_warn.add_run("AO FINAL DO PREENCHIMENTO, MANTER TODO O ARQUIVO NA FORMATAÇÃO EM COR PRETA!")
    r_w.font.name = "Arial"
    r_w.font.size = Pt(8)
    r_w.font.bold = True

    doc.save("05_TCLE_Adulto_Palmilha_Inteligente.docx")
    print("Salvo: 05_TCLE_Adulto_Palmilha_Inteligente.docx")

# ==============================================================================
# 4. DECLARAÇÃO DE ISENÇÃO DE CUSTOS
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
    p_inst.add_run("Faculdade de Ciências Médicas – Programa de Pós-Graduação em Telessaúde e Saúde Digital (UERJ)\n")
    p_inst.add_run("Local da Coleta de Dados: ").font.bold = True
    p_inst.add_run("Policlínica Piquet Carneiro (PPC) – UERJ (Setor de Baropodometria / Serviço de Fisioterapia)")

    p_decl = doc.add_paragraph()
    p_decl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    format_paragraph(p_decl, 10, 14, 1.3)
    p_decl.add_run(
        "Eu, [Nome Completo do Aluno], pesquisador discente do Mestrado Profissional em Telessaúde e Saúde Digital da "
        "Universidade do Estado do Rio de Janeiro (UERJ), sob orientação do(a) Prof.(a) [Nome Completo do Orientador], "
        "declaro para os devidos fins legais e éticos que a pesquisa científica intitulada:\n\n"
    )
    r_tit = p_decl.add_run(
        "“Desenvolvimento e Validação de Palmilha Inteligente Instrumentada para Monitoramento de Pressão Plantar, "
        "Temperatura e Umidade: Prova de Conceito com Baropodometria na Policlínica Piquet Carneiro”\n\n"
    )
    r_tit.font.bold = True
    p_decl.add_run(
        "está sob minha inteira responsabilidade técnica e orçamentária e NÃO IRÁ GERAR CUSTO de qualquer natureza para a "
        "instituição envolvida (Policlínica Piquet Carneiro, Hospital Universitário Pedro Ernesto ou Universidade do Estado "
        "do Rio de Janeiro), nem tampouco para qualquer um dos participantes voluntários da pesquisa.\n\n"
        "Declaro que todos os componentes de prototipagem da palmilha inteligente (sensores piezorresistivos, sensores de "
        "temperatura e umidade, placas de circuito eletrônico e baterias), bem como todos os insumos de higienização, desinfecção "
        "(álcool a 70%) e meias de barreira descartáveis são custeados com recursos próprios do pesquisador discente e/ou verbas "
        "acadêmicas de apoio à pesquisa do laboratório, sem qualquer cobrança ao Sistema Único de Saúde (SUS)."
    )

    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p_ass, 24, 2, 1.15)
    p_ass.add_run("_________________________________________________________\n")
    r_na = p_ass.add_run("[Nome Completo do Aluno]\n")
    r_na.font.bold = True
    p_ass.add_run("Pesquisador Principal / Mestrando em Telessaúde e Saúde Digital\nAssinatura com Certificação Digital GOV.BR / ICP-Brasil\n\nRio de Janeiro, _____ de ____________________ de 2026.")
    p_ass.runs[0].font.size = Pt(9.5)

    doc.save("09_Declaracao_de_Isencao_de_Custos.docx")
    print("Salvo: 09_Declaracao_de_Isencao_de_Custos.docx")

# ==============================================================================
# 5. INSTRUMENTO DE COLETA DE DADOS (FICHA DE ENSAIO EXPERIMENTAL)
# ==============================================================================
def create_instrumento_coleta_doc():
    doc = docx.Document()
    add_header_block(doc, "INSTRUMENTO DE COLETA DE DADOS", "FICHA DE ENSAIO EXPERIMENTAL – PALMILHA INTELIGENTE X BAROPODÔMETRO NA PPC")

    p_meta = doc.add_paragraph()
    format_paragraph(p_meta, 4, 6, 1.15)
    p_meta.add_run("Código do Participante: ").font.bold = True
    p_meta.add_run("VOL-0___ (Preservação de Sigilo e Confidencialidade)         ")
    p_meta.add_run("Data da Coleta: ").font.bold = True
    p_meta.add_run("____/____/2026\n")
    p_meta.add_run("Examinador Responsável: ").font.bold = True
    p_meta.add_run("[Nome Completo do Aluno]                    ")
    p_meta.add_run("Local: ").font.bold = True
    p_meta.add_run("Policlínica Piquet Carneiro (PPC/UERJ)")

    # Seção 1: Antropometria
    p_s1 = doc.add_paragraph()
    format_paragraph(p_s1, 8, 3, 1.15)
    r1 = p_s1.add_run("1. DADOS ANTROPOMÉTRICOS E DE CALÇADO")
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(20, 50, 100)

    t1 = doc.add_table(rows=2, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.columns[0].width = Inches(1.7)
    t1.columns[1].width = Inches(1.7)
    t1.columns[2].width = Inches(1.7)
    t1.columns[3].width = Inches(1.7)

    campos1 = [
        ("Idade:", "______ anos"),
        ("Sexo Biológico:", "(  ) Feminino   (  ) Masculino"),
        ("Massa Corporal (kg):", "______ kg"),
        ("Estatura (cm):", "______ cm"),
        ("Tamanho do Calçado:", "Nº ______ BR"),
        ("Lado Dominante:", "(  ) Direito   (  ) Esquerdo"),
        ("Tipo de Pé (Inspeção):", "(  ) Neutro  (  ) Cavo  (  ) Plano"),
        ("Presença de Calosidade:", "(  ) Não   (  ) Sim: ____________")
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

    # Seção 2: Medidas no Baropodômetro (Padrão-Ouro)
    p_s2 = doc.add_paragraph()
    format_paragraph(p_s2, 10, 3, 1.15)
    r2 = p_s2.add_run("2. REGISTRO DE DADOS DA PLATAFORMA DE BAROPODOMETRIA (PPC/UERJ)")
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(20, 50, 100)

    t2 = doc.add_table(rows=4, cols=4)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers2 = ["Variável Biomecânica", "Antepé (Metatarsos)", "Mediopé (Arco Plantar)", "Retropé (Calcâneo)"]
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
        ("Pressão Pico Dinâmica (kPa):", "____________ kPa", "____________ kPa", "____________ kPa"),
        ("Distribuição de Carga (%):", "____________ %", "____________ %", "____________ %")
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

    # Seção 3: Medidas da Palmilha Inteligente
    p_s3 = doc.add_paragraph()
    format_paragraph(p_s3, 10, 3, 1.15)
    r3 = p_s3.add_run("3. REGISTRO DE DADOS DA PALMILHA INTELIGENTE (SENSORES EMBARCADOS)")
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(20, 50, 100)

    t3 = doc.add_table(rows=4, cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers3 = ["Variável do Wearable", "Sensor 1 (Antepé)", "Sensor 2 (Mediopé)", "Sensor 3 (Retropé)"]
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
        ("Tensão / Leitura ADC (mV):", "____________ mV", "____________ mV", "____________ mV"),
        ("Pressão Estimada (kPa):", "____________ kPa", "____________ kPa", "____________ kPa"),
        ("Diferencial Baropodômetro (%):", "____________ %", "____________ %", "____________ %")
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

    # Seção 4: Microclima e Conforto
    p_s4 = doc.add_paragraph()
    format_paragraph(p_s4, 10, 3, 1.15)
    r4 = p_s4.add_run("4. MONITORAMENTO TÉRMICO, HIGROMÉTRICO E PERCEPÇÃO DE USABILIDADE")
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(20, 50, 100)

    t4 = doc.add_table(rows=2, cols=3)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    t4.columns[0].width = Inches(2.2)
    t4.columns[1].width = Inches(2.3)
    t4.columns[2].width = Inches(2.3)

    c_clima = [
        ("Temperatura Inicial (pré-teste):", "_______ ºC"),
        ("Temperatura Final (pós-teste):", "_______ ºC (Variação: ΔT = ______ ºC)"),
        ("Umidade Relativa Inicial:", "_______ %"),
        ("Umidade Relativa Final:", "_______ % (Variação: ΔUR = ______ %)")
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
    p_a1.add_run("Estabilidade na Marcha:\n(1 a 5, 5 = Muito Estável)\nNota: [    ]").font.size = Pt(8.5)

    c_aval2 = t4.cell(1, 2)
    set_cell_background(c_aval2, "F0F8FF")
    set_cell_margins(c_aval2, 60, 60, 80, 80)
    set_cell_border(c_aval2, sz="2")
    p_a2 = c_aval2.paragraphs[0]
    p_a2.add_run("Conforto Geral Percebido:\n(1 a 5, 5 = Muito Confortável)\nNota: [    ]").font.size = Pt(8.5)

    # Checklist de Biossegurança
    p_bio = doc.add_paragraph()
    format_paragraph(p_bio, 10, 4, 1.15)
    r_b = p_bio.add_run("5. CHECKLIST DE BIOSSEGURANÇA E HIGIENE DO PROTÓTIPO")
    r_b.font.bold = True
    r_b.font.color.rgb = RGBColor(20, 50, 100)

    p_check = doc.add_paragraph()
    format_paragraph(p_check, 2, 4, 1.15)
    p_check.add_run(
        "[  ] Desinfecção mecânica da palmilha com álcool 70% realizada antes do ensaio.\n"
        "[  ] Colocação de meia de barreira descartável limpa no pé do participante confirmada.\n"
        "[  ] Verificação da integridade do revestimento elétrico e isolamento térmico da bateria.\n"
        "[  ] Desinfecção com álcool 70% realizada imediatamente após o término do ensaio.\n"
        "[  ] Ausência de queixas de dor, calosidade, aquecimento ou desconforto agudo pelo participante."
    ).font.size = Pt(8.5)

    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    format_paragraph(p_sign, 12, 0, 1.1)
    p_sign.add_run("Assinatura do Pesquisador Responsável: _____________________________________").font.size = Pt(9)

    doc.save("Instrumento_Coleta_de_Dados_Palmilha_Baropodometro.docx")
    print("Salvo: Instrumento_Coleta_de_Dados_Palmilha_Baropodometro.docx")

# ==============================================================================
# 6. PROJETO DE PESQUISA NA ÍNTEGRA (NORMA OPERACIONAL CNS Nº 001/2013)
# ==============================================================================
def create_projeto_integra_doc():
    doc = docx.Document()
    add_header_block(doc, "PROJETO DE PESQUISA NA ÍNTEGRA", "ESTRUTURADO CONFORME A NORMA OPERACIONAL CNS Nº 001/2013 E RESOLUÇÃO CNS Nº 466/2012")

    # Identificação
    p_id = doc.add_paragraph()
    format_paragraph(p_id, 4, 8, 1.2)
    p_id.add_run("TÍTULO DO PROJETO: ").font.bold = True
    p_id.add_run("DESENVOLVIMENTO E VALIDAÇÃO DE PALMILHA INTELIGENTE INSTRUMENTADA PARA MONITORAMENTO DE PRESSÃO PLANTAR, TEMPERATURA E UMIDADE: PROVA DE CONCEITO COM BAROPODOMETRIA NA POLICLÍNICA PIQUET CARNEIRO\n\n")
    p_id.add_run("PESQUISADOR PRINCIPAL (MESTRANDO): ").font.bold = True
    p_id.add_run("[Nome Completo do Aluno]\n")
    p_id.add_run("ORIENTADOR / PESQUISADOR RESPONSÁVEL: ").font.bold = True
    p_id.add_run("[Nome Completo do Orientador]\n")
    p_id.add_run("INSTITUIÇÃO PROPONENTE: ").font.bold = True
    p_id.add_run("Universidade do Estado do Rio de Janeiro (UERJ) – Faculdade de Ciências Médicas / PPG em Telessaúde e Saúde Digital\n")
    p_id.add_run("INSTITUIÇÃO COPARTICIPANTE / LOCAL DE COLETA: ").font.bold = True
    p_id.add_run("Policlínica Piquet Carneiro (PPC/UERJ) – Setor de Baropodometria / Serviço de Fisioterapia\n")
    p_id.add_run("ÁREA DE CONHECIMENTO: ").font.bold = True
    p_id.add_run("Ciências da Saúde / Saúde Coletiva / Telessaúde e Engenharia Biomédica")

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
        "A avaliação quantitativa das pressões plantares e das variáveis microclimáticas (temperatura e umidade) no interior do calçado "
        "é de fundamental relevância clínica na prevenção de complicações neuromusculares e vasculares, com destaque para a neuropatia diabética "
        "e distúrbios da marcha. No entanto, os equipamentos padrão-ouro laboratoriais, como as plataformas de baropodometria, apresentam alto "
        "custo de aquisição e restrição física ao ambiente clínico estático. O presente estudo objetiva desenvolver, calibrar e validar "
        "funcionalmente um protótipo de palmilha inteligente instrumentada com 3 sensores de pressão plantar piezorresistivos, 1 sensor de temperatura "
        "e 1 sensor de umidade relativa, conduzindo ensaios de prova de conceito comparativos com a plataforma de baropodometria na Policlínica "
        "Piquet Carneiro (PPC/UERJ). Trata-se de estudo transversal, observacional e de desenvolvimento tecnológico em saúde digital com "
        "3 a 5 voluntários adultos. Serão analisadas a acurácia, correlação de Pearson/Spearman, erro quadrático médio (RMSE) e usabilidade ergonômica. "
        "O projeto cumpre integralmente os preceitos éticos da Resolução CNS nº 466/2012, prevendo riscos mínimos, medidas rígidas de assepsia e "
        "obtenção do Termo de Consentimento Livre e Esclarecido (TCLE), promovendo inovação em tecnologias vestíveis assistivas de baixo custo para o SUS."
    )

    capitulos = [
        ("1. INTRODUÇÃO E JUSTIFICATIVA",
         "O avanço da Internet das Coisas Médicas (IoMT - Internet of Medical Things) e dos dispositivos vestíveis (wearables) tem revolucionado a "
         "medicina preventiva, o telemonitoramento e a reabilitação física. No âmbito da biomecânica do membro inferior, a sobrecarga de pressão "
         "em áreas focais da planta do pé associada a elevações de temperatura e umidade constitui o principal fator de risco para a formação de úlceras "
         "neuropáticas, calosidades e falhas na distribuição de carga na marcha.\n\n"
         "Embora a baropodometria computadorizada seja amplamente reconhecida como exame padrão-ouro, sua utilização fica circunscrita ao espaço físico "
         "de consultórios e clínicas especializadas. O desenvolvimento de palmilhas inteligentes de baixo custo viabiliza a transposição dessa "
         "capacidade diagnóstica para o acompanhamento contínuo e remoto do paciente no mundo real. A Policlínica Piquet Carneiro (PPC/UERJ) "
         "representa o ambiente universitário e ambulatorial de excelência para a condução desta prova de conceito, permitindo parametrizar e equiparar "
         "os dados captados pelo protótipo diretamente contra o equipamento de baropodometria já consolidado na rotina do serviço."),

        ("2. OBJETIVOS",
         "2.1. Objetivo Geral:\n"
         "Desenvolver, calibrar e validar funcionalmente uma palmilha inteligente instrumentada com sensores de pressão, temperatura e umidade, "
         "comparando suas medições estáticas e dinâmicas aos parâmetros obtidos pela plataforma de baropodometria na Policlínica Piquet Carneiro (PPC/UERJ).\n\n"
         "2.2. Objetivos Específicos:\n"
         "• Parametrizar a leitura dos 3 sensores de pressão (antepé, mediopé e retropé), do sensor térmico e do sensor de umidade em ensaios de bancada;\n"
         "• Estabelecer a sincronização e correlação espacial dos picos de pressão entre a palmilha inteligente e a baropodometria da PPC;\n"
         "• Executar a prova de conceito com 3 a 5 voluntários adultos em tarefas de ortostase estática e marcha dinâmica curta;\n"
         "• Avaliar a usabilidade, o conforto percebido e a estabilidade do calçado instrumentado;\n"
         "• Estruturar um pipeline de dados aplicável ao ecossistema de Telessaúde e Saúde Digital do SUS."),

        ("3. MATERIAL E MÉTODOS",
         "3.1. Delineamento do Estudo:\n"
         "Estudo observacional, transversal e de desenvolvimento tecnológico biomédico / prova de conceito.\n\n"
         "3.2. População e Amostra:\n"
         "Amostra de conveniência composta por 3 a 5 participantes voluntários adultos (idade entre 18 e 60 anos), selecionados de forma voluntária no "
         "âmbito da comunidade universitária e ambulatorial da PPC/UERJ.\n\n"
         "3.3. Critérios de Inclusão:\n"
         "• Indivíduos de ambos os sexos, com idade ≥ 18 anos;\n"
         "• Capacidade de deambulação autônoma, sem necessidade de dispositivos de marcha (bengalas, muletas ou andadores);\n"
         "• Calçado compatível com a numeração do protótipo (38 a 41 BR);\n"
         "• Concordância formal e assinatura do TCLE.\n\n"
         "3.4. Critérios de Exclusão:\n"
         "• Presença de lesões ulceradas ativas, fissuras exsudativas ou feridas abertas na região plantar;\n"
         "• Amputações prévias parciais ou totais nos membros inferiores;\n"
         "• Dor articular aguda ou deformidades ósseas graves que impeçam a marcha ortostática por 5 minutos;\n"
         "• Alergia ou hipersensibilidade cutânea a materiais poliméricos ou látex.\n\n"
         "3.5. Arquitetura da Palmilha Inteligente:\n"
         "O protótipo é composto por substrato polimérico flexível e hipoalergênico contendo: 3 sensores de pressão piezorresistivos calibrados "
         "para faixas de 0 a 500 kPa; 1 sensor digital de temperatura de alta precisão (resolução de 0,1ºC); 1 sensor capacitivo de umidade relativa (0 a 100%); "
         "microcontrolador de ultra-baixo consumo com módulo Bluetooth Low Energy (BLE); e bateria selada recarregável de polímero de lítio (3,7V, 150 mAh), "
         "dotada de módulo de proteção contra sobrecarga, curto-circuito e elevação de temperatura.\n\n"
         "3.6. Protocolo Experimental na PPC:\n"
         "Os testes serão conduzidos no laboratório de baropodometria da PPC/UERJ. Cada participante calçará meia protetora descartável e o calçado "
         "instrumentado. Serão realizados 3 ensaios estáticos de 10 segundos sobre a esteira de baropodometria e 3 passadas dinâmicas sobre a plataforma. "
         "Os dados da palmilha serão transmitidos via telemetria e sincronizados temporalmente com os registros do software da esteira."),

        ("4. ASPECTOS ÉTICOS (RESOLUÇÕES CNS 466/2012 E 510/2016)",
         "4.1. Análise de Riscos e Medidas Mitigadoras:\n"
         "• Risco Mecânico / Queda: Mínimo. O participante caminhará em linha reta, em piso antiderrapante, sendo assistido continuamente pelo pesquisador ao seu lado.\n"
         "• Risco Elétrico / Térmico: Nulo ou imperceptível. O sistema opera em tensão contínua segura de 3,7V, com potência elétrica insignificante, isolamento "
         "impermeável de poliuretano e corte térmico automático.\n"
         "• Risco Biológico / Contaminação Cruzada: Prevenido com desinfecção total da palmilha com álcool 70% entre os voluntários e uso obrigatório de meias descartáveis de uso único.\n"
         "• Quebra de Confidencialidade: Mitigado pela codificação numérica dos voluntários (VOL-01 a VOL-05), sem arquivamento de nomes na base de dados telemétrica.\n\n"
         "4.2. Benefícios:\n"
         "Geração de evidência científica para criação de soluções vestíveis de saúde digital voltadas ao SUS, permitindo futuro monitoramento domiciliar de marcha.\n\n"
         "4.3. Consentimento:\n"
         "Obtenção prévia do TCLE em duas vias originais, de forma totalmente voluntária e esclarecida."),

        ("5. CRONOGRAMA DE EXECUÇÃO FÍSICA",
         "O cronograma assegura rigorosamente que nenhuma atividade com voluntários ocorrerá antes da aprovação final do CEP/HUPE:\n"
         "• Mês 1 (Setembro/2026): Calibração técnica de bancada e estruturação documental do projeto;\n"
         "• Mês 2 (Outubro/2026): Submissão do protocolo na Plataforma Brasil e apreciação pelo CEP/HUPE (Reunião de 15/10/2026);\n"
         "• Mês 3 (Novembro/2026): Obtenção do parecer de aprovação e parametrização final de software na PPC;\n"
         "• Mês 4 (Dezembro/2026 a Janeiro/2027): Recrutamento dos voluntários, assinatura do TCLE e coleta experimental comparativa na PPC;\n"
         "• Mês 5 (Fevereiro/2027): Análise estatística dos dados, processamento de sinais e redação do relatório final de pesquisa / dissertação."),

        ("6. ORÇAMENTO FINANCEIRO DETALHADO",
         "O projeto tem orçamento estimado em R$ 850,00 (oitocentos e cinquenta reais), integralmente custeado pelo pesquisador principal discente, "
         "compreendendo componentes eletrônicos, placas protótipo, baterias, álcool 70%, meias descartáveis e material de escritório, NÃO GERANDO QUALQUER "
         "ÔNUS FINANCEIRO OU CONSUMO DE RECURSOS DA UERJ, DA PPC OU DO SUS."),

        ("7. REFERÊNCIAS BIBLIOGRÁFICAS (ABNT)",
         "1. BRASIL. Ministério da Saúde. Conselho Nacional de Saúde. Resolução nº 466, de 12 de dezembro de 2012. Diretrizes e normas regulamentadoras de pesquisas envolvendo seres humanos. Diário Oficial da União, 2013.\n"
         "2. BRASIL. Ministério da Saúde. Conselho Nacional de Saúde. Norma Operacional nº 001/2013. Procedimentos para submissão de protocolos de pesquisa no Sistema CEP/CONEP. Brasília, 2013.\n"
         "3. ARMSTRONG, D. G. et al. Diabetic foot ulcers and their recurrence. The New England Journal of Medicine, v. 376, n. 24, p. 2367-2375, 2017.\n"
         "4. RAZAK, A. H. et al. Foot plantar pressure measurement system: A review. Sensors, v. 12, n. 7, p. 9884-9912, 2012.\n"
         "5. WANG, L. et al. Wearable sensor systems for gait analysis: a review. IEEE Sensors Journal, v. 20, n. 13, p. 6808-6823, 2020.")
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

    doc.save("Projeto_de_Pesquisa_Integra_Palmilha_Inteligente.docx")
    print("Salvo: Projeto_de_Pesquisa_Integra_Palmilha_Inteligente.docx")

if __name__ == "__main__":
    print("Iniciando geração de todos os documentos Word...")
    create_portfolio_doc()
    create_anuencia_doc()
    create_tcle_doc()
    create_isencao_doc()
    create_instrumento_coleta_doc()
    create_projeto_integra_doc()
    print("Todos os 6 documentos foram gerados com sucesso!")
