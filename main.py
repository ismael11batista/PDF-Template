import os
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus.paragraph import Paragraph
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak, Frame,
    ListFlowable, ListItem
)

AZUL_ESCURO = HexColor("#002B5B")
AZUL_FAIXA = HexColor("#E3F2FD")
AMARELO = HexColor("#FFB300")
BRANCO = HexColor("#FFFFFF")
CINZA_FOOTER = HexColor("#D3D3D3")
CINZA_CLARO = HexColor("#F5F5F5")
CINZA_MEDIO = HexColor("#CCCCCC")

def wrap_text(text, font_name, font_size, max_width, c):
    lines, line = [], ""
    for word in text.split():
        test_line = f"{line} {word}".strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            line = test_line
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines

def mes_extenso_ptbr(month_number):
    nomes = [
        "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return nomes[month_number]

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self._logo_pequena_path = kwargs.pop('logo_pequena_path', None)
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_report_pages = len(self._saved_page_states)
        for idx, state in enumerate(self._saved_page_states):
            self.__dict__.update(state)
            page_number_logic = idx + 1
            self.draw_page_number(page_number_logic, num_report_pages)
            self._draw_footer()
            super().showPage()
        super().save()

    def draw_page_number(self, page, total_pages):
        width, height = self._pagesize
        self.setFont("Helvetica", 9)
        self.setFillColor(AZUL_ESCURO)
        texto = f"Página {page} de {total_pages}"
        self.drawRightString(width - 2.5*cm, height - 1.3*cm, texto)

    def _draw_footer(self):
        width, _ = self._pagesize
        self.setStrokeColor(CINZA_FOOTER)
        self.setLineWidth(0.5)
        self.line(2.5*cm, 2*cm, width - 2.5*cm, 2*cm)
        logo_width = 2.3*cm
        logo_height = 1.4*cm
        logo_x = width - logo_width - 2.5*cm
        logo_y = 0.65*cm
        lp = self._logo_pequena_path
        if lp and os.path.exists(lp):
            self.drawImage(lp, logo_x, logo_y, width=logo_width, height=logo_height,
                           preserveAspectRatio=True, mask='auto')
        self.setFont("Helvetica", 8)
        self.setFillColor(AZUL_ESCURO)
        self.drawCentredString(width/2, 1.5*cm, "CONFIDENCIAL - USO EXCLUSIVO DA EMPRESA CLIENTE")
        self.drawString(3.5*cm, 1.5*cm, f"Emitido em: {datetime.now().strftime('%d/%m/%Y')}")

class VerificacaoAntecedentesReportTodos:
    def __init__(self, output_filename, logo_grande_path, logo_pequena_path):
        self.output_filename = output_filename
        self.logo_grande_path = logo_grande_path
        self.logo_pequena_path = logo_pequena_path
        self.pagesize = A4
        self.width, self.height = self.pagesize
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self.elements = []
        self.doc = None

    def _setup_custom_styles(self):
        self.styles.add(ParagraphStyle(
            name='TitleAgence',
            fontName='Helvetica-Bold',
            fontSize=19,
            textColor=AZUL_ESCURO,
            spaceAfter=15,
            alignment=TA_CENTER
        ))
        self.styles.add(ParagraphStyle(
            name='Subtitle',
            fontName='Helvetica-Bold',
            fontSize=14,
            textColor=AZUL_ESCURO,
            spaceAfter=10,
            alignment=TA_LEFT
        ))
        self.styles.add(ParagraphStyle(
            name='NormalAgence',
            fontName='Helvetica',
            fontSize=11.5,
            textColor=AZUL_ESCURO,
            spaceAfter=8,
            alignment=TA_JUSTIFY
        ))
        self.styles.add(ParagraphStyle(
            name='ContactFinalOnly',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=AZUL_ESCURO,
            alignment=TA_CENTER,
            spaceBefore=12,
            spaceAfter=24
        ))
        self.styles.add(ParagraphStyle(
            name='TableCell',
            fontName='Helvetica',
            fontSize=9,
            textColor=AZUL_ESCURO,
            alignment=TA_LEFT,
            leading=11.5,
            wordWrap='CJK'
        ))
        self.styles.add(ParagraphStyle(
            name='TableHeader',
            fontName='Helvetica-Bold',
            fontSize=10,
            textColor=AZUL_ESCURO,
            alignment=TA_LEFT,   # Alinhar header à esquerda
            spaceAfter=1
        ))

    def draw_cover_page(self, c, titulo):
        width, height = self.width, self.height
        c.setFillColor(AZUL_ESCURO)
        c.rect(0, 0, width, height, fill=1)
        bar_width = 0.7 * cm
        c.setFillColor(AZUL_FAIXA)
        c.roundRect(0.5*cm, 0.5*cm, bar_width, height-1*cm, 0.35*cm, fill=1, stroke=0)
        logo_max_width = width * 0.5
        logo_max_height = height * 0.15
        logo_x = (width - logo_max_width) / 2
        logo_y = height * 0.60
        if os.path.exists(self.logo_grande_path):
            c.drawImage(self.logo_grande_path, logo_x, logo_y, width=logo_max_width, height=logo_max_height,
                        preserveAspectRatio=True, mask='auto')
        title_font = "Helvetica-Bold"
        title_size = 26
        max_title_width = width * 0.75
        title_lines = wrap_text(titulo, title_font, title_size, max_title_width, c)
        c.setFont(title_font, title_size)
        c.setFillColor(AMARELO)
        title_start_y = logo_y - 2.3*cm
        for i, line in enumerate(title_lines):
            y_line = title_start_y - i * title_size * 1.18
            c.drawCentredString(width/2, y_line, line)
        y_data = 4.7*cm
        agora = datetime.now()
        mes_ptbr = mes_extenso_ptbr(agora.month)
        current_date = f"{mes_ptbr} {agora.year}"
        c.setFont("Helvetica", 12)
        c.setFillColor(BRANCO)
        c.drawCentredString(width/2, y_data, f"Publicado em {current_date}")
        c.setStrokeColor(CINZA_FOOTER)
        c.setLineWidth(0.55)
        c.line(2*cm, 2.05*cm, width - 2*cm, 2.05*cm)
        c.setFont("Helvetica", 12)
        c.setFillColor(CINZA_FOOTER)
        ano_corrente = agora.year
        c.drawCentredString(width/2, 1.5*cm, f"© {ano_corrente} Agence Consultoria · Todos os direitos reservados")
        c.showPage()

    def draw_sobre_empresa(self, c, about_title, about_text, disclaimer_text):
        width, height = self.width, self.height
        c.setFillColor(BRANCO)
        c.rect(0, 0, width, height, fill=1)
        bar_width = 0.7*cm
        c.setFillColor(AZUL_FAIXA)
        c.roundRect(0.5*cm, 0.5*cm, bar_width, height-1*cm, 0.35*cm, fill=1, stroke=0)
        TITLE_Y = height - 5.2*cm
        LINE_Y = TITLE_Y - 0.45*cm
        ESPACAMENTO_EXTRA = 1.2*cm
        c.setFont("Helvetica-Bold", 22)
        c.setFillColor(AZUL_ESCURO)
        c.drawCentredString(width/2, TITLE_Y, about_title.upper())
        c.setStrokeColor(AMARELO)
        c.setLineWidth(1.3)
        c.line(width/2 - 6*cm, LINE_Y, width/2 + 6*cm, LINE_Y)
        estilo_about = ParagraphStyle(
            "SobreEmpresa",
            fontName="Helvetica",
            fontSize=13,
            leading=18,
            textColor=AZUL_ESCURO,
            alignment=TA_JUSTIFY,
            spaceAfter=15,
        )
        estilo_bullet = ParagraphStyle(
            "BulletSobre",
            parent=estilo_about,
            leftIndent=18,
            bulletIndent=7,
            spaceBefore=1,
            spaceAfter=0,
        )
        parags = [
            "Desde 1999, a Agence Consultoria é uma multinacional boutique especializada em soluções tecnológicas sob medida. "
            "Com presença no Brasil, Chile, Colômbia e EUA, tornamo-nos referência no cenário latino-americano de tecnologia e inovação.",
            "Atuamos como parceiro estratégico de empresas que buscam eficiência e transformação digital em seus processos. "
            "Nossas áreas de atuação abrangem:"
        ]
        bullets = [
            "Consultoria estratégica de TI",
            "Desenvolvimento de software personalizado",
            "Automação inteligente de processos (RPA)",
            "Headhunting especializado em tecnologia"
        ]
        parags += [
            "A excelência dos nossos serviços está alicerçada em uma equipe multidisciplinar, experiente e comprometida com o sucesso dos clientes. "
            "Cada projeto recebe atendimento individualizado, com foco na ética, transparência e absoluta confidencialidade.",
            "Ao longo de mais de duas décadas, consolidamos nossa reputação de confiança, inovação e entrega de valor. "
            "Mantenha sua empresa à frente — conte com a Agence para desafios de tecnologia e segurança da informação."
        ]
        story = []
        for idx, p in enumerate(parags):
            story.append(Paragraph(p, estilo_about))
            story.append(Spacer(1, 0.18*cm))
            if idx == 1:
                bullet_list = []
                for b in bullets:
                    bullet_list.append(ListItem(Paragraph(b, estilo_bullet), bulletText='•'))
                story.append(ListFlowable(bullet_list, bulletType='bullet', leftIndent=18))
                story.append(Spacer(1, 0.18*cm))
        frame_x = 3*cm
        frame_y = LINE_Y - ESPACAMENTO_EXTRA
        frame_height = frame_y - 2.8*cm
        frame_width = width - 6*cm
        frame = Frame(frame_x, 2.8*cm, frame_width, frame_height, showBoundary=0)
        frame.addFromList(story, c)
        box_margin_x = 2.5*cm
        box_margin_bottom = 2.55*cm
        box_width = width - 2*box_margin_x
        box_height = 2.8*cm
        box_y = box_margin_bottom
        c.setFillColor(CINZA_CLARO)
        c.roundRect(box_margin_x, box_y, box_width, box_height, 8, fill=1, stroke=0)
        c.setStrokeColor(AZUL_FAIXA)
        c.setLineWidth(1)
        c.roundRect(box_margin_x, box_y, box_width, box_height, 8, fill=0, stroke=1)
        text_margin = 0.5*cm
        tx = c.beginText()
        tx.setTextOrigin(box_margin_x + text_margin, box_y + box_height - 1.0*cm)
        tx.setFont("Helvetica-Bold", 10)
        tx.setFillColor(AZUL_ESCURO)
        wrapped = wrap_text(disclaimer_text, "Helvetica-Bold", 10, box_width - 2*text_margin, c)
        for line in wrapped:
            tx.textLine(line)
        c.drawText(tx)
        logo_sm_width = 2.3*cm
        logo_sm_height = 1.4*cm
        logo_x = width - logo_sm_width - 1.5*cm
        logo_y = 0.65*cm
        if os.path.exists(self.logo_pequena_path):
            c.drawImage(self.logo_pequena_path, logo_x, logo_y, width=logo_sm_width,
                        height=logo_sm_height, preserveAspectRatio=True, mask='auto')
        c.setStrokeColor(CINZA_FOOTER)
        c.setLineWidth(0.55)
        c.line(2*cm, 2.05*cm, width - 2*cm, 2.05*cm)
        c.setFont("Helvetica", 12)
        c.setFillColor(CINZA_FOOTER)
        ano_corrente = datetime.now().year
        c.drawCentredString(
            width/2,
            1.25*cm,
            f"© {ano_corrente} Agence Consultoria · Todos os direitos reservados"
        )
        c.showPage()

    def add_candidate_info(self, name, cpf, other_info=None):
        self.elements.append(Paragraph("Dados do profissional", self.styles['TitleAgence']))
        self.elements.append(Spacer(1, 0.5*cm))
        if cpf and len(cpf) >= 2:
            masked_cpf = "***.***.***-" + cpf[-2:]
        else:
            masked_cpf = cpf
        data = [
            ["Nome completo:", name],
            ["CPF:", masked_cpf]
        ]
        if other_info:
            for key, value in other_info.items():
                data.append([f"{key}:", value])
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), CINZA_CLARO),
            ('TEXTCOLOR', (0, 0), (0, -1), AZUL_ESCURO),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, CINZA_MEDIO),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
        candidate_table = Table(data, colWidths=[4*cm, 10*cm])
        candidate_table.setStyle(table_style)
        self.elements.append(candidate_table)
        self.elements.append(Spacer(1, 1*cm))

    def add_results_table(self, results_data, dynamic_columns):
        self.elements.append(Paragraph("Resultados da verificação", self.styles['TitleAgence']))
        self.elements.append(Spacer(1, 0.5*cm))
        total_width = self.width - 5*cm
        base_col_widths = {
            "Órgão": 3.3*cm,
            "Resultado": 5.5*cm,
            "Status": 3*cm,
            "Data": 2.5*cm,
        }
        col_widths = []
        used_cols = []
        for col in dynamic_columns:
            if col in base_col_widths:
                col_widths.append(base_col_widths[col])
                used_cols.append(col)
            else:
                col_widths.append(3.3*cm)
                used_cols.append(col)
        table_data = [
            [Paragraph(str(col), self.styles['TableHeader']) for col in used_cols]
        ]
        for result in results_data:
            row = []
            for col in used_cols:
                valor = result.get(col, "")
                if valor is None:
                    valor = ""
                row.append(Paragraph(str(valor), self.styles['TableCell']))
            table_data.append(row)
        table_style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), AZUL_FAIXA),
            ('TEXTCOLOR', (0, 0), (-1, 0), AZUL_ESCURO),
            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),       # Ajuste: Header LEFT
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 1), (-1, -1), BRANCO),
            ('TEXTCOLOR', (0, 1), (-1, -1), AZUL_ESCURO),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, CINZA_MEDIO),
            ('BOX', (0, 0), (-1, -1), 1, AZUL_FAIXA),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]
        results_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        results_table.setStyle(TableStyle(table_style_commands))
        self.elements.append(results_table)
        self.elements.append(Spacer(1, 1*cm))

    def add_contact_final_page_only(self):
        contact_text = "Para esclarecimentos adicionais, entre em contato com nossa equipe técnica pelo<br/>" \
                       "<b>e-mail ismael.batista@sp.agence.com.br</b> ou pelo telefone <b>(11) 5286-3220</b>."
        self.elements.append(Spacer(1, 1.2*cm))
        self.elements.append(Paragraph(contact_text, self.styles['ContactFinalOnly']))

    def generate_report(self, all_candidates_data, dynamic_columns):
        disclaimer_text = (
            "AVISO DE CONFIDENCIALIDADE: Este documento contém informações confidenciais e privilegiadas. "
            "Qualquer divulgação, distribuição ou cópia para pessoas não autorizadas é proibida. As informações "
            "contidas neste relatório destinam-se exclusivamente à empresa cliente, em conformidade com a legislação aplicável."
        )
        buff_path = self.output_filename + ".temp"
        c = canvas.Canvas(buff_path, pagesize=self.pagesize)
        self.draw_cover_page(
            c,
            "Relatório de Verificação de Antecedentes Criminais"
        )
        about_title = "Agence Consultoria"
        about_text = ""
        self.draw_sobre_empresa(c, about_title, about_text, disclaimer_text)
        c.save()
        self.elements = []
        total = len(all_candidates_data)
        for idx, candidate in enumerate(all_candidates_data):
            name = candidate.get("candidate_name", "")
            cpf = candidate.get("cpf", "")
            self.add_candidate_info(name, cpf, candidate.get("other_info"))
            self.add_results_table(candidate.get("results", []), dynamic_columns)
            if idx < total - 1:
                self.elements.append(PageBreak())
            else:
                self.add_contact_final_page_only()
        doc = SimpleDocTemplate(
            self.output_filename,
            pagesize=self.pagesize,
            leftMargin=2.5*cm, rightMargin=2.5*cm,
            topMargin=2.5*cm, bottomMargin=2.5*cm
        )
        def my_canvas_creator(*args, **kwargs):
            c = NumberedCanvas(*args, **kwargs, logo_pequena_path=self.logo_pequena_path)
            return c
        doc.build(self.elements, canvasmaker=my_canvas_creator)
        from PyPDF2 import PdfMerger
        merger = PdfMerger()
        merger.append(buff_path)
        merger.append(self.output_filename)
        merger.write(self.output_filename)
        merger.close()
        os.remove(buff_path)
        return self.output_filename

def gerar_relatorio_geral(dados_json, output_filename, logo_grande_path, logo_pequena_path):
    all_candidates = []
    columns = ["Órgão", "Resultado", "Status", "Data"]
    for candidato in dados_json:
        nome = candidato.get('nome', '')
        cpf = candidato.get('cpf', '')
        consultas = candidato.get('consultas', [])
        results = []
        for consulta in consultas:
            result = dict(zip(columns, consulta))
            results.append(result)
        report_data = {
            'candidate_name': nome,
            'cpf': cpf,
            'report_date': datetime.now().strftime("%d/%m/%Y"),
            'results': results
        }
        all_candidates.append(report_data)
    report = VerificacaoAntecedentesReportTodos(
        output_filename,
        logo_grande_path,
        logo_pequena_path
    )
    report.generate_report(all_candidates, columns)
    return output_filename

if __name__ == "__main__":
    import json
    with open("dados.json", "r", encoding="utf-8") as fp:
        dados = json.load(fp)
    output_filename = "relatorio_geral.pdf"
    logo_grande_path = "./images/logo_agence_relatorio_branca_grande.png"
    logo_pequena_path = "./images/logo_agence_relatorio_branca_pequena.png"
    gerar_relatorio_geral(dados, output_filename, logo_grande_path, logo_pequena_path)
    print(f"Relatório único gerado: {output_filename}")
