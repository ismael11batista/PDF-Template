import os
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.colors import HexColor, white
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image,
    PageBreak
)
from reportlab.lib.utils import ImageReader

AZUL_ESCURO   = HexColor("#002B5B")
AZUL_FAIXA    = HexColor("#E3F2FD")
AMARELO       = HexColor("#FFB300")
BRANCO        = HexColor("#FFFFFF")
CINZA_FOOTER  = HexColor("#D3D3D3")
CINZA_CLARO   = HexColor("#F5F5F5")
CINZA_MEDIO   = HexColor("#CCCCCC")

def wrap_text(text, font_name, font_size, max_width, c):
    lines, line = [], ""
    for word in text.split():
        test_line = f"{line} {word}".strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            line = test_line
        else:
            if line: lines.append(line)
            line = word
    if line: lines.append(line)
    return lines

def mes_extenso_ptbr(month_number):
    nomes = [
        "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return nomes[month_number]

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
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
        lp = getattr(self, "_logo_path", None)
        if lp and os.path.exists(lp):
            self.drawImage(lp, logo_x, logo_y, width=logo_width, height=logo_height, mask='auto')
        self.setFont("Helvetica", 8)
        self.setFillColor(AZUL_ESCURO)
        self.drawCentredString(width/2, 1.5*cm, "CONFIDENCIAL - USO EXCLUSIVO DA EMPRESA CLIENTE")
        self.drawString(3.5*cm, 1.5*cm, f"Emitido em: {datetime.now().strftime('%d/%m/%Y')}")

class VerificacaoAntecedentesReport:
    def __init__(self, output_filename, logo_branca_path):
        self.output_filename = output_filename
        self.logo_path = logo_branca_path
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

    def draw_cover_page(self, c, titulo, nome_profissional=None):
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
        if os.path.exists(self.logo_path):
            c.drawImage(self.logo_path, logo_x, logo_y, width=logo_max_width, height=logo_max_height,
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
        # SUBTÍTULO COM WRAP PARA TEXTO LONGO
        if nome_profissional and nome_profissional.strip():
            subtitulo = f"Análise detalhada para o profissional {nome_profissional.strip()}"
            subtitle_font = "Helvetica"
            subtitle_size = 15
            max_subtitle_width = width * 0.75
            subtitle_lines = wrap_text(subtitulo, subtitle_font, subtitle_size, max_subtitle_width, c)
            c.setFont(subtitle_font, subtitle_size)
            c.setFillColor(BRANCO)
            y_sub = title_start_y - len(title_lines)*title_size*1.18 - 1.0*cm
            for i, subline in enumerate(subtitle_lines):
                y_line = y_sub - i * subtitle_size * 1.18
                c.drawCentredString(width/2, y_line, subline)
            y_after_sub = y_sub - len(subtitle_lines)*subtitle_size*1.18 - 1.0*cm
            y_data = min(4.7*cm, y_after_sub)
        else:
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
        c.setFont("Helvetica-Bold", 22)
        c.setFillColor(AZUL_ESCURO)
        c.drawCentredString(width/2, height - 4*cm, about_title.upper())
        c.setStrokeColor(AMARELO)
        c.setLineWidth(1.3)
        c.line(width/2 - 6*cm, height - 4.4*cm, width/2 + 6*cm, height - 4.4*cm)
        body_font = "Helvetica"
        body_size = 13
        max_width_text = width - 6*cm
        y_body = height - 6.5*cm
        dummy_canvas = canvas.Canvas("dummy.pdf", pagesize=A4)
        about_lines = wrap_text(about_text, body_font, body_size, max_width_text, dummy_canvas)
        c.setFont(body_font, body_size)
        c.setFillColor(AZUL_ESCURO)
        line_height = body_size * 1.5
        for line in about_lines:
            c.drawString(3*cm, y_body, line)
            y_body -= line_height
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
        if os.path.exists(self.logo_path):
            c.drawImage(self.logo_path, logo_x, logo_y, width=logo_sm_width, height=logo_sm_height, preserveAspectRatio=True, mask='auto')
        c.showPage()

    def generate_report(self, data):
        disclaimer_text = (
            "AVISO DE CONFIDENCIALIDADE: Este documento contém informações confidenciais e privilegiadas. "
            "Qualquer divulgação, distribuição ou cópia para pessoas não autorizadas é proibida. As informações "
            "contidas neste relatório destinam-se exclusivamente à empresa cliente, em conformidade com a legislação aplicável."
        )
        buff_path = self.output_filename + ".temp"
        c = canvas.Canvas(buff_path, pagesize=self.pagesize)
        nome_profissional = data.get("candidate_name", "")
        self.draw_cover_page(
            c,
            "Relatório de Verificação de Antecedentes Criminais",
            nome_profissional=nome_profissional
        )
        about_title = "Agence Consultoria"
        about_text = (
            "Fundada em 1999, a Agence Consultoria é uma multinacional brasileira líder em soluções tecnológicas sob medida. "
            "Com escritórios no Brasil, Chile, Colômbia e EUA, oferecemos consultoria, desenvolvimento de software, automação de processos e headhunting de TI. "
            "Atuamos com ética, precisão e absoluta confidencialidade "
            "na análise de antecedentes, riscos reputacionais e pesquisas estratégicas. Nosso compromisso é garantir segurança, "
            "integridade e transparência para nossos clientes no ambiente corporativo ou institucional."
        )
        self.draw_sobre_empresa(c, about_title, about_text, disclaimer_text)
        c.save()
        self.elements = []
        self.add_candidate_info(data.get("candidate_name", ""), data.get("cpf", ""), data.get("other_info"))
        self.add_results_table(data.get("results", []))
        self.add_analysis(data.get("analysis", ""))
        self.add_contact_block()
        doc = SimpleDocTemplate(
            self.output_filename,
            pagesize=self.pagesize,
            leftMargin=2.5*cm, rightMargin=2.5*cm,
            topMargin=2.5*cm, bottomMargin=2.5*cm
        )
        def my_canvas_creator(*args, **kwargs):
            c = NumberedCanvas(*args, **kwargs)
            c._logo_path = self.logo_path
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

    def add_candidate_info(self, name, cpf, other_info=None):
        self.elements.append(Paragraph("Dados do candidato", self.styles['TitleAgence']))
        self.elements.append(Spacer(1, 0.5*cm))
        if cpf and len(cpf) >= 11:
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
        ])
        candidate_table = Table(data, colWidths=[4*cm, 10*cm])
        candidate_table.setStyle(table_style)
        self.elements.append(candidate_table)
        self.elements.append(Spacer(1, 1*cm))

    def add_results_table(self, results_data):
        self.elements.append(Paragraph("Resultados da verificação", self.styles['TitleAgence']))
        self.elements.append(Spacer(1, 0.5*cm))
        headers = ["Ocorrências", "Tipo", "Status", "Data da Ocorrência"]
        table_data = [headers]
        for result in results_data:
            row = [
                result.get("Ocorrências", ""),
                result.get("Tipo", ""),
                result.get("Status", ""),
                result.get("Data da Ocorrência", "")
            ]
            table_data.append(row)
        table_style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), AZUL_FAIXA),
            ('TEXTCOLOR', (0, 0), (-1, 0), AZUL_ESCURO),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
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
        ]
        table_style = TableStyle(table_style_commands)
        results_table = Table(table_data, repeatRows=1)
        results_table.setStyle(table_style)
        self.elements.append(results_table)
        self.elements.append(Spacer(1, 1*cm))

    def add_analysis(self, analysis_text):
        self.elements.append(Paragraph("Análise dos resultados", self.styles['TitleAgence']))
        self.elements.append(Spacer(1, 0.5*cm))
        paragraphs = analysis_text.split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                self.elements.append(Paragraph(paragraph, self.styles['NormalAgence']))
                self.elements.append(Spacer(1, 0.3*cm))
        self.elements.append(Spacer(1, 0.7*cm))

    def add_contact_block(self):
        contact_text = """
<b>Para esclarecimentos adicionais</b>, entre em contato com nossa equipe técnica pelo e-mail <a href="mailto:ismael.batista@sp.agence.com.br">ismael.batista@sp.agence.com.br</a> 
ou pelo telefone (11) 5286-3220.
"""    
        contact_style = ParagraphStyle(
            'Contact',
            parent=self.styles['Normal'],
            fontSize=11,
            alignment=TA_CENTER,
            spaceBefore=12,
            spaceAfter=24
        )
        self.elements.append(Paragraph(contact_text, contact_style))

def load_data_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    candidates = {}
    current_candidate = None
    for _, row in df.iterrows():
        if not pd.isna(row['Nome/CPF']):
            current_candidate = row['Nome/CPF']
            candidates[current_candidate] = []
        if current_candidate:
            candidates[current_candidate].append({
                'Ocorrências': row['Ocorrências'],
                'Tipo': row['Tipo'],
                'Status': row['Status'],
                'Data da Ocorrência': row['Data da Ocorrência']
            })
    return candidates

def generate_analysis(results):
    analysis = "Com base nas verificações realizadas, foram identificadas as seguintes ocorrências:\n\n"
    occurrences_count = {}
    for result in results:
        tipo = result.get('Tipo', '')
        if tipo in occurrences_count:
            occurrences_count[tipo] += 1
        else:
            occurrences_count[tipo] = 1
    for tipo, count in occurrences_count.items():
        if tipo == "NADA CONSTA":
            analysis += f"- {tipo}: Não foram encontradas ocorrências em {count} consulta(s).\n\n"
        else:
            analysis += f"- {tipo}: Foram encontradas {count} ocorrência(s) deste tipo.\n\n"
    arquivados = sum(1 for result in results if result.get('Status', '') == 'ARQUIVADO')
    if arquivados > 0:
        analysis += f"Das ocorrências identificadas, {arquivados} estão com status ARQUIVADO, o que indica que os processos foram concluídos.\n\n"
    analysis += "É importante ressaltar que a presença de ocorrências não implica necessariamente em impedimento para contratação, devendo ser analisada caso a caso, considerando a natureza da função a ser exercida e o tempo decorrido desde as ocorrências."
    return analysis

def main(excel_file, output_dir, logo_branca_path):
    os.makedirs(output_dir, exist_ok=True)
    candidates_data = load_data_from_excel(excel_file)
    generated_reports = []
    for candidate, results in candidates_data.items():
        if '\n' in candidate:
            parts = candidate.split('\n')
            name = parts[0]
            cpf = parts[1] if len(parts) > 1 else ''
        else:
            name = candidate
            cpf = ''
        analysis = generate_analysis(results)
        report_data = {
            'candidate_name': name,
            'cpf': cpf,
            'report_date': datetime.now().strftime("%d/%m/%Y"),
            'results': results,
            'analysis': analysis
        }
        safe_name = name.replace(' ', '_').lower()
        output_filename = os.path.join(output_dir, f"relatorio_{safe_name}.pdf")
        report = VerificacaoAntecedentesReport(
            output_filename,
            logo_branca_path
        )
        report.generate_report(report_data)
        generated_reports.append(output_filename)
        print(f"Relatório gerado: {output_filename}")
    return generated_reports

if __name__ == "__main__":
    excel_file = "resultados.xlsx"
    output_dir = "./"
    logo_branca_path = "./images/logo_branca.png"
    reports = main(excel_file, output_dir, logo_branca_path)
    print(f"Total de relatórios gerados: {len(reports)}")
