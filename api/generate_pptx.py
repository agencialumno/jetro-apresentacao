import copy
import math
import io
import re
import unicodedata

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# ===================== Paleta (mesma identidade visual do frontend) =====================
DARK_PURPLE = RGBColor(0x1D, 0x13, 0x39)
PURPLE = RGBColor(0x45, 0x2E, 0x86)
GOLD = RGBColor(0xD0, 0x8E, 0x02)
BLUE = RGBColor(0x44, 0x72, 0xC4)
RED = RGBColor(0xC0, 0x50, 0x4D)
GRAY_TEXT = RGBColor(0x6B, 0x64, 0x80)
BORDER = RGBColor(0xE4, 0xE1, 0xEC)
CARD_BG = RGBColor(0xF7, 0xF7, 0xF7)

DEFAULT_X = Inches(1.3)
DEFAULT_Y = Inches(1.3)
DEFAULT_CHART_HEIGHT = Inches(7.0)


# ===================== Helpers básicos =====================

def fmt_brl(v):
    if v is None:
        return "-"
    try:
        n = float(v)
    except (TypeError, ValueError):
        return "-"
    s = f"{abs(n):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if n < 0 else "") + s


def iter_all_shapes(shapes):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)
        else:
            yield shape


def set_text(shape, text):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    paragraphs = tf.paragraphs
    if paragraphs and paragraphs[0].runs:
        first_run = paragraphs[0].runs[0]
        first_run.text = text
        for extra_run in list(paragraphs[0].runs[1:]):
            extra_run._r.getparent().remove(extra_run._r)
        for extra_p in list(paragraphs[1:]):
            extra_p._p.getparent().remove(extra_p._p)
    else:
        tf.text = text

def _set_text_on_element(shape_el, text):
    """Como set_text(), mas opera DIRETO no elemento XML ainda solto (antes de
    ser anexado a arvore do slide novo). Usamos isso em vez de mutar o shape
    depois de anexado porque, nesse modelo especifico, mutacoes feitas APOS a
    anexacao nao estao sendo persistidas na hora de salvar o .pptx (confirmado
    por diagnostico: o texto aparecia certo em memoria, inclusive apos o
    prs.save(), mas o arquivo fisico gravado continuava com o texto antigo do
    template). Mutar o elemento antes dele entrar na arvore contorna isso."""
    txBody = shape_el.find(qn("p:txBody"))
    if txBody is None:
        return
    paragraphs = txBody.findall(qn("a:p"))
    if not paragraphs:
        return
    first_p = paragraphs[0]
    runs = first_p.findall(qn("a:r"))
    if runs:
        first_run = runs[0]
        t_el = first_run.find(qn("a:t"))
        if t_el is not None:
            t_el.text = text
        for extra_run in runs[1:]:
            first_p.remove(extra_run)
    for extra_p in paragraphs[1:]:
        txBody.remove(extra_p)


def _remap_image_relationships(source_slide, dest_slide):
    root = dest_slide.shapes._spTree.getroottree().getroot()
    for blip in root.iter(qn("a:blip")):
        old_rId = blip.get(qn("r:embed"))
        if not old_rId or old_rId not in source_slide.part.rels:
            continue
        rel = source_slide.part.rels[old_rId]
        new_rId = dest_slide.part.relate_to(rel.target_part, rel.reltype)
        blip.set(qn("r:embed"), new_rId)


def duplicate_slide(prs, index, text_replacements=None):
    """text_replacements: dict {nome_do_shape: novo_texto}, aplicado no
    elemento ainda solto, ANTES de ele ser anexado ao slide novo."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)

    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    for shp in source.shapes:
        el = copy.deepcopy(shp._element)
        if text_replacements and shp.name in text_replacements:
            _set_text_on_element(el, text_replacements[shp.name])
        dest.shapes._spTree.append(el)

    # Copiamos o <p:bg> (cor de fundo solida do slide) manipulando o XML
    # DIRETAMENTE via lxml, sem usar a propriedade dest.background do
    # python-pptx: confirmamos por diagnostico que o simples ACESSO a essa
    # propriedade forcava o python-pptx a reconstruir parte da arvore do slide
    # a partir do layout, revertendo as mutacoes de texto feitas no loop acima.
    src_cSld = source.part._element.find(qn("p:cSld"))
    dest_cSld = dest.part._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg")) if src_cSld is not None else None
    if src_bg is not None and dest_cSld is not None:
        existing_bg = dest_cSld.find(qn("p:bg"))
        if existing_bg is not None:
            dest_cSld.remove(existing_bg)
        new_bg = copy.deepcopy(src_bg)
        dest_cSld.insert(0, new_bg)  # <p:bg> deve vir antes de <p:spTree>

    _remap_image_relationships(source, dest)

    return dest


TITULO_CONTEUDO_PT = 32  # igual a .pv-ct em style.css (32pt = 2.22cqw)


def add_title(slide, titulo):
    for shape in slide.placeholders:
        try:
            if shape.placeholder_format.idx == 0:
                set_text(shape, titulo)
                return
        except Exception:
            continue

    for shape in iter_all_shapes(slide.shapes):
        if shape.has_text_frame and "title" in (shape.name or "").lower():
            set_text(shape, titulo)
            return

    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(17.4), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = titulo
    p = tf.paragraphs[0]
    p.font.size = Pt(TITULO_CONTEUDO_PT)
    p.font.bold = True
    p.font.color.rgb = DARK_PURPLE


# ---- Periodo do relatorio no cabecalho (linha pequena logo abaixo do titulo) ----
_MESES_NUM = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}


def _mes_num(label):
    """'Nov/2025' -> '11/2025' (devolve o proprio texto se nao reconhecer)."""
    m = re.match(r"^\s*([A-Za-zçÇ]{3})[a-z]*\s*/\s*(\d{4})\s*$", label or "")
    if m and m.group(1).lower() in _MESES_NUM:
        return f"{_MESES_NUM[m.group(1).lower()]:02d}/{m.group(2)}"
    return label or ""


def _periodo_meses(meses):
    if not meses:
        return None
    return f"Período: {_mes_num(meses[0])} a {_mes_num(meses[-1])}"


def _periodo_16b(txt):
    """'De 01/10/2026 até 31/10/2026' -> 'Período: 01/10/2026 a 31/10/2026'."""
    m = re.match(r"^\s*De\s+(.+?)\s+at[ée]\s+(.+?)\s*$", txt or "", re.I)
    if m:
        return f"Período: {m.group(1)} a {m.group(2)}"
    return f"Período: {txt.strip()}" if txt and txt.strip() else None


def add_periodo(slide, texto):
    """Linha 'Periodo: ...' abaixo do titulo (x=1,3\", y=1,3\"); igual a .pv-periodo em style.css."""
    if not texto:
        return
    tb = slide.shapes.add_textbox(Inches(1.3), Inches(1.3), Inches(17.4), Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_top = tf.margin_bottom = 0
    tf.margin_left = 0
    tf.text = texto
    p = tf.paragraphs[0]
    p.font.size = Pt(14)
    p.font.bold = False
    p.font.color.rgb = RGBColor(0x5B, 0x4F, 0x86)


# ---- Divisores: fonte automatica (auto-fit). Mesma estimativa de tamDivisor em app.js ----
DIVISOR_CAIXA_W = 15.0  # mais estreita que o slide pra nao encostar nas curvas decorativas dos cantos
DIVISOR_CAIXA_H = 5.0
DIVISOR_PT_MIN = 24


def _larg_char_divisor(ch):
    if ch == " ":
        return 0.28
    if ch.isupper() or ch.isdigit():
        return 0.74
    return 0.62


def _tam_divisor(texto):
    """Maior fonte (60pt -> 24pt, de 2 em 2) em que o titulo cabe na caixa do divisor, sem quebrar palavra no meio."""
    util = DIVISOR_CAIXA_W - 0.5
    palavras = (texto or "").split()
    tam = DIVISOR_PT
    while tam >= DIVISOR_PT_MIN:
        def larg(p):
            return sum(_larg_char_divisor(c) for c in p) * tam / 72
        esp = _larg_char_divisor(" ") * tam / 72
        if all(larg(p) <= util for p in palavras):
            linhas, atual = 1, 0.0
            for p in palavras:
                w = larg(p)
                if atual == 0:
                    atual = w
                elif atual + esp + w <= util:
                    atual += esp + w
                else:
                    linhas += 1
                    atual = w
            if linhas * tam * 1.25 / 72 <= DIVISOR_CAIXA_H - 0.2:
                return tam
        tam -= 2
    return DIVISOR_PT_MIN


def _borda(cell, tag, largura_pt=2.25, cor="46328F"):
    """Borda (a:lnT topo / a:lnB base) da celula; a ordem no tcPr e lnL, lnR, lnT, lnB, depois o preenchimento."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn(tag)):
        tcPr.remove(old)
    ln = tcPr.makeelement(qn(tag), {"w": str(int(largura_pt * 12700)), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
    sf = ln.makeelement(qn("a:solidFill"), {})
    clr = sf.makeelement(qn("a:srgbClr"), {"val": cor})
    sf.append(clr)
    ln.append(sf)
    pos = 0
    for i, ch in enumerate(list(tcPr)):
        if ch.tag in (qn("a:lnL"), qn("a:lnR")) or (tag == "a:lnB" and ch.tag == qn("a:lnT")):
            pos = i + 1
    tcPr.insert(pos, ln)


# ===================== Tabelas =====================

def add_table(slide, x, headers, rows, col_widths_in, font_size=10, height_in=0.35, left_cols=None, y=None, margin_v=None, margin_h=None, sep_rows=None):
    left_cols = left_cols if left_cols is not None else {0}
    top = y if y is not None else DEFAULT_Y

    n_rows = len(rows) + 1
    n_cols = len(headers)
    total_width = Inches(sum(col_widths_in))
    total_height = Inches(height_in * n_rows)

    graphic_frame = slide.shapes.add_table(n_rows, n_cols, x, top, total_width, total_height)
    table = graphic_frame.table

    for i, w in enumerate(col_widths_in):
        table.columns[i].width = Inches(w)
    for i in range(n_rows):
        table.rows[i].height = Inches(height_in)

    for c, htext in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = "" if htext is None else str(htext)
        if margin_v is not None:
            cell.margin_top = cell.margin_bottom = Inches(margin_v)
        if margin_h is not None:
            cell.margin_left = cell.margin_right = Inches(margin_h)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK_PURPLE
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.LEFT if c in left_cols else PP_ALIGN.RIGHT

    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            if isinstance(val, tuple):
                text, bold = val
            else:
                text, bold = val, False
            cell = table.cell(r, c)
            cell.text = "" if text is None else str(text)
            if margin_v is not None:
                cell.margin_top = cell.margin_bottom = Inches(margin_v)
            if margin_h is not None:
                cell.margin_left = cell.margin_right = Inches(margin_h)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(font_size)
            p.font.bold = bool(bold)
            p.alignment = PP_ALIGN.LEFT if c in left_cols else PP_ALIGN.RIGHT
            if sep_rows and (r - 1) in sep_rows:
                _borda(cell, "a:lnT")
                _borda(table.cell(r - 1, c), "a:lnB")  # a base da linha de cima tambem, senao o estilo da tabela esconde a linha

    return graphic_frame


# ===================== Graficos =====================

# Rotulos de valor nos graficos em reais por extenso: "R$ 6.141,12" (separadores seguem o idioma do PowerPoint).
FORMATO_VALOR_REAIS = '"R$ "#,##0.00'
# Barras agrupadas: sem centavos ("R$ 10.485") pra o rotulo horizontal caber na largura da barra
FORMATO_VALOR_REAIS_INTEIRO = '"R$ "#,##0'


def _eixo_valores(maxv):
    """(maximo, passo) do eixo de valores com folga de ~25% no topo pros rotulos. Mesmo algoritmo de
    pvEixoValores em app.js, pra prévia e .pptx terem a mesma escala."""
    if maxv <= 0:
        return 1.0, 1.0
    alvo = maxv * 1.25
    bruto = alvo / 6
    mag = 10 ** math.floor(math.log10(bruto))
    passo = mag
    for m in (1, 2, 2.5, 5, 10):
        passo = m * mag
        if passo >= bruto:
            break
    return math.ceil(alvo / passo - 1e-9) * passo, passo


def _ajustar_eixo(chart, valores, ref_max=None):
    if any(v < 0 for v in valores):
        return
    maximo, passo = _eixo_valores(ref_max if ref_max is not None else (max(valores) if valores else 0))
    va = chart.value_axis
    va.minimum_scale = 0
    va.maximum_scale = maximo
    va.major_unit = passo
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
    va.tick_labels.number_format = '#,##0'
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.size = Pt(CHART_FONTE_EIXO)


def _rotulos_valores(chart, posicao, tamanho, girar=False, formato=FORMATO_VALOR_REAIS):
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_value = True
    dl.number_format = formato
    dl.number_format_is_linked = False
    dl.position = posicao
    dl.font.size = Pt(tamanho)
    if girar:
        # rotulos na vertical (de baixo pra cima): "R$ 10.485,06" nao cabe na horizontal entre barras vizinhas
        bodyPr = dl._element.get_or_add_txPr().find(qn("a:bodyPr"))
        bodyPr.set("rot", "-5400000")
        bodyPr.set("vert", "horz")


# Fontes dos graficos pensadas pra telao: 6 meses por slide nas barras deixam o rotulo de 14pt caber na barra
CHART_FONTE_GERAL = 18
CHART_FONTE_EIXO = 16
CHART_FONTE_ROTULO_BARRAS = 22
CHART_FONTE_ROTULO_LINHA = 20
BARRAS_MESES_POR_SLIDE = 12  # igual a BARRAS_MESES_POR_SLIDE em app.js (12 meses = 1 slide)
LINHA_MESES_POR_SLIDE = 12  # igual a LINHA_MESES_POR_SLIDE em app.js
GAP_BARRAS_PCT = 60  # igual a GAP_BARRAS em app.js (largura da barra = categoria / (series + 0,6))


def add_bar_chart(slide, x, width_in, categories, series, y=None, height=None, ref_max=None):
    # legenda so quando ha mais de uma serie (o quadradinho de cor da legenda e nativo do grafico)
    top = y if y is not None else DEFAULT_Y
    h = height if height is not None else DEFAULT_CHART_HEIGHT

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values, _color in series:
        chart_data.add_series(name, values)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, top, Inches(width_in), h, chart_data
    )
    chart = graphic_frame.chart
    chart.has_title = False
    chart.has_legend = len(series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    chart.font.size = Pt(CHART_FONTE_GERAL)

    for i, (_name, _values, color) in enumerate(series):
        s = chart.series[i]
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = color

    _ajustar_eixo(chart, [v for _n, vals, _c in series for v in vals], ref_max)
    # Barras largas (espaco entre grupos de 20%) + rotulo horizontal de 9pt sem centavos: cada rotulo cabe
    # dentro da largura da propria barra, entao nao colide com a barra vizinha nem com outro mes.
    chart.plots[0].gap_width = GAP_BARRAS_PCT
    chart.plots[0].overlap = 0
    return graphic_frame


def add_line_chart(slide, x, width_in, categories, series_name, values, color, y=None, height=None, ref_max=None):
    top = y if y is not None else DEFAULT_Y
    h = height if height is not None else DEFAULT_CHART_HEIGHT

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, x, top, Inches(width_in), h, chart_data
    )
    chart = graphic_frame.chart
    chart.has_title = False   # serie unica: sem titulo automatico duplicado e sem legenda
    chart.has_legend = False
    chart.font.size = Pt(CHART_FONTE_GERAL)

    s = chart.series[0]
    s.format.line.color.rgb = color
    s.format.line.width = Pt(2.5)
    s.smooth = False
    # marcador da mesma cor da linha (o padrao do PowerPoint era um losango azul)
    s.marker.style = XL_MARKER_STYLE.CIRCLE
    s.marker.size = 8
    s.marker.format.fill.solid()
    s.marker.format.fill.fore_color.rgb = color
    s.marker.format.line.color.rgb = color

    _ajustar_eixo(chart, list(values), ref_max)
    return graphic_frame


# ===================== Texto corrido / bullets =====================

# Texto livre (bullets): fonte grande pra leitura a distancia. Escolhe o maior tamanho em que tudo cabe em 1 slide
# (area de 16,4" x 7,4"); se nem 24pt cabe, divide em varios slides. Mesmo algoritmo de paginarBullets em app.js.
BULLETS_FONTES = [54, 48, 44, 40, 36, 32, 28, 24]
BULLETS_LARGURA_IN = 16.4
BULLETS_ALTURA_IN = 7.4


def _altura_bullets(itens, tam):
    cpl = max(10, int(BULLETS_LARGURA_IN * 72 / (tam * 0.52)))
    return sum(max(1, math.ceil((len(t) + 2) / cpl)) * tam * 1.2 + 10 for t in itens) / 72


def paginar_bullets(itens):
    """[(tamanho_pt, [itens da pagina]), ...]"""
    for tam in BULLETS_FONTES:
        if _altura_bullets(itens, tam) <= BULLETS_ALTURA_IN:
            return [(tam, itens)]
    tam = BULLETS_FONTES[-1]
    paginas, atual = [], []
    for it in itens:
        if atual and _altura_bullets(atual + [it], tam) > BULLETS_ALTURA_IN:
            paginas.append((tam, atual))
            atual = []
        atual.append(it)
    if atual:
        paginas.append((tam, atual))
    return paginas


def add_bullets(slide, x, items, y=None, height_in=3.5, width_in=16.0, font_size=14, color=None):
    top = y if y is not None else DEFAULT_Y
    tb = slide.shapes.add_textbox(x, top, Inches(width_in), Inches(height_in))
    tf = tb.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "- " + str(item)
        p.font.size = Pt(font_size)
        p.space_after = Pt(10)
        if color is not None:
            p.font.color.rgb = color

    return tb


# Linhas de dados por slide nas tabelas de previsao (fonte 11pt, linha 0,27"): 26 linhas + cabecalho
# terminam em ~9,2", antes da decoracao do rodape (~9,7").
# Tabelas grandes de informacao: fonte 16pt, linha de 0,4". Comecam em 1,9" e 17 linhas + cabecalho
# terminam em ~9,1", antes da decoracao do rodape (~9,7"). Mais que isso vira outra pagina.
TABELA_FONTE = 18
TABELA_ALTURA_LINHA = 0.44
TABELA_LINHAS_POR_SLIDE = 15
PREVISAO_LINHAS_POR_SLIDE = TABELA_LINHAS_POR_SLIDE
# Previsao: um slide por grupo (CONTRATOS, PESSOAL...) + um "Resumo" com as linhas de fechamento. Poucas linhas => fonte maior.
PREVISAO_FONTE_GRANDE = 24
PREVISAO_ALTURA_GRANDE = 0.62
PREVISAO_LINHAS_GRANDE = 12  # incluindo o cabecalho

# Tabelas de lancamentos do 16B: mesma geometria da previsao (11pt, linha de 0,27"). A descricao e
# cortada pra nunca quebrar em 2 linhas (o que aumentaria a linha e estouraria o slide).
LANCAMENTOS_LINHAS_POR_SLIDE = TABELA_LINHAS_POR_SLIDE
LANCAMENTOS_DESCRICAO_MAX = 80


def _cortar(texto, maximo):
    t = " ".join(str(texto or "").split())
    return t if len(t) <= maximo else t[: maximo - 1].rstrip() + "…"


# ===================== Slide de pauta (paginado) =====================

# Capacidade em "linhas" por slide (ver paginarPauta em app.js - manter os valores iguais).
# Calibrado medindo o slide no PowerPoint: cada linha ocupa ~0,5"; a lista comeca em 3,5" (1o slide) ou
# 1,9" (continuacao) e a decoracao do rodape comeca em ~9,7". 10 e 13 linhas terminam por volta de
# 8,6" - mais de 1" de folga, que absorve itens que quebram em 2 linhas. Com 22pt, a lista (15,5" de
# largura) comporta ~120 caracteres; usamos 100 pra estimar o peso de itens longos.
PAUTA_LINHAS_PRIMEIRA = 10
PAUTA_LINHAS_CONTINUACAO = 13
PAUTA_CHARS_POR_LINHA = 100
PAUTA_FONT_SIZE = 22
TEXTO_PAUTA_ASSUNTOS = "Assuntos gerais"  # exibicao curta do item 'assuntos' na lista Ordem do dia (manter igual em app.js)


def paginar_pauta(textos):
    """Divide os itens em paginas [[(numero, texto), ...], ...] sem estourar a capacidade.
    Itens longos pesam mais de uma linha. A numeracao e continua entre as paginas.
    Cada pagina leva no minimo 1 item e o corte e feito por indice (inicio:fim), entao todo
    item cai em exatamente uma pagina, mesmo que a estimativa de peso esteja errada."""
    pesos = [max(1, -(-(len(t) + 4) // PAUTA_CHARS_POR_LINHA)) for t in textos]  # ceil
    paginas = []
    inicio = 0
    capacidade = PAUTA_LINHAS_PRIMEIRA
    while inicio < len(textos):
        fim = inicio + 1
        usado = pesos[inicio]
        while fim < len(textos) and usado + pesos[fim] <= capacidade:
            usado += pesos[fim]
            fim += 1
        paginas.append([(i + 1, textos[i]) for i in range(inicio, fim)])
        inicio = fim
        capacidade = PAUTA_LINHAS_CONTINUACAO
    if sum(len(pg) for pg in paginas) != len(textos):
        raise RuntimeError("paginar_pauta perdeu itens da pauta")
    return paginas


# Fonte dinamica por pagina (manter igual a tamanhoFontePauta em app.js). A altura orcada e a mesma
# que ja foi medida no PowerPoint (linhas da capacidade x 0,506" a 22pt => fim do texto em ~8,5").
# Fonte maior => menos caracteres por linha e linhas mais altas, tudo proporcional a size/22.
PAUTA_FONTES = [32, 30, 28, 26, 24, 22]  # do maior pro menor; 22 e o piso (sempre cabe pela paginacao)
PAUTA_ALTURA_LINHA_IN = 0.506  # altura de 1 linha + espacamento entre itens a 22pt


def tamanho_fonte_pauta(textos, primeira):
    """Maior fonte da escala em que os itens DESTA pagina ainda cabem na altura orcada."""
    orcamento = (PAUTA_LINHAS_PRIMEIRA if primeira else PAUTA_LINHAS_CONTINUACAO) * PAUTA_ALTURA_LINHA_IN
    for tam in PAUTA_FONTES:
        escala = tam / PAUTA_FONT_SIZE
        chars = PAUTA_CHARS_POR_LINHA / escala
        linhas = sum(max(1, -(-(len(t) + 4) // int(chars))) for t in textos)
        if linhas * PAUTA_ALTURA_LINHA_IN * escala <= orcamento:
            return tam
    return PAUTA_FONT_SIZE


def linhas_edital(dados):
    """Linhas do cabecalho do edital (mesma regra de linhasEdital em app.js): so entra o trecho
    preenchido; horario invalido/vazio omite a convocacao inteira - nunca sobra placeholder."""
    def hora_ok(h):
        return bool(re.fullmatch(r"\d{1,2}h(\d{2})?", (h or "").strip()))
    data = (dados.get("data") or "").strip()
    local = (dados.get("local") or "").strip()
    c1, c2 = (dados.get("conv1") or "").strip(), (dados.get("conv2") or "").strip()
    linha1 = [x for x in (f"Data: {data}" if data else "", f"Local: {local}" if local else "") if x]
    linha2 = [x for x in (f"1ª convocação: {c1}" if hora_ok(c1) else "",
                          f"2ª convocação: {c2}" if hora_ok(c2) else "") if x]
    return ["   •   ".join(l) for l in (linha1, linha2) if l]


def _int_ou_none(v):
    d = re.sub(r"\D", "", str(v if v is not None else ""))
    return int(d) if d else None


def add_text_line(slide, x, y, width_in, height_in, text, size, bold=False, color=None):
    tb = slide.shapes.add_textbox(x, y, Inches(width_in), Inches(height_in))
    tf = tb.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    return tb


def add_numbered_list(slide, x, y, itens, width_in=15.5, height_in=6.2, font_size=PAUTA_FONT_SIZE):
    """itens: [(numero, texto), ...] - mantem o numero global da pauta."""
    tb = slide.shapes.add_textbox(x, y, Inches(width_in), Inches(height_in))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (numero, texto) in enumerate(itens):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(round(10 * font_size / PAUTA_FONT_SIZE))
        run = p.add_run()
        run.text = f"{numero}. {texto}"
        run.font.size = Pt(font_size)
        run.font.color.rgb = DARK_PURPLE
    return tb


# ===================== Cards de estatistica =====================

def add_stat_cards(slide, x, y, cards, width_total_in=16.0, height_in=1.8, gap_in=0.3):
    n = len(cards)
    if n == 0:
        return
    card_w_in = (width_total_in - gap_in * (n - 1)) / n
    cur_x = x

    for label, value in cards:
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, cur_x, y, Inches(card_w_in), Inches(height_in)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = BORDER
        box.line.width = Pt(1)
        box.shadow.inherit = False

        tf = box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = str(value)
        p1.font.size = Pt(28)
        p1.font.bold = True
        p1.font.color.rgb = PURPLE
        p1.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = str(label)
        p2.font.size = Pt(14)
        p2.font.color.rgb = GRAY_TEXT
        p2.alignment = PP_ALIGN.CENTER

        cur_x = Emu(int(cur_x) + int(Inches(card_w_in)) + int(Inches(gap_in)))


# ===================== Geracao da apresentacao =====================

def _ajustar_fonte_contato(shape, max_chars):
    """Encolhe a fonte do contato (33,3pt no modelo) quando o texto e maior que a barra dourada
    do modelo comporta - sem isso, o site padrao 'www.jetroadministradora.com.br' vazava da barra."""
    texto = shape.text_frame.text
    if len(texto) <= max_chars:
        return
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(round(33.32 * max_chars / len(texto), 1))


DIVISOR_PT = 60  # fonte unica dos slides divisores
PLANO_LINHAS_POR_SLIDE = 17  # igual a LINHAS_POR_SLIDE_PLANO em app.js (13pt, linha de 0,4")
PLANO_FONTE = 14
PLANO_ALTURA_LINHA = 0.4
PLANO_FONTE_GRANDE = 16  # slides com poucas linhas (ate PLANO_LINHAS_GRANDE incl. cabecalho) usam fonte e linha maiores
PLANO_ALTURA_GRANDE = 0.55
PLANO_LINHAS_GRANDE = 12
PLANO_COL_DESC = 4.2  # polegadas; o resto (17,4") e dividido entre os meses


def _dim_tabela(n_linhas, y_in, fim_in=9.3, fonte_base=18, altura_base=0.7):
    """Altura de linha e fonte para uma tabela de n_linhas (com cabecalho) caber entre y_in e fim_in."""
    altura = min(altura_base, (fim_in - y_in) / max(1, n_linhas))
    if altura >= altura_base * 0.85:
        fonte = fonte_base
    elif altura >= altura_base * 0.65:
        fonte = max(12, fonte_base - 2)
    else:
        fonte = 12
    return fonte, altura


def _grupos_plano(contas, secao, offset, n_meses):
    """Plano de contas real do condominio quebrado por grupo (conta de nivel 0 + descendentes), um item por
    slide: [(nome_do_grupo, [((texto, negrito), (mes1, negrito), ... (mesN, negrito)), ...])].
    Contas sem valor nenhum nos meses exibidos ficam de fora. Grupos de uma linha so (conta solta, sem filhas)
    sao reunidos num unico grupo 'Outras contas'. Mesmo algoritmo de gruposPlanoContas em app.js."""
    grupos = []
    for c in contas:
        if c.get("secao") != secao:
            continue
        vals = [v or 0 for v in (c.get("valores") or [])[offset:]]
        vals = (vals + [0] * n_meses)[:n_meses]
        if not any(vals):
            continue
        nivel = max(0, int(c.get("nivel", 0)))
        negrito = c.get("natureza") == "0"
        linha = (("   " * nivel + c["nome"], negrito),) + tuple((fmt_brl(v), negrito) for v in vals)
        if nivel == 0 or not grupos:
            grupos.append([c["nome"], [linha]])
        else:
            grupos[-1][1].append(linha)
    soltos = [g for g in grupos if len(g[1]) == 1]
    if len(soltos) > 1:
        saida, inserido = [], False
        for g in grupos:
            if len(g[1]) == 1:
                if not inserido:
                    saida.append(["Outras contas", [l for s in soltos for l in s[1]]])
                    inserido = True
            else:
                saida.append(g)
        grupos = saida
    return [(g[0], g[1]) for g in grupos]


def _segmentos_previsao(itens):
    """Quebra a lista de itens da previsao em [(titulo, [itens])]: um segmento por grupo (a linha do grupo
    vira o titulo do slide) e um segmento final 'Resumo' com as linhas fora de grupo (Total de pagamentos,
    Garantidora, Sub-Total, Total despesas). Mesmo algoritmo de segmentosPrevisao em app.js."""
    segs, resumo, atual = [], [], None
    for it in itens:
        if it["tipo"] == "grupo":
            atual = [it["nome"], []]
            segs.append(atual)
        elif atual is not None:
            atual[1].append(it)
            if it["tipo"] == "subtotal":
                atual = None
        else:
            resumo.append(it)
    out = [(n, its) for n, its in segs if its]
    if resumo:
        out.append(("Resumo", resumo))
    return out


def _titulo_plano(rotulo, grupo, sufixo):
    """'Receitas' + grupo 'Receitas Ordinarias' => 'Receitas — Ordinarias' (sem repetir o rotulo no nome do grupo);
    grupo que e so 'Receitas' => 'Receitas'. Mesmo algoritmo de tituloPlano em app.js."""
    m = re.match(r"^\s*(receitas?|despesas?)\b[\s:\-\u2013\u2014]*", grupo, re.I)
    resto = (grupo[m.end():] if m else grupo).strip()
    if not resto:
        return f"{rotulo}{sufixo}"
    resto = resto[:1].upper() + resto[1:]
    if len(resto) > 46:
        resto = resto[:45].rstrip() + "\u2026"
    return f"{rotulo} \u2014 {resto}{sufixo}"


def _cabecalho_categoria(nome, maximo=30):
    """Texto do cabecalho da 1a coluna: o nome da categoria do slide (no lugar de 'Descricao')."""
    nome = (nome or "").strip()
    return nome if len(nome) <= maximo else nome[:maximo - 1].rstrip() + "\u2026"


# ---- Paginacao e layout das tabelas, com ajuste manual do usuario ----
# O usuario pode, na previa, mudar "linhas por slide" de uma tabela (dados["ajustes"][chave]["linhas"]).
# Fonte e altura de linha saem do numero de linhas: quanto mais linhas, menor a fonte, ate o piso de 12pt.
# Mesmas regras de paginarTabela / layoutTabela em app.js.
FONTE_PISO = 12
AREA_TABELA_IN = 9.3 - 1.9
TIPOS_TABELA = {
    "meses": {"padrao": 17, "fonte": 14, "altura": 0.40, "fonte_g": 16, "altura_g": 0.55, "grande": 12},
    "previsao": {"padrao": 15, "fonte": 18, "altura": 0.44, "fonte_g": 24, "altura_g": 0.62, "grande": 12},
    "lanc": {"padrao": 15, "fonte": 18, "altura": 0.44, "fonte_g": 18, "altura_g": 0.44, "grande": 0},
}


def _slug(texto):
    t = unicodedata.normalize("NFD", str(texto or "")).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]+", "-", t).strip("-")


def _cap_linhas(tipo):
    c = TIPOS_TABELA[tipo]
    altura_min = c["altura"] * FONTE_PISO / c["fonte"]
    return max(c["padrao"], int(AREA_TABELA_IN / altura_min) - 1)


def _paginar_tabela(n, tipo, chave, ajustes):
    """[(inicio, fim), ...] das paginas de uma tabela de n linhas, mais o limite de linhas por pagina usado.
    Sem ajuste manual: padrao do tipo, e uma ultima pagina com 1-2 linhas e absorvida pelas anteriores
    (cada pagina ganha ate 2 linhas, com fonte um pouco menor). Com ajuste: o numero escolhido (ate o limite)."""
    c = TIPOS_TABELA[tipo]
    cap = _cap_linhas(tipo)
    if n <= 0:
        return [], c["padrao"]
    pedido = (ajustes or {}).get(chave, {}).get("linhas") if isinstance(ajustes, dict) else None
    manual = False
    L = c["padrao"]
    if pedido:
        try:
            L = max(1, min(int(pedido), cap))
            manual = True
        except (TypeError, ValueError):
            L = c["padrao"]
    paginas = -(-n // L)
    if not manual and paginas > 1:
        k = n - (paginas - 1) * L
        if k <= 2:
            L2 = L + -(-k // (paginas - 1))
            if L2 <= cap:
                L = L2
                paginas = -(-n // L)
    return [(i * L, min(n, (i + 1) * L)) for i in range(paginas)], L


# ---- Fusao de blocos adjacentes (usuario arrasta um slide sobre o outro na previa) ----
# dados["ajustes"][chave_do_bloco]["juntarCom"] = chave do bloco imediatamente anterior: os dois viram um slide so
# (titulo "A + B"). So vale se as linhas somadas couberem em um slide ate o piso de legibilidade. Mesmas regras de
# fundirBlocos em app.js.
SEPS_POR_BLOCO = {"meses": 0, "previsao": 1, "lanc": 1}  # linhas de titulo de bloco inseridas ao fundir


def _tam_unidade(partes, tipo):
    if len(partes) == 1:
        return len(partes[0]["itens"])
    return sum(len(p["itens"]) + SEPS_POR_BLOCO[tipo] for p in partes)


def _fundir_blocos(blocos, tipo, ajustes):
    """blocos: [{"chave", "nome", "itens", ...}] em ordem. Devolve unidades [{"chave", "nomes", "itens", "fundido", "partes"}]."""
    cap = _cap_linhas(tipo)
    unidades = []
    for k, b in enumerate(blocos):
        aj = ajustes.get(b["chave"]) if isinstance(ajustes, dict) else None
        cur = unidades[-1] if unidades else None
        if (cur and k > 0 and isinstance(aj, dict) and aj.get("juntarCom") == blocos[k - 1]["chave"]
                and _tam_unidade(cur["partes"] + [b], tipo) <= cap):
            cur["partes"].append(b)
            continue
        unidades.append({"partes": [b]})
    saida = []
    for u in unidades:
        ps = u["partes"]
        if len(ps) == 1:
            itens = list(ps[0]["itens"])
        else:
            itens = []
            for pt in ps:
                if tipo == "previsao":
                    itens.append({"tipo": "grupo", "nome": pt["nome"]})
                elif tipo == "lanc":
                    itens.append({"sep": pt["nome"]})
                itens.extend(pt["itens"])
        saida.append({"chave": ps[0]["chave"], "nomes": [pt["nome"] for pt in ps], "itens": itens, "fundido": len(ps) > 1, "partes": ps})
    return saida


def _juntar_nomes(nomes, palavra="grupos", limite=46):
    j = " + ".join(nomes)
    if len(j) <= limite:
        return j
    n = len(nomes) - 1
    return f"{nomes[0]} + {n} {palavra if n > 1 else palavra[:-1]}"


def _nome_curto_plano(grupo):
    m = re.match(r"^\s*(receitas?|despesas?)\b[\s:\-\u2013\u2014]*", grupo, re.I)
    resto = (grupo[m.end():] if m else grupo).strip()
    return (resto[:1].upper() + resto[1:]) if resto else grupo.strip()


def _layout_tabela(tipo, n_dados):
    """(fonte_pt, altura_linha_in) para uma tabela com n_dados linhas (+ cabecalho)."""
    c = TIPOS_TABELA[tipo]
    n = n_dados + 1
    grande = bool(c["grande"]) and n <= c["grande"]
    fb, hb = (c["fonte_g"], c["altura_g"]) if grande else (c["fonte"], c["altura"])
    altura = min(hb, AREA_TABELA_IN / n)
    fonte = max(FONTE_PISO, int(fb * min(1.0, altura / hb) + 0.5))
    return fonte, altura


def _fonte_tabela_meses(linhas, grande):
    """Fonte da tabela de 12 meses: a maior em que o valor mais longo cabe na coluna de 1,1" (margem lateral 0,04").
    Mesma regra de fontePlanoMeses em app.js. linhas = lista de linhas; cada celula e (texto, negrito) ou texto."""
    maxlen = 0
    for l in linhas:
        for cel in l[1:]:
            t = cel[0] if isinstance(cel, tuple) else cel
            maxlen = max(maxlen, len(str(t)))
    base = PLANO_FONTE_GRANDE if grande else PLANO_FONTE
    if maxlen > 10:
        base -= 2
    elif maxlen > 9:
        base -= 1
    return max(12, base)


# ---- Eleicao de sindico: slide de candidatos (cards por cargo) e slide de resultado ----
ELEICAO_VOTACAO_PADRAO = "A definir na assembleia"


def _tam_nomes_card(nomes, largura_in, altura_in, tamanhos=(32, 28, 26, 24, 22, 20, 18, 16)):
    """Maior fonte em que a lista de nomes cabe no card (estimativa por caracteres por linha)."""
    for tam in tamanhos:
        cpl = max(8, int((largura_in - 0.7) * 72 / (tam * 0.52)))
        linhas = sum(max(1, -(-(len(n) + 2) // cpl)) for n in nomes)
        if (linhas * tam * 1.25 + len(nomes) * 8) / 72 <= altura_in:
            return tam
    return tamanhos[-1]


def _add_card_cargo(slide, x, y, w, h, cargo, nomes, cor_header, vazio="Sem candidatos inscritos", tam_fixo=None, centralizado=False):
    """Card de um cargo: faixa colorida com o cargo no topo e a lista de nomes embaixo."""
    corpo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(w), Inches(h))
    corpo.adjustments[0] = 0.04
    corpo.fill.solid()
    corpo.fill.fore_color.rgb = CARD_BG
    corpo.line.color.rgb = BORDER
    corpo.line.width = Pt(1.25)
    corpo.shadow.inherit = False
    alt_h = 0.95
    topo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(w), Inches(alt_h))
    topo.adjustments[0] = 0.2
    topo.fill.solid()
    topo.fill.fore_color.rgb = cor_header
    topo.line.fill.background()
    topo.shadow.inherit = False
    tf = topo.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = cargo
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = PP_ALIGN.CENTER
    tb = slide.shapes.add_textbox(x + Inches(0.3), y + Inches(alt_h + 0.15), Inches(w - 0.6), Inches(h - alt_h - 0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    if not nomes:
        p = tf.paragraphs[0]
        p.text = vazio
        p.font.size = Pt(22)
        p.font.italic = True
        p.font.color.rgb = GRAY_TEXT
        p.alignment = PP_ALIGN.CENTER if centralizado else PP_ALIGN.LEFT
        return
    tam = tam_fixo or _tam_nomes_card(nomes, w, h - alt_h - 0.3)
    for i, nome in enumerate(nomes):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = nome if centralizado else "\u2022 " + nome
        p.font.size = Pt(tam)
        p.font.bold = centralizado
        p.font.color.rgb = DARK_PURPLE
        p.space_after = Pt(8)
        p.alignment = PP_ALIGN.CENTER if centralizado else PP_ALIGN.LEFT


def _lista_nomes(v):
    if isinstance(v, str):
        v = v.split("\n")
    return [str(x).strip() for x in (v or []) if str(x).strip()]


def _dados_eleicao(el):
    """Normaliza o dict de eleicao; devolve None se nao houver candidatos nem resultado (cai no texto livre)."""
    if not isinstance(el, dict):
        return None
    d = {
        "cand": [("Síndico", _lista_nomes(el.get("candSindico"))), ("Subsíndico", _lista_nomes(el.get("candSubsindico"))), ("Conselho fiscal", _lista_nomes(el.get("candConselho")))],
        "res": [("Síndico eleito", _lista_nomes(el.get("eleitoSindico"))), ("Subsíndico eleito", _lista_nomes(el.get("eleitoSubsindico"))), ("Conselho fiscal eleito", _lista_nomes(el.get("eleitosConselho")))],
        "votacao": (el.get("votacao") or "").strip(),
        "mandato": (el.get("mandato") or "").strip(),
        "obs": _lista_nomes(el.get("obs")),
    }
    d["cand"] = [c for c in d["cand"] if c[1]]
    d["res"] = [r for r in d["res"] if r[1]]
    if not d["cand"] and not d["res"]:
        return None
    return d


def _slides_eleicao(novo_conteudo_fn, d):
    """Slide 1: candidatos por cargo (cards lado a lado). Slide 2: resultado (cards com os eleitos) + mandato e observacoes."""
    votacao = d["votacao"] if d["votacao"] and d["votacao"] != ELEICAO_VOTACAO_PADRAO else ""
    obs_no_fim = d["obs"]
    if d["cand"]:
        s1 = novo_conteudo_fn("Eleição \u2014 Candidatos")
        k = len(d["cand"])
        gap = 0.4
        w = (17.4 - gap * (k - 1)) / k
        altura = 6.0 if (votacao or (not d["res"] and (d["mandato"] or obs_no_fim))) else 6.9
        for i, (cargo, nomes) in enumerate(d["cand"]):
            _add_card_cargo(s1, Inches(1.3 + i * (w + gap)), Inches(1.9), w, altura, f"{cargo} ({len(nomes)})", nomes, PURPLE)
        y = 8.15
        if votacao:
            add_text_line(s1, Inches(1.3), Inches(y), 17.4, 0.6, "Forma de votação: " + votacao, 26, bold=True, color=DARK_PURPLE)
        if not d["res"]:
            extras = ([f"Mandato: {d['mandato']}"] if d["mandato"] else []) + obs_no_fim
            if extras:
                add_text_line(s1, Inches(1.3), Inches(y + (0.65 if votacao else 0)), 17.4, 0.6, "   |   ".join(extras)[:150], 22, color=GRAY_TEXT)
    if d["res"]:
        s2 = novo_conteudo_fn("Eleição \u2014 Resultado")
        k = len(d["res"])
        gap = 0.4
        w = (17.4 - gap * (k - 1)) / k
        for i, (cargo, nomes) in enumerate(d["res"]):
            _add_card_cargo(s2, Inches(1.3 + i * (w + gap)), Inches(1.9), w, 4.4, cargo, nomes, GOLD, tam_fixo=_tam_nomes_card(nomes, w, 4.4 - 0.95 - 0.3, (36, 32, 28, 26, 24, 22, 20)), centralizado=True)
        linhas = ([f"Forma de votação: {votacao}"] if votacao and not d["cand"] else []) + ([f"Mandato: {d['mandato']}"] if d["mandato"] else [])
        y = 6.7
        for ln in linhas:
            add_text_line(s2, Inches(1.3), Inches(y), 17.4, 0.6, ln, 26, bold=True, color=DARK_PURPLE)
            y += 0.7
        for ln in obs_no_fim:
            add_text_line(s2, Inches(1.3), Inches(y), 17.4, 0.55, ln, 22, color=GRAY_TEXT)
            y += 0.55


def _linhas_resumo(a11, offset, n_meses, so_grupos=False):
    """Linhas do resumo 'Prestacao de contas': saldo anterior, TODAS as contas de receita do plano, total de receitas,
    TODAS as de despesa, total de despesas e saldo final. [(nome, [12 valores], negrito, nivel_recuo)].
    so_grupos=True: so os grupos de PRIMEIRO NIVEL (filhos diretos de Receitas/Despesas), sem entrar nos subniveis.
    Regra geral, sem nomes fixos: se a secao tem uma unica raiz (nivel 0, natureza 0), o 1o nivel e o 1; senao, o menor nivel.
    Mesma regra de linhasResumo em app.js."""
    def fat(arr):
        return list(arr[offset:]) if arr else [None] * n_meses
    contas = a11.get("contasPlano") or []
    linhas = [("Saldo anterior", fat(a11.get("saldoAnterior")), False, 0, False)]
    sep = True  # a proxima linha abre um novo bloco (Receitas | Despesas | Saldo final)
    for secao, chave, rotulo in (("receitas", "totalReceitas", "Total de Receitas"), ("despesas", "totalDespesas", "Total de Despesas")):
        da_secao = [c for c in contas if c.get("secao") == secao]
        niveis = [int(c.get("nivel", 0) or 0) for c in da_secao]
        raizes = [c for c in da_secao if int(c.get("nivel", 0) or 0) == 0]
        nivel_alvo = 1 if (len(raizes) == 1 and raizes[0].get("natureza") == "0" and 1 in niveis) else (min(niveis) if niveis else 0)
        for c in da_secao:
            vals = [v or 0 for v in (c.get("valores") or [])[offset:]]
            vals = (vals + [0] * n_meses)[:n_meses]
            niv = int(c.get("nivel", 0) or 0)
            if any(vals) and (not so_grupos or niv == nivel_alvo):
                linhas.append((c["nome"], vals, c.get("natureza") == "0", 1 if so_grupos else 1 + max(0, niv), sep))
                sep = False
        linhas.append((rotulo, fat(a11.get(chave)), True, 0, sep))
        sep = True
    linhas.append(("Saldo Final", fat(a11.get("saldoFinal")), True, 0, True))
    return linhas


def gerar_apresentacao_bytes(dados, template_path="modelo.pptx"):
    prs = Presentation(template_path)

    slide1 = prs.slides[0]
    for shp in iter_all_shapes(slide1.shapes):
        if shp.has_text_frame and shp.text_frame.text.strip() == "Assembleia":
            set_text(shp, dados["tipo"])
            shp.height = Inches(1.5)
            shp.text_frame.word_wrap = True
        elif shp.has_text_frame and "Bosque Primavera" in shp.text_frame.text:
            set_text(shp, dados["condominio"])
            shp.top = shp.top + Inches(1.05)

    divisor_template_idx = 1
    conteudo_template_idx = 2

    pauta = dados["pauta"]

    def novo_divisor(texto):
        s = duplicate_slide(prs, divisor_template_idx, text_replacements={"TextBox 12": texto})
        for shp in iter_all_shapes(s.shapes):
            if shp.name == "TextBox 12" and shp.has_text_frame:
                tf = shp.text_frame
                tf.word_wrap = True

                # A caixa original (~6.3") e a fonte (~83pt) foram feitas sob medida pra
                # "Multiplos beneficios". Para textos de pauta reais (mais longos, e com
                # palavras isoladas compridas tipo "orcamentaria"), alargamos a caixa e
                # calculamos a fonte considerando tambem a MAIOR PALAVRA isolada - e
                # justamente a combinacao palavra-longa + fonte-grande-demais-pra-caixa
                # que causava a quebra no meio da palavra ("orcamentari" / "a").
                shp.left = Inches((20 - DIVISOR_CAIXA_W) / 2)
                shp.width = Inches(DIVISOR_CAIXA_W)
                shp.height = Inches(DIVISOR_CAIXA_H)
                shp.top = Inches((11.25 - DIVISOR_CAIXA_H) / 2)  # centralizado na vertical no slide
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE

                # Fonte unica em todos os divisores (titulos longos apenas quebram em mais linhas); igual a DIVISOR_PT em app.js
                tamanho = _tam_divisor(texto)  # auto-fit: diminui a fonte ate caber (igual a tamDivisor em app.js)

                for p in tf.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
                    p.line_spacing = 1.0
                    for r in p.runs:
                        r.font.size = Pt(tamanho)

                # Rede de seguranca: se mesmo assim nao couber, o PowerPoint encolhe a
                # fonte automaticamente em vez de quebrar a palavra de forma feia.
                try:
                    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                except Exception:
                    pass
        return s

    def novo_conteudo(titulo):
        s = duplicate_slide(prs, conteudo_template_idx)
        add_title(s, titulo)
        return s

    def slides_pauta():
        # Na lista "Ordem do dia" o item de Assuntos aparece abreviado; o texto guardado na pauta nao muda.
        paginas = paginar_pauta([TEXTO_PAUTA_ASSUNTOS if it.get("role") == "assuntos" else it["texto"] for it in pauta])
        for n, itens_pagina in enumerate(paginas, start=1):
            if n == 1:
                s = novo_conteudo("Itens a serem discutidos")
                add_text_line(s, Inches(1.3), Inches(1.4), 15.5, 0.55,
                              f"Edital de Convocação — {dados['tipo']}", 24, bold=True, color=DARK_PURPLE)
                for k, linha in enumerate(linhas_edital(dados)):
                    add_text_line(s, Inches(1.3), Inches(1.95 + 0.42 * k), 15.5, 0.45, linha, 18, color=GRAY_TEXT)
                add_text_line(s, Inches(1.3), Inches(2.85), 15.5, 0.55,
                              "Ordem do dia:", 22, bold=True, color=DARK_PURPLE)
                y_lista = Inches(3.5)
            else:
                s = novo_conteudo("Ordem do dia")
                y_lista = Inches(1.9)
            add_numbered_list(s, Inches(1.3), y_lista, itens_pagina,
                              font_size=tamanho_fonte_pauta([t for _, t in itens_pagina], n == 1))

    if pauta:
        slides_pauta()

    for idx, item in enumerate(pauta, start=1):
        role = item.get("role")
        if role != "assuntos":  # "Assuntos Gerais" tem slide unico (ja numerado), ver o branch abaixo
            novo_divisor(f"{idx}. {item['texto']}")

        if role == "prestacao":
            # Inadimplencia e Resumo Financeiro (16B) nao dependem do 11A; so os slides do 11A precisam dele.
            a11 = dados.get("a11")
            if a11:
                meses = a11["months"][-12:]
                offset = len(a11["months"]) - len(meses)
                periodo_11a = _periodo_meses(meses)

                def fatiar(arr):
                    return (arr[offset:] if arr else [None] * len(meses))

                if a11.get("contasPlano"):
                    col_w_res = [PLANO_COL_DESC] + [(17.4 - PLANO_COL_DESC) / len(meses)] * len(meses)
                    # 1o o resumo so com grupos/subgrupos, depois o detalhamento completo (com as contas-folha)
                    for so_grupos, titulo_res in ((True, "Prestação de contas — Resumo"), (False, "Prestação de contas")):
                        linhas_res = _linhas_resumo(a11, offset, len(meses), so_grupos)
                        chave_res = "resumo-grupos" if so_grupos else "resumo"
                        faixas_res, _l = _paginar_tabela(len(linhas_res), "meses", chave_res, dados.get("ajustes"))
                        pags_res = len(faixas_res)
                        for pr, (i0, i1) in enumerate(faixas_res):
                            lote_r = linhas_res[i0:i1]
                            rows_r = [[(("   " * niv) + nome, neg)] + [(fmt_brl(v), neg) for v in vals] for nome, vals, neg, niv, _sp in lote_r]
                            seps_r = {i for i, l in enumerate(lote_r) if l[4]}
                            grande_r = len(lote_r) + 1 <= TIPOS_TABELA["meses"]["grande"]
                            fonte_r, altura_r = _layout_tabela("meses", len(lote_r))
                            sr = novo_conteudo(titulo_res + (f" ({pr+1}/{pags_res})" if pags_res > 1 else ""))
                            add_table(sr, Inches(1.3), ["Categoria"] + meses, rows_r, col_w_res, font_size=min(fonte_r, _fonte_tabela_meses(rows_r, grande_r)),
                                      height_in=altura_r, y=Inches(1.9), margin_v=0.04, margin_h=0.04, sep_rows=seps_r)
                            add_periodo(sr, periodo_11a)
                else:
                    s = novo_conteudo("Prestação de contas")
                    add_periodo(s, periodo_11a)
                    tbl_rows = [
                        [("Saldo anterior", False)] + [fmt_brl(v) for v in fatiar(a11["saldoAnterior"])],
                        [("Total de Receitas", True)] + [fmt_brl(v) for v in fatiar(a11["totalReceitas"])],
                        [("Total de Despesas", True)] + [fmt_brl(v) for v in fatiar(a11["totalDespesas"])],
                        [("Saldo Final", True)] + [fmt_brl(v) for v in fatiar(a11["saldoFinal"])],
                    ]
                    add_table(s, Inches(1.3), [""] + meses, tbl_rows, [2.6] + [(17.4 - 2.6) / len(meses)] * len(meses), font_size=13, height_in=0.8, y=Inches(3.4))

                rec_v = [v or 0 for v in fatiar(a11["totalReceitas"])]
                desp_v = [v or 0 for v in fatiar(a11["totalDespesas"])]
                ref_max = max(rec_v + desp_v + [0])
                n_pag_barras = (len(meses) + BARRAS_MESES_POR_SLIDE - 1) // BARRAS_MESES_POR_SLIDE
                for pb in range(n_pag_barras):
                    i0, i1 = pb * BARRAS_MESES_POR_SLIDE, (pb + 1) * BARRAS_MESES_POR_SLIDE
                    suf_b = f" ({pb+1}/{n_pag_barras})" if n_pag_barras > 1 else ""
                    s2 = novo_conteudo("Prestação de contas — Receitas x Despesas" + suf_b)
                    add_periodo(s2, periodo_11a)
                    add_bar_chart(s2, Inches(1.8), 16.4, meses[i0:i1], [
                        ("Receitas", rec_v[i0:i1], BLUE),
                        ("Despesas", desp_v[i0:i1], RED),
                    ], y=Inches(2.6), height=Inches(6.8), ref_max=ref_max)

                def _slides_linha(titulo, nome_serie, valores):
                    vals = [v or 0 for v in valores]
                    ref = max(vals + [0])
                    n_pag = (len(meses) + LINHA_MESES_POR_SLIDE - 1) // LINHA_MESES_POR_SLIDE
                    for pl in range(n_pag):
                        i0, i1 = pl * LINHA_MESES_POR_SLIDE, (pl + 1) * LINHA_MESES_POR_SLIDE
                        suf = f" ({pl+1}/{n_pag})" if n_pag > 1 else ""
                        sl_ = novo_conteudo(titulo + suf)
                        add_periodo(sl_, periodo_11a)
                        add_line_chart(sl_, Inches(1.8), 16.4, meses[i0:i1], nome_serie, vals[i0:i1], PURPLE, y=Inches(2.6), height=Inches(6.8), ref_max=ref)

                if a11.get("energiaEletrica"):
                    _slides_linha("Despesas com Energia Elétrica", "Energia Elétrica", fatiar(a11["energiaEletrica"]))

                if a11.get("aguaEsgotoDespesa"):
                    _slides_linha("Despesas com Água / Esgoto", "Água / Esgoto", fatiar(a11["aguaEsgotoDespesa"]))

            inad = dados.get("inadimplencia")
            if inad:
                s5 = novo_conteudo("Inadimplência")
                if inad.get("posicao"):
                    add_periodo(s5, "Posição em " + inad["posicao"])

                # % de inadimplencia = unidades inadimplentes / total de unidades x 100 (calculado aqui;
                # se nao houver o total de unidades, cai no % total digitado manualmente)
                n_inad = _int_ou_none(inad.get("unidades"))
                n_total = _int_ou_none(inad.get("unidades_total"))
                if n_inad is not None and n_total:
                    pct_total = f"{n_inad / n_total * 100:.2f}".replace(".", ",")
                    unidades_txt = f"{n_inad} de {n_total}"
                else:
                    pct_total = (inad.get("total") or "").strip() or None
                    unidades_txt = str(n_inad) if n_inad is not None else "-"

                cards = [
                    ("Acumulado", ("R$ " + inad["valor"]) if inad.get("valor") else "-"),
                    ("Cobranças", inad.get("cobrancas") or "-"),
                    ("Unidades inadimplentes", unidades_txt),
                    ("% de inadimplência", (pct_total + "%") if pct_total else "-"),
                ]
                add_stat_cards(s5, Inches(1.3), Inches(1.6), cards, width_total_in=17.4, height_in=2.2)

                rows_inad = [
                    [("% Inadimplência orgânica", True), ((inad.get("organica") or "") + "%" if inad.get("organica") else "-", True)],
                    [("% Inadimplência garantidora", True), ((inad.get("garantidora") or "") + "%" if inad.get("garantidora") else "-", True)],
                    [("% Inadimplência total (unidades inadimplentes ÷ total de unidades)", True), ((pct_total + "%") if pct_total else "-", True)],
                ]
                add_table(s5, Inches(1.3), ["Indicador", "Valor"], [[(r[0][0], r[0][1]), (r[1][0], r[1][1])] for r in rows_inad], [13.4, 4.0], font_size=18, height_in=0.7, y=Inches(4.6))

            b16 = dados.get("b16")
            if b16:
                periodo_16 = _periodo_16b(dados.get("periodo16b", ""))
                contas16 = b16["contas"]
                headers16 = ["Conta", "Saldo anterior", "Créditos", "Débitos", "Saldo final"]
                rows16 = [[c["nome"], fmt_brl(c["saldoAnt"]), fmt_brl(c["creditos"]), fmt_brl(c["debitos"]), (fmt_brl(c["saldoFinal"]), True)] for c in contas16]
                if len(contas16) > 1:
                    def _soma(campo):
                        return sum((c.get(campo) or 0) for c in contas16)
                    rows16.append([("Total", True), (fmt_brl(_soma("saldoAnt")), True), (fmt_brl(_soma("creditos")), True),
                                   (fmt_brl(_soma("debitos")), True), (fmt_brl(_soma("saldoFinal")), True)])
                # Contas e grupos de saldo na MESMA tabela/slide: contas, (Total), linha de titulo "Grupos de saldo", grupos, Saldo final
                grupos16 = b16.get("grupos") or []
                saldo_geral = b16.get("saldoFinalGeral")
                if grupos16 or saldo_geral is not None:
                    rows16.append([("Grupos de saldo", True), "", "", "", ""])
                    rows16 += [[g["nome"], fmt_brl(g["saldoAnt"]), fmt_brl(g["creditos"]), fmt_brl(g["debitos"]), (fmt_brl(g["saldoFinal"]), True)] for g in grupos16]
                    # Saldo final = mesma logica da linha Total das contas: soma de cada coluna dos grupos (o saldo final do documento, se houver, prevalece)
                    def _soma_g(campo):
                        return sum((g.get(campo) or 0) for g in grupos16)
                    if grupos16 or saldo_geral is not None:
                        sf_final = saldo_geral if saldo_geral is not None else _soma_g("saldoFinal")
                        rows16.append([("Saldo final", True), (fmt_brl(_soma_g("saldoAnt")), True), (fmt_brl(_soma_g("creditos")), True),
                                       (fmt_brl(_soma_g("debitos")), True), (fmt_brl(sf_final), True)])
                POR_SLIDE_16B = 14
                n_pags_16 = max(1, -(-len(rows16) // POR_SLIDE_16B))
                for pg16 in range(n_pags_16):
                    lote16 = rows16[pg16 * POR_SLIDE_16B:(pg16 + 1) * POR_SLIDE_16B]
                    s6 = novo_conteudo("Resumo Financeiro" + (f" ({pg16+1}/{n_pags_16})" if n_pags_16 > 1 else ""))
                    add_periodo(s6, periodo_16)
                    f16, h16 = _dim_tabela(len(lote16) + 1, 2.2, fonte_base=24, altura_base=0.85)
                    add_table(s6, Inches(1.3), headers16, lote16, [6.0, 2.85, 2.85, 2.85, 2.85], font_size=f16, height_in=h16, y=Inches(2.2), margin_v=0.05)

                LINHAS = LANCAMENTOS_LINHAS_POR_SLIDE
                MAX_LANCAMENTOS_POR_CONTA = 150
                blocos_m = []
                for conta in b16["contas"]:
                    movs_completos = conta.get("movimentos") or []
                    if not movs_completos:
                        continue
                    blocos_m.append({"chave": f"lanc:{_slug(conta['nome'])}", "nome": conta["nome"], "itens": movs_completos[:MAX_LANCAMENTOS_POR_CONTA],
                                     "excedente": max(0, len(movs_completos) - MAX_LANCAMENTOS_POR_CONTA)})
                for un in _fundir_blocos(blocos_m, "lanc", dados.get("ajustes")):
                    movs = un["itens"]
                    faixas_m, _l = _paginar_tabela(len(movs), "lanc", un["chave"], dados.get("ajustes"))
                    paginas = len(faixas_m)
                    sm = None
                    for p, (i0, i1) in enumerate(faixas_m):
                        lote = movs[i0:i1]
                        fonte_m, altura_m = _layout_tabela("lanc", len(lote))
                        sufixo = f" ({p+1}/{paginas})" if paginas > 1 else ""
                        titulo_m = f"Lançamentos — {_juntar_nomes(un['nomes'], 'contas')}{sufixo}"
                        sm = novo_conteudo(titulo_m)
                        add_periodo(sm, periodo_16)
                        headersm = ["Data", "Descrição", "Tipo", "Valor"]
                        rowsm = [[(m["sep"], True), "", "", ""] if "sep" in m else
                                 [m["data"], _cortar(m["descricao"], LANCAMENTOS_DESCRICAO_MAX), "Crédito" if m["tipo"] == "C" else "Débito", fmt_brl(m["valor"])] for m in lote]
                        add_table(sm, Inches(1.3), headersm, rowsm, [2.0, 9.4, 2.7, 3.3], font_size=fonte_m, height_in=altura_m, left_cols={0, 1}, y=Inches(1.9), margin_v=0.04)
                    exced = [(pt["nome"], pt["excedente"]) for pt in un["partes"] if pt.get("excedente")]
                    if exced and sm is not None:
                        # nota logo abaixo da tabela da ultima pagina, dentro das margens
                        y_nota = Inches(1.9 + altura_m * (len(lote) + 1) + 0.25)
                        if len(un["partes"]) == 1:
                            nota_txt = [f"{exced[0][1]} lançamento(s) adicional(is) dessa conta não exibidos aqui por limite de tamanho — consulte o extrato completo no sistema."]
                        else:
                            nota_txt = [f"{n} lançamento(s) adicional(is) de {nm} não exibidos aqui por limite de tamanho — consulte o extrato completo no sistema." for nm, n in exced]
                        add_bullets(sm, Inches(1.3), nota_txt, y=y_nota, width_in=17.4, height_in=0.5, font_size=14)

        elif role == "previsao" and dados.get("previsao"):
            prev = dados["previsao"]
            itens = prev.get("itensCompletos") or [{"tipo": "subtotal", "nome": c["nome"], "atualRaw": c["atual"], "previstoRaw": c["previsto"]} for c in prev["categories"]]
            LINHAS = PREVISAO_LINHAS_POR_SLIDE
            colunas = [str(c) for c in (prev.get("colunasValores") or ["Atual", "Previsto"])]
            n_cols = len(colunas)

            def _txt_po(v):
                if isinstance(v, str) and "consumo" in v.lower():
                    return "Conf. Consumo"
                return fmt_brl(v if isinstance(v, (int, float)) else None)

            def _valores_po(it):
                v = list(it["valores"]) if it.get("valores") is not None else [it.get("atualRaw"), it.get("previstoRaw")]
                return (v + [None] * n_cols)[:n_cols]

            # colunas de valor dividem a largura; a descricao fica com o resto dos 17,4"
            w_val = max(2.0, min(3.5, 11.4 / n_cols))
            larguras = [17.4 - w_val * n_cols] + [w_val] * n_cols

            blocos_pv = [{"chave": f"previsao:{_slug(n)}", "nome": n, "itens": its} for n, its in _segmentos_previsao(itens)]
            for un in _fundir_blocos(blocos_pv, "previsao", dados.get("ajustes")):
                nome_seg, itens_seg = un["nomes"][0], un["itens"]
                faixas_p, _l = _paginar_tabela(len(itens_seg), "previsao", un["chave"], dados.get("ajustes"))
                paginas = len(faixas_p)
                for p, (i0, i1) in enumerate(faixas_p):
                    lote = itens_seg[i0:i1]
                    sufixo = f" ({p+1}/{paginas})" if paginas > 1 else ""
                    if un["fundido"]:
                        titulo_pv = f"Previsão — {_juntar_nomes(un['nomes'])}{sufixo}"
                    else:
                        titulo_pv = f"Previsão — {nome_seg}{sufixo}" if len(nome_seg) <= 40 else f"Previsão — {nome_seg[:39].rstrip()}…{sufixo}"
                    s = novo_conteudo(titulo_pv)
                    rows = []
                    for it in lote:
                        bold = it["tipo"] == "subtotal"
                        if it["tipo"] == "grupo":
                            rows.append([(it["nome"], True)] + [""] * n_cols)
                        else:
                            rows.append([(it["nome"], bold)] + [(_txt_po(v), bold) for v in _valores_po(it)])
                    fonte_p, altura_p = _layout_tabela("previsao", len(lote))
                    add_table(s, Inches(1.3), ["Descrição" if un["fundido"] else _cabecalho_categoria(nome_seg)] + colunas, rows, larguras,
                              font_size=fonte_p, height_in=altura_p, y=Inches(1.9), margin_v=0.04)

            rateio_atual = prev.get("rateioAtual")
            rateio_previsto = prev.get("rateioPrevisto")
            if rateio_atual or rateio_previsto:
                s_rateio = novo_conteudo("Previsão orçamentária — Rateio")

                # Layout que preenche o slide: duas tabelas grandes lado a lado + cartoes de destaque embaixo.
                blocos = [b for b in (rateio_atual, rateio_previsto) if b and b.get("itens")]
                n_linhas = (max(len(b["itens"]) for b in blocos) + 2) if blocos else 0  # cabecalho + itens + TOTAL
                fonte_r, altura_r = _dim_tabela(n_linhas, 2.4, fim_in=7.6, fonte_base=18, altura_base=0.62)

                def _add_rateio_table(slide_r, x, bloco, titulo):
                    if not bloco or not bloco.get("itens"):
                        return
                    tb = slide_r.shapes.add_textbox(x, Inches(1.55), Inches(8.5), Inches(0.6))
                    tf = tb.text_frame
                    tf.text = titulo
                    tf.paragraphs[0].font.size = Pt(22)
                    tf.paragraphs[0].font.bold = True
                    tf.paragraphs[0].font.color.rgb = DARK_PURPLE
                    rows = []
                    for it in bloco["itens"]:
                        recibo = "Conf. Consumo" if it.get("recibo") is None else fmt_brl(it.get("recibo"))
                        total_r = "Conf. Consumo" if it.get("total") is None else fmt_brl(it.get("total"))
                        rows.append([it.get("nome", ""), recibo, total_r])
                    rows.append([("TOTAL", True), "", (fmt_brl(bloco.get("total")), True)])
                    add_table(slide_r, x, ["Fração", "Recibo", "Total Rateio"], rows, [3.3, 2.6, 2.6], font_size=fonte_r, height_in=altura_r, y=Inches(2.4))

                _add_rateio_table(s_rateio, Inches(1.3), rateio_atual, "Rateio atual")
                _add_rateio_table(s_rateio, Inches(10.2), rateio_previsto, "Previsão (novo rateio)")

                destaques = []
                for bloco in (rateio_atual, rateio_previsto):
                    if bloco and bloco.get("reajuste") is not None:
                        if bloco.get("reajusteIsPercent"):
                            destaques.append(("Reajuste do rateio", f"{bloco['reajuste']*100:.2f}".replace(".", ",") + "%"))
                        else:
                            destaques.append(("Diferença do rateio", "R$ " + fmt_brl(bloco["reajuste"])))
                if prev.get("totalDespesasAtual") is not None or prev.get("totalDespesasPrevisto") is not None:
                    destaques.append(("Total de despesas (atual \u2192 previsto)", fmt_brl(prev.get("totalDespesasAtual")) + " \u2192 " + fmt_brl(prev.get("totalDespesasPrevisto"))))
                if destaques:
                    add_stat_cards(s_rateio, Inches(1.3), Inches(7.9), destaques, width_total_in=17.4, height_in=1.5, gap_in=0.4)

        elif role == "assuntos":
            # Slide unico no estilo divisor: numero do item + titulo fixo "Assuntos Gerais" (sem bullets);
            # nem o texto da pauta nem conteudo vindo do front aparecem aqui.
            novo_divisor(f"{idx}. Assuntos Gerais")

        elif role in (None, "topico"):
            conteudo = (item.get("conteudo") or "").strip()
            dados_el = _dados_eleicao(item.get("eleicao"))
            if dados_el:
                _slides_eleicao(novo_conteudo, dados_el)
            elif conteudo:
                paginas_b = paginar_bullets([l.strip() for l in conteudo.split("\n") if l.strip()])
                for pb, (tam_b, itens_b) in enumerate(paginas_b, start=1):
                    sufixo_b = f" ({pb}/{len(paginas_b)})" if len(paginas_b) > 1 else ""
                    s = novo_conteudo(item["texto"] + sufixo_b)
                    add_bullets(s, Inches(1.3), itens_b, y=Inches(1.9), height_in=BULLETS_ALTURA_IN, width_in=17.4, font_size=tam_b)

    slide4 = prs.slides[3]
    canais = dados.get("canais", {})
    for shp in iter_all_shapes(slide4.shapes):
        if not shp.has_text_frame:
            continue
        t = shp.text_frame.text
        if "97039" in t:
            set_text(shp, canais.get("whats", t))
            _ajustar_fonte_contato(shp, 15)
            shp.width = Inches(9.0)
        elif "3596-6640" in t:
            set_text(shp, canais.get("tel", t))
            _ajustar_fonte_contato(shp, 15)
            shp.width = Inches(9.0)
        elif "relacionamento@" in t:
            set_text(shp, canais.get("email", t))
            _ajustar_fonte_contato(shp, 31)
            shp.width = Inches(9.0)
        elif "jetroadministradora.com" in t:
            set_text(shp, canais.get("site", t))
            _ajustar_fonte_contato(shp, 26)
            shp.width = Inches(9.0)
        elif "Entre em" in t.replace(chr(10), " ") and "contato" in t:
            # Titulo final "Obrigado pela atenção" (60pt, 2 linhas) + subtitulo "Entre em contato" (menor, abaixo)
            titulo_final = ["Agradecemos a todos", "pela presença!"]
            paras = shp.text_frame.paragraphs
            for i, p in enumerate(paras):
                if i < len(titulo_final) and p.runs:
                    p.runs[0].text = titulo_final[i]
                    for r in p.runs[1:]:
                        r._r.getparent().remove(r._r)
                    p.runs[0].font.size = Pt(44)
                    p.line_spacing = Pt(60)
            sub = copy.deepcopy(paras[1]._p)
            paras[1]._p.addnext(sub)
            p_sub = shp.text_frame.paragraphs[2]
            p_sub.runs[0].text = "Nossos canais de atendimento"
            p_sub.runs[0].font.size = Pt(28)
            p_sub.line_spacing = Pt(40)
            p_sub.space_before = Pt(10)
            shp.width = Inches(8.6)
            shp.height = Inches(4.2)
        elif "Barra da Tijuca" in t:
            endereco = canais.get("endereco", "")
            partes = [p.strip() for p in endereco.replace(" - ", ",").split(",") if p.strip()]
            tf = shp.text_frame
            paras = tf.paragraphs
            for i, p in enumerate(paras):
                novo = partes[i] if i < len(partes) else ""
                if p.runs:
                    p.runs[0].text = novo
                    for r in p.runs[1:]:
                        r._r.getparent().remove(r._r)

    xml_slides = prs.slides._sldIdLst
    slide4_elm = list(xml_slides)[3]
    xml_slides.remove(slide4_elm)
    xml_slides.append(slide4_elm)

    for idx in sorted([1, 2], reverse=True):
        elm = list(xml_slides)[idx]
        xml_slides.remove(elm)

    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()


def gerar_apresentacao(dados, template_path="modelo.pptx", output_path="saida.pptx"):
    data = gerar_apresentacao_bytes(dados, template_path)
    with open(output_path, "wb") as f:
        f.write(data)
    return output_path