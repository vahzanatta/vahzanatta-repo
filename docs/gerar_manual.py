from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

OUTPUT = os.path.join(os.path.dirname(__file__), "Manual_ModuloPendencias.pdf")

doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=A4,
    leftMargin=2.5*cm,
    rightMargin=2.5*cm,
    topMargin=2.5*cm,
    bottomMargin=2.5*cm,
    title="Manual de Uso — ModuloPendencias",
    author="Vahzanatta",
)

W = A4[0] - 5*cm  # usable width

# ── Estilos ────────────────────────────────────────────────────────────────
base = getSampleStyleSheet()

titulo = ParagraphStyle(
    "titulo",
    parent=base["Title"],
    fontSize=22,
    textColor=colors.HexColor("#1F3864"),
    spaceAfter=6,
    alignment=TA_CENTER,
)
subtitulo = ParagraphStyle(
    "subtitulo",
    parent=base["Normal"],
    fontSize=11,
    textColor=colors.HexColor("#5A5A5A"),
    spaceAfter=20,
    alignment=TA_CENTER,
)
h1 = ParagraphStyle(
    "h1",
    parent=base["Heading1"],
    fontSize=14,
    textColor=colors.HexColor("#1F3864"),
    spaceBefore=18,
    spaceAfter=6,
    borderPad=0,
)
h2 = ParagraphStyle(
    "h2",
    parent=base["Heading2"],
    fontSize=11,
    textColor=colors.HexColor("#2E5FA3"),
    spaceBefore=12,
    spaceAfter=4,
)
corpo = ParagraphStyle(
    "corpo",
    parent=base["Normal"],
    fontSize=10,
    leading=15,
    spaceAfter=6,
    alignment=TA_JUSTIFY,
)
nota = ParagraphStyle(
    "nota",
    parent=base["Normal"],
    fontSize=9,
    leading=13,
    textColor=colors.HexColor("#5A5A5A"),
    spaceAfter=4,
    leftIndent=12,
)
code_style = ParagraphStyle(
    "code_style",
    parent=base["Normal"],
    fontName="Courier",
    fontSize=9,
    leading=13,
    backColor=colors.HexColor("#F4F4F4"),
    textColor=colors.HexColor("#333333"),
    leftIndent=8,
    rightIndent=8,
    spaceAfter=4,
    spaceBefore=4,
    borderPad=4,
)
aviso = ParagraphStyle(
    "aviso",
    parent=base["Normal"],
    fontSize=9,
    leading=13,
    textColor=colors.HexColor("#7B3F00"),
    backColor=colors.HexColor("#FFF3CD"),
    leftIndent=8,
    rightIndent=8,
    spaceAfter=6,
    spaceBefore=4,
    borderPad=6,
)
bullet = ParagraphStyle(
    "bullet",
    parent=base["Normal"],
    fontSize=10,
    leading=15,
    leftIndent=16,
    spaceAfter=3,
)

def hr():
    return HRFlowable(width="100%", thickness=0.5,
                      color=colors.HexColor("#CCCCCC"), spaceAfter=4)

def sp(n=6):
    return Spacer(1, n)

# ── Estilos de tabela ──────────────────────────────────────────────────────
HEADER_BG  = colors.HexColor("#1F3864")
ROW_ALT    = colors.HexColor("#EBF0F8")
ROW_WHITE  = colors.white
BORDER_CLR = colors.HexColor("#AAAAAA")

def table_style(n_data_rows):
    cmds = [
        ("BACKGROUND",  (0, 0), (-1,  0),  HEADER_BG),
        ("TEXTCOLOR",   (0, 0), (-1,  0),  colors.white),
        ("FONTNAME",    (0, 0), (-1,  0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1,  0),  9),
        ("ALIGN",       (0, 0), (-1, -1),  "LEFT"),
        ("VALIGN",      (0, 0), (-1, -1),  "MIDDLE"),
        ("FONTNAME",    (0, 1), (-1, -1),  "Helvetica"),
        ("FONTSIZE",    (0, 1), (-1, -1),  9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [ROW_WHITE, ROW_ALT]),
        ("GRID",        (0, 0), (-1, -1),  0.4, BORDER_CLR),
        ("TOPPADDING",  (0, 0), (-1, -1),  4),
        ("BOTTOMPADDING",(0,0), (-1, -1),  4),
        ("LEFTPADDING", (0, 0), (-1, -1),  6),
        ("RIGHTPADDING",(0, 0), (-1, -1),  6),
    ]
    return TableStyle(cmds)

def make_table(headers, rows, col_widths):
    data = [headers] + rows
    t = Table(data, colWidths=col_widths)
    t.setStyle(table_style(len(rows)))
    return t

# ── Conteúdo ───────────────────────────────────────────────────────────────
story = []

# Capa
story += [
    sp(30),
    Paragraph("Manual de Uso", titulo),
    Paragraph("Módulo VBA — Relatório de Pendências SAP", subtitulo),
    hr(),
    sp(4),
    Paragraph(
        "Este documento explica como importar e utilizar o módulo <b>ModuloPendencias</b> "
        "no Excel, e como preencher corretamente a tabela de regras (aba <b>BASE</b>) "
        "para identificação automática das áreas responsáveis por partidas em aberto "
        "exportadas do SAP (transação FBL3H).",
        corpo,
    ),
    sp(30),
]

# 1 — Visão geral
story += [
    Paragraph("1. Visão Geral", h1),
    hr(),
    Paragraph(
        "O módulo processa um relatório de partidas em aberto do SAP e, para cada linha, "
        "identifica a <b>Área Responsável</b> pela contabilização com base em regras "
        "definidas pelo usuário. Ao final, três colunas são adicionadas automaticamente "
        "ao relatório:",
        corpo,
    ),
    Paragraph("• <b>Data Base</b> — data em que o processamento foi executado.", bullet),
    Paragraph("• <b>Dias em Aberto</b> — diferença em dias entre a Data Base e a Data de Lançamento.", bullet),
    Paragraph("• <b>Área Responsável</b> — área identificada pelas regras da tabela BASE.", bullet),
    sp(4),
    Paragraph(
        "Linhas com o campo <b>Tipo de Documento</b> em branco são excluídas automaticamente "
        "antes do processamento.",
        nota,
    ),
]

# 2 — Estrutura do arquivo
story += [
    Paragraph("2. Estrutura do Arquivo Excel", h1),
    hr(),
    Paragraph(
        "O arquivo deve conter <b>duas abas</b> com os nomes exatos indicados abaixo "
        "(incluindo maiúsculas e minúsculas).",
        corpo,
    ),

    Paragraph("2.1 Aba SAP — Dados do relatório FBL3H", h2),
    Paragraph(
        "Cole aqui o relatório exportado diretamente do SAP. "
        "A linha 1 deve ser o cabeçalho e os dados devem começar na linha 2.",
        corpo,
    ),
    sp(4),
    make_table(
        ["Coluna", "Campo", "Tipo"],
        [
            ["A", "Empresa", "Texto"],
            ["B", "Moeda da empresa", "Texto"],
            ["C", "Conta do Razão", "Texto / Número  ← usado na busca de regras"],
            ["D", "Data de documento", "Data"],
            ["E", "Data de lançamento", "Data  ← base para Dias em Aberto"],
            ["F", "Nº documento", "Texto / Número"],
            ["G", "Tipo de documento", "Texto  ← linhas em branco são excluídas"],
            ["H", "Valor em moeda da empresa", "Número"],
            ["I", "Chave de lançamento", "Texto"],
            ["J", "Atribuição", "Texto"],
            ["K", "Texto", "Texto  ← usado na busca de regras"],
            ["L", "Data Base ❆ gerada pela macro", "Data"],
            ["M", "Dias em Aberto ❆ gerada pela macro", "Número"],
            ["N", "Área Responsável ❆ gerada pela macro", "Texto"],
        ],
        [1.2*cm, 5.5*cm, 6.5*cm],
    ),
    sp(8),

    Paragraph("2.2 Aba BASE — Tabela de regras", h2),
    Paragraph(
        "Esta aba contém as regras usadas para identificar a área responsável. "
        "A linha 1 deve ser o cabeçalho e as regras começam na linha 2.",
        corpo,
    ),
    sp(4),
    make_table(
        ["Coluna", "Campo", "Descrição"],
        [
            ["A", "Conta do Razão", "Conta contábil a filtrar. Deixe em branco para aplicar a regra a qualquer conta."],
            ["B", "Começa em", "O campo Texto do SAP deve começar com este valor. Deixe em branco para ignorar."],
            ["C", "Contém", "O campo Texto do SAP deve conter este valor. Deixe em branco para ignorar."],
            ["D", "Regex", "Expressão regular aplicada ao Texto do SAP. Deixe em branco para ignorar."],
            ["E", "Área Responsável", "Nome da área que será preenchido no relatório. Obrigatório."],
        ],
        [1.2*cm, 3.2*cm, 8.8*cm],
    ),
]

# 3 — Como preencher a BASE
story += [
    Paragraph("3. Como Preencher a Tabela BASE", h1),
    hr(),

    Paragraph("3.1 Lógica de pontuação (score)", h2),
    Paragraph(
        "Para cada linha do SAP, a macro percorre todas as regras da BASE e calcula "
        "uma pontuação de 0 a 4. A regra com <b>maior pontuação</b> vence. "
        "Em caso de empate, prevalece a <b>primeira regra na ordem da tabela</b>.",
        corpo,
    ),
    sp(4),
    make_table(
        ["Condição avaliada", "Pontos"],
        [
            ["Conta do Razão em branco na BASE  ou  bate com a do SAP", "+1"],
            ['Campo "Começa em" não vazio  e  o Texto do SAP começa com ele', "+1"],
            ['Campo "Contém" não vazio  e  o Texto do SAP contém ele', "+1"],
            ['Campo "Regex" não vazio  e  o Texto do SAP corresponde ao padrão', "+1"],
        ],
        [11*cm, 2.2*cm],
    ),
    sp(6),
    Paragraph(
        "<b>Eliminação:</b> se um campo não estiver em branco e não bater com o SAP, "
        "a pontuação da regra vai a zero e ela é descartada.",
        nota,
    ),

    Paragraph("3.2 Exemplos práticos", h2),
    Paragraph(
        "Tabela BASE com as regras cadastradas:",
        corpo,
    ),
    sp(4),
    make_table(
        ["Conta (A)", "Começa em (B)", "Contém (C)", "Regex (D)", "Área (E)"],
        [
            ["28001107", "TED-TRANSF", "OPELLA", "", "CONTAS A RECEBER"],
            ["28001310", "", "VISA", "", "CONTABILIDADE"],
            ["", "", "", r"^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}$", "GESTÃO DE CAIXA"],
            ["28001213", "", "", "", "TRANSPORTADORA"],
        ],
        [2.2*cm, 2.5*cm, 2.0*cm, 4.5*cm, 2.0*cm],
    ),
    sp(8),
    Paragraph("Resultados para linhas do relatório SAP:", corpo),
    sp(4),
    make_table(
        ["Conta SAP", "Texto SAP", "Score", "Área Responsável"],
        [
            ["28001107", "TED-TRANSF OPELLA HEALTHCARE 03/06", "3", "CONTAS A RECEBER"],
            ["28001310", "PAGTO FATURA VISA 06/2026", "2", "CONTABILIDADE"],
            ["28001107", "0003305-48.2020.8.26.0565", "2", "GESTÃO DE CAIXA"],
            ["28001213", "NF 001234 TRANSP EXPRESS", "1", "TRANSPORTADORA"],
            ["28001107", "DEVOLUCAO CLIENTE 999", "0", "⚠ Não identificado"],
        ],
        [2.2*cm, 5.5*cm, 1.5*cm, 4.0*cm],
    ),
    sp(8),

    Paragraph("3.3 Casos especiais", h2),
    Paragraph(
        "<b>Conta em branco na BASE</b> — a regra se aplica a qualquer conta do SAP "
        "(útil para padrões de texto que aparecem em múltiplas contas).",
        bullet,
    ),
    Paragraph(
        "<b>Regex</b> — use para textos com estrutura numérica variável, como "
        "números de processo (ex: 0003305-48.2020.8.26.0565). O padrão "
        r"^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}$"
        " corresponde a esse formato.",
        bullet,
    ),
    Paragraph(
        "<b>Apenas conta preenchida</b> — a regra identifica todos os textos "
        "daquela conta. Use para contas com uma única área responsável.",
        bullet,
    ),
    Paragraph(
        "<b>Nenhuma regra bate</b> — a célula de Área Responsável recebe "
        "\"⚠ Não identificado\". Adicione uma regra na BASE para cobrir esse caso.",
        bullet,
    ),
    sp(4),
    Paragraph(
        "Todas as comparações de texto (Começa em, Contém) são insensíveis a "
        "maiúsculas/minúsculas. O campo Regex usa a flag IgnoreCase.",
        nota,
    ),

    Paragraph("3.4 Boas práticas", h2),
    Paragraph("• Coloque as regras <b>mais específicas primeiro</b> na tabela BASE.", bullet),
    Paragraph("• Use \"Começa em\" quando o texto sempre inicia com um padrão fixo.", bullet),
    Paragraph("• Use \"Contém\" quando o padrão pode aparecer em qualquer posição do texto.", bullet),
    Paragraph("• Use \"Regex\" para textos com estrutura numérica sem parte fixa.", bullet),
    Paragraph("• Deixe a Conta do Razão em branco apenas quando o texto for único o suficiente para não gerar conflitos.", bullet),
    Paragraph("• Revise as linhas marcadas como \"⚠ Não identificado\" e adicione as regras faltantes.", bullet),
]

# 4 — Importando o módulo
story += [
    Paragraph("4. Como Importar o Módulo VBA", h1),
    hr(),
    Paragraph(
        "Siga os passos abaixo para adicionar o módulo ao seu arquivo Excel. "
        "Este procedimento é feito <b>uma única vez</b> por arquivo.",
        corpo,
    ),
    sp(4),
    make_table(
        ["Passo", "Ação"],
        [
            ["1", "Abra o arquivo Excel que contém as abas SAP e BASE."],
            ["2", "Pressione Alt + F11 para abrir o Editor Visual Basic (VBE)."],
            ["3", "No menu superior do VBE, clique em Arquivo → Importar Arquivo..."],
            ["4", "Localize e selecione o arquivo ModuloPendencias.bas e clique em Abrir."],
            ["5", "Confirme que o módulo \"ModuloPendencias\" aparece na lista à esquerda."],
            ["6", "Feche o VBE (Alt + Q ou clique no X da janela)."],
            ["7", "Salve o arquivo Excel no formato .xlsm (Excel com Macros Habilitadas)."],
        ],
        [1.5*cm, 11.7*cm],
    ),
    sp(6),
    Paragraph(
        "⚠  Ao abrir o arquivo futuramente, o Excel pode exibir uma barra amarela "
        "pedindo para \"Habilitar Conteúdo\". Clique nesse botão para que as macros "
        "funcionem.",
        aviso,
    ),
]

# 5 — Executando a macro
story += [
    Paragraph("5. Executando as Macros", h1),
    hr(),

    Paragraph("5.1 ProcessarPendencias — macro principal", h2),
    Paragraph(
        "Esta macro realiza todo o processamento: exclui linhas inválidas, "
        "identifica as áreas responsáveis e preenche as colunas L, M e N.",
        corpo,
    ),
    sp(4),
    make_table(
        ["Passo", "Ação"],
        [
            ["1", "Certifique-se de que as abas se chamam exatamente SAP e BASE."],
            ["2", "Pressione Alt + F8 para abrir a janela de macros."],
            ["3", "Selecione ProcessarPendencias na lista e clique em Executar."],
            ["4", "Aguarde. Para arquivos grandes (100 mil+ linhas) o processamento pode levar alguns minutos."],
            ["5", "Ao final, uma mensagem exibirá o resumo: linhas analisadas, atribuídas e não identificadas."],
        ],
        [1.5*cm, 11.7*cm],
    ),
    sp(8),

    Paragraph("5.2 LimparResultados — macro auxiliar", h2),
    Paragraph(
        "Apaga o conteúdo das colunas L, M e N (incluindo os cabeçalhos) para "
        "permitir reprocessamento com novos dados ou regras.",
        corpo,
    ),
    sp(4),
    make_table(
        ["Passo", "Ação"],
        [
            ["1", "Pressione Alt + F8 para abrir a janela de macros."],
            ["2", "Selecione LimparResultados e clique em Executar."],
            ["3", "Confirme que as colunas L, M e N foram esvaziadas."],
        ],
        [1.5*cm, 11.7*cm],
    ),
    sp(6),
    Paragraph(
        "Execute LimparResultados antes de ProcessarPendencias sempre que substituir "
        "os dados do SAP ou atualizar as regras da BASE.",
        nota,
    ),
]

# 6 — Mensagens de erro
story += [
    Paragraph("6. Mensagens e Situações Comuns", h1),
    hr(),
    make_table(
        ["Mensagem / Situação", "Causa", "Solução"],
        [
            [
                "Aba 'SAP' não encontrada",
                "A aba com os dados não existe ou tem nome diferente.",
                "Renomeie a aba para SAP (exatamente, com maiúsculas).",
            ],
            [
                "Aba 'BASE' não encontrada",
                "A tabela de regras não existe ou tem nome diferente.",
                "Renomeie a aba para BASE (exatamente, com maiúsculas).",
            ],
            [
                "Nenhum dado encontrado na aba 'SAP'",
                "A aba SAP está vazia ou só tem o cabeçalho.",
                "Cole os dados do FBL3H a partir da linha 2.",
            ],
            [
                "BASE vazia",
                "A aba BASE não tem nenhuma regra cadastrada.",
                "Adicione pelo menos uma regra a partir da linha 2.",
            ],
            [
                "Todas as linhas foram excluídas",
                "Todas as linhas do SAP tinham Tipo de Documento em branco.",
                "Verifique se o relatório foi exportado corretamente.",
            ],
            [
                "⚠ Não identificado na coluna N",
                "Nenhuma regra da BASE corresponde àquela linha.",
                "Adicione uma regra na BASE para cobrir o texto/conta daquela linha.",
            ],
        ],
        [4*cm, 4.2*cm, 5*cm],
    ),
]

# 7 — Fluxo resumido
story += [
    Paragraph("7. Fluxo Resumido de Uso", h1),
    hr(),
    Paragraph("Siga este fluxo sempre que receber um novo relatório do SAP:", corpo),
    sp(4),
    Paragraph("<b>1.</b> Exporte o relatório FBL3H do SAP.", bullet),
    Paragraph("<b>2.</b> Cole os dados na aba <b>SAP</b> a partir da linha 2 (preservando o cabeçalho na linha 1).", bullet),
    Paragraph("<b>3.</b> Verifique e atualize as regras na aba <b>BASE</b> se necessário.", bullet),
    Paragraph("<b>4.</b> Execute <b>LimparResultados</b> para apagar resultados anteriores.", bullet),
    Paragraph("<b>5.</b> Execute <b>ProcessarPendencias</b> para processar o novo relatório.", bullet),
    Paragraph("<b>6.</b> Revise as linhas marcadas como \"⚠ Não identificado\" e adicione regras se necessário.", bullet),
    Paragraph("<b>7.</b> Repita os passos 4–6 até que todas as linhas estejam identificadas.", bullet),
]

doc.build(story)
print(f"PDF gerado em: {OUTPUT}")
