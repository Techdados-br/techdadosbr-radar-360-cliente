import os
import html
import textwrap
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak


st.set_page_config(
    page_title="TechDadosBR Radar 360",
    page_icon="📡",
    layout="wide",
)

st.markdown('''
<style>
.radar-card-value.valor-negativo {
    color: #D85C5C !important;
}

.radar-card-value.valor-positivo {
    color: #28A879 !important;
}

.radar-card-value.valor-atencao {
    color: #D6A23A !important;
}

.radar-card-value.valor-neutro {
    color: #4C93AD !important;
}

.radar-card-value {
    font-weight: 850 !important;
    opacity: 1 !important;
}
</style>
''', unsafe_allow_html=True)

st.markdown('''
<style>
.radar-card-value {
    text-shadow: none !important;
}
</style>
''', unsafe_allow_html=True)

st.markdown('''
<style>
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #14324A 0%, #102634 100%) !important;
}
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
    color: #F4F8F9 !important;
}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
    background: rgba(194, 138, 44, 0.18) !important;
    border-left: 3px solid #C28A2C !important;
    border-radius: 6px !important;
}
</style>
''', unsafe_allow_html=True)

DIRETORIO_APP = os.path.dirname(os.path.abspath(__file__))

# Funciona localmente e também no Streamlit Cloud.
if os.path.basename(DIRETORIO_APP) == "03_App_Cliente":
    DIRETORIO_PROJETO = os.path.dirname(DIRETORIO_APP)
else:
    DIRETORIO_PROJETO = DIRETORIO_APP

CAMINHO_BASE = os.path.join(
    DIRETORIO_PROJETO,
    "01_Base_Teste",
    "Base_Teste_TechDadosBR_Radar_360.xlsx",
)

CAMINHO_ACOMPANHAMENTO = os.path.join(
    DIRETORIO_PROJETO,
    "01_Base_Teste",
    "Acompanhamento_Acoes_Radar_360.xlsx",
)

ID_PLANILHA_ACOMPANHAMENTO = (
    "1WjInAguXhrkdPFpCKYKA5RmWTTKm02S6KkXxMucWrC0"
)

NOME_ABA_ACOMPANHAMENTO = "Acompanhamento_Acoes"

ESCOPO_GOOGLE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def conectar_planilha_acompanhamento():
    credenciais = Credentials.from_service_account_info(
        st.secrets["google_service_account"],
        scopes=ESCOPO_GOOGLE,
    )

    cliente_google = gspread.authorize(credenciais)
    planilha = cliente_google.open_by_key(
        ID_PLANILHA_ACOMPANHAMENTO
    )

    return planilha.worksheet(
        NOME_ABA_ACOMPANHAMENTO
    )


def verificar_acesso_cliente():
    senha_configurada = str(
        st.secrets.get("acesso_cliente", {}).get("senha", "")
    ).strip()

    if not senha_configurada:
        st.error(
            "A senha do cliente ainda não foi configurada nos Secrets do Streamlit."
        )
        st.stop()

    if st.session_state.get("cliente_autenticado", False):
        return

    st.markdown("### Acesso ao Painel do Cliente")
    senha_informada = st.text_input(
        "Senha",
        type="password",
        key="senha_acesso_cliente",
    )

    if st.button("Entrar", type="primary"):
        if senha_informada == senha_configurada:
            st.session_state["cliente_autenticado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")

    st.stop()


def moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data
def carregar_base(caminho):
    return pd.read_excel(caminho, sheet_name=None)


def carregar_acompanhamento(caminho=None):
    try:
        aba = conectar_planilha_acompanhamento()
        registros = aba.get_all_records()

        if not registros:
            return pd.DataFrame()

        df = pd.DataFrame(registros)
        return df.dropna(how="all")

    except Exception as erro:
        st.error(
            f"Não foi possível carregar o acompanhamento do Google Sheets: {erro}"
        )
        return pd.DataFrame()


def salvar_acompanhamento(caminho, registro):
    aba = conectar_planilha_acompanhamento()

    cabecalhos = aba.row_values(1)
    id_acao = str(registro["ID_Acao"])

    coluna_id = cabecalhos.index("ID_Acao") + 1
    valores_id = aba.col_values(coluna_id)

    linha_encontrada = None

    for numero_linha, valor_id in enumerate(valores_id[1:], start=2):
        if str(valor_id) == id_acao:
            linha_encontrada = numero_linha
            break

    linha_registro = []

    for cabecalho in cabecalhos:
        valor = registro.get(cabecalho, "")

        if valor is None:
            valor = ""
        elif isinstance(valor, pd.Timestamp):
            valor = valor.strftime("%d/%m/%Y %H:%M:%S")
        elif hasattr(valor, "strftime"):
            valor = valor.strftime("%d/%m/%Y")
        elif pd.isna(valor):
            valor = ""

        linha_registro.append(valor)

    if linha_encontrada is None:
        aba.append_row(
            linha_registro,
            value_input_option="USER_ENTERED",
        )
    else:
        ultima_celula = gspread.utils.rowcol_to_a1(
            linha_encontrada,
            len(cabecalhos),
        )
        intervalo = f"A{linha_encontrada}:{ultima_celula}"

        aba.update(
            intervalo,
            [linha_registro],
            value_input_option="USER_ENTERED",
        )


def gerar_pdf_executivo(imoveis, proprietarios, acompanhamento):
    """Gera um relatório executivo visual em duas páginas."""
    buffer = BytesIO()

    AZUL = colors.HexColor("#163A5F")
    AZUL_MEDIO = colors.HexColor("#2B6F89")
    AZUL_CLARO = colors.HexColor("#EAF2F8")
    VERDE = colors.HexColor("#178A64")
    VERDE_CLARO = colors.HexColor("#E7F5EF")
    VERMELHO = colors.HexColor("#B84343")
    VERMELHO_CLARO = colors.HexColor("#FCEBEC")
    DOURADO = colors.HexColor("#C28A2C")
    DOURADO_CLARO = colors.HexColor("#FFF4D8")
    CINZA = colors.HexColor("#66788A")
    CINZA_CLARO = colors.HexColor("#F4F7F9")
    BORDA = colors.HexColor("#D7E0E8")
    TEXTO = colors.HexColor("#203746")
    BRANCO = colors.white

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=15 * mm,
        title="TechDadosBR Radar 360 - Relatório Executivo",
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TituloVisual",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=21,
        textColor=AZUL,
        alignment=TA_LEFT,
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        name="SubtituloVisual",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=CINZA,
        spaceAfter=7,
    ))
    styles.add(ParagraphStyle(
        name="SecaoVisual",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=14,
        textColor=AZUL,
        spaceBefore=13,
        spaceAfter=9,
    ))
    styles.add(ParagraphStyle(
        name="CardTitulo",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=8.5,
        textColor=CINZA,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name="CardValorRisco",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=13.2,
        leading=15,
        textColor=VERMELHO,
    ))
    styles.add(ParagraphStyle(
        name="CardValorPositivo",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=13.2,
        leading=15,
        textColor=VERDE,
    ))
    styles.add(ParagraphStyle(
        name="CardSub",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=6.4,
        leading=8,
        textColor=CINZA,
    ))
    styles.add(ParagraphStyle(
        name="TextoVisual",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.2,
        leading=10.6,
        textColor=TEXTO,
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="TextoBox",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=7.6,
        leading=9.5,
        textColor=TEXTO,
    ))
    styles.add(ParagraphStyle(
        name="TituloBoxRisco",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10,
        textColor=VERMELHO,
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="TituloBoxOportunidade",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10,
        textColor=VERDE,
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="TextoTabelaVisual",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=6.8,
        leading=8.2,
        textColor=TEXTO,
    ))
    styles.add(ParagraphStyle(
        name="TextoPequenoVisual",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=6.8,
        leading=8.5,
        textColor=CINZA,
    ))

    # -------------------- Cálculos --------------------
    contratos = imoveis[imoveis["Dias_Para_Vencimento"].between(0, 90)]
    atrasos = imoveis[imoveis["Perfil_Pagamento"] == "Atraso recorrente"]
    perda_vacancia = float(imoveis["Perda_Mensal_Vacancia"].sum())

    mascara_receita = (
        imoveis["Dias_Para_Vencimento"].between(0, 90)
        | imoveis["Perfil_Pagamento"].eq("Atraso recorrente")
    )
    receita_exposta = float(
        imoveis.loc[mascara_receita, "Receita_Mensal_Imobiliaria"].sum()
    )
    valor_locativo = float(perda_vacancia + atrasos["Aluguel_Mensal"].sum())

    alto_risco = imoveis[
        imoveis["Classificacao_Risco"].isin(["Alto", "Crítico"])
    ]
    proprietarios_risco = proprietarios[
        proprietarios["Risco_Relacionamento"] >= 70
    ]

    top_vagos = imoveis[
        imoveis["Status_Imovel"] == "Vago"
    ].sort_values("Perda_Mensal_Vacancia", ascending=False).head(5)

    concentracao = 0.0
    if perda_vacancia > 0:
        concentracao = float(
            top_vagos["Perda_Mensal_Vacancia"].sum()
            / perda_vacancia
            * 100
        )

    oportunidades = imoveis.copy()
    oportunidades["Potencial_Locativo"] = (
        oportunidades["Valor_Mercado_Estimado"]
        - oportunidades["Aluguel_Mensal"]
    ).clip(lower=0)

    # Usa exatamente os mesmos critérios da tela "Radar de oportunidades".
    mascara_oportunidade_ocupado = (
        (
            (oportunidades["Status_Imovel"] == "Ocupado")
            & (oportunidades["Potencial_Locativo"] >= 200)
            & (oportunidades["Meses_Sem_Reajuste"] >= 18)
        )
        |
        (
            (oportunidades["Status_Imovel"] == "Ocupado")
            & (oportunidades["Potencial_Locativo"] >= 300)
            & (oportunidades["Meses_Sem_Reajuste"] < 18)
        )
    )

    oportunidades_ocupadas = oportunidades[
        mascara_oportunidade_ocupado
    ].copy()

    oportunidades_vagas = oportunidades[
        (oportunidades["Status_Imovel"] == "Vago")
        & (oportunidades["Dias_Vago"] >= 90)
        & (oportunidades["Diferenca_Preco_Mercado"] >= 0.10)
    ].copy()

    potencial_locativo = float(
        oportunidades_ocupadas["Potencial_Locativo"].sum()
    )
    potencial_imobiliaria = float(
        (
            oportunidades_ocupadas["Potencial_Locativo"]
            * oportunidades_ocupadas["Taxa_Administracao"]
        ).sum()
    )

    prioridades = imoveis.copy()
    prioridades["Impacto"] = pd.to_numeric(
        prioridades["Perda_Mensal_Vacancia"],
        errors="coerce",
    ).fillna(0).astype(float)

    mascara_impacto = (
        prioridades["Perfil_Pagamento"].eq("Atraso recorrente")
        | prioridades["Dias_Para_Vencimento"].between(0, 90)
    )
    prioridades.loc[mascara_impacto, "Impacto"] += prioridades.loc[
        mascara_impacto, "Receita_Mensal_Imobiliaria"
    ]

    def acao_pdf(row):
        if row["Dias_Vago"] >= 90 and row["Diferenca_Preco_Mercado"] >= 0.10:
            return "Revisar preço e reposicionar anúncio"
        if row["Perfil_Pagamento"] == "Atraso recorrente":
            return "Iniciar cobrança preventiva"
        if pd.notna(row["Dias_Para_Vencimento"]) and 0 <= row["Dias_Para_Vencimento"] <= 90:
            return "Iniciar renovação contratual"
        if row["Meses_Sem_Reajuste"] >= 18:
            return "Revisar reajuste contratual"
        if row["Chamados_12m"] >= 5:
            return "Analisar custo operacional"
        return "Manter monitoramento"

    def prazo_pdf(row):
        indice = row["Indice_Risco_Imovel"]
        if indice >= 80:
            return "Até 2 dias"
        if indice >= 60:
            return "Até 5 dias"
        if indice >= 30:
            return "Até 10 dias"
        return "Até 30 dias"

    def faixa_pdf(indice):
        if indice >= 80:
            return "Crítico"
        if indice >= 60:
            return "Alto"
        if indice >= 30:
            return "Atenção"
        return "Baixo"

    prioridades["Ação_PDF"] = prioridades.apply(acao_pdf, axis=1)
    prioridades["Prazo_PDF"] = prioridades.apply(prazo_pdf, axis=1)
    prioridades["Faixa_PDF"] = prioridades["Indice_Risco_Imovel"].apply(faixa_pdf)
    prioridades = prioridades[
        prioridades["Indice_Risco_Imovel"] >= 30
    ].sort_values(
        ["Indice_Risco_Imovel", "Impacto"],
        ascending=[False, False],
    ).head(7)

    total_recuperado = 0.0
    nao_iniciadas = 0
    em_andamento = 0
    concluidas = 0
    atrasadas_acomp = 0

    if acompanhamento is not None and not acompanhamento.empty:
        acompanhamento = acompanhamento.copy()
        acompanhamento["Resultado_Recuperado"] = pd.to_numeric(
            acompanhamento["Resultado_Recuperado"],
            errors="coerce",
        ).fillna(0)
        total_recuperado = float(
            acompanhamento["Resultado_Recuperado"].sum()
        )
        status = acompanhamento["Status"].fillna("Não iniciado")
        nao_iniciadas = int((status == "Não iniciado").sum())
        em_andamento = int((status == "Em andamento").sum())
        concluidas = int((status == "Concluído").sum())
        datas_limite = pd.to_datetime(
            acompanhamento["Data_Limite"],
            errors="coerce",
        )
        atrasadas_acomp = int(
            (
                datas_limite.notna()
                & (datas_limite < pd.Timestamp.today().normalize())
                & ~status.isin(["Concluído", "Cancelado"])
            ).sum()
        )

    distribuicao = (
        imoveis["Classificacao_Risco"]
        .value_counts()
        .reindex(["Baixo", "Atenção", "Alto", "Crítico"], fill_value=0)
    )

    data_emissao = pd.Timestamp.now().strftime("%d/%m/%Y")

    # -------------------- Helpers visuais --------------------
    def card(titulo, valor, subtitulo, cor_fundo, estilo_valor):
        conteudo = [
            [Paragraph(titulo, styles["CardTitulo"])],
            [Paragraph(valor, styles[estilo_valor])],
            [Paragraph(subtitulo, styles["CardSub"])],
        ]
        tabela = Table(conteudo, colWidths=[43.5 * mm], rowHeights=[11 * mm, 11 * mm, 11 * mm])
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cor_fundo),
            ("BOX", (0, 0), (-1, -1), 0.6, BORDA),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        return tabela

    def faixa_titulo(texto, cor):
        tabela = Table([[Paragraph(texto, ParagraphStyle(
            name=f"Faixa{texto}",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=BRANCO,
        ))]], colWidths=[180 * mm])
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cor),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        return tabela

    # -------------------- Página 1 --------------------
    story = []

    topo = Table(
        [[
            Paragraph("TechDadosBR Radar 360", styles["TituloVisual"]),
            Paragraph(
                f"<b>Relatório Executivo da Carteira Imobiliária</b><br/>"
                f"Data de emissão: {data_emissao}",
                styles["SubtituloVisual"],
            ),
        ]],
        colWidths=[105 * mm, 75 * mm],
    )
    topo.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, -1), 1.2, AZUL),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(topo)
    story.append(Spacer(1, 18))

    cards = Table(
        [[
            card(
                "VALOR LOCATIVO EM RISCO",
                moeda(valor_locativo),
                "Vacância mensal + aluguéis ligados a atraso recorrente",
                VERMELHO_CLARO,
                "CardValorRisco",
            ),
            card(
                "RECEITA DA IMOBILIÁRIA EM RISCO",
                moeda(receita_exposta),
                "Vencimento em 90 dias ou atraso recorrente",
                VERMELHO_CLARO,
                "CardValorRisco",
            ),
            card(
                "POTENCIAL LOCATIVO MENSAL",
                moeda(potencial_locativo),
                "Imóveis ocupados abaixo do mercado",
                VERDE_CLARO,
                "CardValorPositivo",
            ),
            card(
                "POTENCIAL ADICIONAL DA IMOBILIÁRIA",
                moeda(potencial_imobiliaria),
                "Possível aumento das taxas de administração",
                VERDE_CLARO,
                "CardValorPositivo",
            ),
        ]],
        colWidths=[45 * mm] * 4,
    )
    cards.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    story.append(cards)
    story.append(Spacer(1, 20))

    story.append(faixa_titulo("LEITURA EXECUTIVA", AZUL_MEDIO))
    leitura_tbl = Table(
        [[Paragraph(
            f"<b>Valor locativo em risco:</b> soma a perda mensal por vacância "
            f"aos aluguéis dos imóveis com atraso recorrente.<br/>"
            f"<b>Risco concentrado:</b> os cinco imóveis com maior perda representam "
            f"{concentracao:.1f}% da vacância mensal.<br/>"
            f"<b>Receita contratual:</b> {len(contratos)} contratos vencem em até 90 dias, "
            f"associados a {moeda(float(contratos['Receita_Mensal_Imobiliaria'].sum()))} "
            f"em taxas mensais.<br/>"
            f"<b>Oportunidade:</b> a carteira apresenta potencial locativo mensal de "
            f"{moeda(potencial_locativo)} e potencial adicional de "
            f"{moeda(potencial_imobiliaria)} para a imobiliária.",
            styles["TextoVisual"],
        )]],
        colWidths=[180 * mm],
    )
    leitura_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_CLARO),
        ("BOX", (0, 0), (-1, -1), 0.5, BORDA),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]))
    story.append(leitura_tbl)
    story.append(Spacer(1, 20))

    riscos_texto = (
        f"<b>Vacância:</b> {len(top_vagos)} imóveis concentram as maiores perdas.<br/>"
        f"<b>Pagamento:</b> {len(atrasos)} imóveis possuem atraso recorrente, "
        f"associados a {moeda(float(atrasos['Aluguel_Mensal'].sum()))}.<br/>"
        f"<b>Relacionamento:</b> {len(proprietarios_risco)} proprietário(s) "
        f"exigem contato gerencial prioritário."
    )
    oportunidades_texto = (
        f"<b>Reajuste:</b> {len(oportunidades_ocupadas)} imóvel(is) ocupado(s) "
        f"apresentam potencial de adequação.<br/>"
        f"<b>Vacância:</b> {len(oportunidades_vagas)} imóvel(is) exigem revisão "
        f"imediata de preço e posicionamento.<br/>"
        f"<b>Resultado:</b> {moeda(total_recuperado)} já foi registrado como recuperado."
    )

    box_risco = Table(
        [[Paragraph("PRINCIPAIS RISCOS IDENTIFICADOS", styles["TituloBoxRisco"])],
         [Paragraph(riscos_texto, styles["TextoBox"])]],
        colWidths=[87 * mm],
    )
    box_risco.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), VERMELHO_CLARO),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#E7B6BA")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    box_oportunidade = Table(
        [[Paragraph("PRINCIPAIS OPORTUNIDADES", styles["TituloBoxOportunidade"])],
         [Paragraph(oportunidades_texto, styles["TextoBox"])]],
        colWidths=[87 * mm],
    )
    box_oportunidade.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), VERDE_CLARO),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#A9D8C6")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    dois_boxes = Table(
        [[box_risco, box_oportunidade]],
        colWidths=[90 * mm, 90 * mm],
    )
    dois_boxes.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    story.append(dois_boxes)
    story.append(Spacer(1, 22))

    story.append(Paragraph("Distribuição do risco da carteira", styles["SecaoVisual"]))
    dados_risco = [
        [
            Paragraph("<b>Baixo</b>", styles["TextoBox"]),
            Paragraph("<b>Atenção</b>", styles["TextoBox"]),
            Paragraph("<b>Alto</b>", styles["TextoBox"]),
            Paragraph("<b>Crítico</b>", styles["TextoBox"]),
        ],
        [
            str(int(distribuicao["Baixo"])),
            str(int(distribuicao["Atenção"])),
            str(int(distribuicao["Alto"])),
            str(int(distribuicao["Crítico"])),
        ],
    ]
    tabela_risco = Table(dados_risco, colWidths=[45 * mm] * 4)
    tabela_risco.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#DCEEFE")),
        ("BACKGROUND", (1, 0), (1, -1), DOURADO_CLARO),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#FADBD8")),
        ("BACKGROUND", (3, 0), (3, -1), colors.HexColor("#F5B7B1")),
        ("TEXTCOLOR", (0, 0), (-1, -1), TEXTO),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.3),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDA),
    ]))
    story.append(tabela_risco)
    story.append(Spacer(1, 20))

    acompanhamento_cards = Table(
        [[
            ["Não iniciadas", str(nao_iniciadas), CINZA_CLARO],
            ["Em andamento", str(em_andamento), AZUL_CLARO],
            ["Concluídas", str(concluidas), VERDE_CLARO],
            ["Atrasadas", str(atrasadas_acomp), VERMELHO_CLARO],
            ["Recuperado", moeda(total_recuperado), VERDE_CLARO],
        ]],
        colWidths=[36 * mm] * 5,
    )
    # Converte cada célula em mini-card
    for coluna in range(5):
        titulo, valor, fundo = acompanhamento_cards._cellvalues[0][coluna]
        acompanhamento_cards._cellvalues[0][coluna] = Table(
            [[Paragraph(titulo, styles["CardTitulo"])],
             [Paragraph(
                 valor,
                 styles["CardValorPositivo"] if coluna in [2, 4]
                 else styles["CardValorRisco"] if coluna == 3
                 else styles["TextoVisual"],
             )]],
            colWidths=[34 * mm],
        )
        acompanhamento_cards._cellvalues[0][coluna].setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), fundo),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDA),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
    acompanhamento_cards.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.2),
    ]))
    story.append(Paragraph("Acompanhamento das ações", styles["SecaoVisual"]))
    story.append(acompanhamento_cards)

    story.append(PageBreak())

    # -------------------- Página 2 --------------------
    story.append(Paragraph("Plano de Ação Prioritário", styles["TituloVisual"]))
    story.append(Paragraph(
        "Ações recomendadas por urgência, impacto financeiro e acompanhamento.",
        styles["SubtituloVisual"],
    ))

    story.append(faixa_titulo("PRIORIDADES DA CARTEIRA", AZUL))
    dados_prioridades = [[
        "Imóvel", "Índice", "Faixa", "Fator", "Impacto", "Prazo", "Ação prioritária"
    ]]

    for _, row in prioridades.iterrows():
        dados_prioridades.append([
            row["ID_Imovel"],
            str(int(row["Indice_Risco_Imovel"])),
            row["Faixa_PDF"],
            row["Fator_Principal"],
            moeda(float(row["Impacto"])),
            row["Prazo_PDF"],
            Paragraph(row["Ação_PDF"], styles["TextoTabelaVisual"]),
        ])

    tabela_prioridades = Table(
        dados_prioridades,
        colWidths=[19 * mm, 13 * mm, 18 * mm, 22 * mm, 25 * mm, 22 * mm, 61 * mm],
        repeatRows=1,
    )
    tabela_prioridades.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), BRANCO),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.35, BORDA),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, CINZA_CLARO]),
    ]))

    # Cores das faixas na coluna correspondente
    for linha, (_, row) in enumerate(prioridades.iterrows(), start=1):
        faixa = row["Faixa_PDF"]
        cor = {
            "Crítico": VERMELHO_CLARO,
            "Alto": colors.HexColor("#FCE4D6"),
            "Atenção": DOURADO_CLARO,
            "Baixo": AZUL_CLARO,
        }[faixa]
        tabela_prioridades.setStyle(TableStyle([
            ("BACKGROUND", (2, linha), (2, linha), cor),
            ("FONTNAME", (2, linha), (2, linha), "Helvetica-Bold"),
        ]))

    story.append(tabela_prioridades)
    story.append(Spacer(1, 21))

    story.append(faixa_titulo("3 AÇÕES MAIS IMPORTANTES", DOURADO))
    top3 = prioridades.head(3)
    destaques = []
    for _, row in top3.iterrows():
        destaques.append(Table(
            [[Paragraph(
                f"<b>{row['ID_Imovel']}</b> - {row['Fator_Principal']}<br/>"
                f"{row['Ação_PDF']}<br/>"
                f"<b>Prazo:</b> {row['Prazo_PDF']} &nbsp;&nbsp; "
                f"<b>Impacto:</b> {moeda(float(row['Impacto']))}",
                styles["TextoBox"],
            )]],
            colWidths=[57 * mm],
        ))
        destaques[-1].setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), DOURADO_CLARO),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#E5C678")),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
    while len(destaques) < 3:
        destaques.append("")
    tabela_destaques = Table([destaques], colWidths=[60 * mm] * 3)
    tabela_destaques.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    story.append(tabela_destaques)
    story.append(Spacer(1, 21))

    story.append(faixa_titulo("OPORTUNIDADES PRIORIZADAS", VERDE))
    oportunidades_top = oportunidades_ocupadas.sort_values(
        "Potencial_Locativo", ascending=False
    ).head(5)

    dados_oportunidades = [[
        "Imóvel", "Bairro", "Potencial locativo", "Potencial imobiliária", "Ação"
    ]]
    for _, row in oportunidades_top.iterrows():
        potencial_taxa = (
            float(row["Potencial_Locativo"])
            * float(row["Taxa_Administracao"])
        )
        acao = (
            "Revisar reajuste e negociar adequação gradual"
            if row["Meses_Sem_Reajuste"] >= 18
            else "Avaliar adequação na próxima negociação"
        )
        dados_oportunidades.append([
            row["ID_Imovel"],
            row["Bairro"],
            moeda(float(row["Potencial_Locativo"])),
            moeda(potencial_taxa),
            Paragraph(acao, styles["TextoTabelaVisual"]),
        ])

    if len(oportunidades_top) == 0:
        dados_oportunidades.append(
            ["-", "-", moeda(0), moeda(0), "Nenhuma oportunidade relevante"]
        )

    tabela_oportunidades = Table(
        dados_oportunidades,
        colWidths=[20 * mm, 31 * mm, 31 * mm, 31 * mm, 67 * mm],
        repeatRows=1,
    )
    tabela_oportunidades.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), VERDE),
        ("TEXTCOLOR", (0, 0), (-1, 0), BRANCO),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.35, BORDA),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, VERDE_CLARO]),
    ]))
    story.append(tabela_oportunidades)
    story.append(Spacer(1, 21))

    story.append(faixa_titulo("RECOMENDAÇÕES EXECUTIVAS", AZUL_MEDIO))
    recomendacoes = [
        "Concentrar a atuação nos imóveis com maior impacto financeiro e maior índice de risco.",
        "Antecipar renovação dos contratos relevantes e cobrança preventiva antes da perda efetiva.",
        "Revisar preço e posicionamento dos imóveis com vacância prolongada e preço acima do mercado.",
        "Acompanhar mensalmente o resultado recuperado para comprovar o retorno do plano de ação.",
    ]
    if len(oportunidades_vagas) > 0:
        recomendacoes.append(
            f"{len(oportunidades_vagas)} imóveis vagos exigem revisão imediata de preço."
        )

    rec_texto = "<br/>".join([f"- {texto}" for texto in recomendacoes])
    rec_tbl = Table(
        [[Paragraph(rec_texto, styles["TextoVisual"])]],
        colWidths=[180 * mm],
    )
    rec_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_CLARO),
        ("BOX", (0, 0), (-1, -1), 0.5, BORDA),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(rec_tbl)
    story.append(Spacer(1, 15))
    story.append(Paragraph(
        "Observação: os valores apresentados são estimativas gerenciais baseadas "
        "nos dados fornecidos. Não representam garantia de recuperação ou receita.",
        styles["TextoPequenoVisual"],
    ))

    def adicionar_rodape(canvas, doc_obj):
        canvas.saveState()
        largura, _ = A4
        canvas.setStrokeColor(BORDA)
        canvas.setLineWidth(0.4)
        canvas.line(12 * mm, 11 * mm, largura - 12 * mm, 11 * mm)
        canvas.setFont("Helvetica", 7.2)
        canvas.setFillColor(CINZA)
        canvas.drawString(
            12 * mm,
            7 * mm,
            "TechDadosBR Radar 360 - Relatório Executivo",
        )
        canvas.drawRightString(
            largura - 12 * mm,
            7 * mm,
            f"Página {doc_obj.page}",
        )
        canvas.restoreState()

    doc.build(
        story,
        onFirstPage=adicionar_rodape,
        onLaterPages=adicionar_rodape,
    )
    buffer.seek(0)
    return buffer.getvalue()


def preparar_dados(abas):
    imoveis = abas["Imoveis_Contratos"].copy()
    proprietarios = abas["Proprietarios"].copy()

    imoveis["Fim_Contrato"] = pd.to_datetime(imoveis["Fim_Contrato"], errors="coerce")
    imoveis["Ultimo_Reajuste"] = pd.to_datetime(imoveis["Ultimo_Reajuste"], errors="coerce")

    hoje = pd.Timestamp.today().normalize()

    imoveis["Dias_Para_Vencimento"] = (imoveis["Fim_Contrato"] - hoje).dt.days
    imoveis["Meses_Sem_Reajuste"] = (
        (hoje - imoveis["Ultimo_Reajuste"]).dt.days / 30.4
    ).fillna(0)

    # ÍNDICE DE RISCO DO IMÓVEL — 0 a 100
    # O objetivo não é repetir ocorrências operacionais, mas combinar sinais
    # financeiros, contratuais, comerciais e operacionais em uma leitura gerencial.

    # 1) Financeiro — até 35 pontos
    imoveis["Risco_Financeiro"] = 0
    imoveis.loc[
        imoveis["Perfil_Pagamento"] == "Atraso recorrente",
        "Risco_Financeiro",
    ] += 25
    imoveis.loc[
        imoveis["Perfil_Pagamento"] == "Atraso ocasional",
        "Risco_Financeiro",
    ] += 12
    imoveis.loc[
        imoveis["Perda_Mensal_Vacancia"] >= 4000,
        "Risco_Financeiro",
    ] += 10
    imoveis.loc[
        imoveis["Perda_Mensal_Vacancia"].between(2000, 3999.99),
        "Risco_Financeiro",
    ] += 6
    imoveis["Risco_Financeiro"] = imoveis["Risco_Financeiro"].clip(upper=35)

    # 2) Contratual — até 25 pontos
    imoveis["Risco_Contratual"] = 0
    imoveis.loc[
        imoveis["Dias_Para_Vencimento"].between(0, 30),
        "Risco_Contratual",
    ] += 25
    imoveis.loc[
        imoveis["Dias_Para_Vencimento"].between(31, 60),
        "Risco_Contratual",
    ] += 20
    imoveis.loc[
        imoveis["Dias_Para_Vencimento"].between(61, 90),
        "Risco_Contratual",
    ] += 15
    imoveis.loc[
        imoveis["Meses_Sem_Reajuste"] >= 24,
        "Risco_Contratual",
    ] += 10
    imoveis.loc[
        imoveis["Meses_Sem_Reajuste"].between(18, 23.99),
        "Risco_Contratual",
    ] += 6
    imoveis["Risco_Contratual"] = imoveis["Risco_Contratual"].clip(upper=25)

    # 3) Comercial — até 25 pontos
    imoveis["Risco_Comercial"] = 0
    imoveis.loc[imoveis["Dias_Vago"] >= 120, "Risco_Comercial"] += 25
    imoveis.loc[
        imoveis["Dias_Vago"].between(90, 119),
        "Risco_Comercial",
    ] += 20
    imoveis.loc[
        imoveis["Dias_Vago"].between(45, 89),
        "Risco_Comercial",
    ] += 12
    imoveis.loc[
        imoveis["Diferenca_Preco_Mercado"] >= 0.15,
        "Risco_Comercial",
    ] += 10
    imoveis.loc[
        imoveis["Diferenca_Preco_Mercado"].between(0.10, 0.149999),
        "Risco_Comercial",
    ] += 6
    imoveis.loc[
        (imoveis["Status_Imovel"] == "Vago")
        & (imoveis["Visitas_60d"] == 0),
        "Risco_Comercial",
    ] += 8
    imoveis.loc[
        (imoveis["Status_Imovel"] == "Vago")
        & (imoveis["Visitas_60d"] >= 5)
        & (imoveis["Propostas_60d"] == 0),
        "Risco_Comercial",
    ] += 8
    imoveis["Risco_Comercial"] = imoveis["Risco_Comercial"].clip(upper=25)

    # 4) Operacional — até 15 pontos
    imoveis["Risco_Operacional"] = 0
    imoveis.loc[imoveis["Chamados_12m"] >= 7, "Risco_Operacional"] = 15
    imoveis.loc[
        imoveis["Chamados_12m"].between(5, 6),
        "Risco_Operacional",
    ] = 10
    imoveis.loc[
        imoveis["Chamados_12m"].between(3, 4),
        "Risco_Operacional",
    ] = 5

    imoveis["Indice_Risco_Imovel"] = (
        imoveis["Risco_Financeiro"]
        + imoveis["Risco_Contratual"]
        + imoveis["Risco_Comercial"]
        + imoveis["Risco_Operacional"]
    ).clip(upper=100)

    dimensoes = {
        "Financeiro": "Risco_Financeiro",
        "Contratual": "Risco_Contratual",
        "Comercial": "Risco_Comercial",
        "Operacional": "Risco_Operacional",
    }

    imoveis["Fator_Principal"] = imoveis.apply(
        lambda linha: max(
            dimensoes,
            key=lambda nome: linha[dimensoes[nome]],
        ),
        axis=1,
    )

    imoveis["Classificacao_Risco"] = pd.cut(
        imoveis["Indice_Risco_Imovel"],
        bins=[-1, 29, 59, 79, 100],
        labels=["Baixo", "Atenção", "Alto", "Crítico"],
    )

    return imoveis, proprietarios


def aplicar_tema_dataframe(df, tema_atual):
    if tema_atual == "Escuro":
        fundo = "#14324A"
        fundo_alterno = "#1A405A"
        cabecalho = "#1B4965"
        texto = "#F4F8F9"
        borda = "#486B7A"
    else:
        fundo = "#EEF5F7"
        fundo_alterno = "#DDECEF"
        cabecalho = "#163A5F"
        texto = "#14324A"
        borda = "#B8CDD8"

    return (
        df.style
        .set_properties(**{
            "background-color": fundo,
            "color": texto,
            "border-color": borda,
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", cabecalho),
                    ("color", "#FFFFFF"),
                    ("border-color", borda),
                    ("font-weight", "700"),
                ],
            },
            {
                "selector": "tbody tr:nth-child(even) td",
                "props": [
                    ("background-color", fundo_alterno),
                ],
            },
        ])
    )


def renderizar_tabela_azul(df, tema_atual, altura_max=420):
    if df is None or df.empty:
        st.info("Nenhum registro disponível.")
        return

    escuro = tema_atual == "Escuro"

    cabecalho = "#1B4965" if escuro else "#163A5F"
    linha_1 = "#14324A" if escuro else "#EEF5F7"
    linha_2 = "#1A405A" if escuro else "#DDECEF"
    texto = "#F4F8F9" if escuro else "#14324A"
    borda = "#486B7A" if escuro else "#B8CDD8"

    cabecalhos = "".join(
        f"<th>{html.escape(str(coluna))}</th>"
        for coluna in df.columns
    )

    linhas = []
    for indice, (_, row) in enumerate(df.iterrows()):
        fundo = linha_1 if indice % 2 == 0 else linha_2
        celulas = "".join(
            f'<td style="background:{fundo};">{html.escape(str(valor)) if pd.notna(valor) else "-"}</td>'
            for valor in row
        )
        linhas.append(f"<tr>{celulas}</tr>")

    tabela = f"""
    <div class="tabela-azul-wrap" style="
        max-height:{altura_max}px;
        overflow:auto;
        border:1px solid {borda};
        border-radius:10px;
        background:{linha_1};
    ">
        <table class="tabela-azul" style="
            width:100%;
            min-width:980px;
            border-collapse:collapse;
            table-layout:auto;
        ">
            <thead>
                <tr style="background:{cabecalho};">
                    {cabecalhos}
                </tr>
            </thead>
            <tbody>
                {''.join(linhas)}
            </tbody>
        </table>
    </div>
    <style>
        .tabela-azul th {{
            position:sticky;
            top:0;
            z-index:2;
            background:{cabecalho} !important;
            color:#FFFFFF !important;
            text-align:left;
            font-weight:700;
            padding:11px 10px;
            border-right:1px solid {borda};
            border-bottom:1px solid {borda};
            white-space:nowrap;
            font-size:0.88rem;
        }}
        .tabela-azul td {{
            color:{texto} !important;
            padding:9px 10px;
            border-right:1px solid {borda};
            border-bottom:1px solid {borda};
            vertical-align:top;
            font-size:0.84rem;
            line-height:1.3;
            white-space:normal;
            overflow-wrap:anywhere;
        }}
        .tabela-azul th:last-child,
        .tabela-azul td:last-child {{
            min-width:150px;
            white-space:nowrap;
        }}
    </style>
    """

    st.markdown(textwrap.dedent(tabela), unsafe_allow_html=True)


def card_indicador(titulo, valor, subtitulo, destaque):
    titulo_normalizado = titulo.lower()

    if any(
        termo in titulo_normalizado
        for termo in [
            "risco",
            "crítico",
            "críticas",
            "exposto",
            "impacto mensal priorizado",
        ]
    ):
        classe_valor = "valor-negativo"
    elif any(
        termo in titulo_normalizado
        for termo in [
            "potencial",
            "recuperado",
            "concluída",
            "concluídas",
        ]
    ):
        classe_valor = "valor-positivo"
    elif any(
        termo in titulo_normalizado
        for termo in [
            "atenção",
            "alta prioridade",
            "vacâncias para revisão",
            "proprietários",
        ]
    ):
        classe_valor = "valor-atencao"
    else:
        classe_valor = "valor-neutro"

    st.markdown(
        f"""
        <div class="radar-card" style="border-top: 4px solid {destaque};">
            <div class="radar-card-title">{titulo}</div>
            <div class="radar-card-value {classe_valor}">{valor}</div>
            <div class="radar-card-subtitle">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def formatar_multilinha(texto):
    if pd.isna(texto) or str(texto).strip() == "":
        return "-"

    resumos = {
        "Contrato próximo do vencimento": "Contrato vencendo",
        "Excesso de manutenção": "Manutenção recorrente",
        "Vacância prolongada": "Vacância prolongada",
        "Atrasos recorrentes": "Atrasos recorrentes",
        "Preço acima do mercado": "Preço acima do mercado",
        "Iniciar contato preventivo e revisar histórico de pagamento": "Contato preventivo e revisão do histórico",
        "Iniciar negociação de renovação": "Negociar renovação",
        "Analisar causa recorrente e custo operacional": "Revisar causa e custo operacional",
        "Reavaliar preço com o proprietário": "Reavaliar preço com o proprietário",
        "Revisar preço, anúncio e estratégia comercial": "Revisar preço e estratégia comercial",
        "Manter monitoramento": "Manter monitoramento",
    }

    partes = [p.strip() for p in str(texto).split("|") if p.strip()]
    if not partes:
        return "-"

    partes = [resumos.get(parte, parte) for parte in partes]
    return "<br>".join([f"• {html.escape(p)}" for p in partes])


def renderizar_tabela_prioridades(df):
    linhas_html = []

    for _, row in df.iterrows():
        linha = f"""
<tr>
    <td>{html.escape(str(row["Imóvel"]))}</td>
    <td>{html.escape(str(row["Cidade"]))}</td>
    <td>{html.escape(str(row["Bairro"]))}</td>
    <td class="col-risco">{html.escape(str(row["Índice"]))}</td>
    <td class="col-faixa">{html.escape(str(row["Faixa"]))}</td>
    <td class="col-fator">{html.escape(str(row["Fator principal"]))}</td>
    <td class="col-motivo">{formatar_multilinha(row["Motivo"])}</td>
    <td class="col-acao">{formatar_multilinha(row["Ação recomendada"])}</td>
    <td class="col-impacto">{html.escape(str(row["Impacto mensal"]))}</td>
</tr>
"""
        linhas_html.append(textwrap.dedent(linha).strip())

    tabela_html = f"""
<div class="prioridades-html-wrap">
<table class="prioridades-html-table">
<thead>
<tr>
    <th class="col-imovel">Imóvel</th>
    <th class="col-cidade">Cidade</th>
    <th class="col-bairro">Bairro</th>
    <th class="col-risco">Índice</th>
    <th class="col-faixa">Faixa</th>
    <th class="col-fator">Fator principal</th>
    <th class="col-motivo">Motivo</th>
    <th class="col-acao">Ação recomendada</th>
    <th class="col-impacto">Impacto mensal</th>
</tr>
</thead>
<tbody>
{"".join(linhas_html)}
</tbody>
</table>
</div>
"""

    st.markdown(textwrap.dedent(tabela_html).strip(), unsafe_allow_html=True)



st.markdown(
    """
    <style>
    .block-container {
        padding-top: 0 !important;
        padding-bottom: 2rem;
        margin-top: -2.15rem !important;
        max-width: 100%;
    }

    [data-testid="stHeader"],
    header[data-testid="stHeader"] {
        height: 1.55rem !important;
        min-height: 1.55rem !important;
    }

    @media (min-width: 1200px) {
        .block-container {
            padding-left: 2.2rem;
            padding-right: 2.2rem;
        }
    }

    @media (max-width: 1199px) {
        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }
    }

    h1, h2, h3 {
        color: #EAF2F8;
    }

    .radar-card {
        background: linear-gradient(145deg, #182536, #111B28);
        border: 1px solid #2D4055;
        border-radius: 14px;
        padding: 17px 18px 15px 18px;
        min-height: 112px;
        box-shadow: 0 6px 18px rgba(0, 0, 0, 0.20);
    }

    .radar-card-title {
        color: #AFC2D4;
        font-size: 0.83rem;
        font-weight: 600;
        margin-bottom: 8px;
    }

    .radar-card-value {
        color: #FFFFFF;
        font-size: 1.65rem;
        font-weight: 750;
        line-height: 1.1;
        margin-bottom: 7px;
    }

    .radar-card-subtitle {
        color: #7F96AA;
        font-size: 0.73rem;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid #26384A;
        border-radius: 10px;
        overflow: auto;
        max-width: 100%;
    }

    .diagnostico-fechamento {
        background: #F4F8F9;
        border: 1px solid #D9E2EC;
        border-radius: 10px;
        padding: 12px 14px;
        color: #66788A;
        font-size: 0.88rem;
        line-height: 1.32;
    }

    .diagnostico-card {
        background: #F4F8F9;
        border: 1px solid #D9E2EC;
        border-left: 4px solid #2B6F89;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 12px;
        color: #203746;
    }

    .diagnostico-card-titulo {
        font-weight: 700;
        font-size: 1rem;
        margin-bottom: 5px;
    }

    .diagnostico-card-texto {
        font-size: 0.92rem;
        line-height: 1.45;
        margin-bottom: 6px;
    }

    .diagnostico-card-acao {
        font-size: 0.88rem;
        color: #5F6F82;
    }

    .consolidado-linha {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        align-items: center;
        background: #F4F8F9;
        border: 1px solid #D9E2EC;
        border-radius: 9px;
        padding: 10px 12px;
        margin-bottom: 14px;
        color: #203746;
        font-size: 0.88rem;
    }

    .consolidado-linha span {
        white-space: nowrap;
    }

    .pagina-interna-topo {
        height: 0 !important;
        min-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 0 !important;
    }

    .cabecalho-interno {
        display: flex;
        align-items: center;
        gap: 9px;
        margin: -2.10rem 0 2px 0;
        color: #203746;
        line-height: 1.1;
    }

    .cabecalho-interno-marca {
        font-size: 1.55rem;
        font-weight: 760;
    }

    .cabecalho-interno-separador {
        color: #8FA2AB;
        font-size: 1.08rem;
    }

    .cabecalho-interno-pagina {
        font-size: 1.55rem;
        font-weight: 720;
    }

    .cabecalho-interno-subtitulo {
        color: #7A8794;
        font-size: 0.88rem;
        margin: 0 0 12px 0;
    }

    .cabecalho-visao {
        display: flex;
        align-items: center;
        gap: 10px;
        margin: -1.15rem 0 4px 0;
        line-height: 1.1;
    }

    .cabecalho-visao-icone {
        font-size: 1.55rem;
    }

    .cabecalho-visao-titulo {
        font-size: 2rem;
        font-weight: 760;
        color: #163A5F;
    }

    .indice-linha {
        display: flex;
        flex-wrap: wrap;
        gap: 14px;
        align-items: center;
        background: #F4F8F9;
        border: 1px solid #D9E2EC;
        border-radius: 9px;
        padding: 10px 12px;
        margin-top: 6px;
        color: #203746;
        font-size: 0.88rem;
    }

    .indice-linha span {
        white-space: nowrap;
    }

    .indice-linha-lateral {
        margin-top: 0;
        position: relative;
        top: -8px;
        height: 50px;
        min-height: 50px;
        box-sizing: border-box;
        padding: 0 14px;
        display: flex;
        align-items: center;
    }

    @media (max-width: 1100px) {
        .indice-linha-lateral {
            top: 0;
        }
    }

    .indice-legenda {
        background: #F7FAFC;
        border: 1px solid #D8E1EA;
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 16px;
        color: #203746;
    }

    .legenda-faixas {
        display: flex;
        flex-wrap: wrap;
        gap: 18px;
        margin: 8px 0 6px 0;
        font-size: 0.92rem;
    }

    .legenda-explicacao {
        color: #66788A;
        font-size: 0.86rem;
        line-height: 1.45;
    }

    .prioridades-html-wrap {
        width: 100%;
        overflow-x: auto;
        border: 1px solid #9DB7CF;
        border-radius: 12px;
        background: #EEF5F7;
    }

    .prioridades-html-table {
        width: 100%;
        min-width: 1040px;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 0.86rem;
    }

    .prioridades-html-table th,
    .prioridades-html-table td {
        border-bottom: 1px solid #E5E7EB;
        border-right: 1px solid #E5E7EB;
        padding: 8px 8px;
        vertical-align: top;
        text-align: left;
        color: #14324A;
        background: #EEF5F7;
        line-height: 1.35;
        word-wrap: break-word;
        overflow-wrap: break-word;
        white-space: normal;
    }

    .prioridades-html-table th:last-child,
    .prioridades-html-table td:last-child {
        border-right: none;
    }

    .prioridades-html-table thead th {
        background: #163A5F;
        color: #FFFFFF;
        font-weight: 700;
        position: sticky;
        top: 0;
        z-index: 1;
    }

    .prioridades-html-table tbody tr:nth-child(even) td {
        background: #DDECEF;
    }

    .prioridades-html-table tbody tr:hover td {
        background: #CFE2F3;
    }

    .prioridades-html-table .col-imovel { width: 7%; }
    .prioridades-html-table .col-cidade { width: 8%; }
    .prioridades-html-table .col-bairro { width: 10%; }
    .prioridades-html-table .col-risco  { width: 5%; text-align: center; white-space: nowrap; }
    .prioridades-html-table .col-faixa  { width: 6%; }
    .prioridades-html-table .col-fator  { width: 8%; }
    .prioridades-html-table .col-motivo { width: 16%; font-size: 0.78rem; line-height: 1.25; }
    .prioridades-html-table .col-acao   { width: 22%; font-size: 0.78rem; line-height: 1.25; }
    .prioridades-html-table .col-impacto{ width: 18%; min-width: 145px; white-space: nowrap; }

    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('''
<style>
:root {
    --radar-azul-petroleo: #163A5F;
    --radar-azul-medio: #2B6F89;
    --radar-esmeralda: #178A64;
    --radar-dourado: #C28A2C;
    --radar-vermelho: #B84343;
}

.radar-card {
    border-width: 1px !important;
}

.radar-card[style*="#B84343"] {
    box-shadow: 0 8px 20px rgba(184, 67, 67, 0.14) !important;
}

.radar-card[style*="#178A64"] {
    box-shadow: 0 8px 20px rgba(23, 138, 100, 0.14) !important;
}

.radar-card[style*="#C28A2C"] {
    box-shadow: 0 8px 20px rgba(194, 138, 44, 0.14) !important;
}

.radar-card[style*="#2B6F89"] {
    box-shadow: 0 8px 20px rgba(43, 111, 137, 0.14) !important;
}
</style>
''', unsafe_allow_html=True)


if not os.path.exists(CAMINHO_BASE):
    st.error(
        "Base não encontrada. Confirme se a planilha está na pasta "
        "01_Base_Teste com o nome correto."
    )
    st.stop()

try:
    abas = carregar_base(CAMINHO_BASE)
    imoveis, proprietarios = preparar_dados(abas)
except Exception as erro:
    st.error(f"Não foi possível carregar a base: {erro}")
    st.stop()

verificar_acesso_cliente()

tema = st.sidebar.radio(
    "Tema",
    ["Claro", "Escuro"],
    index=0,
    horizontal=True,
)

if tema == "Escuro":
    st.markdown(
        """
        <style>
        .stApp {
            background: #102634;
            color: #F4F8F9;
        }

        [data-testid="stHeader"],
        header[data-testid="stHeader"] {
            background: #102634 !important;
            border-bottom: 0 !important;
            height: 2rem !important;
            min-height: 2rem !important;
        }

        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"] {
            background: transparent !important;
        }

        [data-testid="stHeader"] * {
            color: #E4EEF1 !important;
        }

        .block-container {
            padding-top: 0 !important;
            margin-top: -0.85rem !important;
        }

        .pagina-interna-topo {
            height: 0 !important;
        }

        .cabecalho-interno {
            margin-top: -2.10rem !important;
        }

        .cabecalho-interno-subtitulo {
            margin-bottom: 10px !important;
        }

        [data-testid="stSidebar"] {
            background: #143142;
            border-right: 1px solid #263449;
        }

        [data-testid="stSidebar"] * {
            color: #F4F8F9 !important;
        }

        h1, h2, h3, h4, h5, h6,
        p, label, span {
            color: #F4F8F9;
        }

        .cabecalho-visao-titulo {
            color: #F4F8F9 !important;
        }

        .cabecalho-visao-icone {
            color: #F4F8F9 !important;
        }

        [data-testid="stCaptionContainer"] p,
        .stCaption,
        .cabecalho-interno-subtitulo {
            color: #B8CAD6 !important;
        }

        hr {
            border-color: #263449 !important;
        }

        /* Cards */
        .radar-card {
            background: linear-gradient(145deg, #172235, #143142) !important;
            border-color: #3A5D6B !important;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28) !important;
        }

        .radar-card-title,
        .radar-card-subtitle {
            color: #C9D8DE !important;
        }

        .radar-card-value {
            font-weight: 800 !important;
        }

        .radar-card[style*="#B84343"],
        .radar-card[style*="#B84343"] {
            border-top-color: #D66A6A !important;
        }

        .radar-card[style*="#178A64"],
        .radar-card[style*="#178A64"] {
            border-top-color: #2FB27C !important;
        }

        .radar-card[style*="#C28A2C"],
        .radar-card[style*="#C28A2C"] {
            border-top-color: #D9A441 !important;
        }

        .radar-card[style*="#2B6F89"],
        .radar-card[style*="#2B6F89"] {
            border-top-color: #4C8FA6 !important;
        }

        /* Blocos internos */
        .indice-legenda,
        .indice-linha,
        .consolidado-linha,
        .diagnostico-card,
        .diagnostico-fechamento,
        .prioridades-html-wrap {
            background: #14324A !important;
            border-color: #3A5D6B !important;
            color: #F4F8F9 !important;
        }

        .diagnostico-card {
            border-left-color: #2B6F89 !important;
        }

        .cabecalho-interno,
        .cabecalho-interno-marca,
        .cabecalho-interno-pagina,
        .diagnostico-card-titulo,
        .diagnostico-card-texto {
            color: #F4F8F9 !important;
        }

        .cabecalho-interno-subtitulo,
        .legenda-explicacao,
        .diagnostico-card-acao,
        .diagnostico-fechamento {
            color: #C9D8DE !important;
        }

        /* Tabela HTML */
        .prioridades-html-table th {
            background: #1B4965 !important;
            color: #E4EEF1 !important;
            border-color: #3A5D6B !important;
        }

        .prioridades-html-table td {
            background: #14324A !important;
            color: #F4F8F9 !important;
            border-color: #3A5D6B !important;
        }

        .prioridades-html-table tbody tr:hover td {
            background: #1A405A !important;
        }

        /* Dataframes */
        [data-testid="stDataFrame"] {
            background: #14324A !important;
            border-color: #3A5D6B !important;
        }

        [data-testid="stDataFrame"] > div {
            background: #14324A !important;
        }

        [data-testid="stDataFrame"] canvas {
            filter: brightness(0.80) contrast(1.12);
        }

        /* Expanders */
        [data-testid="stExpander"] {
            background: #14324A !important;
            border-color: #3A5D6B !important;
        }

        [data-testid="stExpander"] details {
            background: #14324A !important;
        }

        [data-testid="stExpander"] summary {
            background: #1A405A !important;
            color: #F4F8F9 !important;
            border-color: #3A5D6B !important;
        }

        [data-testid="stExpander"] summary:hover {
            background: #1B4965 !important;
        }

        [data-testid="stExpander"] summary svg {
            fill: #E4EEF1 !important;
            color: #E4EEF1 !important;
        }

        /* Inputs */
        [data-baseweb="select"] > div,
        [data-baseweb="input"] > div,
        [data-baseweb="textarea"] > div,
        [data-testid="stDateInput"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stTextInput"] input,
        [data-testid="stTextArea"] textarea {
            background: #1B4965 !important;
            color: #F4F8F9 !important;
            border-color: #557581 !important;
        }

        [data-baseweb="select"] * {
            color: #F4F8F9 !important;
        }

        [data-baseweb="popover"] {
            background: #1B4965 !important;
        }

        [role="listbox"] {
            background: #1B4965 !important;
            color: #F4F8F9 !important;
        }

        [role="option"] {
            color: #F4F8F9 !important;
        }

        [role="option"]:hover {
            background: #3A5D6B !important;
        }

        input::placeholder,
        textarea::placeholder {
            color: #8FA2AB !important;
            opacity: 1 !important;
        }

        /* Botões */
        [data-testid="stFormSubmitButton"] button,
        [data-testid="stDownloadButton"] button,
        [data-testid="stButton"] button {
            background: #2B6F89 !important;
            color: #FFFFFF !important;
            border: 1px solid #2B6F89 !important;
            font-weight: 600 !important;
        }

        [data-testid="stFormSubmitButton"] button:hover,
        [data-testid="stDownloadButton"] button:hover,
        [data-testid="stButton"] button:hover {
            background: #1D4ED8 !important;
            border-color: #4C8FA6 !important;
            color: #FFFFFF !important;
        }

        [data-testid="stFormSubmitButton"] button:disabled,
        [data-testid="stDownloadButton"] button:disabled,
        [data-testid="stButton"] button:disabled {
            background: #3A5D6B !important;
            border-color: #557581 !important;
            color: #8FA2AB !important;
            opacity: 1 !important;
        }

        /* Alertas */
        [data-testid="stAlert"] {
            background: #1A405A !important;
            color: #F4F8F9 !important;
            border-color: #3A5D6B !important;
        }

        /* Métricas nativas */
        [data-testid="stMetric"] {
            background: #14324A !important;
            border-color: #3A5D6B !important;
        }

        [data-testid="stMetricLabel"],
        [data-testid="stMetricValue"] {
            color: #F4F8F9 !important;
        }

        /* Legendas e captions */
        [data-testid="stCaptionContainer"] p,
        small {
            color: #8FA2AB !important;
        }

        /* Menus suspensos e popovers */
        div[data-baseweb="popover"],
        div[data-baseweb="menu"],
        ul[role="listbox"] {
            background: #1B4965 !important;
            border: 1px solid #557581 !important;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.35) !important;
        }

        li[role="option"],
        div[role="option"] {
            background: #1B4965 !important;
            color: #F4F8F9 !important;
        }

        li[role="option"] *,
        div[role="option"] * {
            color: #F4F8F9 !important;
        }

        li[role="option"]:hover,
        div[role="option"]:hover,
        li[role="option"][aria-selected="true"],
        div[role="option"][aria-selected="true"] {
            background: #3A5D6B !important;
            color: #FFFFFF !important;
        }

        /* Dataframes do Streamlit no tema escuro */
        [data-testid="stDataFrame"] {
            background: #14324A !important;
            border: 1px solid #3A5D6B !important;
            border-radius: 10px !important;
            overflow: hidden !important;
        }

        [data-testid="stDataFrame"] iframe,
        [data-testid="stDataFrame"] > div,
        [data-testid="stDataFrame"] [role="grid"],
        [data-testid="stDataFrame"] [role="row"],
        [data-testid="stDataFrame"] [role="gridcell"] {
            background: #14324A !important;
            color: #E4EEF1 !important;
            border-color: #3A5D6B !important;
        }

        [data-testid="stDataFrame"] [role="columnheader"] {
            background: #1B4965 !important;
            color: #F4F8F9 !important;
            border-color: #3A5D6B !important;
        }

        [data-testid="stDataFrame"] [role="columnheader"] {
            background: #1B4965 !important;
            color: #C9D8DE !important;
        }

        [data-testid="stDataFrame"] canvas {
            background: #162A44 !important;
            filter: none !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <style>
        .stApp {
            background: #FFFFFF;
            color: #1B4052;
        }

        [data-testid="stSidebar"] {
            background: #EDF3F5;
        }

        [data-testid="stSidebar"] * {
            color: #1B4052 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


pagina = st.sidebar.radio(
    "Menu",
    ["Visão Executiva", "Plano de Ação", "Diagnóstico", "Relatório"],
    index=0,
)

if pagina == "Relatório":
    st.markdown('<div class="pagina-interna-topo"></div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="cabecalho-interno">
            <div class="cabecalho-interno-marca">📡 TechDadosBR Radar 360</div>
            <div class="cabecalho-interno-separador">•</div>
            <div class="cabecalho-interno-pagina">Relatório Executivo</div>
        </div>
        <div class="cabecalho-interno-subtitulo">
            Síntese das informações relevantes da Visão Executiva, Diagnóstico,
            Plano de Ação, Oportunidades e Acompanhamento.
        </div>
        """,
        unsafe_allow_html=True,
    )

    acompanhamento_pdf = carregar_acompanhamento(CAMINHO_ACOMPANHAMENTO)
    pdf_bytes = gerar_pdf_executivo(
        imoveis,
        proprietarios,
        acompanhamento_pdf,
    )

    st.markdown(
        """
        <div class="diagnostico-card">
            <div class="diagnostico-card-titulo">Relatório consolidado</div>
            <div class="diagnostico-card-texto">
                O PDF reune exposição financeira, índice de risco, diagnóstico da
                carteira, prioridades com prazo, oportunidades, acompanhamento das
                ações e resultado recuperado.
            </div>
            <div class="diagnostico-card-acao">
                O detalhamento completo permanece no aplicativo para evitar um
                relatório operacional extenso.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.download_button(
        "Baixar Relatório Executivo em PDF",
        data=pdf_bytes,
        file_name="Relatorio_Executivo_Radar_360.pdf",
        mime="application/pdf",
        use_container_width=False,
    )

    st.caption(
        "O PDF utiliza os dados atuais da base e os registros de acompanhamento salvos."
    )
    st.stop()

if pagina == "Diagnóstico":
    st.markdown('<div class="pagina-interna-topo"></div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="cabecalho-interno">
            <div class="cabecalho-interno-marca">📡 TechDadosBR Radar 360</div>
            <div class="cabecalho-interno-separador">•</div>
            <div class="cabecalho-interno-pagina">Diagnóstico Executivo</div>
        </div>
        <div class="cabecalho-interno-subtitulo">
            Leitura gerencial dos principais riscos, perdas e oportunidades da carteira.
        </div>
        """,
        unsafe_allow_html=True,
    )

    perda_vacancia_diag = imoveis["Perda_Mensal_Vacancia"].sum()

    contratos_diag = imoveis[
        imoveis["Dias_Para_Vencimento"].between(0, 90)
    ]

    atrasos_diag = imoveis[
        imoveis["Perfil_Pagamento"] == "Atraso recorrente"
    ]

    imoveis_alto_risco_diag = imoveis[
        imoveis["Classificacao_Risco"].isin(["Alto", "Crítico"])
    ]

    proprietarios_diag = proprietarios[
        proprietarios["Risco_Relacionamento"] >= 70
    ]

    mascara_receita_diag = (
        imoveis["Dias_Para_Vencimento"].between(0, 90)
        | imoveis["Perfil_Pagamento"].eq("Atraso recorrente")
    )

    receita_imobiliaria_diag = imoveis.loc[
        mascara_receita_diag,
        "Receita_Mensal_Imobiliaria",
    ].sum()

    valor_locativo_diag = (
        perda_vacancia_diag
        + atrasos_diag["Aluguel_Mensal"].sum()
    )

    d1, d2, d3 = st.columns(3)

    with d1:
        card_indicador(
            "Valor locativo exposto",
            moeda(valor_locativo_diag),
            "Vacância mensal + aluguéis ligados a atraso recorrente",
            "#C28A2C",
        )

    with d2:
        card_indicador(
            "Receita da imobiliária exposta",
            moeda(receita_imobiliaria_diag),
            "Vencimento em 90 dias ou atraso recorrente",
            "#B84343",
        )

    with d3:
        card_indicador(
            "Imóveis com risco elevado",
            str(len(imoveis_alto_risco_diag)),
            "Faixas Alto e Crítico do índice",
            "#B84343",
        )

    st.divider()

    st.subheader("Leitura executiva")

    diagnosticos_executivos = []

    if perda_vacancia_diag > 0:
        top_vacancia_diag = imoveis[
            imoveis["Status_Imovel"] == "Vago"
        ].sort_values(
            "Perda_Mensal_Vacancia",
            ascending=False,
        ).head(5)

        concentracao_diag = (
            top_vacancia_diag["Perda_Mensal_Vacancia"].sum()
            / perda_vacancia_diag
            * 100
        )

        diagnosticos_executivos.append(
            (
                "Vacância concentrada",
                f"Os cinco imóveis com maior perda concentram "
                f"{concentracao_diag:.1f}% da perda mensal por vacância.",
                "Revisar preço e estratégia comercial desses imóveis primeiro.",
            )
        )

    if len(contratos_diag) > 0:
        diagnosticos_executivos.append(
            (
                "Receita contratual em risco",
                f"{len(contratos_diag)} contratos vencem nos próximos 90 dias, "
                f"com impacto mensal potencial de "
                f"{moeda(contratos_diag['Receita_Mensal_Imobiliaria'].sum())}.",
                "Iniciar renovação pelos contratos de maior impacto financeiro.",
            )
        )

    if len(atrasos_diag) > 0:
        diagnosticos_executivos.append(
            (
                "Deterioração de pagamento",
                f"{len(atrasos_diag)} imóveis apresentam atraso recorrente, "
                f"associados a {moeda(atrasos_diag['Aluguel_Mensal'].sum())} "
                f"em valor locativo mensal.",
                "Atuar preventivamente antes da inadimplência se agravar.",
            )
        )

    if len(proprietarios_diag) == 1:
        diagnosticos_executivos.append(
            (
                "Relacionamentos a proteger",
                "1 proprietário apresenta risco elevado de relacionamento.",
                "Priorizar contato gerencial com esse proprietário.",
            )
        )
    elif len(proprietarios_diag) > 1:
        diagnosticos_executivos.append(
            (
                "Relacionamentos a proteger",
                f"{len(proprietarios_diag)} proprietários apresentam "
                f"risco elevado de relacionamento.",
                "Priorizar contato gerencial com os proprietários mais relevantes.",
            )
        )

    for titulo, leitura, acao in diagnosticos_executivos:
        st.markdown(
            f"""
            <div class="diagnostico-card">
                <div class="diagnostico-card-titulo">{titulo}</div>
                <div class="diagnostico-card-texto">{leitura}</div>
                <div class="diagnostico-card-acao"><strong>Ação:</strong> {acao}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.divider()

    st.markdown(
        """
        <div class="diagnostico-fechamento">
            O detalhamento por imóvel e as ações específicas permanecem na página
            <strong>Plano de Ação</strong>. Esta página apresenta apenas a leitura
            consolidada da carteira, evitando repetição de informações.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.stop()

if pagina == "Plano de Ação":
    st.markdown('<div class="pagina-interna-topo"></div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="cabecalho-interno">
            <div class="cabecalho-interno-marca">📡 TechDadosBR Radar 360</div>
            <div class="cabecalho-interno-separador">•</div>
            <div class="cabecalho-interno-pagina">Plano de Ação</div>
        </div>
        <div class="cabecalho-interno-subtitulo">
            Ações priorizadas por urgência e impacto financeiro.
        </div>
        """,
        unsafe_allow_html=True,
    )

    plano = imoveis.copy()

    plano["Impacto_Plano"] = pd.to_numeric(
        plano["Perda_Mensal_Vacancia"],
        errors="coerce",
    ).fillna(0).astype(float)

    mascara_receita_plano = (
        plano["Perfil_Pagamento"].eq("Atraso recorrente")
        | plano["Dias_Para_Vencimento"].between(0, 90)
    )

    plano.loc[
        mascara_receita_plano,
        "Impacto_Plano",
    ] += plano.loc[
        mascara_receita_plano,
        "Receita_Mensal_Imobiliaria",
    ]

    def definir_acao_principal(row):
        if row["Dias_Vago"] >= 90 and row["Diferenca_Preco_Mercado"] >= 0.10:
            return "Revisar preço e reposicionar o anúncio"
        if row["Perfil_Pagamento"] == "Atraso recorrente":
            return "Iniciar contato preventivo de cobrança"
        if pd.notna(row["Dias_Para_Vencimento"]) and 0 <= row["Dias_Para_Vencimento"] <= 90:
            return "Iniciar negociação de renovação"
        if row["Meses_Sem_Reajuste"] >= 18:
            return "Revisar reajuste contratual"
        if row["Chamados_12m"] >= 5:
            return "Analisar causa e custo das manutenções"
        if row["Diferenca_Preco_Mercado"] >= 0.10:
            return "Reavaliar preço com o proprietário"
        return "Manter monitoramento"

    def definir_prazo(row):
        if row["Indice_Risco_Imovel"] >= 80:
            return "Até 2 dias"
        if row["Indice_Risco_Imovel"] >= 60:
            return "Até 5 dias"
        if row["Indice_Risco_Imovel"] >= 30:
            return "Até 10 dias"
        return "Até 30 dias"

    plano["Ação prioritária"] = plano.apply(definir_acao_principal, axis=1)
    plano["Prazo sugerido"] = plano.apply(definir_prazo, axis=1)

    plano_priorizado = plano[
        plano["Indice_Risco_Imovel"] >= 30
    ].sort_values(
        ["Indice_Risco_Imovel", "Impacto_Plano"],
        ascending=[False, False],
    ).head(12)

    total_impacto_plano = plano_priorizado["Impacto_Plano"].sum()
    criticas_plano = (plano_priorizado["Indice_Risco_Imovel"] >= 80).sum()
    altas_plano = plano_priorizado["Indice_Risco_Imovel"].between(60, 79).sum()

    p1, p2, p3 = st.columns(3)

    with p1:
        card_indicador(
            "Impacto mensal priorizado",
            moeda(total_impacto_plano),
            "Vacância + receita da imobiliária associada a atraso ou vencimento",
            "#A63D40",
        )

    with p2:
        if int(criticas_plano) == 0:
            card_indicador(
                "Ações críticas",
                "Nenhuma",
                "Nenhuma ação crítica identificada",
                "#8E2F3F",
            )
        else:
            card_indicador(
                "Ações críticas",
                str(int(criticas_plano)),
                "Prazo sugerido de até 2 dias",
                "#B84343",
            )

    with p3:
        card_indicador(
            "Ações de alta prioridade",
            str(int(altas_plano)),
            "Prazo sugerido de até 5 dias",
            "#C28A2C",
        )

    st.divider()

    tabela_plano = plano_priorizado[
        [
            "ID_Imovel",
            "Cidade",
            "Bairro",
            "Indice_Risco_Imovel",
            "Fator_Principal",
            "Ação prioritária",
            "Prazo sugerido",
            "Impacto_Plano",
        ]
    ].copy()

    tabela_plano.columns = [
        "Imóvel",
        "Cidade",
        "Bairro",
        "Índice",
        "Fator principal",
        "Ação prioritária",
        "Prazo sugerido",
        "Impacto mensal",
    ]

    tabela_plano["Impacto mensal"] = tabela_plano[
        "Impacto mensal"
    ].apply(moeda)

    renderizar_tabela_azul(
        tabela_plano,
        tema,
        altura_max=470,
    )

    st.divider()

    acompanhamento_consolidado = carregar_acompanhamento(
        CAMINHO_ACOMPANHAMENTO
    )

    with st.expander("Ver acompanhamento consolidado", expanded=False):
        if acompanhamento_consolidado.empty:
            st.info("Ainda não existem ações registradas para acompanhamento.")
        else:
            acompanhamento_consolidado["Data_Limite"] = pd.to_datetime(
                acompanhamento_consolidado["Data_Limite"],
                errors="coerce",
            )
            acompanhamento_consolidado["Resultado_Recuperado"] = pd.to_numeric(
                acompanhamento_consolidado["Resultado_Recuperado"],
                errors="coerce",
            ).fillna(0)

            hoje_acompanhamento = pd.Timestamp.today().normalize()

            status_series = acompanhamento_consolidado["Status"].fillna(
                "Não iniciado"
            )

            qtd_nao_iniciado = int(
                (status_series == "Não iniciado").sum()
            )
            qtd_em_andamento = int(
                (status_series == "Em andamento").sum()
            )
            qtd_concluido = int(
                (status_series == "Concluído").sum()
            )

            mascara_atrasadas = (
                acompanhamento_consolidado["Data_Limite"].notna()
                & (
                    acompanhamento_consolidado["Data_Limite"]
                    < hoje_acompanhamento
                )
                & ~status_series.isin(["Concluído", "Cancelado"])
            )

            qtd_atrasadas = int(mascara_atrasadas.sum())
            total_recuperado = float(
                acompanhamento_consolidado["Resultado_Recuperado"].sum()
            )

            st.markdown(
                f"""
                <div class="consolidado-linha">
                    <span><b>Não iniciadas</b> {qtd_nao_iniciado}</span>
                    <span><b>Em andamento</b> {qtd_em_andamento}</span>
                    <span><b>Concluídas</b> {qtd_concluido}</span>
                    <span><b>Atrasadas</b> {qtd_atrasadas}</span>
                    <span><b>Resultado recuperado</b> {moeda(total_recuperado)}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            f1, f2 = st.columns([1, 1])

            responsaveis_disponiveis = sorted(
                [
                    str(valor)
                    for valor in acompanhamento_consolidado[
                        "Responsavel"
                    ].dropna().unique()
                    if str(valor).strip()
                ]
            )

            with f1:
                filtro_responsavel = st.selectbox(
                    "Responsável",
                    options=["Todos"] + responsaveis_disponiveis,
                    key="filtro_responsavel_acompanhamento",
                )

            with f2:
                filtro_status = st.selectbox(
                    "Status",
                    options=[
                        "Todos",
                        "Não iniciado",
                        "Em andamento",
                        "Aguardando",
                        "Concluído",
                        "Cancelado",
                        "Atrasadas",
                    ],
                    key="filtro_status_acompanhamento",
                )

            acompanhamento_filtrado = acompanhamento_consolidado.copy()

            if filtro_responsavel != "Todos":
                acompanhamento_filtrado = acompanhamento_filtrado[
                    acompanhamento_filtrado["Responsavel"].astype(str)
                    == filtro_responsavel
                ]

            if filtro_status == "Atrasadas":
                acompanhamento_filtrado = acompanhamento_filtrado[
                    mascara_atrasadas
                ]
            elif filtro_status != "Todos":
                acompanhamento_filtrado = acompanhamento_filtrado[
                    acompanhamento_filtrado["Status"].fillna(
                        "Não iniciado"
                    )
                    == filtro_status
                ]

            tabela_acompanhamento = acompanhamento_filtrado[
                [
                    "ID_Imovel",
                    "Acao_Prioritaria",
                    "Responsavel",
                    "Status",
                    "Data_Limite",
                    "Resultado_Recuperado",
                ]
            ].copy()

            tabela_acompanhamento.columns = [
                "Imóvel",
                "Ação",
                "Responsável",
                "Status",
                "Data limite",
                "Resultado recuperado",
            ]

            tabela_acompanhamento["Data limite"] = pd.to_datetime(
                tabela_acompanhamento["Data limite"],
                errors="coerce",
            ).dt.strftime("%d/%m/%Y")

            tabela_acompanhamento[
                "Resultado recuperado"
            ] = tabela_acompanhamento[
                "Resultado recuperado"
            ].apply(moeda)

            quantidade_linhas = len(tabela_acompanhamento)
            altura_tabela = min(
                300,
                max(120, 42 + quantidade_linhas * 38),
            )

            renderizar_tabela_azul(
                tabela_acompanhamento,
                tema,
                altura_max=altura_tabela,
            )

    st.divider()

    with st.expander("Registrar acompanhamento de uma ação", expanded=False):
        opcoes_acoes = plano_priorizado["ID_Imovel"].tolist()

        imovel_acao = st.selectbox(
            "Selecione o imóvel",
            options=opcoes_acoes,
            key="acao_imovel_selecionado",
        )

        dados_acao = plano_priorizado.loc[
            plano_priorizado["ID_Imovel"] == imovel_acao
        ].iloc[0]

        acompanhamento = carregar_acompanhamento(
            CAMINHO_ACOMPANHAMENTO
        )

        id_acao = f"ACAO-{imovel_acao}"

        registro_existente = pd.DataFrame()

        if not acompanhamento.empty and "ID_Acao" in acompanhamento.columns:
            registro_existente = acompanhamento[
                acompanhamento["ID_Acao"].astype(str) == id_acao
            ]

        existente = (
            registro_existente.iloc[0]
            if not registro_existente.empty
            else None
        )

        status_padrao = (
            str(existente["Status"])
            if existente is not None
            and pd.notna(existente.get("Status"))
            else "Não iniciado"
        )

        responsavel_padrao = (
            str(existente["Responsavel"])
            if existente is not None
            and pd.notna(existente.get("Responsavel"))
            else ""
        )

        observacao_padrao = (
            str(existente["Observacao"])
            if existente is not None
            and pd.notna(existente.get("Observacao"))
            else ""
        )

        resultado_padrao = (
            float(existente["Resultado_Recuperado"])
            if existente is not None
            and pd.notna(existente.get("Resultado_Recuperado"))
            else 0.0
        )

        data_limite_padrao = pd.Timestamp.today().date()

        if (
            existente is not None
            and pd.notna(existente.get("Data_Limite"))
        ):
            data_limite_padrao = pd.to_datetime(
                existente["Data_Limite"]
            ).date()

        st.markdown(
            f"**Ação:** {dados_acao['Ação prioritária']}  \n"
            f"**Índice:** {int(dados_acao['Indice_Risco_Imovel'])} | "
            f"**Impacto mensal:** {moeda(dados_acao['Impacto_Plano'])}"
        )

        with st.form("form_acompanhamento_acao"):
            c1, c2, c3 = st.columns([1.2, 1, 1])

            with c1:
                responsavel = st.text_input(
                    "Responsável",
                    value=responsavel_padrao,
                )

            with c2:
                opcoes_status = [
                    "Não iniciado",
                    "Em andamento",
                    "Aguardando",
                    "Concluído",
                    "Cancelado",
                ]
                indice_status = (
                    opcoes_status.index(status_padrao)
                    if status_padrao in opcoes_status
                    else 0
                )
                status = st.selectbox(
                    "Status",
                    options=opcoes_status,
                    index=indice_status,
                )

            with c3:
                data_limite = st.date_input(
                    "Data limite",
                    value=data_limite_padrao,
                    format="DD/MM/YYYY",
                )

            observacao = st.text_area(
                "Observação",
                value=observacao_padrao,
                height=90,
            )

            resultado_recuperado = st.number_input(
                "Resultado recuperado",
                min_value=0.0,
                value=resultado_padrao,
                step=100.0,
                format="%.2f",
            )

            salvar = st.form_submit_button(
                "Salvar acompanhamento",
                use_container_width=False,
            )

        if salvar:
            agora = pd.Timestamp.now().to_pydatetime()

            data_criacao = agora
            if (
                existente is not None
                and pd.notna(existente.get("Data_Criacao"))
            ):
                data_criacao = pd.to_datetime(
                    existente["Data_Criacao"]
                ).to_pydatetime()

            data_conclusao = None
            if status == "Concluído":
                data_conclusao = agora

            registro = {
                "ID_Acao": id_acao,
                "ID_Imovel": imovel_acao,
                "Acao_Prioritaria": dados_acao["Ação prioritária"],
                "Fator_Principal": dados_acao["Fator_Principal"],
                "Indice_Risco": int(
                    dados_acao["Indice_Risco_Imovel"]
                ),
                "Impacto_Mensal": float(
                    dados_acao["Impacto_Plano"]
                ),
                "Prazo_Sugerido": dados_acao["Prazo sugerido"],
                "Responsavel": responsavel.strip(),
                "Status": status,
                "Data_Criacao": data_criacao,
                "Data_Limite": data_limite,
                "Observacao": observacao.strip(),
                "Data_Conclusao": data_conclusao,
                "Resultado_Recuperado": float(
                    resultado_recuperado
                ),
            }

            try:
                salvar_acompanhamento(
                    CAMINHO_ACOMPANHAMENTO,
                    registro,
                )
                st.success("Acompanhamento salvo com sucesso.")
            except PermissionError:
                st.error(
                    "Não foi possível salvar. Feche a planilha "
                    "Acompanhamento_Acoes_Radar_360.xlsx no Excel "
                    "e tente novamente."
                )
            except Exception as erro:
                st.error(
                    f"Não foi possível salvar o acompanhamento: {erro}"
                )

    st.caption(
        "Os registros desta etapa são salvos na planilha Google compartilhada "
        "Acompanhamento Ações Radar 360."
    )


    st.stop()

st.markdown(
    """
    <div class="cabecalho-visao">
        <span class="cabecalho-visao-icone">📡</span>
        <span class="cabecalho-visao-titulo">TechDadosBR Radar 360</span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption("Riscos, impacto financeiro, prioridades e oportunidades da carteira imobiliária.")

perda_vacancia = imoveis["Perda_Mensal_Vacancia"].sum()

contratos_criticos = imoveis[
    imoveis["Dias_Para_Vencimento"].between(0, 90)
]

imoveis_atraso_recorrente = imoveis[
    imoveis["Perfil_Pagamento"] == "Atraso recorrente"
]

# Impacto bruto para a carteira/proprietários:
# aluguel potencial perdido por vacância + aluguéis ligados a atrasos recorrentes.
valor_locativo_em_risco = (
    perda_vacancia
    + imoveis_atraso_recorrente["Aluguel_Mensal"].sum()
)

# Impacto específico para a imobiliária:
# considera uma única vez cada imóvel com contrato próximo do vencimento
# ou atraso recorrente, evitando duplicidade.
mascara_receita_risco = (
    imoveis["Dias_Para_Vencimento"].between(0, 90)
    | (imoveis["Perfil_Pagamento"] == "Atraso recorrente")
)

receita_imobiliaria_em_risco = imoveis.loc[
    mascara_receita_risco,
    "Receita_Mensal_Imobiliaria",
].sum()

imoveis_criticos = imoveis[
    imoveis["Classificacao_Risco"].isin(["Alto", "Crítico"])
]

proprietarios_risco = proprietarios[
    proprietarios["Risco_Relacionamento"] >= 70
]

c1, c2, c3, c4 = st.columns(4)

with c1:
    card_indicador(
        "Valor locativo em risco",
        moeda(valor_locativo_em_risco),
        "Vacância mensal + aluguéis ligados a atraso recorrente",
        "#A63D40",
    )

with c2:
    card_indicador(
        "Receita da imobiliária em risco",
        moeda(receita_imobiliaria_em_risco),
        "Taxas de imóveis com vencimento em 90 dias ou atraso recorrente",
        "#A63D40",
    )

with c3:
    card_indicador(
        "Imóveis com risco elevado",
        str(len(imoveis_criticos)),
        "Imóveis nas faixas Alto e Crítico",
        "#8E2F3F",
    )

with c4:
    card_indicador(
        "Proprietários em risco",
        str(len(proprietarios_risco)),
        "Relacionamentos a proteger",
        "#C28A2C",
    )

st.divider()

st.subheader("Central de prioridades")

st.markdown(
    """
    <div class="indice-legenda">
        <div><strong>Como ler o Índice de Risco do Imóvel</strong></div>
        <div class="legenda-faixas">
            <span><b>0–29</b> Baixo</span>
            <span><b>30–59</b> Atenção</span>
            <span><b>60–79</b> Alto</span>
            <span><b>80–100</b> Crítico</span>
        </div>
        <div class="legenda-explicacao">
            A nota combina quatro dimensões: Financeiro (até 35 pontos),
            Contratual (até 25), Comercial (até 25) e Operacional (até 15).
            O fator principal mostra onde está a maior concentração do risco.
            O impacto mensal considera a perda por vacância e, quando aplicável,
            a receita mensal da imobiliária ligada a atraso recorrente ou contrato
            com vencimento em até 90 dias.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

prioridades = imoveis.copy()
prioridades["Impacto_Financeiro"] = pd.to_numeric(
    prioridades["Perda_Mensal_Vacancia"],
    errors="coerce",
).fillna(0).astype(float)

mascara_impacto_receita = (
    prioridades["Perfil_Pagamento"].eq("Atraso recorrente")
    | prioridades["Dias_Para_Vencimento"].between(0, 90)
)

prioridades.loc[
    mascara_impacto_receita,
    "Impacto_Financeiro",
] += prioridades.loc[
    mascara_impacto_receita,
    "Receita_Mensal_Imobiliaria",
]


def motivo(row):
    motivos = []

    if row["Dias_Vago"] >= 90:
        motivos.append("Vacância prolongada")
    if row["Perfil_Pagamento"] == "Atraso recorrente":
        motivos.append("Atrasos recorrentes")
    if pd.notna(row["Dias_Para_Vencimento"]) and 0 <= row["Dias_Para_Vencimento"] <= 90:
        motivos.append("Contrato próximo do vencimento")
    if row["Diferenca_Preco_Mercado"] >= 0.10:
        motivos.append("Preço acima do mercado")
    if row["Chamados_12m"] >= 5:
        motivos.append("Excesso de manutenção")

    return " | ".join(motivos) if motivos else "Monitoramento"


def acao_recomendada(row):
    acoes = []

    if row["Dias_Vago"] >= 90:
        acoes.append("Revisar preço, anúncio e estratégia comercial")
    if row["Perfil_Pagamento"] == "Atraso recorrente":
        acoes.append("Iniciar contato preventivo e revisar histórico de pagamento")
    if pd.notna(row["Dias_Para_Vencimento"]) and 0 <= row["Dias_Para_Vencimento"] <= 90:
        acoes.append("Iniciar negociação de renovação")
    if row["Diferenca_Preco_Mercado"] >= 0.10:
        acoes.append("Reavaliar preço com o proprietário")
    if row["Chamados_12m"] >= 5:
        acoes.append("Analisar causa recorrente e custo operacional")

    if not acoes:
        return "Manter monitoramento"

    return " | ".join(acoes)


prioridades["Motivo"] = prioridades.apply(motivo, axis=1)
prioridades["Acao_Recomendada"] = prioridades.apply(acao_recomendada, axis=1)

ranking = prioridades[
    prioridades["Indice_Risco_Imovel"] >= 30
].sort_values(
    ["Indice_Risco_Imovel", "Impacto_Financeiro"],
    ascending=[False, False],
).head(10)

ranking_exibicao = ranking[
    [
        "ID_Imovel",
        "Cidade",
        "Bairro",
        "Indice_Risco_Imovel",
        "Classificacao_Risco",
        "Fator_Principal",
        "Motivo",
        "Acao_Recomendada",
        "Impacto_Financeiro",
    ]
].copy()

ranking_exibicao.columns = [
    "Imóvel",
    "Cidade",
    "Bairro",
    "Índice",
    "Faixa",
    "Fator principal",
    "Motivo",
    "Ação recomendada",
    "Impacto mensal",
]

ranking_exibicao["Impacto mensal"] = ranking_exibicao[
    "Impacto mensal"
].apply(moeda)

renderizar_tabela_prioridades(ranking_exibicao)

with st.expander("Ver composição do índice de um imóvel", expanded=False):
    col_select, col_resumo = st.columns([0.34, 0.66], gap="medium", vertical_alignment="bottom")

    with col_select:
        opcoes_imoveis = ranking["ID_Imovel"].tolist()

        imovel_selecionado = st.selectbox(
            "Selecione o imóvel",
            options=opcoes_imoveis,
            index=0,
            key="detalhe_indice_imovel",
        )

    detalhe = ranking.loc[
        ranking["ID_Imovel"] == imovel_selecionado
    ].iloc[0]

    with col_resumo:
        st.markdown(
            f"""
            <div class="indice-linha indice-linha-lateral">
                <strong>Índice {int(detalhe['Indice_Risco_Imovel'])}</strong>
                <span>Faixa: {detalhe['Classificacao_Risco']}</span>
                <span>Fator principal: {detalhe['Fator_Principal']}</span>
                <span>Financeiro {int(detalhe['Risco_Financeiro'])}</span>
                <span>Contratual {int(detalhe['Risco_Contratual'])}</span>
                <span>Comercial {int(detalhe['Risco_Comercial'])}</span>
                <span>Operacional {int(detalhe['Risco_Operacional'])}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("Onde a perda por vacância está concentrada")

    vacancia_bairro = (
        imoveis.groupby("Bairro", as_index=False)["Perda_Mensal_Vacancia"]
        .sum()
        .sort_values("Perda_Mensal_Vacancia", ascending=False)
        .head(8)
    )

    fig = px.bar(
        vacancia_bairro,
        x="Perda_Mensal_Vacancia",
        y="Bairro",
        orientation="h",
        labels={
            "Perda_Mensal_Vacancia": "Perda mensal",
            "Bairro": "",
        },
    )

    fig.update_layout(
        height=390,
        showlegend=False,
        yaxis={"categoryorder": "total ascending"},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"color": "#DCE7F1"},
    )

    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("Distribuição do risco da carteira")

    risco = (
        imoveis["Classificacao_Risco"]
        .value_counts()
        .rename_axis("Classificação")
        .reset_index(name="Quantidade")
    )

    risco = risco[risco["Quantidade"] > 0].copy()
    total_risco = risco["Quantidade"].sum()
    risco["Percentual"] = (
        risco["Quantidade"] / total_risco * 100
    )

    risco["Rotulo"] = risco["Percentual"].apply(
        lambda valor: f"{valor:.1f}%".replace(".", ",")
        if valor >= 5
        else ""
    )

    fig = px.pie(
        risco,
        names="Classificação",
        values="Quantidade",
        hole=0.62,
        color="Classificação",
        color_discrete_map={
            "Baixo": "#1F5D7A",
            "Atenção": "#7FB6C9",
            "Alto": "#C65353",
            "Crítico": "#8E2F3F",
        },
    )

    fig.update_traces(
        text=risco["Rotulo"],
        textinfo="text",
        textposition="inside",
        hovertemplate="<b>%{label}</b><br>%{percent:.1%}<extra></extra>",
        sort=False,
    )

    cor_texto_grafico = "#EAF2F4" if tema == "Escuro" else "#203746"

    fig.update_layout(
        height=390,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"color": cor_texto_grafico},
        legend={
            "font": {"color": cor_texto_grafico, "size": 12},
            "bgcolor": "rgba(0,0,0,0)",
        },
        uniformtext_minsize=11,
        uniformtext_mode="hide",
    )

    st.plotly_chart(fig, use_container_width=True)

st.divider()

st.subheader("Radar de oportunidades")
st.caption(
    "Prioriza ganhos potenciais identificados pelo cruzamento entre preço, "
    "mercado, vacância, reajuste e taxa de administração."
)

oportunidades = imoveis.copy()

# Diferença positiva entre valor estimado de mercado e aluguel atual.
oportunidades["Potencial_Locativo_Mensal"] = (
    oportunidades["Valor_Mercado_Estimado"]
    - oportunidades["Aluguel_Mensal"]
).clip(lower=0)

oportunidades["Potencial_Receita_Imobiliaria"] = (
    oportunidades["Potencial_Locativo_Mensal"]
    * oportunidades["Taxa_Administracao"]
)

def identificar_oportunidade(row):
    oportunidades_item = []

    if (
        row["Status_Imovel"] == "Ocupado"
        and row["Potencial_Locativo_Mensal"] >= 200
        and row["Meses_Sem_Reajuste"] >= 18
    ):
        oportunidades_item.append("Aluguel abaixo do mercado e reajuste pendente")

    if (
        row["Status_Imovel"] == "Ocupado"
        and row["Potencial_Locativo_Mensal"] >= 300
        and row["Meses_Sem_Reajuste"] < 18
    ):
        oportunidades_item.append("Aluguel abaixo do mercado")

    if (
        row["Status_Imovel"] == "Vago"
        and row["Dias_Vago"] >= 90
        and row["Diferenca_Preco_Mercado"] >= 0.10
    ):
        oportunidades_item.append("Vacância prolongada com preço acima do mercado")

    return " | ".join(oportunidades_item)


def acao_oportunidade(row):
    acoes = []

    if (
        row["Status_Imovel"] == "Ocupado"
        and row["Potencial_Locativo_Mensal"] >= 200
        and row["Meses_Sem_Reajuste"] >= 18
    ):
        acoes.append("Revisar reajuste e negociar adequação gradual do aluguel")

    elif (
        row["Status_Imovel"] == "Ocupado"
        and row["Potencial_Locativo_Mensal"] >= 300
    ):
        acoes.append("Avaliar adequação contratual na próxima janela de negociação")

    if (
        row["Status_Imovel"] == "Vago"
        and row["Dias_Vago"] >= 90
        and row["Diferenca_Preco_Mercado"] >= 0.10
    ):
        acoes.append("Revisar preço e reposicionar o anúncio para reduzir a vacância")

    return " | ".join(acoes)


oportunidades["Oportunidade"] = oportunidades.apply(
    identificar_oportunidade,
    axis=1,
)
oportunidades["Acao_Oportunidade"] = oportunidades.apply(
    acao_oportunidade,
    axis=1,
)

oportunidades_validas = oportunidades[
    oportunidades["Oportunidade"].ne("")
].copy()

potencial_locativo = oportunidades_validas.loc[
    oportunidades_validas["Status_Imovel"] == "Ocupado",
    "Potencial_Locativo_Mensal",
].sum()

potencial_receita_imobiliaria = oportunidades_validas.loc[
    oportunidades_validas["Status_Imovel"] == "Ocupado",
    "Potencial_Receita_Imobiliaria",
].sum()

vacancias_revisao = oportunidades_validas[
    oportunidades_validas["Oportunidade"].str.contains(
        "Vacância prolongada",
        na=False,
    )
]

oc1, oc2, oc3 = st.columns(3)

with oc1:
    card_indicador(
        "Potencial locativo mensal",
        moeda(potencial_locativo),
        "Diferença estimada em imóveis ocupados abaixo do mercado",
        "#178A64",
    )

with oc2:
    card_indicador(
        "Potencial adicional da imobiliária",
        moeda(potencial_receita_imobiliaria),
        "Possível aumento mensal das taxas de administração",
        "#178A64",
    )

with oc3:
    card_indicador(
        "Vacâncias para revisão imediata",
        str(len(vacancias_revisao)),
        "Preço acima do mercado e vacância prolongada",
        "#C28A2C",
    )

if oportunidades_validas.empty:
    st.info("Nenhuma oportunidade relevante foi identificada nesta base.")
else:
    ranking_oportunidades = oportunidades_validas.sort_values(
        [
            "Potencial_Locativo_Mensal",
            "Perda_Mensal_Vacancia",
        ],
        ascending=False,
    ).head(8)

    tabela_oportunidades = ranking_oportunidades[
        [
            "ID_Imovel",
            "Cidade",
            "Bairro",
            "Oportunidade",
            "Acao_Oportunidade",
            "Potencial_Locativo_Mensal",
            "Potencial_Receita_Imobiliaria",
        ]
    ].copy()

    tabela_oportunidades.columns = [
        "Imóvel",
        "Cidade",
        "Bairro",
        "Oportunidade identificada",
        "Ação recomendada",
        "Potencial locativo",
        "Potencial da imobiliária",
    ]

    tabela_oportunidades["Potencial locativo"] = tabela_oportunidades[
        "Potencial locativo"
    ].apply(moeda)

    tabela_oportunidades["Potencial da imobiliária"] = tabela_oportunidades[
        "Potencial da imobiliária"
    ].apply(moeda)

    with st.expander("Ver oportunidades priorizadas", expanded=False):
        renderizar_tabela_azul(
            tabela_oportunidades,
            tema,
            altura_max=360,
        )

st.divider()

st.subheader("Diagnóstico executivo")

diagnosticos = []

diagnosticos.append(
    f"O valor locativo em risco é de {moeda(valor_locativo_em_risco)}, "
    f"enquanto a receita mensal da imobiliária potencialmente comprometida "
    f"é de {moeda(receita_imobiliaria_em_risco)}."
)

if perda_vacancia > 0:
    top_vagos = imoveis[
        imoveis["Status_Imovel"] == "Vago"
    ].sort_values(
        "Perda_Mensal_Vacancia",
        ascending=False,
    ).head(5)

    concentracao = (
        top_vagos["Perda_Mensal_Vacancia"].sum()
        / perda_vacancia
        * 100
    )

    diagnosticos.append(
        f"Os cinco imóveis com maior perda concentram "
        f"{concentracao:.1f}% da perda mensal por vacância."
    )

if len(contratos_criticos) > 0:
    diagnosticos.append(
        f"Existem {len(contratos_criticos)} contratos com vencimento "
        f"nos próximos 90 dias e impacto mensal potencial de "
        f"{moeda(contratos_criticos['Receita_Mensal_Imobiliaria'].sum())}."
    )

if len(proprietarios_risco) == 1:
    diagnosticos.append(
        "1 proprietário apresenta risco elevado de relacionamento "
        "e deve receber contato gerencial prioritário."
    )
elif len(proprietarios_risco) > 1:
    diagnosticos.append(
        f"{len(proprietarios_risco)} proprietários apresentam risco elevado "
        f"de relacionamento e devem receber contato gerencial prioritário."
    )

for texto in diagnosticos:
    st.markdown(
        f"""
        <div style="
            background:#E8F2FF;
            color:#0059B3;
            padding:18px 20px;
            border-radius:10px;
            margin-bottom:12px;
            font-size:1rem;
        ">
            {texto}
        </div>
        """,
        unsafe_allow_html=True,
    )
