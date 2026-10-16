# -*- coding: utf-8 -*-
"""
SLC Agrícola – Desafio Prático: Analista de Suprimentos com foco em IA
=======================================================================
Aplicativo local em Python (Tkinter + ttkbootstrap + openpyxl) para alterar a
planilha de Suprimentos ("base_cadastro_e_requisições*.xlsx"), composta pelas
abas BASE_CADASTRO_ITENS (cadastro de itens) e BASE_RCs (requisições abertas).

Operações:

    1) Redirecionamento de RC por exceção
       - o usuário seleciona RCs diretamente na tabela (uma ou várias) ou as
         adiciona manualmente; depois informa o novo Grupo Comprador;
       - altera o Grupo Comprador das RCs na aba BASE_RCs;
       - NÃO altera o cadastro dos itens;
       - registra a exceção na aba EXCECOES_RCs (persistente após fechar o app).

    2) Correção do cadastro de um item
       - altera o Grupo Comprador do item na aba BASE_CADASTRO_ITENS;
       - atualiza as linhas da aba BASE_RCs que usam esse item;
       - NÃO altera RCs que possuem redirecionamento por exceção.

Regra principal: uma exceção cadastrada para uma RC SEMPRE prevalece sobre o
cadastro do item. Ao carregar o arquivo, as exceções registradas são reaplicadas
às RCs (caso a planilha base tenha sido atualizada externamente).

Como executar:
    pip install -r requirements.txt
    python app_suprimentos.py
"""

import os
import re
import sys
import traceback
import unicodedata
from datetime import datetime

import tkinter as tk

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.dialogs import Messagebox
except ImportError:
    sys.stdout.write("ERRO: biblioteca 'ttkbootstrap' nao encontrada.\n"
                     "Instale com:  pip install ttkbootstrap\n")
    raise SystemExit(1)

import openpyxl
from openpyxl.styles import Font

# ============================== Constantes ==============================

ABA_ITENS = "BASE_CADASTRO_ITENS"
ABA_RCS = "BASE_RCs"
ABA_EXCECOES = "EXCECOES_RCs"

# Colunas fixas (índice 1-based) conforme a planilha fornecida
COL_ITENS_CODIGO = 1
COL_ITENS_DESCRICAO = 2
COL_ITENS_GRUPO = 3
COL_RCS_RC = 1
COL_RCS_LINHA = 2
COL_RCS_CODIGO = 3
COL_RCS_DESCRICAO = 4
COL_RCS_GRUPO = 5
COL_EXC_RC = 1
COL_EXC_GRUPO = 2
COL_EXC_DATA = 3

# Recursos empacotados junto com o aplicativo (ícone da janela e logotipo)
CAMINHO_ICONE = os.path.join("assets", "slc.ico")
CAMINHO_LOGO = os.path.join("assets", "slc_brand_name.png")


# ============================== Funções auxiliares ======================

def _to_int(valor):
    """Converte um valor (int/float/str) para int. Retorna None se não for possível."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, int):
        return valor
    if isinstance(valor, float):
        return int(valor) if valor.is_integer() else None
    s = str(valor).strip()
    if not s:
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    try:
        f = float(s)
        return int(f) if f.is_integer() else None
    except ValueError:
        return None


def _lista_ints(texto):
    """
    Separa uma lista de números inteiros informada em texto livre.
    Aceita separação por vírgula, ponto e vírgula, espaço ou quebra de linha.
    Retorna (lista_de_ints, tokens_invalidos).
    """
    tokens = re.split(r"[,;\s]+", texto.strip())
    valores, invalidos = [], []
    for token in tokens:
        token = token.strip().strip("'\"")
        if not token:
            continue
        v = _to_int(token)
        if v is None:
            invalidos.append(token)
        else:
            valores.append(v)
    return valores, invalidos


def _norm(texto):
    """Texto minúsculo e sem acentos, para buscas e comparações."""
    t = unicodedata.normalize("NFD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.lower()


def pasta_do_app():
    """Pasta onde o aplicativo está instalado (compatível com PyInstaller)."""
    if getattr(sys, "frozen", False):  # quando empacotado com PyInstaller
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def caminho_recurso(relativo):
    """
    Caminho absoluto de um recurso empacotado (pasta assets).
    Funciona rodando pelo Python e também como executável do PyInstaller:
    no modo onefile os recursos são extraídos para sys._MEIPASS.
    """
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relativo)


def localizar_arquivo_base(pasta):
    """
    Procura na pasta informada um .xlsx cujo nome corresponda à planilha base
    (ex.: 'base_cadastro_e_requisições.xlsx'). Ignora arquivos temporários (~$).
    Se não houver candidato com esse padrão, usa um .xlsx único na pasta, se existir.
    """
    candidatos, outros = [], []
    try:
        arquivos = os.listdir(pasta)
    except OSError:
        return None
    for nome in arquivos:
        if nome.startswith("~$") or not nome.lower().endswith((".xlsx", ".xlsm")):
            continue
        n = _norm(nome)
        if n.startswith("base") and "cadastro" in n:
            candidatos.append(nome)
        else:
            outros.append(nome)
    if candidatos:
        return os.path.join(pasta, sorted(candidatos)[0])
    if len(outros) == 1:
        return os.path.join(pasta, outros[0])
    return None


def _habilitar_dpi():
    """Deixa a interface nítida em telas com escala alta (Windows)."""
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


# ============================== Camada de dados =========================

class ExcelService:
    """
    Encapsula o acesso à planilha Excel e as duas operações de negócio:
      • redirecionar_rcs      – redirecionamento de RC por exceção;
      • corrigir_cadastro_item – correção do Grupo Comprador do item.
    """

    def __init__(self, caminho):
        self.caminho = caminho
        self.wb = None

    # ------------------------------ infraestrutura ----------------------

    def carregar(self):
        self.wb = openpyxl.load_workbook(self.caminho)
        for aba in (ABA_ITENS, ABA_RCS):
            if aba not in self.wb.sheetnames:
                raise ValueError(f"A aba '{aba}' não foi encontrada no arquivo.")

    def salvar(self):
        try:
            self.wb.save(self.caminho)
        except PermissionError:
            # Descarta as alterações em memória para evitar inconsistências.
            try:
                self.wb = openpyxl.load_workbook(self.caminho)
            except Exception:
                self.wb = None
            raise

    def _aba(self, nome):
        return self.wb[nome]

    def _linhas_preenchidas(self, aba):
        """Gera os números das linhas preenchidas (ignorando o cabeçalho)."""
        ws = self._aba(aba)
        for ln in range(2, ws.max_row + 1):
            if any(ws.cell(row=ln, column=c).value is not None
                   for c in range(1, ws.max_column + 1)):
                yield ln

    def resumo(self):
        """Contagens simples para exibição na barra de status."""
        return {
            "itens": sum(1 for _ in self._linhas_preenchidas(ABA_ITENS)),
            "linhas_rc": sum(1 for _ in self._linhas_preenchidas(ABA_RCS)),
            "excecoes": len(self.obter_excecoes()),
        }

    # ------------------------------ leituras p/ tabelas ------------------

    def linhas_itens(self):
        """Lista de tuplas (código, descrição, grupo) da BASE_CADASTRO_ITENS."""
        ws = self._aba(ABA_ITENS)
        return [
            (ws.cell(row=ln, column=COL_ITENS_CODIGO).value,
             ws.cell(row=ln, column=COL_ITENS_DESCRICAO).value,
             ws.cell(row=ln, column=COL_ITENS_GRUPO).value)
            for ln in self._linhas_preenchidas(ABA_ITENS)
        ]

    def linhas_rcs(self):
        """Lista de tuplas (RC, linha, código item, descrição, grupo) da BASE_RCs."""
        ws = self._aba(ABA_RCS)
        return [
            (ws.cell(row=ln, column=COL_RCS_RC).value,
             ws.cell(row=ln, column=COL_RCS_LINHA).value,
             ws.cell(row=ln, column=COL_RCS_CODIGO).value,
             ws.cell(row=ln, column=COL_RCS_DESCRICAO).value,
             ws.cell(row=ln, column=COL_RCS_GRUPO).value)
            for ln in self._linhas_preenchidas(ABA_RCS)
        ]

    def linhas_excecoes(self):
        """Lista de tuplas (RC, grupo exceção, data/hora) da aba EXCECOES_RCs."""
        if ABA_EXCECOES not in self.wb.sheetnames:
            return []
        ws = self._aba(ABA_EXCECOES)
        return [
            (ws.cell(row=ln, column=COL_EXC_RC).value,
             ws.cell(row=ln, column=COL_EXC_GRUPO).value,
             ws.cell(row=ln, column=COL_EXC_DATA).value)
            for ln in self._linhas_preenchidas(ABA_EXCECOES)
        ]

    # ------------------------------ buscas ------------------------------

    def _buscar_linha_item(self, codigo):
        """Retorna o nº da linha do item na BASE_CADASTRO_ITENS ou None."""
        ws = self._aba(ABA_ITENS)
        for ln in self._linhas_preenchidas(ABA_ITENS):
            if _to_int(ws.cell(row=ln, column=COL_ITENS_CODIGO).value) == codigo:
                return ln
        return None

    def _linhas_da_rc(self, rc):
        """Lista com os nºs das linhas da RC na aba BASE_RCs."""
        ws = self._aba(ABA_RCS)
        linhas = []
        for ln in self._linhas_preenchidas(ABA_RCS):
            if _to_int(ws.cell(row=ln, column=COL_RCS_RC).value) == rc:
                linhas.append(ln)
        return linhas

    # ------------------------------ exceções ----------------------------

    def _garantir_aba_excecoes(self):
        """Cria a aba EXCECOES_RCs, se ainda não existir."""
        if ABA_EXCECOES not in self.wb.sheetnames:
            ws = self.wb.create_sheet(ABA_EXCECOES)
            ws.append(["RC", "Grupo Comprador Exceção", "Data/Hora do Registro"])
            for col in ("A", "B", "C"):
                ws[f"{col}1"].font = Font(bold=True)
            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 24
            ws.freeze_panes = "A2"
        return self._aba(ABA_EXCECOES)

    def obter_excecoes(self):
        """Retorna um dicionário {RC: Grupo Comprador de exceção}."""
        excecoes = {}
        if ABA_EXCECOES not in self.wb.sheetnames:
            return excecoes
        ws = self._aba(ABA_EXCECOES)
        for ln in self._linhas_preenchidas(ABA_EXCECOES):
            rc = _to_int(ws.cell(row=ln, column=COL_EXC_RC).value)
            grupo = _to_int(ws.cell(row=ln, column=COL_EXC_GRUPO).value)
            if rc is not None and grupo is not None:
                excecoes[rc] = grupo
        return excecoes

    def _registrar_excecao(self, rc, grupo):
        """Insere ou atualiza o registro da exceção de uma RC."""
        ws = self._garantir_aba_excecoes()
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        for ln in self._linhas_preenchidas(ABA_EXCECOES):
            if _to_int(ws.cell(row=ln, column=COL_EXC_RC).value) == rc:
                ws.cell(row=ln, column=COL_EXC_GRUPO).value = grupo
                ws.cell(row=ln, column=COL_EXC_DATA).value = agora
                return
        ws.append([rc, grupo, agora])

    def reaplicar_excecoes(self):
        """
        Reaplica às RCs os grupos definidos nas exceções registradas. Garante que,
        mesmo que a planilha base seja atualizada externamente, as RCs com exceção
        permaneçam com o Grupo Comprador delegado. Retorna o nº de células alteradas.
        """
        excecoes = self.obter_excecoes()
        if not excecoes:
            return 0
        ws = self._aba(ABA_RCS)
        alteracoes = 0
        for ln in self._linhas_preenchidas(ABA_RCS):
            rc = _to_int(ws.cell(row=ln, column=COL_RCS_RC).value)
            if rc in excecoes:
                celula = ws.cell(row=ln, column=COL_RCS_GRUPO)
                if _to_int(celula.value) != excecoes[rc]:
                    celula.value = excecoes[rc]
                    alteracoes += 1
        if alteracoes:
            self.salvar()
        return alteracoes

    # ------------------------------ operação 1 --------------------------

    def redirecionar_rcs(self, rcs, novo_grupo, confirmar_substituicao=None):
        """
        Redireciona as RCs informadas para o novo Grupo Comprador, sem alterar o
        cadastro dos itens, e registra a exceção (persistente) para cada RC.

        Parâmetros:
            rcs        – lista de RCs (ints), com ou sem repetição;
            novo_grupo – novo Grupo Comprador (int);
            confirmar_substituicao – callback([(rc, grupo_atual), ...], novo_grupo)
                         chamado quando a RC já possui exceção para outro grupo.
                         Se None, RCs que já possuem exceção são mantidas.

        Retorna um dicionário com o resultado detalhado da operação.
        """
        res = {
            "atualizadas": [],      # RCs redirecionadas nesta execução
            "substituidas": [],     # RCs cuja exceção anterior foi substituída
            "mantidas": [],         # RCs mantidas com a exceção anterior
            "ja_mesmo_grupo": [],   # RCs que já tinham exceção para o novo grupo
            "nao_encontradas": [],  # RCs inexistentes na BASE_RCs
        }
        unicas = list(dict.fromkeys(rcs))
        excecoes = self.obter_excecoes()

        a_processar, a_confirmar = [], []
        for rc in unicas:
            if not self._linhas_da_rc(rc):
                res["nao_encontradas"].append(rc)
                continue
            grupo_excecao_atual = excecoes.get(rc)
            if grupo_excecao_atual == novo_grupo:
                res["ja_mesmo_grupo"].append(rc)
                continue
            if grupo_excecao_atual is not None:
                a_confirmar.append((rc, grupo_excecao_atual))
                continue
            a_processar.append(rc)

        # Tratamento de "RC que já possui exceção": pede confirmação para substituir
        if a_confirmar:
            substituir = (confirmar_substituicao is not None
                          and bool(confirmar_substituicao(a_confirmar, novo_grupo)))
            if substituir:
                res["substituidas"] = [rc for rc, _ in a_confirmar]
                a_processar.extend(res["substituidas"])
            else:
                res["mantidas"] = [rc for rc, _ in a_confirmar]

        ws_r = self._aba(ABA_RCS)
        modificou = False
        for rc in a_processar:
            for ln in self._linhas_da_rc(rc):
                ws_r.cell(row=ln, column=COL_RCS_GRUPO).value = novo_grupo
            self._registrar_excecao(rc, novo_grupo)
            res["atualizadas"].append(rc)
            modificou = True

        # RCs que já possuíam exceção para o mesmo grupo: apenas garante consistência
        for rc in res["ja_mesmo_grupo"]:
            for ln in self._linhas_da_rc(rc):
                celula = ws_r.cell(row=ln, column=COL_RCS_GRUPO)
                if _to_int(celula.value) != novo_grupo:
                    celula.value = novo_grupo
                    modificou = True

        if modificou:
            self.salvar()
        return res

    # ------------------------------ operação 2 --------------------------

    def corrigir_cadastro_item(self, codigo, novo_grupo):
        """
        Corrige o Grupo Comprador de um item no cadastro e atualiza as linhas da
        BASE_RCs que usam esse item — EXCETO as RCs que possuem exceção registrada
        (regra principal: a exceção sempre prevalece sobre o cadastro do item).
        """
        res = {
            "item_encontrado": False,
            "linhas_atualizadas": 0,
            "rcs_preservadas": [],  # [(RC, Linha RC, grupo mantido)]
        }
        linha_item = self._buscar_linha_item(codigo)
        if linha_item is None:
            return res
        res["item_encontrado"] = True

        ws_itens = self._aba(ABA_ITENS)
        ws_itens.cell(row=linha_item, column=COL_ITENS_GRUPO).value = novo_grupo

        excecoes = self.obter_excecoes()
        ws_r = self._aba(ABA_RCS)
        for ln in self._linhas_preenchidas(ABA_RCS):
            if _to_int(ws_r.cell(row=ln, column=COL_RCS_CODIGO).value) != codigo:
                continue
            rc = _to_int(ws_r.cell(row=ln, column=COL_RCS_RC).value)
            if rc in excecoes:
                linha_rc = ws_r.cell(row=ln, column=COL_RCS_LINHA).value
                res["rcs_preservadas"].append((rc, linha_rc, excecoes[rc]))
            else:
                ws_r.cell(row=ln, column=COL_RCS_GRUPO).value = novo_grupo
                res["linhas_atualizadas"] += 1

        self.salvar()
        return res


# ============================== Componentes de interface =================

class Tabela(ttk.Frame):
    """Painel de tabela (Treeview) com busca, ordenação por coluna e contagem."""

    def __init__(self, master, colunas, bootstyle="primary", acao=None, ao_duplo_clique=None):
        super().__init__(master)
        self.colunas = colunas            # [{id, texto, largura, ancora, esticar}]
        self.ao_duplo_clique = ao_duplo_clique
        self.dados = []                   # lista de tuplas (valores originais)
        self._visiveis = []              # linhas atualmente exibidas
        self.filtro = ""
        self.coluna_ordenada = None
        self.ascendente = True

        # ----- barra de ferramentas: busca, contagem e ação específica
        barra = ttk.Frame(self)
        barra.pack(fill="x", pady=(0, 6))
        ttk.Label(barra, text="Buscar:").pack(side="left")
        self.ent_busca = ttk.Entry(barra)
        self.ent_busca.pack(side="left", fill="x", expand=True, padx=(6, 10))
        self.ent_busca.bind("<KeyRelease>", self._ao_buscar)
        self.lbl_contagem = ttk.Label(barra, text="0 linhas")
        self.lbl_contagem.pack(side="left", padx=(0, 10))
        if acao is not None:
            texto, estilo, comando = acao
            ttk.Button(barra, text=texto, bootstyle=estilo, command=comando).pack(side="right")

        # ----- tabela com rolagem
        caixa = ttk.Frame(self)
        caixa.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(caixa, columns=[c["id"] for c in colunas],
                                 show="headings", bootstyle=bootstyle)
        rolagem = ttk.Scrollbar(caixa, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=rolagem.set)
        self.tree.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

        for col in colunas:
            self.tree.heading(col["id"], text=col["texto"],
                              command=lambda i=col["id"]: self._ordenar_por(i))
            self.tree.column(col["id"], width=col.get("largura", 100),
                             anchor=col.get("ancora", "w"),
                             stretch=col.get("esticar", False),
                             minwidth=col.get("largura_min", 40))
        self.tree.tag_configure("zebra", background="#f3f5f8")
        if ao_duplo_clique is not None:
            self.tree.bind("<Double-1>", self._ao_duplo_clique)

    # ------------------------------ API ---------------------------------

    def definir_dados(self, dados):
        self.dados = list(dados)
        self._render()

    def linhas_selecionadas(self):
        """Tuplas de dados das linhas selecionadas (Ctrl/Shift para várias)."""
        return [self._visiveis[self.tree.index(iid)]
                for iid in self.tree.selection()
                if self.tree.index(iid) < len(self._visiveis)]

    def limpar_selecao(self):
        for iid in self.tree.selection():
            self.tree.selection_remove(iid)

    # ------------------------------ internos -----------------------------

    def _ao_buscar(self, _evento=None):
        self.filtro = self.ent_busca.get()
        self._render()

    def _chave_ordenacao(self, valor):
        n = _to_int(valor)
        if n is not None:
            return (0, n, "")
        return (1, _norm(valor), "")

    def _ordenar_por(self, col_id):
        if self.coluna_ordenada == col_id:
            self.ascendente = not self.ascendente
        else:
            self.coluna_ordenada = col_id
            self.ascendente = True
        self._render()

    def _corresponde(self, linha, consulta):
        if not consulta:
            return True
        texto = " ".join(str(v) for v in linha)
        return _norm(consulta) in _norm(texto)

    def _render(self):
        consulta = self.filtro.strip()
        visiveis = [l for l in self.dados if self._corresponde(l, consulta)]
        if self.coluna_ordenada:
            idx = [c["id"] for c in self.colunas].index(self.coluna_ordenada)
            visiveis.sort(key=lambda l: self._chave_ordenacao(l[idx]),
                          reverse=not self.ascendente)
        self._visiveis = visiveis
        self.tree.delete(*self.tree.get_children())
        for i, linha in enumerate(visiveis):
            tags = ("zebra",) if i % 2 else ()
            self.tree.insert("", "end", values=linha, tags=tags)
        if consulta:
            self.lbl_contagem.configure(text=f"{len(visiveis)} de {len(self.dados)} linhas")
        else:
            self.lbl_contagem.configure(text=f"{len(self.dados)} linhas")
        self._atualizar_setas()

    def _atualizar_setas(self):
        for col in self.colunas:
            texto = col["texto"]
            if col["id"] == self.coluna_ordenada:
                texto += "  ▲" if self.ascendente else "  ▼"
            self.tree.heading(col["id"], text=texto)

    def _ao_duplo_clique(self, evento):
        iid = self.tree.identify_row(evento.y)
        if not iid:
            return
        idx = self.tree.index(iid)
        if 0 <= idx < len(self._visiveis) and self.ao_duplo_clique:
            self.ao_duplo_clique(self._visiveis[idx])


class ListaRcs(ttk.Frame):
    """
    Lista visual (chips) das RCs selecionadas para o redirecionamento.
    Cada chip pode ser removido com um clique no '×'; a lista se reorganiza
    automaticamente conforme a largura disponível.
    """

    ALTURA_CHIP = 28

    def __init__(self, master, on_alteracao=None):
        super().__init__(master, style="Card.TFrame")
        self.valores = []
        self.on_alteracao = on_alteracao
        self._largura_render = None
        self.area = ttk.Frame(self, style="Card.TFrame")
        self.area.pack(fill="x")
        self.area.bind("<Configure>", self._ao_redimensionar)
        self._render()

    # ------------------------------ API ---------------------------------

    def adicionar(self, rcs):
        """Adiciona RCs (ignora repetidas, inclusive dentro do mesmo lote).
        Retorna as RCs que entraram."""
        existentes = set(self.valores)
        novos = []
        for rc in rcs:
            if rc not in existentes:
                novos.append(rc)
                existentes.add(rc)
        if not novos:
            return []
        self.valores.extend(novos)
        self._render()
        self._notificar()
        return novos

    def remover(self, rc):
        if rc in self.valores:
            self.valores.remove(rc)
            self._render()
            self._notificar()

    def limpar(self):
        if self.valores:
            self.valores = []
            self._render()
            self._notificar()

    # ------------------------------ internos -----------------------------

    def _notificar(self):
        if self.on_alteracao:
            self.on_alteracao(self.valores)

    def _ao_redimensionar(self, evento):
        if evento.width > 1 and evento.width != self._largura_render:
            self._render()

    def _render(self):
        largura = self.area.winfo_width()
        if largura <= 1:
            largura = 420
        self._largura_render = largura
        for filho in self.area.winfo_children():
            filho.destroy()

        if not self.valores:
            ttk.Label(self.area, style="CardHint.TLabel", justify="left",
                      text="Nenhuma RC selecionada — selecione linhas na aba BASE_RCs, "
                           "dê duplo clique em uma linha ou adicione manualmente abaixo.",
                      wraplength=largura - 8).place(x=0, y=0, anchor="nw")
            self.area.configure(height=20)
            return

        chips = []
        for rc in self.valores:
            chip = ttk.Frame(self.area, style="Chip.TFrame")
            ttk.Label(chip, text=str(rc), style="Chip.TLabel").pack(side="left", padx=(9, 0), pady=2)
            fechar = ttk.Label(chip, text=" ×", style="ChipX.TLabel", cursor="hand2")
            fechar.pack(side="left", padx=(3, 6), pady=2)
            fechar.bind("<Button-1>", lambda _e, r=rc: self.remover(r))
            fechar.bind("<Enter>", lambda _e, b=fechar: b.configure(style="ChipXHover.TLabel"))
            fechar.bind("<Leave>", lambda _e, b=fechar: b.configure(style="ChipX.TLabel"))
            chips.append(chip)

        self.update_idletasks()  # mede a largura de cada chip
        x = y = 0
        for chip in chips:
            w = chip.winfo_reqwidth() + 4
            if x > 0 and x + w > largura:
                x = 0
                y += self.ALTURA_CHIP + 4
            chip.place(x=x, y=y)
            x += w
        self.area.configure(height=y + self.ALTURA_CHIP)


# ============================== Interface principal =====================

class App(ttk.Window):
    """Janela principal do aplicativo."""

    def __init__(self):
        _habilitar_dpi()
        super().__init__(themename="minty")
        self.title("SLC Agrícola · Suprimentos — Redirecionamento de RCs e Cadastro de Itens")
        self.geometry("1240x740")
        self.minsize(1060, 620)
        self._aplicar_icone()
        self.service = None
        self._rcs_validas = set()
        self._montar_estilos()
        self._montar_interface()
        self._carregar_arquivo_inicial()

    # ------------------------------ ícone --------------------------------

    def _aplicar_icone(self):
        """
        Aplica o ícone da SLC (assets/slc.ico) na janela: barra de título,
        Alt+Tab e barra de tarefas.
        """
        icone = caminho_recurso(CAMINHO_ICONE)
        if not os.path.isfile(icone):
            return
        try:
            # No Windows, iconbitmap aplica o .ico na janela e na barra de tarefas.
            self.iconbitmap(icone)
        except tk.TclError:
            # Fallback: converte o .ico em PhotoImage (via Pillow) e usa iconphoto.
            try:
                from PIL import Image, ImageTk
                self._icone_img = ImageTk.PhotoImage(Image.open(icone))
                self.iconphoto(True, self._icone_img)
            except Exception:
                pass  # segue com o ícone padrão, sem interromper o aplicativo

    # ------------------------------ estilos ------------------------------

    def _montar_estilos(self):
        estilo = ttk.Style()
        estilo.configure(".", font=("Segoe UI", 10))
        for prefixo in ("", "Primary.", "Secondary.", "Success.", "Warning.", "Danger.", "Info."):
            estilo.configure(prefixo + "Treeview", rowheight=26, font=("Segoe UI", 10))
            estilo.configure(prefixo + "Treeview.Heading", font=("Segoe UI", 9, "bold"))

        primaria = estilo.colors.primary
        estilo.configure("Header.TFrame", background=primaria)
        estilo.configure("Header.TLabel", background=primaria, foreground="#ffffff",
                         font=("Segoe UI", 15, "bold"))
        estilo.configure("HeaderSub.TLabel", background=primaria, foreground="#c5d1e0",
                         font=("Segoe UI", 9))

        estilo.configure("Sidebar.TFrame", background="#eef1f5")
        estilo.configure("Card.TFrame", background="#ffffff")
        estilo.configure("Card.TLabel", background="#ffffff", foreground="#212529")
        estilo.configure("CardTitulo.TLabel", background="#ffffff", foreground="#212529",
                         font=("Segoe UI", 11, "bold"))
        estilo.configure("CardCampo.TLabel", background="#ffffff", foreground="#495057",
                         font=("Segoe UI", 9))
        estilo.configure("CardHint.TLabel", background="#ffffff", foreground="#868e96",
                         font=("Segoe UI", 8))

        estilo.configure("Chip.TFrame", background="#e2e8f0")
        estilo.configure("Chip.TLabel", background="#e2e8f0", foreground="#334155",
                         font=("Segoe UI", 9, "bold"))
        estilo.configure("ChipX.TLabel", background="#e2e8f0", foreground="#64748b",
                         font=("Segoe UI", 10, "bold"))
        estilo.configure("ChipXHover.TLabel", background="#ef4444", foreground="#ffffff",
                         font=("Segoe UI", 10, "bold"))

        estilo.configure("Status.TFrame", background="#212529")
        estilo.configure("Status.TLabel", background="#212529", foreground="#ced4da",
                         font=("Segoe UI", 9))
        estilo.configure("StatusOk.TLabel", background="#212529", foreground="#69db7c",
                         font=("Segoe UI", 9))
        estilo.configure("StatusErr.TLabel", background="#212529", foreground="#ff8787",
                         font=("Segoe UI", 9))

    # ------------------------------ diálogos ----------------------------

    def _info(self, mensagem, titulo="Concluído"):
        Messagebox.show_info(message=mensagem, title=titulo, parent=self)

    def _aviso(self, mensagem, titulo="Atenção"):
        Messagebox.show_warning(message=mensagem, title=titulo, parent=self)

    def _erro(self, mensagem, titulo="Erro"):
        Messagebox.show_error(message=mensagem, title=titulo, parent=self)

    def _sim_nao(self, mensagem, titulo="Confirmar"):
        resposta = Messagebox.yesno(message=mensagem, title=titulo, parent=self)
        return str(resposta).strip().lower() in ("yes", "sim")

    # ------------------------------ interface ----------------------------

    def _montar_interface(self):
        # ================== cabeçalho
        cabecalho = ttk.Frame(self, style="Header.TFrame")
        cabecalho.pack(fill="x")
        bloco = ttk.Frame(cabecalho, style="Header.TFrame")
        bloco.pack(side="left", padx=18, pady=10)
        
        from PIL import Image, ImageTk
        imagem_original = Image.open(caminho_recurso(CAMINHO_LOGO))
        
        # 2. Redimensionar a imagem para caber no cabeçalho (ajuste os valores 150x40 conforme achar melhor)
        imagem_redimensionada = imagem_original.resize((150, 17), Image.Resampling.LANCZOS)
        
        # 3. Converter para o formato que o Tkinter entende
        # IMPORTANTE: Tem que salvar no 'self' (self.logo_img) senão a imagem some da tela!
        self.logo_img = ImageTk.PhotoImage(imagem_redimensionada)

        # 4. Criar o Label usando a imagem no lugar do texto
        self.lbl_logo = ttk.Label(
            bloco, # Substitua pelo nome da sua variável do frame do cabeçalho
            image=self.logo_img, 
            style="Header.TLabel"
        )
        
        # 5. Posicionar no canto esquerdo (se quiser no direito de fato, mude para side="right")
        self.lbl_logo.pack(anchor="w")
        ttk.Label(bloco, text="Redirecionamento de RCs por exceção  ·  Correção de cadastro de itens",
                  style="HeaderSub.TLabel").pack(anchor="w")


        # ================== barra de status (embaixo)
        self.barra_status = ttk.Frame(self, style="Status.TFrame")
        self.barra_status.pack(side="bottom", fill="x")
        self.lbl_arq_status = ttk.Label(self.barra_status, text="●  Nenhum arquivo carregado",
                                        style="StatusErr.TLabel")
        self.lbl_arq_status.pack(side="left", padx=(16, 0), pady=6)
        self.lbl_contagens = ttk.Label(self.barra_status, text="", style="Status.TLabel")
        self.lbl_contagens.pack(side="right", padx=(0, 16))

        # ================== corpo: coluna esquerda (operações) + coluna direita (tabelas)
        corpo = ttk.Frame(self)
        corpo.pack(fill="both", expand=True, padx=12, pady=12)

        sidebar = ttk.Frame(corpo, style="Sidebar.TFrame")
        sidebar.pack(side="left", fill="y", padx=(0, 12))
        sidebar.configure(width=470)
        sidebar.pack_propagate(False)

        conteudo = ttk.Frame(corpo)
        conteudo.pack(side="left", fill="both", expand=True)

        self._montar_sidebar(sidebar)
        self._montar_abas(conteudo)

    def _montar_sidebar(self, master):
        # ---------- cartão: arquivo
        card_arq = ttk.Frame(master, style="Card.TFrame", padding=12)
        card_arq.pack(fill="x")
        ttk.Label(card_arq, text="Arquivo Excel", style="CardTitulo.TLabel").pack(anchor="w")
        linha_arq = ttk.Frame(card_arq, style="Card.TFrame")
        linha_arq.pack(fill="x", pady=(6, 6))
        self.var_caminho = tk.StringVar()
        ttk.Entry(linha_arq, textvariable=self.var_caminho).pack(
            side="left", fill="x", expand=True)
        ttk.Button(linha_arq, text="Procurar…", bootstyle="secondary-outline",
                   command=self._procurar_arquivo).pack(side="left", padx=(6, 0))
        ttk.Button(card_arq, text="Carregar / Recarregar", bootstyle="secondary-outline",
                   command=self._recarregar).pack(fill="x")

        # ---------- cartão: operação 1
        card1 = ttk.Frame(master, style="Card.TFrame", padding=12)
        card1.pack(fill="x", pady=(12, 0))
        ttk.Label(card1, text="1 · Redirecionamento de RC por exceção",
                  style="CardTitulo.TLabel").pack(anchor="w")

        topo_chips = ttk.Frame(card1, style="Card.TFrame")
        topo_chips.pack(fill="x", pady=(8, 0))
        self.lbl_qtd_rcs = ttk.Label(topo_chips, text="RCs selecionadas: 0",
                                     style="CardCampo.TLabel", font=("Segoe UI", 9, "bold"))
        self.lbl_qtd_rcs.pack(side="left")
        ttk.Button(topo_chips, text="Limpar", bootstyle="danger-link",
                   command=self._limpar_selecao_rcs).pack(side="right")

        self.chips = ListaRcs(card1, on_alteracao=self._ao_mudar_chips)
        self.chips.pack(fill="x", pady=(6, 2))

        ttk.Separator(card1).pack(fill="x", pady=(8, 6))

        ttk.Label(card1, text="Adicionar manualmente (uma RC ou lista, separada por vírgula/espaço):",
                  style="CardCampo.TLabel").pack(anchor="w")
        linha_manual = ttk.Frame(card1, style="Card.TFrame")
        linha_manual.pack(fill="x", pady=(4, 0))
        self.ent_rcs_manual = ttk.Entry(linha_manual)
        self.ent_rcs_manual.pack(side="left", fill="x", expand=True)
        self.ent_rcs_manual.bind("<Return>", lambda _e: self._adicionar_rcs_manual())
        ttk.Button(linha_manual, text="+ Adicionar", bootstyle="secondary-outline",
                   command=self._adicionar_rcs_manual).pack(side="left", padx=(6, 0))

        ttk.Label(card1, style="CardHint.TLabel", justify="left", wraplength=430,
                  text="Selecione linhas na aba BASE_RCs (Ctrl/Shift para várias) e use "
                       "“Adicionar RCs selecionadas”. Duplo clique também adiciona. "
                       "Remova uma RC clicando no × do chip."
                  ).pack(anchor="w", pady=(6, 0))

        linha_grupo = ttk.Frame(card1, style="Card.TFrame")
        linha_grupo.pack(fill="x", pady=(10, 0))
        ttk.Label(linha_grupo, text="Novo Grupo Comprador:", style="CardCampo.TLabel").pack(side="left")
        self.ent_grupo_exc = ttk.Entry(linha_grupo, width=10)
        self.ent_grupo_exc.pack(side="left", padx=(8, 0))
        self.ent_grupo_exc.bind("<Return>", lambda _e: self._executar_redirecionamento())
        ttk.Button(card1, text="Executar redirecionamento", bootstyle="primary",
                   command=self._executar_redirecionamento).pack(fill="x", pady=(10, 0))
        ttk.Label(card1, style="CardHint.TLabel", justify="left", wraplength=430,
                  text="O cadastro dos itens não é alterado. A exceção é registrada na aba "
                       "EXCECOES_RCs e prevalece sobre futuras correções de cadastro."
                  ).pack(anchor="w", pady=(6, 0))

        # ---------- cartão: operação 2
        card2 = ttk.Frame(master, style="Card.TFrame", padding=12)
        card2.pack(fill="x", pady=(12, 0))
        ttk.Label(card2, text="2 · Correção do cadastro de um item",
                  style="CardTitulo.TLabel").pack(anchor="w")
        linha_item = ttk.Frame(card2, style="Card.TFrame")
        linha_item.pack(fill="x", pady=(8, 0))
        ttk.Label(linha_item, text="Código do item:", style="CardCampo.TLabel").pack(side="left")
        self.ent_item = ttk.Entry(linha_item, width=18)
        self.ent_item.pack(side="left", padx=(8, 0))
        linha_grupo2 = ttk.Frame(card2, style="Card.TFrame")
        linha_grupo2.pack(fill="x", pady=(6, 0))
        ttk.Label(linha_grupo2, text="Novo Grupo Comprador:", style="CardCampo.TLabel").pack(side="left")
        self.ent_grupo_item = ttk.Entry(linha_grupo2, width=10)
        self.ent_grupo_item.pack(side="left", padx=(8, 0))
        self.ent_grupo_item.bind("<Return>", lambda _e: self._executar_correcao())
        ttk.Button(card2, text="Executar correção", bootstyle="success",
                   command=self._executar_correcao).pack(fill="x", pady=(10, 0))
        ttk.Label(card2, style="CardHint.TLabel", justify="left", wraplength=430,
                  text="Selecione um item na aba BASE_CADASTRO_ITENS e use “Usar item "
                       "selecionado”. RCs com exceção registrada NÃO são alteradas."
                  ).pack(anchor="w", pady=(6, 0))

    def _montar_abas(self, master):
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill="both", expand=True)

        # ----- aba BASE_RCs
        aba_rcs = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(aba_rcs, text="  BASE_RCs  ")
        self.tab_rcs = Tabela(
            aba_rcs,
            colunas=[
                {"id": "rc", "texto": "RC", "largura": 95, "ancora": "e"},
                {"id": "linha", "texto": "Linha", "largura": 70, "ancora": "center"},
                {"id": "item", "texto": "Código item", "largura": 100, "ancora": "e"},
                {"id": "descricao", "texto": "Descrição", "largura": 300, "esticar": True},
                {"id": "grupo", "texto": "Grupo Comprador", "largura": 140, "ancora": "center"},
            ],
            acao=("＋ Adicionar RCs selecionadas", "primary", self._adicionar_rcs_selecao),
            ao_duplo_clique=self._duplo_clique_rc,
        )
        self.tab_rcs.pack(fill="both", expand=True)

        # ----- aba BASE_CADASTRO_ITENS
        aba_itens = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(aba_itens, text="  BASE_CADASTRO_ITENS  ")
        self.tab_itens = Tabela(
            aba_itens,
            colunas=[
                {"id": "item", "texto": "Código item", "largura": 110, "ancora": "e"},
                {"id": "descricao", "texto": "Descrição", "largura": 300, "esticar": True},
                {"id": "grupo", "texto": "Grupo Comprador", "largura": 150, "ancora": "center"},
            ],
            acao=("Usar item selecionado", "secondary", self._usar_item_selecionado),
            ao_duplo_clique=self._duplo_clique_item,
        )
        self.tab_itens.pack(fill="both", expand=True)

        # ----- aba EXCECOES_RCs
        aba_exc = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(aba_exc, text="  EXCECOES_RCs  ")
        self.tab_exc = Tabela(
            aba_exc,
            colunas=[
                {"id": "rc", "texto": "RC", "largura": 110, "ancora": "e"},
                {"id": "grupo", "texto": "Grupo Comprador Exceção", "largura": 190, "ancora": "center"},
                {"id": "data", "texto": "Registrado em", "largura": 190, "ancora": "center"},
            ],
        )
        self.tab_exc.pack(fill="both", expand=True)

        # ----- aba Histórico
        aba_log = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(aba_log, text="  Histórico  ")
        topo_log = ttk.Frame(aba_log)
        topo_log.pack(fill="x")
        ttk.Label(topo_log, text="Registro das operações executadas nesta sessão:").pack(side="left")
        ttk.Button(topo_log, text="Limpar histórico", bootstyle="secondary-outline",
                   command=self._limpar_log).pack(side="right")
        caixa_log = ttk.Frame(aba_log)
        caixa_log.pack(fill="both", expand=True, pady=(6, 0))
        self.txt_log = tk.Text(caixa_log, background="#212529", foreground="#b9c0cc",
                               insertbackground="#b9c0cc", relief="flat",
                               font=("Consolas", 10), state="disabled", wrap="word",
                               padx=12, pady=10)
        rolagem = ttk.Scrollbar(caixa_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=rolagem.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

    # ------------------------------ utilidades ---------------------------

    def _log(self, mensagem):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {mensagem}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _arquivo_ok(self):
        return self.service is not None and self.service.wb is not None

    def _exigir_arquivo(self):
        if not self._arquivo_ok():
            self._aviso("Carregue o arquivo Excel antes de executar operações.",
                        "Nenhum arquivo carregado")
            return False
        return True

    def _erro_ao_salvar(self):
        self._log("ERRO ao salvar: o arquivo parece estar aberto no Microsoft Excel. "
                  "Nenhuma alteração foi gravada.")
        self._erro("Não foi possível salvar a planilha.\n"
                   "Verifique se o arquivo não está aberto no Microsoft Excel e tente novamente.",
                   "Erro ao salvar")
        self._atualizar_status()

    def _atualizar_status(self):
        if not self._arquivo_ok():
            self.lbl_arq_status.configure(text="●  Nenhum arquivo carregado", style="StatusErr.TLabel")
            self.lbl_contagens.configure(text="")
            return
        r = self.service.resumo()
        self.lbl_arq_status.configure(
            text=f"●  {os.path.basename(self.service.caminho)}", style="StatusOk.TLabel")
        self.lbl_contagens.configure(
            text=f"{r['itens']} itens   ·   {r['linhas_rc']} linhas de RC   ·   "
                 f"{r['excecoes']} exceção(ões) registrada(s)")

    def _atualizar_tabelas(self):
        if not self._arquivo_ok():
            self.tab_rcs.definir_dados([])
            self.tab_itens.definir_dados([])
            self.tab_exc.definir_dados([])
            self._rcs_validas = set()
            return
        rcs = self.service.linhas_rcs()
        self.tab_rcs.definir_dados(rcs)
        self._rcs_validas = {_to_int(l[0]) for l in rcs if _to_int(l[0]) is not None}
        self.tab_itens.definir_dados(self.service.linhas_itens())
        self.tab_exc.definir_dados(self.service.linhas_excecoes())

    # ------------------------------ arquivo ------------------------------

    def _carregar_arquivo_inicial(self):
        caminho = localizar_arquivo_base(pasta_do_app())
        if caminho:
            self.var_caminho.set(caminho)
            self._recarregar(silencioso=True)
            self._log(f"Arquivo base localizado automaticamente: {os.path.basename(caminho)}")
        else:
            self._log("Nenhum arquivo base encontrado automaticamente. "
                      "Use “Procurar…” para selecionar a planilha.")

    def _procurar_arquivo(self):
        from tkinter import filedialog
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha base",
            filetypes=[("Planilha Excel", "*.xlsx *.xlsm"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.var_caminho.set(caminho)
            self._recarregar()

    def _recarregar(self, silencioso=False):
        """Carrega (ou recarrega) o arquivo indicado no campo 'Arquivo Excel'."""
        caminho = self.var_caminho.get().strip().strip('"')
        if not caminho:
            self._aviso("Informe ou selecione o arquivo Excel.", "Arquivo não informado")
            return
        if not os.path.isfile(caminho):
            self._erro(f"Arquivo não encontrado:\n{caminho}", "Arquivo inválido")
            return
        try:
            service = ExcelService(caminho)
            service.carregar()
            reaplicados = service.reaplicar_excecoes()
        except PermissionError:
            self._erro_ao_salvar()
            return
        except Exception as exc:
            self._erro(f"Não foi possível abrir o arquivo:\n{exc}", "Erro ao carregar")
            return
        self.service = service
        self._atualizar_tabelas()
        self._atualizar_status()

        # remove da seleção RCs que não existem mais na base
        obsoletas = [rc for rc in list(self.chips.valores) if rc not in self._rcs_validas]
        for rc in obsoletas:
            self.chips.remover(rc)
        if obsoletas:
            self._log("RC(s) removida(s) da seleção por não existirem mais na BASE_RCs: "
                      + ", ".join(map(str, obsoletas)))

        r = service.resumo()
        self._log(f"Arquivo carregado: {os.path.basename(caminho)} "
                  f"({r['itens']} itens, {r['linhas_rc']} linhas de RC, {r['excecoes']} exceções)")
        if reaplicados:
            self._log(f"{reaplicados} linha(s) de RC atualizada(s) automaticamente para "
                      "respeitar as exceções registradas.")

    # ------------------------------ seleção de RCs ----------------------

    def _ao_mudar_chips(self, valores):
        self.lbl_qtd_rcs.configure(text=f"RCs selecionadas: {len(valores)}")

    def _limpar_selecao_rcs(self):
        self.chips.limpar()
        self._log("Seleção de RCs limpa.")

    def _adicionar_rcs_selecao(self):
        if not self._exigir_arquivo():
            return
        linhas = self.tab_rcs.linhas_selecionadas()
        if not linhas:
            self._aviso("Selecione ao menos uma linha na tabela BASE_RCs.\n"
                        "Use Ctrl ou Shift para selecionar várias.", "Nenhuma linha selecionada")
            return
        rcs = [_to_int(l[0]) for l in linhas]
        rcs = [rc for rc in rcs if rc is not None]
        novos = self.chips.adicionar(rcs)
        if novos:
            self._log(f"{len(novos)} RC(s) adicionada(s) à seleção pela tabela: "
                      + ", ".join(map(str, novos)))

    def _duplo_clique_rc(self, linha):
        rc = _to_int(linha[0])
        if rc is not None and self.chips.adicionar([rc]):
            self._log(f"RC {rc} adicionada à seleção (duplo clique).")

    def _adicionar_rcs_manual(self):
        if not self._exigir_arquivo():
            return
        texto = self.ent_rcs_manual.get().strip()
        if not texto:
            self._aviso("Digite uma RC ou uma lista de RCs no campo.", "Nada informado")
            return
        valores, invalidos = _lista_ints(texto)
        avisos = []
        if invalidos:
            avisos.append("Valor(es) inválido(s) ignorado(s): " + ", ".join(invalidos))
        validas = [rc for rc in valores if rc in self._rcs_validas]
        inexistentes = [rc for rc in valores if rc not in self._rcs_validas]
        if inexistentes:
            avisos.append("RC(s) não encontrada(s) na BASE_RCs (não adicionada[s]): "
                          + ", ".join(map(str, inexistentes)))
        novos = self.chips.adicionar(validas)
        self.ent_rcs_manual.delete(0, "end")
        if novos:
            self._log(f"{len(novos)} RC(s) adicionada(s) manualmente: " + ", ".join(map(str, novos)))
        if avisos:
            self._aviso("\n".join(avisos), "Atenção")
        elif not novos:
            self._log("Nenhuma RC nova adicionada (já estavam na seleção).")

    def _usar_item_selecionado(self):
        linhas = self.tab_itens.linhas_selecionadas()
        if not linhas:
            self._aviso("Selecione um item na tabela BASE_CADASTRO_ITENS.",
                        "Nenhum item selecionado")
            return
        item = _to_int(linhas[0][0])
        if item is None:
            self._erro("Não foi possível ler o código do item selecionado.", "Erro")
            return
        self.ent_item.delete(0, "end")
        self.ent_item.insert(0, str(item))
        self.ent_grupo_item.focus_set()
        self._log(f"Item {item} levado para a operação 2 (seleção na tabela).")

    def _duplo_clique_item(self, linha):
        item = _to_int(linha[0])
        if item is None:
            return
        self.ent_item.delete(0, "end")
        self.ent_item.insert(0, str(item))
        self._log(f"Item {item} levado para a operação 2 (duplo clique).")

    # ------------------------------ operação 1 --------------------------

    def _executar_redirecionamento(self):
        if not self._exigir_arquivo():
            return
        rcs = list(self.chips.valores)
        if not rcs:
            self._aviso("Selecione ou adicione ao menos uma RC antes de executar o "
                        "redirecionamento.", "RCs não informadas")
            return
        novo_grupo = _to_int(self.ent_grupo_exc.get())
        if novo_grupo is None:
            self._aviso("Informe o novo Grupo Comprador (número inteiro).",
                        "Grupo Comprador não informado")
            return

        try:
            res = self.service.redirecionar_rcs(rcs, novo_grupo,
                                                confirmar_substituicao=self._confirmar_substituicao)
        except PermissionError:
            self._erro_ao_salvar()
            return
        except Exception as exc:
            self._erro(f"Ocorreu um erro ao executar a operação:\n{exc}", "Erro")
            self._log(f"ERRO no redirecionamento: {exc}")
            return

        fmt = lambda lista: ", ".join(str(v) for v in lista)
        partes = []
        if res["atualizadas"]:
            partes.append(f"RC(s) redirecionada(s) para o grupo {novo_grupo}: "
                          f"{fmt(res['atualizadas'])}")
        if res["substituidas"]:
            partes.append(f"RC(s) cuja exceção anterior foi substituída: {fmt(res['substituidas'])}")
        if res["ja_mesmo_grupo"]:
            partes.append(f"RC(s) que já possuíam exceção para o grupo {novo_grupo} "
                          f"(nenhuma alteração): {fmt(res['ja_mesmo_grupo'])}")
        if res["mantidas"]:
            partes.append(f"RC(s) mantidas com a exceção anterior (substituição cancelada): "
                          f"{fmt(res['mantidas'])}")
        if res["nao_encontradas"]:
            partes.append(f"RC(s) NÃO ENCONTRADA(S) na aba {ABA_RCS}: {fmt(res['nao_encontradas'])}")

        resumo = "\n".join(partes)
        self._log("Redirecionamento por exceção → " + " | ".join(partes))

        # RCs já tratadas saem da seleção; as mantidas/não encontradas continuam
        for rc in res["atualizadas"] + res["ja_mesmo_grupo"]:
            self.chips.remover(rc)

        self._atualizar_tabelas()
        self._atualizar_status()
        if res["atualizadas"] or res["ja_mesmo_grupo"]:
            self._info(resumo + "\n\nAs alterações foram salvas na planilha.",
                       "Redirecionamento concluído")
        else:
            self._aviso("Nenhuma RC foi alterada.\n\n" + resumo, "Redirecionamento")

    def _confirmar_substituicao(self, pares, novo_grupo):
        detalhe = "\n".join(f"  • RC {rc}: exceção atual = grupo {grupo}" for rc, grupo in pares)
        return self._sim_nao(
            "As RCs abaixo já possuem exceção registrada:\n\n" + detalhe +
            f"\n\nDeseja substituir a exceção pelo novo grupo {novo_grupo}?\n\n"
            "Sim = substitui e redireciona as RCs\nNão = mantém a exceção atual",
            "RCs já possuem exceção")

    # ------------------------------ operação 2 --------------------------

    def _executar_correcao(self):
        if not self._exigir_arquivo():
            return
        codigo = _to_int(self.ent_item.get())
        if codigo is None:
            self._aviso("Informe o código do item (número inteiro).", "Item não informado")
            return
        novo_grupo = _to_int(self.ent_grupo_item.get())
        if novo_grupo is None:
            self._aviso("Informe o novo Grupo Comprador (número inteiro).",
                        "Grupo Comprador não informado")
            return

        try:
            res = self.service.corrigir_cadastro_item(codigo, novo_grupo)
        except PermissionError:
            self._erro_ao_salvar()
            return
        except Exception as exc:
            self._erro(f"Ocorreu um erro ao executar a operação:\n{exc}", "Erro")
            self._log(f"ERRO na correção de cadastro: {exc}")
            return

        if not res["item_encontrado"]:
            self._log(f"Correção de cadastro → item {codigo} NÃO encontrado na aba {ABA_ITENS}.")
            self._aviso(f"O item {codigo} não foi encontrado na aba {ABA_ITENS}.",
                        "Item não encontrado")
            return

        partes = [
            f"Cadastro do item {codigo} atualizado para o grupo {novo_grupo}.",
            f"Linhas de RC atualizadas na aba {ABA_RCS}: {res['linhas_atualizadas']}.",
        ]
        if res["rcs_preservadas"]:
            detalhe = "\n".join(
                f"  • RC {rc} (linha {linha}): permanece no grupo {grupo} — exceção registrada"
                for rc, linha, grupo in res["rcs_preservadas"])
            partes.append(f"Linhas preservadas por exceção: {len(res['rcs_preservadas'])}\n{detalhe}")

        mensagem = "\n".join(partes)
        self._log("Correção de cadastro → " + mensagem.replace("\n", " | "))

        self.ent_item.delete(0, "end")
        self.ent_grupo_item.delete(0, "end")
        self._atualizar_tabelas()
        self._atualizar_status()
        self._info(mensagem + "\n\nAs alterações foram salvas na planilha.",
                   "Correção concluída")


# ============================== Execução ================================

def main():
    _habilitar_dpi()
    try:
        app = App()
        app.mainloop()
    except Exception as exc:  # fallback com detalhes, caso a interface não suba
        traceback.print_exc()
        try:
            import tkinter.messagebox as mb
            mb.showerror("Erro", f"Erro ao iniciar o aplicativo:\n{exc}")
        except Exception:
            pass


if __name__ == "__main__":
    main()
