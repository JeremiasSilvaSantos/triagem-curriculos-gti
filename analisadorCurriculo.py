import os
import re
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


class CurriculoProcessor:
    """Camada de Regra de Negócio e Processamento de Currículos."""

    @staticmethod
    def extrair_texto_pdf(caminho_pdf: Path) -> str:
        """Extrai texto bruto de um arquivo PDF com tratamento de exceções."""
        if not PdfReader:
            return ""
        try:
            reader = PdfReader(caminho_pdf)
            texto_total = ""
            for pagina in reader.pages:
                conteudo = pagina.extract_text()
                if conteudo:
                    texto_total += conteudo
            return texto_total.lower()
        except Exception as e:
            print(f"[Log Técnico] Erro ao ler PDF {caminho_pdf.name}: {e}")
            return ""

    @staticmethod
    def analisar(texto: str, nome_arquivo: str, termos_extras: list) -> dict:
        """Aplica o motor de pontuação ponderada baseado nos critérios da vaga de GTI."""
        pontuacao = 0
        detalhes = {
            "Arquivo": nome_arquivo,
            "Anos de Experiência": 0,
            "Formação GTI": "Não",
            "Pós-Graduação": "Não",
            "Senioridade": "Júnior/Não especificado",
            "Nível de Inglês": "Básico/Intermediário",
            "Qtd Cursos/Certificados": 0,
            "Competências Extras": "",
            "Pontuação Total": 0,
        }

        if not texto:
            return detalhes

        # 1. Anos de Experiência (Regex)
        matches_exp = re.findall(r"(\d+)\s*(?:anos?|ans?)\s*(?:de)?\s*experiência", texto)
        if matches_exp:
            try:
                anos = max([int(x) for x in matches_exp])
                detalhes["Anos de Experiência"] = anos
                if anos >= 5:
                    pontuacao += 30
                elif anos >= 3:
                    pontuacao += 15
            except ValueError:
                pass

        # 2. Formação Acadêmica (GTI)
        if "gestão da tecnologia da informação" in texto or "gti" in texto:
            detalhes["Formação GTI"] = "Sim"
            pontuacao += 20

        # 3. Pós-Graduação / Especialização
        if any(termo in texto for termo in ["pós-graduação", "pos-graduacao", "especialização", "mestrado", "mba"]):
            detalhes["Pós-Graduação"] = "Sim"
            pontuacao += 15

        # 4. Senioridade
        if "sênior" in texto or "senior" in texto:
            detalhes["Senioridade"] = "Sênior"
            pontuacao += 25
        elif "pleno" in texto:
            detalhes["Senioridade"] = "Pleno"
            pontuacao += 15

        # 5. Idioma (Inglês Avançado/Fluente)
        if any(termo in texto for termo in ["inglês fluente", "ingles fluente", "advanced english", "fluent english", "inglês avançado", "ingles avancado", "c1", "c2"]):
            detalhes["Nível de Inglês"] = "Avançado/Fluente"
            pontuacao += 20

        # 6. Portfólio / Cursos Tecnológicos Base
        palavras_cursos = ["certificação", "certificacao", "curso", "alura", "udemy", "aws", "azure", "scrum", "itil", "sql"]
        qtd_cursos = sum(texto.count(palavra) for palavra in palavras_cursos)
        detalhes["Qtd Cursos/Certificados"] = qtd_cursos
        if qtd_cursos >= 10:
            pontuacao += 20
        elif qtd_cursos >= 5:
            pontuacao += 10

        # 7. Competências Adicionais Customizadas
        encontrados = []
        if termos_extras:
            for termo in termos_extras:
                termo_limpo = termo.strip().lower()
                if termo_limpo and termo_limpo in texto:
                    encontrados.append(termo.strip())
                    pontuacao += 5

        detalhes["Competências Extras"] = ", ".join(encontrados) if encontrados else "Nenhuma"
        detalhes["Pontuação Total"] = pontuacao
        return detalhes


class ExcelReportBuilder:
    """Camada de Persistência e Estilização de Planilhas Corporativas."""

    @staticmethod
    def gerar_relatorio(caminho_saida: str, dados: list) -> bool:
        """Gera o arquivo Excel estruturado e aplica design profissional via openpyxl."""
        try:
            df = pd.DataFrame(dados)
            df.to_excel(caminho_saida, index=False)

            wb = openpyxl.load_workbook(caminho_saida)
            ws = wb.active

            # Paleta de Cores Corporativa
            AZUL_CABECALHO = "1F4E78"
            BRANCO = "FFFFFF"
            VERDE_DESTAQUE = "C6EFCE"
            VERDE_TEXTO = "006100"
            CINZA_ZEBRA = "F2F2F2"

            font_header = Font(name="Calibri", size=11, bold=True, color=BRANCO)
            fill_header = PatternFill(start_color=AZUL_CABECALHO, end_color=AZUL_CABECALHO, fill_type="solid")
            border_thin = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
            fill_verde = PatternFill(start_color=VERDE_DESTAQUE, end_color=VERDE_DESTAQUE, fill_type="solid")
            font_verde = Font(name="Calibri", size=11, bold=True, color=VERDE_TEXTO)
            fill_zebra = PatternFill(start_color=CINZA_ZEBRA, end_color=CINZA_ZEBRA, fill_type="solid")

            # Identifica coluna de pontuação
            col_pontuacao_idx = None
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col).value == "Pontuação Total":
                    col_pontuacao_idx = col
                    break

            # Formata Cabeçalho
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border_thin
            ws.row_dimensions[1].height = 25

            # Formata Linhas de Dados
            for row in range(2, ws.max_row + 1):
                is_zebra = (row % 2 == 0)
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border_thin
                    cell.alignment = Alignment(horizontal="left" if col <= 2 else "center", vertical="center")
                    
                    cell.fill = fill_verde
                    if col_pontuacao_idx and col == col_pontuacao_idx:
                        cell.font = font_verde
                ws.row_dimensions[row].height = 20

            # Auto-dimensionamento de colunas
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

            wb.save(caminho_saida)
            return True

        except PermissionError:
            messagebox.showerror("Erro de Permissão", f"O arquivo '{caminho_saida}' está aberto no Excel!\nFeche-o e tente novamente.")
            return False
        except Exception as e:
            messagebox.showerror("Erro de Exportação", f"Falha ao gerar planilha:\n{e}")
            return False


class CurriculoAppUI:
    """Camada de Apresentação (Interface Gráfica Tkinter)."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Sistema de Triagem de Currículos | GTI - TCC")
        self.root.geometry("640x620")
        self.root.resizable(False, False)

        self.caminho_pasta_var = tk.StringVar()
        self.caminho_saida_var = tk.StringVar(value="relatorio_candidatos_gti.xlsx")
        self.competencias_var = tk.StringVar(value="python, cloud, scrum, itil, sql")
        self.nota_corte_var = tk.IntVar(value=50)

        self._construir_interface()

    def _construir_interface(self):
        # Título
        lbl_titulo = tk.Label(self.root, text="🎓 Sistema de Triagem de Currículos (GTI)", font=("Calibri", 16, "bold"), fg="#1F4E78")
        lbl_titulo.pack(pady=12)

        # 1. Pasta
        f_pasta = tk.LabelFrame(self.root, text=" 1. Diretório de Currículos (PDF) ", font=("Calibri", 10, "bold"), padx=10, pady=6)
        f_pasta.pack(fill="x", padx=20, pady=5)
        tk.Entry(f_pasta, textvariable=self.caminho_pasta_var, width=54, font=("Calibri", 10)).pack(side="left", padx=5)
        tk.Button(f_pasta, text="Procurar...", command=self._selecionar_pasta, bg="#1F4E78", fg="white", font=("Calibri", 9, "bold"), width=10).pack(side="right", padx=5)

        # 2. Competências
        f_comp = tk.LabelFrame(self.root, text=" 2. Competências / Palavras-Chave Adicionais (Vírgula) ", font=("Calibri", 10, "bold"), padx=10, pady=6)
        f_comp.pack(fill="x", padx=20, pady=5)
        tk.Entry(f_comp, textvariable=self.competencias_var, width=68, font=("Calibri", 10)).pack(padx=5, pady=2)

        # 3. Nota de Corte
        f_corte = tk.LabelFrame(self.root, text=" 3. Nota de Corte Mínima (Filtro de Desempenho) ", font=("Calibri", 10, "bold"), padx=10, pady=6)
        f_corte.pack(fill="x", padx=20, pady=5)
        tk.Spinbox(f_corte, from_=0, to=300, textvariable=self.nota_corte_var, width=12, font=("Calibri", 11, "bold")).pack(side="left", padx=5)
        tk.Label(f_corte, text="(Apenas candidatos com pontuação igual ou superior)", font=("Calibri", 9, "italic"), fg="gray").pack(side="left", padx=10)

        # 4. Saída
        f_saida = tk.LabelFrame(self.root, text=" 4. Caminho do Relatório Excel de Saída ", font=("Calibri", 10, "bold"), padx=10, pady=6)
        f_saida.pack(fill="x", padx=20, pady=5)
        tk.Entry(f_saida, textvariable=self.caminho_saida_var, width=54, font=("Calibri", 10)).pack(side="left", padx=5)
        tk.Button(f_saida, text="Salvar Como", command=self._selecionar_saida, bg="#595959", fg="white", font=("Calibri", 9, "bold"), width=10).pack(side="right", padx=5)

        # Botão Executar
        self.btn_executar = tk.Button(self.root, text="🚀 EXECUTAR ANÁLISE DE CURRÍCULOS", command=self._iniciar_processamento_thread, bg="#2E7D32", fg="white", font=("Calibri", 11, "bold"), height=2, width=38)
        self.btn_executar.pack(pady=15)

        # Barra de Progresso e Status
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=580, mode="indeterminate")
        self.progress.pack(pady=5)

        self.lbl_status = tk.Label(self.root, text="Status: Aguardando diretório de entrada...", font=("Calibri", 9, "italic"), fg="gray")
        self.lbl_status.pack(side="bottom", pady=8)

    def _selecionar_pasta(self):
        dir_path = filedialog.askdirectory(title="Selecionar Pasta com Currículos PDF")
        if dir_path:
            self.caminho_pasta_var.set(dir_path)
            self.lbl_status.config(text=f"Diretório selecionado: {dir_path}")

    def _selecionar_saida(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], title="Salvar Relatório Como")
        if file_path:
            self.caminho_saida_var.set(file_path)

    def _iniciar_processamento_thread(self):
        self.btn_executar.config(state="disabled")
        self.progress.start(12)
        thread = threading.Thread(target=self._processar_pipeline)
        thread.daemon = True
        thread.start()

    def _processar_pipeline(self):
        try:
            pasta = self.caminho_pasta_var.get().strip()
            saida = self.caminho_saida_var.get().strip()
            extras_str = self.competencias_var.get().strip()
            nota_corte = self.nota_corte_var.get()

            if not pasta or not os.path.isdir(pasta):
                messagebox.showerror("Erro", "O diretório de currículos informado é inválido!")
                return

            if not saida:
                messagebox.showerror("Erro", "Defina o caminho de saída para o arquivo Excel!")
                return

            termos_extras = [t.strip() for t in extras_str.split(",") if t.strip()]
            pasta_path = Path(pasta)
            arquivos_pdf = list(pasta_path.glob("*.pdf"))

            if not arquivos_pdf:
                messagebox.showwarning("Aviso", "Nenhum arquivo PDF encontrado na pasta selecionada.")
                return

            self.lbl_status.config(text="Status: Lendo e pontuando currículos...")
            self.root.update_idletasks()

            resultados = []
            erros = 0

            for pdf in arquivos_pdf:
                texto = CurriculoProcessor.extrair_texto_pdf(pdf)
                if not texto.strip():
                    erros += 1
                    continue
                analise = CurriculoProcessor.analisar(texto, pdf.name, termos_extras)
                resultados.append(analise)

            if not resultados:
                messagebox.showwarning("Aviso", "Nenhum currículo pôde ser extraído (arquivos vazios ou protegidos).")
                return

            df = pd.DataFrame(resultados)
            df = df[df["Pontuação Total"] >= nota_corte]

            if df.empty:
                messagebox.showwarning("Corte de Nota", f"Nenhum candidato atingiu a nota mínima de {nota_corte} pontos.")
                return

            df = df.sort_values(by="Pontuação Total", ascending=False).reset_index(drop=True)

            sucesso = ExcelReportBuilder.gerar_relatorio(saida, df.to_dict(orient="records"))

            if sucesso:
                msg_erros = f"\n({erros} PDF(s) ilegíveis ignorados)" if erros > 0 else ""
                self.lbl_status.config(text=f"Status: Concluído com sucesso! {len(df)} candidatos aprovados.")
                messagebox.showinfo("Sucesso", f"Processo finalizado!\n\n{len(df)} candidatos aprovados salvos em:\n{saida}{msg_erros}")

        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
            self.lbl_status.config(text="Status: Erro na execução.")
        finally:
            self.progress.stop()
            self.btn_executar.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = CurriculoAppUI(root)
    root.mainloop()