# Importa a biblioteca customtkinter, versão personalizada do Tkinter
import customtkinter as ctk
# Importa messagebox do Tkinter para exibir mensagens de alerta e informação
from tkinter import messagebox
# Importa load_workbook da openpyxl para manipular arquivos Excel
from openpyxl import load_workbook
# Importa Playwright para automação de navegador
from playwright.sync_api import sync_playwright
# Importa pyperclip para copiar conteúdo para o clipboard
import pyperclip
# Importa biblioteca os para manipulação de arquivos e caminhos
import os
# Importa expressões regulares para buscar padrões em texto
import re
# Importa win32com.client para controlar o Outlook via COM
import win32com.client as win32
# Importa datetime para manipulação de datas
from datetime import datetime

# -------------------- Configurações SAP Web -------------------- #
SAP_WEB_URL = "https://ps0.wdisp.bosch.com/sap/bc/gui/sap/its/webgui#"
STORAGE_STATE_PATH = "sap_session.json"

# -------------------- Função para converter mês/ano -------------------- #
def mes_ano_para_formato_curto(mes_ano):
    try:
        mes_nome, ano = mes_ano.split("/")
        mes_nome = mes_nome.strip().lower()
        meses = {
            "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4,
            "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
            "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
        }
        mes_num = meses.get(mes_nome, 0)
        ano_curto = str(ano.strip())[-2:]
        return f"{mes_num:02d}.{ano_curto}" if mes_num else mes_ano
    except Exception as e:
        print(f"Erro ao converter mês/ano: {e}")
        return mes_ano

# -------------------- Atualizar planilha -------------------- #
# COMENTADO - Procura pelo arquivo Excel desabilitada
# caminho_excel = r"C:\Users\ajl8ca\Desktop\HRS_Projects_Dev\cobranca_r&s\Controle_Cobranca_PES.XLSX"

# def atualizar_planilha(encontrados, numero):
#     try:
#         wb = load_workbook(caminho_excel)
#         ws = wb.active
#         cabecalho = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
#         idx_id = cabecalho.index("ID Vaga")
#         idx_mes = cabecalho.index("Mês/Ano")
#         idx_cobranca = cabecalho.index("Número Cobrança")

#         for row in ws.iter_rows(min_row=2):
#             id_vaga = row[idx_id].value
#             mes_ano = row[idx_mes].value
#             for d in encontrados:
#                 if d["id"] == id_vaga and d["mes"] == mes_ano:
#                     row[idx_cobranca].value = numero

#         wb.save(caminho_excel)
#         wb.close()
#         print(f"Planilha atualizada com o número {numero}")

#     except Exception as e:
#         messagebox.showerror("Erro Excel", f"Falha ao atualizar planilha:\n{e}")

# -------------------- Abrir SAP Web -------------------- #
def abrir_sap_web(mes_ano, encontrados):
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=False,
                channel="chrome",
                args=[
                    "--enable-features=ClipboardReadWrite",
                    "--disable-features=IsolateOrigins,site-per-process",
                    "--unsafely-treat-insecure-origin-as-secure=https://ps0.wdisp.bosch.com"
                ]
            )

            if os.path.exists(STORAGE_STATE_PATH):
                context = browser.new_context(
                    storage_state=STORAGE_STATE_PATH,
                    permissions=["clipboard-read", "clipboard-write"]
                )
            else:
                context = browser.new_context(
                    permissions=["clipboard-read", "clipboard-write"]
                )

            page = context.new_page()
            page.set_viewport_size({"width": 1366, "height": 768})
            page.goto(SAP_WEB_URL)
            page.evaluate("document.body.style.zoom='80%'")
            
            # 👇 se aparecer a tela de login (campo sap-user)
            try:
                login_user = page.locator('//*[@id="sap-user"]')
                if login_user.is_visible():
                    print("Tela de login detectada — aguarde login manual...")
                    page.wait_for_selector('//*[@id="sap-password"]', timeout=15000)
                    # 
                    # login_user.fill("seu_usuario")
                    # page.fill('//*[@id="sap-password"]', "sua_senha")
                    # page.press('//*[@id="sap-password"]', 'Enter')
                    # ou apenas esperar manualmente
                    page.wait_for_function("document.querySelector('#sap-user') === null", timeout=120000)
                    print("Login realizado, continuando...")
            except Exception:
                pass
            
            page.wait_for_timeout(5000)

            from playwright.sync_api import TimeoutError

            for tentativa in range(3):
                try:
                    page.wait_for_selector('//*[@id="ToolbarOkCode"]', timeout=10000)
                    page.fill('//*[@id="ToolbarOkCode"]', '/nKB31N')
                    page.press('//*[@id="ToolbarOkCode"]', 'Enter')
                    page.wait_for_timeout(2000)  # espera processar
                    break
                except TimeoutError:
                    print(f"Tentativa {tentativa+1} falhou, tentando novamente...")

            page.wait_for_timeout(3000)

            # -------------------- VERIFICAÇÃO DO CAMPO EMPRESA -------------------- #
            try:
                campo_empresa = page.locator('//*[@id="M1:46:1::0:21"]')
                if campo_empresa.is_visible():
                    campo_empresa.fill('0010')
                    campo_empresa.press('Enter')
                    page.wait_for_timeout(3000)
                    print("Campo empresa encontrado e preenchido.")
                else:
                    print("Tela de empresa não apareceu. Pulando...")
            except Exception as e:
                print(f"Campo empresa não encontrado. Pulando... ({e})")
            # ---------------------------------------------------------------------- #

            page.click('//*[@id="M0:46:2:1::13:35"]')
            page.wait_for_timeout(3000)

            page.fill('//*[@id="M0:46:1:1:2B256::4:12"]', f"Recuperação Custo {mes_ano} R&S - PES")
            page.press('//*[@id="M0:46:1:1:2B256::4:12"]', 'Enter')
            page.wait_for_timeout(2000)
            
            page.click('//*[@id="M0:36::btn[11]"]')

            mensagem = page.inner_text('xpath=//*[@id="wnd[0]/sbar_msg-txt"]')
            numero = re.search(r"\d+", mensagem).group()

            pyperclip.copy(numero)
            print(f"Número copiado: {numero}")

            formato_curto = mes_ano_para_formato_curto(mes_ano)
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "Alessandro.Garbelini@br.bosch.com"
            mail.Subject = f"Apontamentos HRS - MBR{formato_curto}"
            mail.Display()
            mail.HTMLBody = f"""
            Olá!<br><br>
            Segue a chave referente ao apontamento de {mes_ano}.<br>
            Nº {numero}<br><br>
            Atenciosamente,<br>
            {mail.HTMLBody}
            """

            messagebox.showinfo("SAP Web", f"Processo Finalizado.\nChave: {numero}")
            # COMENTADO - Atualização da planilha desabilitada
            # atualizar_planilha(encontrados, numero)

            context.storage_state(path=STORAGE_STATE_PATH)
            browser.close()
            app.destroy()

    except Exception as e:
        messagebox.showerror("Erro SAP Web", f"Não foi possível seguir com o SAP Web.\n")
        app.destroy()

# -------------------- Excel - Carregar dados -------------------- #
# COMENTADO - Carregamento de dados do Excel desabilitado
# def carregar_dados():
#     try:
#         wb = load_workbook(caminho_excel, read_only=True, data_only=True)
#         ws = wb.active
#         cabecalho = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
#         colunas_necessarias = ["ID Vaga", "Nome do Aprovado", "Centro cst", "Mês/Ano", "Número Cobrança"]
#         for col in colunas_necessarias:
#             if col not in cabecalho:
#                 raise ValueError(f"Coluna '{col}' não encontrada.")

#         idx_id = cabecalho.index("ID Vaga")
#         idx_nome = cabecalho.index("Nome do Aprovado")
#         idx_centro = cabecalho.index("Centro cst")
#         idx_mes = cabecalho.index("Mês/Ano")
#         idx_cobranca = cabecalho.index("Número Cobrança")

#         dados = []
#         meses_unicos = []
#         for row in ws.iter_rows(min_row=2, values_only=True):
#             if row[idx_mes] is None:
#                 continue
#             mes = str(row[idx_mes])
#             if mes not in meses_unicos:
#                 meses_unicos.append(mes)
#             dados.append({
#                 "mes": mes,
#                 "id": row[idx_id],
#                 "nome": row[idx_nome],
#                 "centro": row[idx_centro],
#                 "cobranca": row[idx_cobranca]
#             })

#         wb.close()
#         return meses_unicos, dados

#     except Exception as e:
#         messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{e}")
#         return [], []

# -------------------- UI -------------------- #
# TEMPORÁRIO - Dados de exemplo (substituir quando tiver novo arquivo)
meses_anos = ["Janeiro/2026", "Fevereiro/2026", "Março/2026"]
dados_planilha = []

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Faturamento R&S")

ctk.CTkLabel(app, text="Faturamento R&S", font=("Arial", 18, "bold")).pack(pady=10)
ctk.CTkLabel(app, text="Selecione o Mês/Ano:").pack(pady=0)
combo = ctk.CTkComboBox(app, values=meses_anos, width=200, justify="center")
combo.pack(pady=5)

frame_confirmar = ctk.CTkFrame(app, fg_color="transparent", height=50)
frame_confirmar.pack(fill="x", pady=10)

btn_confirmar = ctk.CTkButton(frame_confirmar, text="Confirmar", width=150)
btn_confirmar.place(relx=0.5, rely=0.5, anchor="center")

filtro_var = ctk.BooleanVar(value=False)
check_filtro = ctk.CTkCheckBox(frame_confirmar, text="Filtrar vazias", variable=filtro_var)
check_filtro.place(relx=0.8, rely=0.5, anchor="center")

scrollable_frame = ctk.CTkScrollableFrame(app)
scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)

headers = ["ID Vaga", "Nome do Aprovado", "Centro de Custo", "Número Cobrança"]
for col, h in enumerate(headers):
    ctk.CTkLabel(scrollable_frame, text=h, font=("Arial", 12, "bold")).grid(
        row=0, column=col, padx=5, pady=5, sticky="nsew"
    )
    scrollable_frame.grid_columnconfigure(col, weight=1)

executar_btn = ctk.CTkButton(app, text="Executar", command=lambda: executar())
executar_btn.pack(pady=10)

# -------------------- Funções Auxiliares -------------------- #
def preparar_clipboard(dados_filtrados):
    linhas = []
    for d in dados_filtrados:
        centro = str(d["centro"])
        indice = "HRSR26"
        qtd = "1"
        texto = "Tipo de serviço PES"
        linha = f"{centro}\t{indice}\t{qtd}\t{texto}"
        linhas.append(linha)
    conteudo = "\n".join(linhas)
    pyperclip.copy(conteudo)
    return conteudo

# -------------------- Funções UI -------------------- #
def atualizar_tabela():
    escolhido = combo.get()
    for widget in scrollable_frame.winfo_children()[len(headers):]:
        widget.destroy()

    encontrados = [d for d in dados_planilha if d["mes"] == escolhido]
    if filtro_var.get():
        encontrados = [d for d in encontrados if d["cobranca"] in (None, "")]

    for row_idx, d in enumerate(encontrados, start=1):
        bg_color = "#f5f5f5" if row_idx % 2 == 0 else "#ffffff"
        ctk.CTkLabel(scrollable_frame, text=d["id"], bg_color=bg_color).grid(row=row_idx, column=0, padx=5, pady=2, sticky="nsew")
        ctk.CTkLabel(scrollable_frame, text=d["nome"], bg_color=bg_color).grid(row=row_idx, column=1, padx=5, pady=2, sticky="nsew")
        ctk.CTkLabel(scrollable_frame, text=d["centro"], bg_color=bg_color).grid(row=row_idx, column=2, padx=5, pady=2, sticky="nsew")
        ctk.CTkLabel(scrollable_frame, text=d["cobranca"], bg_color=bg_color).grid(row=row_idx, column=3, padx=5, pady=2, sticky="nsew")

filtro_var.trace_add("write", lambda *args: atualizar_tabela())

def executar():
    escolhido = combo.get()
    if not escolhido:
        messagebox.showwarning("Atenção", "Selecione um Mês/Ano antes de executar.")
        return

    encontrados = [d for d in dados_planilha if d["mes"] == escolhido]
    if filtro_var.get():
        encontrados = [d for d in encontrados if d["cobranca"] in (None, "")]

    if not encontrados:
        messagebox.showinfo("Info", f"Nenhum dado encontrado para {escolhido}")
        return

    preparar_clipboard(encontrados)

    resposta = messagebox.askyesno(
        "Confirmação",
        f"Foram encontrados {len(encontrados)} registros.\n\nDeseja mesmo executar a cobrança\nreferente a {escolhido}?"
    )

    if resposta:
        abrir_sap_web(escolhido, encontrados)
    else:
        messagebox.showinfo("Cancelado", "Operação cancelada.")

def confirmar():
    atualizar_tabela()

btn_confirmar.configure(command=confirmar)

# Maximizar a janela após todos os widgets serem criados
app.after(0, lambda: app.state('zoomed'))

app.mainloop()
