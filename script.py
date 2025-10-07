import pandas as pd
import win32com.client as win32
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from datetime import datetime
import traceback
import time

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SaleMachine - Enviador de Emails")
        self.root.geometry("900x750")
        
        # Variáveis
        self.csv_path = tk.StringVar()
        self.assunto = tk.StringVar(value="Teste Robozinho From SaleMachine rsss")
        self.corpo_email = tk.StringVar()
        self.enviando = False
        self.bdEmail = None
        self.anexos_por_pessoa = {}  # Dicionário para armazenar anexos por pessoa
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="SaleMachine - Enviador de Emails", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Seção 1: Arquivo CSV
        frame_csv = ttk.LabelFrame(main_frame, text="1. Base de Dados", padding="10")
        frame_csv.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        frame_csv.columnconfigure(1, weight=1)
        
        ttk.Label(frame_csv, text="Arquivo CSV:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(frame_csv, textvariable=self.csv_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(frame_csv, text="Procurar", command=self.procurar_csv).grid(row=0, column=2, padx=(10, 0))
        
        # Botões para gerenciar dados
        frame_botoes_csv = ttk.Frame(frame_csv)
        frame_botoes_csv.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(frame_botoes_csv, text="Visualizar Dados", command=self.visualizar_dados).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(frame_botoes_csv, text="Editar Dados", command=self.editar_dados).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(frame_botoes_csv, text="Salvar Alterações", command=self.salvar_csv).grid(row=0, column=2)
        
        # Seção 2: Configurações do Email
        frame_email = ttk.LabelFrame(main_frame, text="2. Configurações do Email", padding="10")
        frame_email.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        frame_email.columnconfigure(1, weight=1)
        
        ttk.Label(frame_email, text="Assunto:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(frame_email, textvariable=self.assunto, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Frame para gerenciamento de anexos
        frame_anexos = ttk.LabelFrame(frame_email, text="Gerenciamento de Anexos", padding="10")
        frame_anexos.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        frame_anexos.columnconfigure(1, weight=1)
        
        ttk.Button(frame_anexos, text="Gerenciar Anexos por Pessoa", 
                  command=self.gerenciar_anexos, width=25).grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # Status dos anexos
        self.status_anexos_var = tk.StringVar(value="Nenhum anexo configurado")
        ttk.Label(frame_anexos, textvariable=self.status_anexos_var, 
                 font=("Arial", 9), foreground="blue").grid(row=1, column=0, columnspan=3, sticky=tk.W)
        
        # Corpo do Email
        ttk.Label(frame_email, text="Corpo do Email (HTML):").grid(row=3, column=0, sticky=tk.NW, padx=(0, 10), pady=(10, 0))
        
        self.texto_corpo = scrolledtext.ScrolledText(frame_email, width=70, height=12)
        self.texto_corpo.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Texto padrão no corpo do email
        texto_padrao = """<p>Olá {primeiro_nome},</p>
<p>Escrever E-mail.</p>
<p>Escrever e-mail (segundo paragrafo).</p>
<p><i>Escrever e-mail (terceiro paragrafo).</i></p>
<b><i>Conclusão do e-mail.</i></b>
<p><b>Saudações e conclusões do assunto.</b></p>"""
        
        self.texto_corpo.insert("1.0", texto_padrao)
        
        # Variáveis de placeholder
        ttk.Label(frame_email, text="Variáveis disponíveis: {primeiro_nome}, {nome_completo}, {email}", 
                 font=("Arial", 8), foreground="blue").grid(row=4, column=1, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Seção 3: Controles
        frame_controles = ttk.LabelFrame(main_frame, text="3. Controles", padding="10")
        frame_controles.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.btn_enviar = ttk.Button(frame_controles, text="Iniciar Envio de Emails", 
                                   command=self.iniciar_envio, state="disabled")
        self.btn_enviar.grid(row=0, column=0, padx=(0, 10))
        
        self.btn_parar = ttk.Button(frame_controles, text="Parar Envio", 
                                  command=self.parar_envio, state="disabled")
        self.btn_parar.grid(row=0, column=1, padx=(0, 10))
        
        self.btn_limpar = ttk.Button(frame_controles, text="Limpar Log", command=self.limpar_log)
        self.btn_limpar.grid(row=0, column=2)
        
        # Seção 4: Log
        frame_log = ttk.LabelFrame(main_frame, text="4. Log de Execução", padding="10")
        frame_log.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        frame_log.columnconfigure(0, weight=1)
        frame_log.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(frame_log, width=80, height=15)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Progresso
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status
        self.status_var = tk.StringVar(value="Pronto para começar")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=6, column=0, columnspan=3)
        
        self.log("Aplicação iniciada. Selecione um arquivo CSV para começar.")
    
    def procurar_csv(self):
        filename = filedialog.askopenfilename(
            title="Selecione o arquivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_path.set(filename)
            self.carregar_csv()
    
    def carregar_csv(self):
        try:
            # Tentar diferentes encodings e separadores
            try:
                self.bdEmail = pd.read_csv(self.csv_path.get(), sep=";", encoding="ANSI", engine='python')
            except:
                try:
                    self.bdEmail = pd.read_csv(self.csv_path.get(), sep=",", encoding="ANSI", engine='python')
                except:
                    self.bdEmail = pd.read_csv(self.csv_path.get(), sep=";", encoding="utf-8", engine='python')
            
            # Verificar se há dados no CSV
            if self.bdEmail.empty:
                messagebox.showerror("Erro", "O arquivo CSV está vazio!")
                self.log("ERRO: CSV está vazio")
                self.btn_enviar.config(state="disabled")
                return
            
            # Verificar se as colunas necessárias existem
            colunas_necessarias = ['Nome', 'Email']
            colunas_faltantes = [col for col in colunas_necessarias if col not in self.bdEmail.columns]
            
            if colunas_faltantes:
                messagebox.showerror("Erro", f"Colunas faltantes no CSV: {', '.join(colunas_faltantes)}\n\nColunas encontradas: {', '.join(self.bdEmail.columns)}")
                self.log(f"ERRO: CSV não contém as colunas necessárias: {colunas_faltantes}")
                self.log(f"Colunas disponíveis: {list(self.bdEmail.columns)}")
                self.btn_enviar.config(state="disabled")
                return
            
            # Verificar se há emails válidos
            emails_validos = self.bdEmail['Email'].notna() & (self.bdEmail['Email'] != '')
            if not emails_validos.any():
                messagebox.showerror("Erro", "Não há emails válidos no arquivo CSV!")
                self.log("ERRO: Não há emails válidos no CSV")
                self.btn_enviar.config(state="disabled")
                return
            
            numero_linhas = len(self.bdEmail)
            self.log(f"CSV carregado com sucesso: {numero_linhas} registros")
            self.log(f"Colunas encontradas: {', '.join(self.bdEmail.columns)}")
            
            # Mostrar prévia dos primeiros registros
            self.log("Prévia dos dados:")
            for i in range(min(3, numero_linhas)):
                nome = self.bdEmail["Nome"].iloc[i]
                email = self.bdEmail["Email"].iloc[i]
                self.log(f"  {i+1}: {nome} - {email}")
            
            if numero_linhas > 3:
                self.log(f"  ... e mais {numero_linhas - 3} registros")
            
            # Habilitar botão de envio independentemente do número de registros
            self.btn_enviar.config(state="normal")
            self.log("Botão de envio habilitado - pronto para iniciar!")
            
            # Atualizar status dos anexos
            self.atualizar_status_anexos()
            
        except Exception as e:
            error_msg = f"Erro ao carregar CSV: {str(e)}"
            messagebox.showerror("Erro", error_msg)
            self.log(f"ERRO: {error_msg}")
            self.log(f"Detalhes: {traceback.format_exc()}")
            self.btn_enviar.config(state="disabled")
    
    def editar_dados(self):
        """Abre janela para editar os dados do CSV"""
        if self.bdEmail is None:
            messagebox.showwarning("Aviso", "Por favor, carregue um arquivo CSV primeiro.")
            return
        
        # Criar janela de edição
        janela_edicao = tk.Toplevel(self.root)
        janela_edicao.title("Editar Dados do CSV")
        janela_edicao.geometry("800x500")
        
        # Frame principal
        main_frame = ttk.Frame(janela_edicao, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview para edição
        frame_tree = ttk.Frame(main_frame)
        frame_tree.pack(fill=tk.BOTH, expand=True)
        
        # Criar Treeview com colunas editáveis
        tree = ttk.Treeview(frame_tree, columns=list(self.bdEmail.columns), show="headings")
        
        # Configurar cabeçalhos
        for col in self.bdEmail.columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor=tk.W)
        
        # Adicionar dados
        for i, row in self.bdEmail.iterrows():
            tree.insert("", tk.END, values=list(row), iid=str(i))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame para botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill=tk.X, pady=10)
        
        def adicionar_linha():
            # Adicionar nova linha vazia
            nova_linha = [""] * len(self.bdEmail.columns)
            tree.insert("", tk.END, values=nova_linha, iid=str(len(self.bdEmail)))
            
            # Atualizar DataFrame
            novo_df = pd.DataFrame([nova_linha], columns=self.bdEmail.columns)
            self.bdEmail = pd.concat([self.bdEmail, novo_df], ignore_index=True)
            
            self.log("Nova linha adicionada para edição")
        
        def remover_linha():
            selecionado = tree.selection()
            if not selecionado:
                messagebox.showwarning("Aviso", "Selecione uma linha para remover.")
                return
            
            # Confirmar remoção
            if messagebox.askyesno("Confirmar", "Deseja realmente remover a linha selecionada?"):
                # Remover do Treeview
                for item in selecionado:
                    tree.delete(item)
                
                # Atualizar DataFrame
                indices_remover = [int(item) for item in selecionado]
                self.bdEmail = self.bdEmail.drop(indices_remover).reset_index(drop=True)
                
                # Reconstruir Treeview com novos IDs
                tree.delete(*tree.get_children())
                for i, row in self.bdEmail.iterrows():
                    tree.insert("", tk.END, values=list(row), iid=str(i))
                
                self.log(f"{len(selecionado)} linha(s) removida(s)")
        
        def salvar_alteracoes():
            # Coletar dados do Treeview
            novos_dados = []
            for item in tree.get_children():
                valores = tree.item(item)['values']
                novos_dados.append(valores)
            
            # Atualizar DataFrame
            self.bdEmail = pd.DataFrame(novos_dados, columns=self.bdEmail.columns)
            
            self.log("Alterações salvas na memória (use 'Salvar Alterações' para gravar no arquivo)")
            messagebox.showinfo("Sucesso", "Alterações salvas na memória!\n\nUse o botão 'Salvar Alterações' na tela principal para gravar no arquivo CSV.")
        
        # Função para edição em linha
        def editar_celula(event):
            item = tree.identify_row(event.y)
            coluna = tree.identify_column(event.x)
            
            if not item or not coluna:
                return
            
            # Converter coluna para índice
            col_idx = int(coluna[1:]) - 1
            
            # Obter valor atual
            valor_atual = tree.item(item, 'values')[col_idx]
            
            # Criar janela de edição
            popup = tk.Toplevel(janela_edicao)
            popup.title(f"Editar {self.bdEmail.columns[col_idx]}")
            popup.geometry("300x100")
            popup.transient(janela_edicao)
            popup.grab_set()
            
            ttk.Label(popup, text=f"Editar {self.bdEmail.columns[col_idx]}:").pack(pady=5)
            
            entry_var = tk.StringVar(value=valor_atual)
            entry = ttk.Entry(popup, textvariable=entry_var, width=30)
            entry.pack(pady=5)
            entry.focus()
            entry.select_range(0, tk.END)
            
            def confirmar_edicao():
                novo_valor = entry_var.get()
                
                # Atualizar Treeview
                valores = list(tree.item(item, 'values'))
                valores[col_idx] = novo_valor
                tree.item(item, values=valores)
                
                popup.destroy()
            
            def cancelar_edicao():
                popup.destroy()
            
            frame_botoes_popup = ttk.Frame(popup)
            frame_botoes_popup.pack(pady=5)
            
            ttk.Button(frame_botoes_popup, text="OK", command=confirmar_edicao).pack(side=tk.LEFT, padx=5)
            ttk.Button(frame_botoes_popup, text="Cancelar", command=cancelar_edicao).pack(side=tk.LEFT, padx=5)
            
            popup.bind('<Return>', lambda e: confirmar_edicao())
            popup.bind('<Escape>', lambda e: cancelar_edicao())
        
        # Botões
        ttk.Button(frame_botoes, text="Adicionar Linha", command=adicionar_linha).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Remover Linha", command=remover_linha).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Salvar na Memória", command=salvar_alteracoes).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Fechar", command=janela_edicao.destroy).pack(side=tk.LEFT, padx=5)
        
        # Vincular evento de duplo clique para edição
        tree.bind("<Double-1>", editar_celula)
    
    def salvar_csv(self):
        """Salva as alterações no arquivo CSV original"""
        if self.bdEmail is None:
            messagebox.showwarning("Aviso", "Não há dados para salvar.")
            return
        
        if not self.csv_path.get():
            messagebox.showwarning("Aviso", "Nenhum arquivo CSV carregado.")
            return
        
        try:
            # Tentar salvar com as mesmas configurações do arquivo original
            self.bdEmail.to_csv(self.csv_path.get(), sep=";", encoding="ANSI", index=False)
            self.log(f"Alterações salvas no arquivo: {self.csv_path.get()}")
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso no arquivo CSV!")
        except Exception as e:
            try:
                # Tentar com encoding UTF-8 se ANSI falhar
                self.bdEmail.to_csv(self.csv_path.get(), sep=";", encoding="utf-8", index=False)
                self.log(f"Alterações salvas com UTF-8: {self.csv_path.get()}")
                messagebox.showinfo("Sucesso", "Alterações salvas com sucesso no arquivo CSV!")
            except Exception as e2:
                error_msg = f"Erro ao salvar CSV: {str(e2)}"
                messagebox.showerror("Erro", error_msg)
                self.log(f"ERRO: {error_msg}")
    
    def gerenciar_anexos(self):
        """Janela para gerenciar a associação de anexos com pessoas"""
        if self.bdEmail is None:
            messagebox.showwarning("Aviso", "Por favor, carregue um arquivo CSV primeiro.")
            return
        
        # Criar janela de gerenciamento
        janela_anexos = tk.Toplevel(self.root)
        janela_anexos.title("Gerenciar Anexos por Pessoa")
        janela_anexos.geometry("800x600")
        
        # Frame principal
        main_frame = ttk.Frame(janela_anexos, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Lista de pessoas
        ttk.Label(main_frame, text="Pessoas:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        frame_pessoas = ttk.Frame(main_frame)
        frame_pessoas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        lista_pessoas = tk.Listbox(frame_pessoas, width=30, height=20)
        scroll_pessoas = ttk.Scrollbar(frame_pessoas, orient=tk.VERTICAL, command=lista_pessoas.yview)
        lista_pessoas.configure(yscrollcommand=scroll_pessoas.set)
        
        # Preencher lista de pessoas
        for i, nome in enumerate(self.bdEmail["Nome"]):
            lista_pessoas.insert(tk.END, f"{i+1}. {nome}")
        
        lista_pessoas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_pessoas.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame para anexos
        frame_controles_anexos = ttk.Frame(main_frame)
        frame_controles_anexos.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(frame_controles_anexos, text="Anexos da Pessoa Selecionada:", 
                 font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # Lista de anexos da pessoa selecionada
        frame_anexos_pessoa = ttk.Frame(frame_controles_anexos)
        frame_anexos_pessoa.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        lista_anexos_pessoa = tk.Listbox(frame_anexos_pessoa, width=40, height=8)
        scroll_anexos_pessoa = ttk.Scrollbar(frame_anexos_pessoa, orient=tk.VERTICAL, command=lista_anexos_pessoa.yview)
        lista_anexos_pessoa.configure(yscrollcommand=scroll_anexos_pessoa.set)
        
        lista_anexos_pessoa.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_anexos_pessoa.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões para gerenciar anexos
        frame_botoes_anexos = ttk.Frame(frame_controles_anexos)
        frame_botoes_anexos.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        def adicionar_anexo():
            pessoa_idx = lista_pessoas.curselection()
            if not pessoa_idx:
                messagebox.showwarning("Aviso", "Selecione uma pessoa primeiro.")
                return
            
            # Procurar arquivo PDF ou DOCX
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo PDF ou DOCX",
                filetypes=[("PDF files", "*.pdf"), ("DOCX files", "*.docx"), ("All files", "*.*")]
            )
            
            if arquivo:
                pessoa_nome = self.bdEmail["Nome"].iloc[pessoa_idx[0]]
                nome_arquivo = os.path.basename(arquivo)
                
                if pessoa_nome not in self.anexos_por_pessoa:
                    self.anexos_por_pessoa[pessoa_nome] = []
                
                # Adicionar caminho completo do arquivo
                if arquivo not in self.anexos_por_pessoa[pessoa_nome]:
                    self.anexos_por_pessoa[pessoa_nome].append(arquivo)
                    lista_anexos_pessoa.insert(tk.END, nome_arquivo)
                    self.log(f"Anexo '{nome_arquivo}' associado a '{pessoa_nome}'")
                    self.atualizar_status_anexos()
        
        def remover_anexo():
            pessoa_idx = lista_pessoas.curselection()
            anexo_idx = lista_anexos_pessoa.curselection()
            
            if not pessoa_idx or not anexo_idx:
                return
            
            pessoa_nome = self.bdEmail["Nome"].iloc[pessoa_idx[0]]
            anexo_caminho = self.anexos_por_pessoa[pessoa_nome][anexo_idx[0]]
            nome_arquivo = os.path.basename(anexo_caminho)
            
            if pessoa_nome in self.anexos_por_pessoa:
                self.anexos_por_pessoa[pessoa_nome].pop(anexo_idx[0])
                lista_anexos_pessoa.delete(anexo_idx[0])
                self.log(f"Anexo '{nome_arquivo}' removido de '{pessoa_nome}'")
                self.atualizar_status_anexos()
        
        def adicionar_anexo_multiplas_pessoas():
            # Procurar arquivo PDF ou DOCX
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo PDF ou DOCX",
                filetypes=[("PDF files", "*.pdf"), ("DOCX files", "*.docx"), ("All files", "*.*")]
            )
            
            if not arquivo:
                return
            
            # Selecionar pessoas para associar o anexo
            janela_selecao = tk.Toplevel(janela_anexos)
            janela_selecao.title("Selecionar Pessoas para o Anexo")
            janela_selecao.geometry("400x300")
            
            frame_selecao = ttk.Frame(janela_selecao, padding="10")
            frame_selecao.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(frame_selecao, text=f"Selecione as pessoas para receber: {os.path.basename(arquivo)}", 
                     font=("Arial", 10, "bold")).pack(pady=(0, 10))
            
            # Lista de seleção múltipla
            frame_lista = ttk.Frame(frame_selecao)
            frame_lista.pack(fill=tk.BOTH, expand=True)
            
            lista_selecao = tk.Listbox(frame_lista, selectmode=tk.MULTIPLE, height=10)
            scroll_selecao = ttk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=lista_selecao.yview)
            lista_selecao.configure(yscrollcommand=scroll_selecao.set)
            
            # Preencher lista de pessoas
            for i, nome in enumerate(self.bdEmail["Nome"]):
                lista_selecao.insert(tk.END, f"{i+1}. {nome}")
            
            lista_selecao.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll_selecao.pack(side=tk.RIGHT, fill=tk.Y)
            
            def confirmar_selecao():
                selecoes = lista_selecao.curselection()
                if not selecoes:
                    messagebox.showwarning("Aviso", "Selecione pelo menos uma pessoa.")
                    return
                
                for idx in selecoes:
                    pessoa_nome = self.bdEmail["Nome"].iloc[idx]
                    if pessoa_nome not in self.anexos_por_pessoa:
                        self.anexos_por_pessoa[pessoa_nome] = []
                    
                    if arquivo not in self.anexos_por_pessoa[pessoa_nome]:
                        self.anexos_por_pessoa[pessoa_nome].append(arquivo)
                        self.log(f"Anexo '{os.path.basename(arquivo)}' associado a '{pessoa_nome}'")
                
                self.atualizar_status_anexos()
                janela_selecao.destroy()
                # Atualizar lista se a pessoa selecionada estiver na lista
                pessoa_atual_idx = lista_pessoas.curselection()
                if pessoa_atual_idx:
                    pessoa_selecionada(event=None)
            
            frame_botoes_selecao = ttk.Frame(frame_selecao)
            frame_botoes_selecao.pack(fill=tk.X, pady=10)
            
            ttk.Button(frame_botoes_selecao, text="Confirmar", command=confirmar_selecao).pack(side=tk.LEFT, padx=5)
            ttk.Button(frame_botoes_selecao, text="Cancelar", command=janela_selecao.destroy).pack(side=tk.LEFT, padx=5)
        
        def pessoa_selecionada(event):
            # Limpar lista de anexos
            lista_anexos_pessoa.delete(0, tk.END)
            
            pessoa_idx = lista_pessoas.curselection()
            if not pessoa_idx:
                return
            
            pessoa_nome = self.bdEmail["Nome"].iloc[pessoa_idx[0]]
            
            # Mostrar anexos já associados
            if pessoa_nome in self.anexos_por_pessoa:
                for anexo_caminho in self.anexos_por_pessoa[pessoa_nome]:
                    nome_arquivo = os.path.basename(anexo_caminho)
                    lista_anexos_pessoa.insert(tk.END, nome_arquivo)
        
        ttk.Button(frame_botoes_anexos, text="Adicionar Anexo", 
                  command=adicionar_anexo, width=20).pack(side=tk.LEFT, padx=2)
        ttk.Button(frame_botoes_anexos, text="Remover Anexo", 
                  command=remover_anexo, width=20).pack(side=tk.LEFT, padx=2)
        ttk.Button(frame_botoes_anexos, text="Anexo para Múltiplas Pessoas", 
                  command=adicionar_anexo_multiplas_pessoas, width=25).pack(side=tk.LEFT, padx=2)
        
        # Resumo
        ttk.Label(frame_controles_anexos, text="Resumo:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky=tk.W, pady=(10, 5))
        
        self.resumo_anexos_var = tk.StringVar(value="Nenhum anexo configurado")
        ttk.Label(frame_controles_anexos, textvariable=self.resumo_anexos_var, 
                 font=("Arial", 9), foreground="green").grid(row=4, column=0, sticky=tk.W)
        
        # Botão fechar
        ttk.Button(frame_controles_anexos, text="Fechar", 
                  command=janela_anexos.destroy, width=20).grid(row=5, column=0, pady=(10, 0))
        
        # Vincular evento de seleção
        lista_pessoas.bind('<<ListboxSelect>>', pessoa_selecionada)
        
        # Atualizar resumo
        self.atualizar_resumo_anexos()
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
    
    def atualizar_status_anexos(self):
        """Atualiza o status dos anexos na interface principal"""
        total_pessoas_com_anexos = len(self.anexos_por_pessoa)
        total_anexos = sum(len(anexos) for anexos in self.anexos_por_pessoa.values())
        
        if total_anexos == 0:
            self.status_anexos_var.set("Nenhum anexo configurado")
        else:
            self.status_anexos_var.set(f"{total_pessoas_com_anexos} pessoa(s) com anexos | Total: {total_anexos} arquivo(s)")
    
    def atualizar_resumo_anexos(self):
        """Atualiza o resumo na janela de anexos"""
        total_pessoas_com_anexos = len(self.anexos_por_pessoa)
        total_anexos = sum(len(anexos) for anexos in self.anexos_por_pessoa.values())
        
        if total_anexos == 0:
            self.resumo_anexos_var.set("Nenhum anexo configurado")
        else:
            self.resumo_anexos_var.set(f"Total: {total_pessoas_com_anexos} pessoa(s) com anexos | {total_anexos} arquivo(s)")
    
    def visualizar_dados(self):
        if self.bdEmail is None:
            messagebox.showwarning("Aviso", "Por favor, carregue um arquivo CSV primeiro.")
            return
        
        try:
            numero_linhas = len(self.bdEmail)
            
            # Criar janela de visualização
            janela_dados = tk.Toplevel(self.root)
            janela_dados.title(f"Visualizar Dados - {numero_linhas} registros")
            janela_dados.geometry("600x400")
            
            # Treeview para mostrar dados
            frame_tree = ttk.Frame(janela_dados)
            frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            tree = ttk.Treeview(frame_tree)
            
            # Definir colunas
            tree["columns"] = list(self.bdEmail.columns)
            tree.column("#0", width=0, stretch=tk.NO)
            
            for col in self.bdEmail.columns:
                tree.column(col, anchor=tk.W, width=100)
                tree.heading(col, text=col, anchor=tk.W)
            
            # Adicionar dados
            for i, row in self.bdEmail.iterrows():
                tree.insert("", tk.END, values=list(row))
            
            # Scrollbar
            scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao visualizar dados:\n{str(e)}")
    
    def iniciar_envio(self):
        if self.bdEmail is None:
            messagebox.showwarning("Aviso", "Por favor, carregue um arquivo CSV primeiro.")
            return
        
        # Verificação adicional para garantir que há dados
        if len(self.bdEmail) == 0:
            messagebox.showwarning("Aviso", "Não há dados para enviar no arquivo CSV.")
            return
        
        self.enviando = True
        self.btn_enviar.config(state="disabled")
        self.btn_parar.config(state="normal")
        
        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=self.enviar_emails)
        thread.daemon = True
        thread.start()
    
    def parar_envio(self):
        self.enviando = False
        self.btn_enviar.config(state="normal")
        self.btn_parar.config(state="disabled")
        self.log("Envio interrompido pelo usuário")
        self.status_var.set("Envio interrompido")
    
    def limpar_log(self):
        self.log_text.delete("1.0", tk.END)
    
    def log(self, mensagem):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {mensagem}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def extrair_primeiro_nome(self, nome_completo):
        """Extrai o primeiro nome de um nome completo"""
        try:
            if pd.isna(nome_completo) or nome_completo == '':
                return "Prezado(a)"
            if ' ' in str(nome_completo):
                return str(nome_completo).split(' ')[0]
            else:
                return nome_completo
        except:
            return "Prezado(a)"
    
    def enviar_emails(self):
        try:
            numero_linhas = len(self.bdEmail)
            
            self.log(f"Iniciando envio de {numero_linhas} emails...")
            self.log("⚠️  Delay de 6 segundos entre cada email para evitar spam")
            self.status_var.set(f"Enviando 0 de {numero_linhas} emails")
            
            # Configurar barra de progresso
            self.progress["maximum"] = numero_linhas
            self.progress["value"] = 0
            
            emails_enviados = 0
            emails_erro = 0
            
            for i in range(numero_linhas):
                if not self.enviando:
                    break
                
                try:
                    # Variáveis para os parâmetros do email
                    nome_completo = str(self.bdEmail["Nome"].iloc[i]).strip()
                    email_destinatario = str(self.bdEmail["Email"].iloc[i]).strip()
                    
                    # Verificar se os dados são válidos
                    if not nome_completo or not email_destinatario or pd.isna(nome_completo) or pd.isna(email_destinatario):
                        self.log(f"AVISO: Dados inválidos na linha {i+1} - Nome: '{nome_completo}', Email: '{email_destinatario}'")
                        emails_erro += 1
                        continue
                    
                    # Extrair primeiro nome
                    primeiro_nome = self.extrair_primeiro_nome(nome_completo)
                    
                    self.log(f"Processando email {i+1}/{numero_linhas}: {nome_completo} -> {email_destinatario}")
                    
                    # Criando integração com o Outlook
                    outlook = win32.Dispatch('outlook.application')
                    email = outlook.CreateItem(0)
                    
                    # Configurar informações do email
                    email.To = email_destinatario
                    email.Subject = self.assunto.get()
                    
                    # Adicionar anexos específicos para esta pessoa
                    anexos_adicionados = 0
                    if nome_completo in self.anexos_por_pessoa:
                        for anexo_caminho in self.anexos_por_pessoa[nome_completo]:
                            if os.path.exists(anexo_caminho):
                                email.Attachments.Add(anexo_caminho)
                                anexos_adicionados += 1
                                nome_arquivo = os.path.basename(anexo_caminho)
                                self.log(f"  📎 Anexo adicionado: {nome_arquivo}")
                            else:
                                nome_arquivo = os.path.basename(anexo_caminho)
                                self.log(f"  ⚠️ Aviso: Anexo não encontrado - {nome_arquivo}")
                    
                    if anexos_adicionados == 0:
                        self.log(f"  ℹ️ Nenhum anexo configurado para {nome_completo}")
                    
                    # Corpo do email com substituição de variáveis
                    corpo_texto = self.texto_corpo.get("1.0", tk.END)
                    corpo_formatado = corpo_texto.replace("{primeiro_nome}", primeiro_nome)
                    corpo_formatado = corpo_formatado.replace("{nome_completo}", nome_completo)
                    corpo_formatado = corpo_formatado.replace("{email}", email_destinatario)
                    
                    email.HTMLBody = corpo_formatado
                    email.Send()
                    
                    emails_enviados += 1
                    self.log(f"  ✓ Email enviado com sucesso! ({anexos_adicionados} anexos)")
                    
                    # Aplicar delay de 6 segundos APÓS o envio (exceto para o último email)
                    if i < numero_linhas - 1 and self.enviando:
                        self.log("  ⏳ Aguardando 6 segundos antes do próximo envio...")
                        for segundos_restantes in range(6, 0, -1):
                            if not self.enviando:
                                break
                            self.status_var.set(f"Aguardando {segundos_restantes}s... ({i + 1} de {numero_linhas} enviados)")
                            time.sleep(1)
                        
                except Exception as e:
                    emails_erro += 1
                    error_details = traceback.format_exc()
                    self.log(f"  ✗ ERRO no email {i+1}: {str(e)}")
                    self.log(f"  Detalhes: {error_details.splitlines()[-1]}")  # Mostra apenas a última linha do erro
                    
                    # Aplicar delay mesmo em caso de erro (exceto para o último email)
                    if i < numero_linhas - 1 and self.enviando:
                        self.log("  ⏳ Aguardando 6 segundos antes do próximo envio...")
                        for segundos_restantes in range(6, 0, -1):
                            if not self.enviando:
                                break
                            self.status_var.set(f"Aguardando {segundos_restantes}s... ({i + 1} de {numero_linhas} processados)")
                            time.sleep(1)
                
                # Atualizar progresso
                self.progress["value"] = i + 1
                if self.enviando:
                    self.status_var.set(f"Enviando {i + 1} de {numero_linhas} emails")
            
            # Resultado final
            if self.enviando:
                self.status_var.set(f"Concluído! {emails_enviados} enviados, {emails_erro} erros")
                self.log(f"Processo finalizado: {emails_enviados} emails enviados, {emails_erro} erros")
                messagebox.showinfo("Concluído", f"Envio finalizado!\n{emails_enviados} emails enviados\n{emails_erro} erros")
            else:
                self.status_var.set(f"Interrompido: {emails_enviados} enviados, {emails_erro} erros")
            
            self.btn_enviar.config(state="normal")
            self.btn_parar.config(state="disabled")
            self.enviando = False
            
        except Exception as e:
            error_details = traceback.format_exc()
            self.log(f"ERRO CRÍTICO: {str(e)}")
            self.log(f"Detalhes: {error_details}")
            messagebox.showerror("Erro", f"Erro durante o envio:\n{str(e)}")
            self.btn_enviar.config(state="normal")
            self.btn_parar.config(state="disabled")
            self.enviando = False

def main():
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()