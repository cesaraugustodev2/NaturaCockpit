
#  Com a licença GNU General Public License (GPL) versão 3, o texto de copyright para o projeto **Natura Backlog** pode ser redigido da seguinte maneira:
#
#  ---
#
#  **Copyright © [2024] César Augusto. Todos os direitos reservados.**
#
#  Este projeto é licenciado sob a Licença Pública Geral GNU, versão 3 (GPL-3.0). Você é livre para usar, copiar, modificar e distribuir este software, desde que siga os termos da licença.
#
#
#  - **Liberdade de Uso**: Você pode usar o software para qualquer propósito.
#  - **Liberdade de Modificação**: Você pode modificar o software e distribuir suas modificações sob os mesmos termos.
#  - **Liberdade de Distribuição**: Você pode distribuir cópias do software original ou modificado, garantindo que todos os destinatários tenham as mesmas liberdades que você recebeu.
#
#  Este software é fornecido "como está", sem garantia de qualquer tipo, expressa ou implícita, incluindo, mas não se limitando a garantias de comercialização ou adequação a um propósito específico.
#
#  Para mais informações, consulte o arquivo LICENSE incluído neste repositório.
#
#  ---

import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import re
import datetime
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from difflib import SequenceMatcher
import tkinter.font as tkFont



''' Classe principal que irá conter todo o código e widgets'''
class BacklogViewer:
    

    def __init__(self, root):
        self.root = root #root é o objeto pai 
        self.root.title("Natura Cockpit 1.1")
        self.root.state('zoomed') #full screen
        #icones e imagem da Natura
        self.root.iconbitmap("img/icon.ico")
        icon = tk.PhotoImage(file="img/icon.png")
        logo_natura = tk.PhotoImage(file="img/natura.png")
        logo_natura_small = logo_natura.subsample(18)
        #Font padronizada para melhor legibilidade
        self.roman_font = tkFont.Font(family="Arial", size=10, slant=tkFont.ROMAN)
        #Tena Personalizado Natura Dark
        s=ttk.Style()
        s.theme_create('dark_orange', parent='clam', settings={
        'TButton': {
            'configure': {
                'foreground': 'white',
                'background': '#ff6b00',
                'font': ('Arial', 10),
                'padding': 5
            },
            'map': {
                'background': [('active', '#ff8c00'), ('pressed', '#cc5500')]
            }
        },
        'TFrame': {
            'configure': {
                'foreground': 'white',
                'background': '#000'
            }
        },
         'Vertical.TScrollbar': {
            'configure': {
                'foreground': 'white',
                'background': '#000',
                'troughcolor': '#000',
                'thumbcolor': '#000',
            }
        },

         'Treeview': {
            'configure': {
                'foreground': 'white',
                'background': '#000',
                'headingbackground':'#000',
                'headingforeground':'white',
                'font': ('Arial', 11),
            }
        },
        'TEntry': {
            'configure': {
                'foreground': 'white',
                'background': '##000',
                'fieldbackground': '#000',
                'font': ('Arial', 10),
                'padding': 5
            }
        },
        'TLabel': {
            'configure': {
                'foreground': 'white',
                'background': '#000',
                'font': ('Arial', 10)
            }
        },
        'TCombobox': {
            'configure': {
                'foreground': 'white',
                'background': '#444444',
                'fieldbackground': '#444444',
            '   arrowcolor': 'white',
                'font': ('Arial', 10),
            '   padding': 5
            }
        },
        'TMenubutton': {
            'configure': {
                'foreground': 'white',
                'background': '#ff6b00',
                'font': ('Arial', 10),
                'padding': 5
            }
        },
        
    })
        
 
        #Inicio do Menu superior
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        file_menu = tk.Menu(self.menu_bar,font=self.roman_font)
        #menu arquivo
        self.menu_bar.add_cascade(label="Backlog", menu=file_menu)
        file_menu.add_command(label="Importar Backlog", command=self.upload_backlog)
        file_menu.add_command(label="Exportar Filtrados", command=self.export_to_excel)
        file_menu.add_command(label="Exportar Selecionados", command=self.export_to_excel_selected)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=root.quit)
        #menu editar
        edit_menu = tk.Menu(self.menu_bar)
        self.menu_bar.add_cascade(label="Editar", menu=edit_menu,font=self.roman_font)
        edit_menu.add_command(label='Copiar ChamadoID', command=self.copy_chamado)
        edit_menu.add_command(label='Copiar ProblemID', command=self.copy_problema)
        filter_menu = tk.Menu(self.menu_bar)
        self.menu_bar.add_cascade(label="Filtro Avançado", menu=filter_menu,font=self.roman_font)
        #menu filtros avançados
        filter_menu.add_command(label="HyperCare Natura MX", command=self.hypercare_MX)
        filter_menu.add_command(label="HyperCare ELO Fase", command=self.hypercare_BR)
        filter_menu.add_command(label="Escalation (Aging Alto)", command=self.escalation_aging)
        filter_menu.add_command(label="Só Resolvidos", command=self.filter_fechado)
        filter_menu.add_command(label="Só Abertos", command=self.filter_abertos)
        filter_menu.add_separator()
        filter_menu.add_command(label="Sem Problem", command=self.filter_problemid)
        filter_menu.add_command(label="Com Problem", command=self.keep_problemid)
        range_menu = tk.Menu(self.menu_bar)
        self.menu_bar.add_cascade(label="Analise Temporal", menu=range_menu,font=self.roman_font)
        #menu analise temporal
        range_menu.add_command(label="Abertos D-1", command=self.aberto_dayminusone)
        range_menu.add_command(label="Abertos D-7", command=self.aberto_dayminusseven)
        range_menu.add_command(label="Abertos D-30", command=self.aberto_dayminusthirty)
        range_menu.add_command(label="Abertos D-60", command=self.aberto_dayminussixty)
        range_menu.add_command(label="Abertos D-N", command=self.aberto_dayminuscustom)
        range_menu.add_separator()
        range_menu.add_command(label="Resolvidos D-1", command=self.fechado_dayminusone)
        range_menu.add_command(label="Resolvidos D-7", command=self.fechado_dayminusseven)
        range_menu.add_command(label="Resolvidos D-30", command=self.fechado_dayminusthirty)
        range_menu.add_command(label="Resolvidos D-60", command=self.fechado_dayminussixty)
        range_menu.add_command(label="Resolvidos D-N", command=self.fechado_dayminuscustom)
        #menu ajuda
        self.help_menu = tk.Menu(self.menu_bar)
        self.menu_bar.add_cascade(label="Ajuda", menu=self.help_menu)
        self.help_menu.add_command(label="Achei um bug", command=self.help_popup)
        self.help_menu.add_command(label="Manual", command=self.manual_popup)
        #menu aparencia
        self.theme_menu = tk.Menu(self.menu_bar)
        self.menu_bar.add_cascade(label="Aparencia", menu=self.theme_menu)
        self.theme_menu.add_command(label="Natura Classico", command=self.vista_mode)
        self.theme_menu.add_command(label="Natura Dark", command=self.dark_mode)
        self.theme_menu.add_command(label="Natura Retro", command=self.clam_mode)
        
        #Frame principal que irá conter a arvore de chamados
        self.list_frame = ttk.Frame(self.root, style='info.TFrame')
        self.tree_frame = ttk.Frame(self.list_frame)

        self.tree = ttk.Treeview(self.tree_frame, columns=("DT_ABERTURA", "DT_SOLUÇÃO", "CHAMADO", "PROBLEMA", "GRUPO", "STATUS", "TIPO", "RESUMO", "AGING_IN_DAYS", "LOCALIDADE", "SLA_VIOLADO", "DESCRICAO"), show="headings", selectmode="extended")

        self.tree_scrollbar_v = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)

        self.tree.pack(side="left", fill="both")
        self.tree.tag_configure('selected', background='#ff6b00', foreground='white')

        #Scrollbar verifical
        self.tree_scrollbar_v.pack(side="left", fill="y")

        self.tree.configure(yscrollcommand=self.tree_scrollbar_v.set)
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=100)
        self.list_frame.pack(fill="both", expand=True)        
        #frame do upload primeira linha superior botão importar backlog, pesquisa e os filtros de data 
        self.upload_frame = ttk.Frame(self.list_frame)
        self.upload_frame.pack(fill="x")

        self.logo_natura_label = tk.Label(self.upload_frame, image=logo_natura_small)
        self.logo_natura_label.image = logo_natura_small
        self.logo_natura_label.pack(side="left")
    

        self.upload_button = ttk.Button(self.upload_frame, text="Importar Backlog", command=self.upload_backlog,style='TButton')
        

        self.upload_button.pack(side="left", padx=10, pady=5)

        self.search_label=tk.Label(self.upload_frame, text="Pesquisar")
        self.search_label.pack(side="left", padx=5, pady=5)
        self.search_entry1 = ttk.Entry(self.upload_frame)
        self.search_entry1.pack(side="left", padx=5, pady=5, fill="both", expand=False)

        self.operator_list = ["e", "ou"]
        
        self.operator = ttk.Combobox(self.upload_frame, values=self.operator_list)
        self.operator.pack(side="left", padx=5, pady=5)
        
        self.search_entry2 = ttk.Entry(self.upload_frame)
        self.search_entry2.pack(side="left", padx=5, pady=5, fill="both", expand=False)

        self.search_button = ttk.Button(self.upload_frame, text="Procurar", command=self.search)
        self.search_button.pack(side="left", padx=5, pady=5)
        
        self.dt_abertura_label = ttk.Label(self.upload_frame, text="Aberto em:",)
        self.dt_abertura_label.pack(side="left", padx=5, pady=2)
        
        self.dt_abertura_filter = ttk.Entry(self.upload_frame, event=None)
        self.dt_abertura_filter.pack(side="left", fill="x", padx=5, pady=2)
        
        self.dt_fechamento_label = ttk.Label(self.upload_frame, text="Fechado em:")
        self.dt_fechamento_label.pack(side="left", padx=5, pady=2)
        
        self.dt_fechamento_filter = ttk.Entry(self.upload_frame, event=None)
        self.dt_fechamento_filter.pack(side="left", fill="x", padx=5, pady=2)
        
        self.problema_label = ttk.Label(self.upload_frame, text="Problema:")
        self.problema_label.pack(side="left", padx=5, pady=2)

        self.problema_filter = ttk.Entry(self.upload_frame)
        self.problema_filter.pack(side="left", fill="x", padx=5, pady=2)

        self.search_filter_frame = ttk.Frame(self.list_frame)
        self.search_filter_frame.pack(fill="x")

        self.tree_frame.pack(fill="both", expand=True, side="left")
        #cabeçalho da lista de chamados aqui é aplicado o comando de ordenação clicando em cada titulo da coluna
        self.tree.heading("DT_ABERTURA", text="Aberto em", command=lambda: self.sort_column("DT_ABERTURA", False))
        self.tree.heading("DT_SOLUÇÃO", text="Fechado em", command=lambda: self.sort_column("DT_SOLUÇÃO", False))
        self.tree.heading("CHAMADO", text="Chamado", command=lambda: self.sort_column("CHAMADO", False))
        self.tree.heading("PROBLEMA", text="Problema", command=lambda: self.sort_column("PROBLEMA", False))
        self.tree.heading("GRUPO", text="Grupo", command=lambda: self.sort_column("GRUPO", False))
        self.tree.heading("STATUS", text="Status", command=lambda: self.sort_column("STATUS", False))
        self.tree.heading("TIPO", text="Tipo", command=lambda: self.sort_column("TIPO", False))
        self.tree.heading("RESUMO", text="Resumo", command=lambda: self.sort_column("RESUMO", False))
        self.tree.heading("AGING_IN_DAYS", text="Aging", command=lambda: self.sort_column("AGING_IN_DAYS", False))
        self.tree.heading("LOCALIDADE", text="Pais", command=lambda: self.sort_column("LOCALIDADE", False))
        self.tree.heading("SLA_VIOLADO", text="SLA", command=lambda: self.sort_column("SLA_VIOLADO", False))
        self.tree.heading("DESCRICAO", text="Descrição", command=lambda: self.sort_column("DESCRICAO", False))
        #largura padrão de cada coluna
        self.tree.column("DT_ABERTURA", width=100)
        self.tree.column("DT_SOLUÇÃO", width=100)
        self.tree.column("CHAMADO", width=50)
        self.tree.column("PROBLEMA", width=50)
        self.tree.column("GRUPO", width=100)
        self.tree.column("STATUS", width=80)
        self.tree.column("AGING_IN_DAYS", width=40)
        self.tree.column("TIPO", width=100)
        self.tree.column("RESUMO", width=100)
        self.tree.column("LOCALIDADE", width=100)
        self.tree.column("SLA_VIOLADO", width=100)
        self.tree.column("DESCRICAO", width=0,stretch=tk.NO) #remove a coluna descrição do frontend
        
        #funçao que abre o menu auxiliar
        def popup_menu(self, event=None):
            try:
             m.tk_popup(self.x_root, self.y_root)
            finally:
             m.grab_release()
             
        
        #Todas as teclas de atalho e eventos bind
        self.root.bind("<Control-o>", self.upload_backlog)
        self.root.bind("<Alt-c>", self.copy_chamado)
        self.root.bind("<Alt-p>", self.copy_problema)
        self.root.bind("<Control-s>", self.export_to_excel)
        self.root.bind("<Control-e>", self.export_to_excel_selected)
        self.root.bind("<<TreeviewSelect>>", self.display_description)
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.root.bind("<Button-3>", popup_menu)
        self.dt_abertura_filter.bind("<KeyRelease>", self.format_date_abertura)
        self.dt_fechamento_filter.bind("<KeyRelease>", self.format_date_fechamento)
        self.root.bind("<Return>", self.filter_data)
        self.root.bind("<Alt-i>", self.export_stats)
        self.root.bind("<Alt-a>", self.filter_abertos)
        self.root.bind("<Alt-f>", self.filter_fechado)
        self.root.bind("<Alt-1>", self.aberto_dayminusseven)
        self.root.bind("<Alt-h>", self.hypercare_MX)
        #Menu auxiliar acionado com o botão direito do mouse
        m = tk.Menu(root, tearoff=0)
        m.add_command(label="Pedir Prioridade", command=self.ask_priority)
        m.add_command(label="Top Ofensores", command=self.ranking_top10)
        m.add_command(label="Aging", command=self.range_aging)
        m.add_separator()
        m.add_command(label="Analisar ProblemID", command=self.locate_by_problem)
        m.add_command(label="Localizar Similares (Resumo)", command=self.locate_similar_resumo)
        m.add_command(label="Localizar Similares (Descrição)", command=self.locate_similar_desc)
        m.add_command(label="Exportar Selecionados", command=self.export_to_excel_selected)
        m.add_separator()
        m.add_command(label="Copiar ChamadoID", command=self.copy_chamado)
        m.add_command(label="Copiar ProblemID", command=self.copy_problema)
        m.add_command(label="Copiar Resumo", command=self.copy_resumo)
        m.add_command(label="Copiar Descrição", command=self.copy_desc)
        m.add_command(label="Exportar Estatisticas", command=self.handle_stats)
        m.add_separator()
        m.add_command(label="Fechar", command=root.destroy)
        #self.df é a lista de chamados imporadas sem qualquer filtro, referencia para os filtros e caso precise ler a lista full
        self.df = pd.DataFrame()

        self.tree.pack(expand=True, fill="y")
        #Filter Frame irá conter todos os combobox dos filtros
        self.filter_frame = ttk.Frame(self.search_filter_frame)
        self.filter_frame.pack(side="top")        
        self.tipo_label = ttk.Label(self.filter_frame, text="Tipo:")
        self.tipo_label.pack(side="left", padx=5, pady=2)

        self.tipo_filter = ttk.Combobox(self.filter_frame)
        self.tipo_filter.pack(side="left", fill="x", padx=5, pady=2)
    
        self.status_label = ttk.Label(self.filter_frame, text="Status:")
        self.status_label.pack(side="left", padx=5, pady=2, expand=False)

        self.status_filter = ttk.Combobox(self.filter_frame)
        self.status_filter.pack(side="left", fill="x", padx=5, pady=2)

        self.status_label = ttk.Label(self.filter_frame, text="SLA:")
        self.status_label.pack(side="left", padx=5, pady=2)

        self.sla_filter = ttk.Combobox(self.filter_frame)
        self.sla_filter.pack(side="left", fill="x", padx=5, pady=2)

        self.localidade_label = ttk.Label(self.filter_frame, text="Localidade:")
        self.localidade_label.pack(side="left", padx=5, pady=2)

        self.localidade_filter = ttk.Combobox(self.filter_frame)
        self.localidade_filter.pack(side="left", fill="x", padx=5, pady=2)

        self.grupo_label = ttk.Label(self.filter_frame, text="Grupo:")
        self.grupo_label.pack(side="left", padx=5, pady=2)

        self.grupo_filter = ttk.Combobox(self.filter_frame)
        self.grupo_filter.pack(side="left", fill="x", padx=5, pady=2)

        self.filter_button = ttk.Button(self.filter_frame, text="Filtrar", command=self.filter_data)
        self.filter_button.pack(side="bottom", pady=5)
        
        #labels que guardam as informações estatisicas, são atualizados de forma dinamica com o metodo update_indo
        self.count_tkt_label = ttk.Label(self.list_frame)
        self.count_tkt_label.pack(side="bottom", padx=10, pady=5)
        
        self.tkt_por_status_label = ttk.Label(self.list_frame)
        self.tkt_por_status_label.pack(side="bottom", padx=10, pady=5)
        
        self.chamado_com_problema_count_label = ttk.Label(self.list_frame)
        self.chamado_com_problema_count_label.pack(side="bottom", padx=10, pady=5)

        self.sla_count_label = ttk.Label(self.list_frame)
        self.sla_count_label.pack(side="bottom", padx=1, pady=2)
        

        self.tkt_por_pais_label = ttk.Label(self.list_frame)
        self.tkt_por_pais_label.pack(side="bottom", padx=1, pady=2)
        
        self.tkt_por_tipo_label = ttk.Label(self.list_frame)
        self.tkt_por_tipo_label.pack(side="bottom", padx=1, pady=2)

        self.tkt_term_label = ttk.Label(self.list_frame)
        self.tkt_term_label.pack(side="bottom", padx=1, pady=2)

        self.description_frame = ttk.Frame(self.list_frame)
        self.description_frame.pack(fill="x", expand=True)

        self.description_label = ttk.Label(self.description_frame, text="Descrição:")
        self.description_label.pack(side="top", padx=5, pady=5)

        self.description_text = tk.Text(self.description_frame, wrap=tk.WORD)
        self.description_text.pack(side="top", padx=5, pady=5, fill="both", expand=True)
        
        self.scroll_frame = ttk.Frame(self.list_frame)
        self.scroll_frame.pack(fill="both", expand=True)

        self.tree_frame = ttk.Frame(self.scroll_frame)
        self.tree_frame.pack(fill="both", expand=True)
        #método para centralizar a janela quando windowed ou full screen
    def center_window(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = 800
        height = 600
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.root.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
        #método para selecinar um tema (aparencia)
    def dark_mode(self,event=None):
         s=ttk.Style()
         s.theme_use('dark_orange') #tema custom (Natura Dark)
    def clam_mode(self,event=None):
         s=ttk.Style()
         s.theme_use('clam') # tema padrão do thinkter (Natura Retro)
    def vista_mode(self,event=None):
         s=ttk.Style()
         s.theme_use('vista') #tema default do windows (Natura Classico)
         
        #método para alterar entre windowwed e full screen
    def toggle_fullscreen(self):
        if self.is_fullscreen:
            self.root.attributes("-fullscreen", False)
            self.is_fullscreen = False
        else:
            self.root.attributes("-fullscreen", True)
            self.is_fullscreen = True
        #método para importar o backlog (caixa de dialogo)
    def upload_backlog(self, event=None):
         file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")]) #só é aceito arquivos csv separado por virgula
         if file_path:
            self.load_backlog(file_path) #chama o metodo load_backlog que irá fazer o processo de ETL do CSV e parser das informações
         else:
            messagebox.showinfo("Aviso", "Importação cancelada")
        #cria um formato para o Entry de data de abertura DD/MM/YYYY
    def format_date_abertura(self, event=None):
     date_str = self.dt_abertura_filter.get()
     if len(date_str) == 2 and event.keysym.lower() != "backspace":
        self.dt_abertura_filter.insert(tk.END, "/")
     elif len(date_str) == 5 and event.keysym.lower() != "backspace":
        self.dt_abertura_filter.insert(tk.END, "/")
     elif len(date_str) > 10 and event.keysym.lower() != "backspace":
      messagebox.showwarning("Aviso", "Data Invalida, por favor revisar") #Valida formato
        #cria um formato para o Entry de data de fechamento DD/MM/YYYY
    def format_date_fechamento(self, event=None):
     date_str = self.dt_fechamento_filter.get()
     if len(date_str) == 2 and event.keysym.lower() != "backspace":
        self.dt_fechamento_filter.insert(tk.END, "/")
     elif len(date_str) == 5 and event.keysym.lower() != "backspace":
        self.dt_fechamento_filter.insert(tk.END, "/")
     elif len(date_str) > 10 and event.keysym.lower() != "backspace":
      messagebox.showwarning("Aviso", "Data Invalida, por favor revisar")
      
    #método ETL que carrega dos dados do CSV para a memoria do proegrama e cria o Dataframe self.df
    def load_backlog(self, file_path):
        try:
            self.df = pd.read_csv(file_path, sep=",", encoding="UTF-8")
            self.df = self.df.replace([pd.NA, np.nan, 'nan'], '')
            self.df["DT_ABERTURA"] = pd.to_datetime(self.df["DT_ABERTURA"], errors='coerce') #corrige erros no formato de data
            self.df["DT_ABERTURA"] = self.df["DT_ABERTURA"].dt.strftime('%d/%m/%Y') #transforma a string para um date
            self.df["DT_SOLUÇÃO"] = pd.to_datetime(self.df["DT_SOLUÇÃO"], errors='coerce')
            self.df["DT_SOLUÇÃO"] = self.df["DT_SOLUÇÃO"].dt.strftime('%d/%m/%Y')
            self.df["DT_SOLUÇÃO"] = self.df["DT_SOLUÇÃO"].fillna('') #converte as linhas NaN para vazio no campo data de solução
            self.df["RESUMO"] = self.df["RESUMO"].astype(str)
            self.df["CHAMADO"] = self.df["CHAMADO"].astype(str)
            self.df['PROBLEMA'] = self.df['PROBLEMA'].apply(lambda x: str(x).replace('.0', '')) #converte as linhas do problem para int removendo um .0 que o Pandas não consegue tratar automaticamente
            self.df["AGING_IN_DAYS"] = self.df["AGING_IN_DAYS"].astype(int) #converte a string para int para calculos
            self.df["GRUPO"] = self.df["GRUPO"].astype("category") #converte o tipo string para o tipo category do pandas
            #mapa dos status para facilitar a leitura
            self.df["STATUS"] = self.df["STATUS"].astype(str).map({"Fechada":"Resolvido", "Resolvida":"Resolvido", "Cancelada":"Cancelado", "Aberta":"Aberto", "Em andamento":"Analise N2", "Aguardando resposta do usuário final":"Ag. usuário", "Aguardando fornecedor":"Ag. Forn", "Mudança em Andamento":"Ag.SM", "Reaberto":"Reaberto", "Aguardando Resposta do usuário final":"Ag.N1", "Aprovada":"Aprovada"})
            self.df["STATUS"] = self.df["STATUS"].astype('category')
            #mapa dos paises para facilitar a leitura
            self.df["LOCALIDADE"] = self.df["LOCALIDADE"].map({"BR": "Brasil", "CL": "Chile", "CO": "Colombia", "PE": "Peru", "AR-NAT": "Argentina", "MX": "México", "MY": "Malasia","":"Não Informado",np.nan:"Não informado","AR-AVON":"Argentina"})
            self.df["LOCALIDADE"] = self.df["LOCALIDADE"].astype('category')
            self.df['SLA_VIOLADO'] = self.df['SLA_VIOLADO'].map({0: "Não Violado", 1: "Violado", 2:"Violado",3:"Violado",4:"Violado",5:"Violado"}) #converte o campo SLA_Violado para um texto explicativo
            self.df['TIPO'] = self.df['TIPO'].map({"ERRO":"Erro", "SOLICITAÇÃO":"Solicitação", "DÚVIDA":"Dúvida", "INDISPONIBILIDADE":"Indisponibilidade", "JOBS":"Jobs", "LENTIDÃO":"Lentidão", "MONITORAÇÃO":"Monitoração","N/A":'Não Informado',np.nan:'Não Informado','':'Não Informado'})

            self.df["TIPO"] = self.df["TIPO"].astype('category')
            #cria os widgets combobox de filtro
            self.populate_status_filter()
            self.populate_localidade_filter()
            self.populate_sla_filter()
            self.populate_grupo_filter()
            self.populate_tipo_filter()
            #cria o widget na tela para mostrar os dados
            self.display_data()
            #instancia o widget que apresenta a mensagem de projeto beta (será removido em versões futuras)
            self.help_popup()
        except Exception as e:
         messagebox.showerror("Erro", f"Erro ao abrir backlog: {str(e)}, por favor revisar manual do usuário, baixar o backlog do Tableau, e abrir no Google Planilhas para converter em CSV separado por virgulas")  
        return self.df
        #metodo que atualiza os dados estatisticos conforme os filtros são selecionados
    def update_info(self, df):
     if df is not None:
        pais = df['LOCALIDADE'].value_counts().to_dict()
        tipo = df['TIPO'].value_counts().to_dict()
        status = df['STATUS'].value_counts().to_dict()
        dentro_count = len(df[df["SLA_VIOLADO"] == "Não Violado"])
        violado_count = len(df[df["SLA_VIOLADO"] == "Violado"])

        self.count_tkt_label.configure(
            text=f"Tickets Filtrados: {len(df)}",
            anchor=tk.CENTER,
            wraplength=300  
        )
        
        self.sla_count_label.configure(
            text=f"SLA no prazo: {dentro_count}  SLA Violado: {violado_count}",
            anchor=tk.CENTER,
            wraplength=300  
        )
        
        self.chamados_sem_problema = len(df[df["PROBLEMA"].str.strip() == ""])
        self.tkt_problema = len(df) - self.chamados_sem_problema
        
        self.chamado_com_problema_count_label.configure(
            text=f"Tickets com problema atrelado: {self.tkt_problema}",
            anchor=tk.CENTER,
            wraplength=500  
        )
        
        
        pais_text = " | ".join(f"{country}: {count}" for country, count in pais.items()if count > 0)
        self.tkt_por_pais_label.configure(
            text=pais_text,
            anchor=tk.CENTER,
            wraplength=450 
        )
        
        tipo_text = " | ".join(f"{tipo_name}: {count}" for tipo_name, count in tipo.items()if count > 0)
        self.tkt_por_tipo_label.configure(
            text=tipo_text,
            anchor=tk.CENTER,
            wraplength=450 
        )
        
        status_text = " | ".join(f"{status_name}: {count}" for status_name, count in status.items()if count > 0)
        self.tkt_por_status_label.configure(
            text=status_text,
            anchor=tk.CENTER,
            wraplength=400
        )
        return self.tkt_problema
        #metodo captura e formata as estatisticas, exportando para a area de transferencia os dados em um formato legivel e organizado
    def export_stats(self,event=None):
     try:
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=["DT_ABERTURA", "DT_SOLUÇÃO", "CHAMADO", "PROBLEMA", "GRUPO", "STATUS", "TIPO", "RESUMO", "AGING_IN_DAYS", "LOCALIDADE", "SLA_VIOLADO", "DESCRICAO"])
        stats = (
            f"Tickets Filtrados: {len(df)}\n"
            f"SLA no prazo: {len(df[df['SLA_VIOLADO'] == 'Não Violado'])}  SLA Violado: {len(df[df['SLA_VIOLADO'] == 'Violado'])}\n"
            f"Tickets com problema atrelado: {self.tkt_problema}\n"
            f"{' | '.join(f'{country}: {count}' for country, count in df['LOCALIDADE'].value_counts().to_dict().items() if count > 0)}\n"
            f"{' | '.join(f'{tipo_name}: {count}' for tipo_name, count in df['TIPO'].value_counts().to_dict().items() if count > 0)}\n"
            f"{' | '.join(f'{status_name}: {count}' for status_name, count in df['STATUS'].value_counts().to_dict().items() if count > 0)}"
        )
        self.root.clipboard_clear()
        self.root.clipboard_append(stats)
        messagebox.showinfo('Ok', 'Estatisticas copiadas por favor colar no bloco de notas Google Documentos')
     except Exception as e:
         messagebox.showerror('Erro','Não foi Possível copiar os dados, por favor  tente novamente!')
        #popular as opçoes do combobox de filtro, de forma dinamica conforme o objeto self.df que foi instanciado anteriormenete
    def populate_status_filter(self):
        self.status_filter["values"] = ["Todos"] + self.df["STATUS"].unique().tolist()
        self.status_filter.set("Todos")

    def populate_grupo_filter(self):
        self.grupo_filter["values"] = ["Todos"] + self.df["GRUPO"].unique().tolist()
        self.grupo_filter.set("Todos")

    def populate_sla_filter(self):
        self.sla_filter["values"] = ["Todos"] + self.df["SLA_VIOLADO"].unique().tolist()
        self.sla_filter.set("Todos")

    def populate_localidade_filter(self):
        self.localidade_filter["values"] = ["Todos"] + self.df["LOCALIDADE"].unique().tolist()
        self.localidade_filter.set("Todos")

    def populate_tipo_filter(self):
        self.tipo_filter["values"] = ["Todos"] + self.df["TIPO"].unique().tolist()
        self.tipo_filter.set("Todos")
    #metodo que popula a arvore de chamados conforme os filtros e base importada
    def display_data(self, df=None):
     self.tree.delete(*self.tree.get_children())
     if df is None:
        df = self.df
     for index, row in df.iterrows():
        self.tree.insert("", "end", iid=index, values=(row["DT_ABERTURA"], row['DT_SOLUÇÃO'], row["CHAMADO"], row["PROBLEMA"], row["GRUPO"], row["STATUS"], row['TIPO'], row["RESUMO"], row["AGING_IN_DAYS"], row["LOCALIDADE"], row["SLA_VIOLADO"], row['DESCRICAO']))
        self.description_text.delete(1.0, tk.END)
    def on_select(self,event=None):
        selected_item = self.tree.selection()
        for item in self.tree.get_children():
            self.tree.item(item, tags=())
        for item in selected_item:
            self.tree.item(item, tags=('selected',))
    #metod que exporta os chamados filtrados na arvore, invocado atraves do botão do menu "exportar fitrrados"
    def export_to_excel(self, event=None):
     file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Exportar como Excel")
    
     if file_path:

        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=["DT_ABERTURA", "DT_SOLUÇÃO", "CHAMADO", "PROBLEMA", "GRUPO", "STATUS", "TIPO", "RESUMO", "AGING_IN_DAYS", "LOCALIDADE", "SLA_VIOLADO", "DESCRICAO"])
        try:
            df.to_excel(file_path, index=False, sheet_name='Backlog')
            messagebox.showinfo = ("Ok", f"Arquivo {file_path.title} Exportado com sucesso")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar os dados: {str(e)}")
    #logica dos filtros sao combinados e podem filtrar multiplas vezes depedendo das escolhas do usuário
    def filter_data(self, event=None):
     terms1 = self.search_entry1.get()
     terms2 = self.search_entry2.get()
     problema = self.problema_filter.get()
     terms1 = ''.join(self.remove_special_chars(char) for char in terms1)
     terms2 = ''.join(self.remove_special_chars(char) for char in terms2)
     operator = self.operator.get()
     status = self.status_filter.get()
     localidade = self.localidade_filter.get()
     sla = self.sla_filter.get()
     grupo = self.grupo_filter.get()
     tipo = self.tipo_filter.get()
     dt_abertura = self.dt_abertura_filter.get()
     dt_solucao = self.dt_fechamento_filter.get()
     filtered_df = self.df
     #Lógica da pesquisa por termos, irá ler as colunas resumo e descriçao de forma separada e depois juntar em um unico resultado removendo os dupkicados
     if terms1 and not terms2:
      term_list = terms1.split(",")
      filtered_df = self.df[((self.df["DESCRICAO"].str.contains("|".join(term_list), case=False)) | 
                              (self.df["DESCRICAO"].str.contains("|".join(term_list), case=False))) | 
                             ((self.df["RESUMO"].str.contains("|".join(term_list), case=False)) | 
                              (self.df["RESUMO"].str.contains("|".join(term_list), case=False)))]
      filtered_df = filtered_df.drop_duplicates()

     elif terms1 and terms2 and operator == "ou":
      term_list1 = terms1.split(",")
      term_list2 = terms2.split(",")
      filtered_df = self.df[((self.df["DESCRICAO"].str.contains("|".join(term_list1), case=False)) | 
                              (self.df["DESCRICAO"].str.contains("|".join(term_list2), case=False))) | 
                             ((self.df["RESUMO"].str.contains("|".join(term_list1), case=False)) | 
                              (self.df["RESUMO"].str.contains("|".join(term_list2), case=False)))]
      filtered_df = filtered_df.drop_duplicates()
     elif terms1 and terms2 and operator == "e":
      term_list1 = terms1.split(",")
      term_list2 = terms2.split(",")
      
      combinations = [(term1, term2) for term1 in term_list1 for term2 in term_list2] #testa todas as possibilidades dos operadores quando os 2 termos são inseridos
    
      conditions = []
      for term1, term2 in combinations:
        conditions.append(
            ((self.df["DESCRICAO"].str.contains(term1, case=False)) & 
             (self.df["DESCRICAO"].str.contains(term2, case=False))) | 
            ((self.df["RESUMO"].str.contains(term1, case=False)) & 
             (self.df["RESUMO"].str.contains(term2, case=False)))
        )
    
      filtered_df = self.df[np.logical_or.reduce(conditions)] #reduz os resultados baseado no retorno
      filtered_df = filtered_df.drop_duplicates()   #remove chamados duplicados
     #testa cada um dos combobox e filtra os chamados anterioremante já filtrados
     if status != "Todos":
            filtered_df = filtered_df[filtered_df["STATUS"] == status]
     if localidade != "Todos":
            filtered_df = filtered_df[filtered_df["LOCALIDADE"] == localidade]
            
     if sla != "Todos":
            filtered_df = filtered_df[filtered_df["SLA_VIOLADO"] == sla]
     if grupo != "Todos":
            filtered_df = filtered_df[filtered_df["GRUPO"] == grupo]
     if tipo != "Todos":
            filtered_df = filtered_df[filtered_df["TIPO"] == tipo]
     if dt_abertura:
            filtered_df = filtered_df[filtered_df["DT_ABERTURA"].str.contains(dt_abertura, case=False)]
     if dt_solucao:
            filtered_df = filtered_df[filtered_df["DT_SOLUÇÃO"].str.contains(dt_solucao, case=False)]
     if problema:
            filtered_df = filtered_df[filtered_df["PROBLEMA"].str.contains(problema, case=False)]
            #caso o objeto filtrado não retorne resultado será exibido uma mensagem indicando que não há resultados
     if len(filtered_df) == 0:
        messagebox.showinfo("Aviso", "Não foram encontrados resultados para os filtros selecionados")
     else: # caso exista será retornado no objeto filtered_df e atualiza a arvore e os dados de estatistica
      self.display_data(filtered_df)
      self.update_info(filtered_df)
     return filtered_df
         #metodo que localiza chamados similares basedo em um modelo de Machine Learning e le o campo resumo
    def locate_similar_resumo(self):
     try:
        selected_item = self.tree.selection()[0]
        selected_summary = self.tree.item(selected_item)['values'][7] #captura o campo resumo de um chamado selecionado

        def preprocess_text(text): #normaliza e converte o texto para minusculo para evitar sobrecarregar o modelo
            text = self.remove_special_chars(text)
            return text.lower()
        

        processed_selected_summary = preprocess_text(selected_summary)
        self.df['processed_resumo'] = self.df['RESUMO'].apply(preprocess_text) #aplica o preprocessamento

        all_summaries = [processed_selected_summary] + list(self.df['processed_resumo']) #cria uma lista com todos os resumos para validar
        #aplica o modelo de ML testado e separando os resumos
        vectorizer = TfidfVectorizer()
        tfidf_matrix = vectorizer.fit_transform(all_summaries)

        cosine_similarities = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:]).flatten()

        threshold = 0.9 #parametro de refinamento (score do modelo deve ser de no minimo 0.9 para ser valido)
        similar_indices = np.where(cosine_similarities >= threshold)[0] #localiza chamados que tenham o campo resumo dentro do parametro de treeshold

        if similar_indices.size > 0: #caso encontre atualiza a arvore e as estatisticas
            filtered_df = self.df.iloc[similar_indices]
            self.update_info(filtered_df)
            self.display_data(filtered_df)
        else: #caso não encontre uma excessão será levantada
            raise ValueError("Não foram encontrados chamados similares na base fornecida")

     except Exception as e:
        messagebox.showinfo("Aviso", "Não foram encontrados chamados similares na base fornecida")

        #metodo que localiza chamados similares basedo em um modelo de Machine Learning e le o campo descrição, para referecias leia os comentarios do metodo acima
    def locate_similar_desc(self):
     try:
        selected_item = self.tree.selection()[0]
        selected_summary = self.tree.item(selected_item)['values'][-1]

        def preprocess_text(text):
            text = self.remove_special_chars(text)
            return text.lower()

        processed_selected_summary = preprocess_text(selected_summary)
        self.df['processed_resumo'] = self.df['DESCRICAO'].apply(preprocess_text)

        all_summaries = [processed_selected_summary] + list(self.df['processed_resumo'])

        vectorizer = TfidfVectorizer()
        tfidf_matrix = vectorizer.fit_transform(all_summaries)

        cosine_similarities = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:]).flatten()

        threshold = 0.9
        similar_indices = np.where(cosine_similarities >= threshold)[0]

        if similar_indices.size > 0:
            filtered_df = self.df.iloc[similar_indices]
            self.update_info(filtered_df)
            self.display_data(filtered_df)
        else:
            raise ValueError("No similar summaries found")

     except Exception as e:
        messagebox.showinfo("Aviso", "Não foram encontrados chamados similares na base fornecida")
        #metodo que cria uma arvore de camados especial, que não irá respeitar os filtros se selecionados porém segue uma regra de escalation do N1:
        #Aging acima de 9 dias, sem problem atrelado e status diferente de resolvido
    def escalation_aging(self):
         status_to_filter = ['Resolvido', 'Cancelado', 'Cancelada']
         filtered_df = self.df[(self.df['AGING_IN_DAYS'] > 9) & (self.df['PROBLEMA'] == '') & (~self.df['STATUS'].isin(status_to_filter))]
         self.display_data(filtered_df)
         self.update_info(filtered_df)
        #metodo que faz uma analise temporal baseado em chamados abertos, ao longo de um intervalo de tempo conforme a variavel ndays receber do usuario
    def aberto_dayminuscustom(self):
     try:
        ndays = tk.simpledialog.askinteger(title="Olá",
                                            prompt="Deseja ver o histórico de abertos de quantos dias?:")
        if ndays is None or ndays < 0:
            messagebox.showinfo("Aviso", "Por favor, insira um número válido de dias.")
            return
        filtered_df = self.filter_data(self.df)
        
        start_date = datetime.now() - timedelta(days=ndays+1)
        end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        filtered_df = filtered_df[
            (pd.to_datetime(filtered_df['DT_ABERTURA'], format='%d/%m/%Y') >= start_date) & 
            (pd.to_datetime(filtered_df['DT_ABERTURA'], format='%d/%m/%Y') < end_date)
        ]
        
        if filtered_df.empty:
            messagebox.showinfo("Aviso", f"Não há chamados em D-{ndays} na base fornecida.")
        else:
            self.display_data(filtered_df)
            self.update_info(filtered_df)
    
     except Exception as e:
        messagebox.showinfo("Erro", f"Ocorreu um erro: {str(e)}")
    def aberto_dayminusone(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus1 = datetime.now() - timedelta(days=2)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') >= dayminus1) & 
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') <= datetime.now())]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados D-1 na base fornecida")
    
    def aberto_dayminusseven(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus7 = datetime.now() - timedelta(days=8)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') >= dayminus7) & 
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') <= datetime.now())
    ]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados em  D-7 na base fornecida")
    
    def aberto_dayminusthirty(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus30 = datetime.now() - timedelta(days=31)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') >= dayminus30) & 
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') <= datetime.now())]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados abertos em  D-30 em na base fornecida")

    def aberto_dayminussixty(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus60 = datetime.now() - timedelta(days=61)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') >= dayminus60) & 
        (pd.to_datetime(self.df['DT_ABERTURA'], format='%d/%m/%Y') <= datetime.now())]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados abertos em  D-60 na base fornecida")
        #metodo que faz uma analise temporal baseado em chamados resolvidos, ao longo de um intervalo de tempo conforme a variavel ndays receber do usuario
    def fechado_dayminuscustom(self):
     try:
        ndays = tk.simpledialog.askinteger(title="Cockpit da Operação",
                                            prompt="Deseja ver o histórico de chamados resolvidos de quantos dias?:")
        if ndays is None or ndays < 0:
            messagebox.showinfo("Aviso", "Por favor, insira um número válido de dias.")
            return
        filtered_df = self.filter_data(self.df)
        
        start_date = datetime.now() - timedelta(days=ndays+1)
        end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        filtered_df = filtered_df[
            (pd.to_datetime(filtered_df['DT_SOLUÇÃO'], format='%d/%m/%Y') >= start_date) & 
            (pd.to_datetime(filtered_df['DT_SOLUÇÃO'], format='%d/%m/%Y') < end_date)
        ]
        
        if filtered_df.empty:
            messagebox.showinfo("Aviso", f"Não há chamados em D-{ndays} na base fornecida.")
        else:
            self.display_data(filtered_df)
            self.update_info(filtered_df)
     except Exception as e:
        messagebox.showerror("Erro",f"Erro ao processar dados {e}")

    
    def fechado_dayminusone(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus1 = datetime.now() - timedelta(days=2)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') >= dayminus1) & 
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') <= datetime.now())]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados D-1 na base fornecida")
    
    def fechado_dayminusseven(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus7 = datetime.now() - timedelta(days=8)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') >= dayminus7) & 
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') <= datetime.now()) ]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados fechados em  D-7 na base fornecida")
    
    def fechado_dayminusthirty(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus30 = datetime.now() - timedelta(days=31)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') >= dayminus30) & 
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') <= datetime.now())]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados fechados em  D-30 na base fornecida")

    def fechado_dayminussixty(self):
     try:
        filtered_df=self.filter_data(self.df)
        dayminus60 = datetime.now() - timedelta(days=61)
        filtered_df = filtered_df[
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') >= dayminus60) & 
        (pd.to_datetime(self.df['DT_SOLUÇÃO'], format='%d/%m/%Y') <= datetime.now())]
        self.display_data(filtered_df)
        self.update_info(filtered_df)
     except Exception as e:
        messagebox.showinfo("Aviso", "Não há chamados fechados em  D-60 na base fornecida")

    def filter_fechado(self, event=None):
     try:
      filtered_df=self.df
      filtered_df=self.filter_data(self.df)
      status_to_include = ['Resolvido', 'Cancelado']  
      filtered_df = filtered_df[filtered_df['STATUS'].isin(status_to_include)]
      
      if len(filtered_df) == 0:
            messagebox.showinfo("Aviso", "Não foram encontrados chamados fechados na base fornecida")
      else:
            self.update_info(filtered_df)
            self.display_data(filtered_df)
     except:
        messagebox.showinfo("Aviso", "Não foram encontrados chamados fechados na base fornecida")

    def filter_abertos(self, event=None):
     try:
      filtered_df=self.df
      filtered_df=self.filter_data(self.df)
      status_to_exclude = ['Resolvido', 'Cancelado']  
      filtered_df = filtered_df[~filtered_df['STATUS'].isin(status_to_exclude)]
      
      if len(filtered_df) == 0:
            messagebox.showinfo("Aviso", "Não foram encontrados chamados abertos na base fornecida")
      else:
            self.update_info(filtered_df)
            self.display_data(filtered_df)
     except:
        messagebox.showinfo("Aviso", "Não foram encontrados chamados abertos na base fornecida")
        #filtro especial que será utilizado para hypercares ou por necessidade da operação
    def hypercare_MX(self, event=None):
     try:
      filtered_df=self.df
      filtered_df=self.filter_data(self.df)
      filtered_df = filtered_df[(filtered_df['GRUPO'].str.contains('HC')) & (filtered_df['LOCALIDADE'] == "México")] #HC Natura MX
      
      if len(filtered_df) == 0:
            messagebox.showinfo("Aviso", "Não foram encontrados chamados de hypercare MX na base fornecida")
      else:
            self.update_info(filtered_df)
            self.display_data(filtered_df)
     except:
        messagebox.showinfo("Aviso", "Não foram encontrados chamados de hypercare MX na base fornecida")

    def hypercare_BR(self, event=None):
     try:
      filtered_df=self.df
      filtered_df=self.filter_data(self.df)
      filtered_df = filtered_df[(filtered_df['GRUPO'].str.contains('HC')) & (filtered_df['LOCALIDADE'] == "Brasil")] #HC Elo Fase 2
      
      if len(filtered_df) == 0:
            messagebox.showinfo("Aviso", "Não foram encontrados chamados ELO Fase 2 na base fornecida")
      else:
            self.update_info(filtered_df)
            self.display_data(filtered_df)
     except:
        messagebox.showinfo("Aviso", "Não foram encontrados chamados Elo Fase 2 na base fornecida")
        
    #metodo que utiliza um problem fornecido no Entry problema_label para localizar os chamados atrelados
    def filter_problemid(self, event=None):
        try:
         filtered_df=self.filter_data(self.df)
         filtered_df = filtered_df[filtered_df["PROBLEMA"] == '']
         self.update_info(filtered_df)
         self.display_data(filtered_df)
        except:
         messagebox.showinfo("Aviso", "Não foram encontrados chamados sem problem atrelado na base fornecida") 
    #metoco que utiliza um problem selecionado na arvore para localizar os chamados atrelados
    def locate_by_problem(self, event=None):
        selected_item = self.tree.selection()[0]
        filtered_df = self.df
        problema_value = self.tree.item(selected_item)['values'][3]
        try:
         if problema_value:
          filtered_df = self.df[self.df["PROBLEMA"] == str(problema_value)]
          self.display_data(filtered_df)
         else:
            messagebox.showinfo("Opa!", "Chamado não possui problemID, por favor selecione novamente")
        except KeyError as e:
            messagebox.showerror("Aviso!", f"The column '{e.args[0]}' does not exist in the DataFrame.")
        self.update_info(filtered_df)
        self.display_data(filtered_df)
        #metodo que filtra os chamados que possuem problem atrelado
    def keep_problemid(self, event=None):
        try:
         filtered_df=self.filter_data(self.df)
         filtered_df = filtered_df[filtered_df["PROBLEMA"] > "1"]
         self.update_info(filtered_df)
         self.display_data(filtered_df)
        except:
         messagebox.showinfo("Aviso", "Não foram encontrados chamados com problem atrelado na base fornecida")
    #metodo que exporta um relatorio com chamados filtrados e cria uma nova coluna ao final com o range de aging baseado em uma lista personalziada
    def range_aging(self, event=None):
     try:
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        
        df = pd.DataFrame(data, columns=["DT_ABERTURA", "DT_SOLUÇÃO", "CHAMADO", "PROBLEMA", "GRUPO", "STATUS", "TIPO", "RESUMO", "AGING_IN_DAYS", "LOCALIDADE", "SLA_VIOLADO", "DESCRICAO"])
        
        bins = [0, 4, 11, 21, float('inf')] #lista personaizada de range, alterar conforme necessario
        labels = ['0-3 dias', '4-10 dias', '11-20 dias', '20+ dias'] #texto que representa cada faixa de aging
        df['AGING_IN_DAYS']=df['AGING_IN_DAYS'].astype(int) #converter a nova lista o campo aging_in_days para um int
        
        df['AGING_RANGE'] = pd.cut(df['AGING_IN_DAYS'], bins=bins, labels=labels) #cria a coluna usando o metodo cut para fatiar a lista baseado nas faixas
        df['AGING_RANGE'] = df['AGING_RANGE'].fillna('0-3 dias') #converte os chamados que possuem 0 dias para o range 0-3 dias, caso essa linha seja removida teremos linhas em branco
    
        
        df.to_clipboard(index=False)
        messagebox.showinfo("Ok", "Aging copiado, por favor cole no Google Planilhas ou Excel")
     except Exception as e:
        messagebox.showinfo("Aviso", f"Erro ao gerar dados: {e}")
    
    #metodo que atualiza um label de estatistica
    def search(self):
        term1 = self.search_entry1.get()
        term2 = self.search_entry2.get()
        operator = self.operator.get()
        self.tkt_term_label.config(text=f"Tickets sobre {term1} {operator} {term2}")
        self.filter_data()
    #metodo utilizado em conjunto com pre procesamento do texto para o modelo de ML
    def remove_special_chars(self, text):
        return re.sub(r'[^a-zA-Z0-9áéíóúàèìòùâêîôûãõäëïöüÁÉÍÓÚÀÈÌÒÙÂÊÎÔÛÃÕÄËÇÏÖÜ~]\[]', '', text) #remove caracteres especiais que não são utlizados na lingua portuguesa
    #logica para ordenar as colunas da arvore de acordo com o cabeçalho
    def sort_column(self, col, reverse):
        data_list = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        data_list.sort(reverse=reverse)
        for index, (val, k) in enumerate(data_list):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))
        #metodo que copia o numero do chamado da seleção, funciona com 1 ou mais chamados selecionados
    def copy_chamado(self, event=None):
     selected_items = self.tree.selection()
     if selected_items:
        chamado_values = [self.tree.item(item)['values'][2] for item in selected_items] #caso a coluna mudar de posição no DataFrame alterar conforme necessário
        self.root.clipboard_clear() #limpa a area de transferencia
        self.root.clipboard_append('\n'.join(map(str, chamado_values))) #copia o chamado
        self.root.update() #atualiza a tela
        messagebox.showinfo("OK!", "Chamado(s) copiado(s). Pode pesquisar no TD Interativa")
     else:
        messagebox.showwarning("Nenhum item selecionado", "Por favor, selecione um ou mais itens na tabela.")
        #metodo que cria um texto padronizado para pedir prioridade para o fornecedor, verifique manual para um exemplo completo do texto
    def ask_priority(self, event=None):
     selected_items = self.tree.selection() #captua as linhas selecionadas
     if selected_items:
        priority_messages = [] #cria uma lista vazia para servir de container
        for item in selected_items: #separa em variaveis cada campo para melhor leitura
            chamado = self.tree.item(item)['values'][2]
            descricao = self.tree.item(item)['values'][7]
            aging = self.tree.item(item)['values'][8]
            sla = self.tree.item(item)['values'][10]
            pais = self.tree.item(item)['values'][9]
            status = self.tree.item(item)['values'][5]
            priority_message = f"Pedido de Prioridade: {pais}\nTime por favor priorizar chamado: {chamado}\nDescrição: {descricao}\nAging: {aging} dias\nSLA: {sla}\nStatus atual: {status}\nObrigado\n=====" #mensagem que será gerada, alterar conforme necessário
            priority_messages.append(priority_message) #caso tenha mais de uma linha selecionada o metodo irá add em um loop
        
        priority_text = '\n\n'.join(priority_messages) #pula 2 linhas entre uma mensagem e outra para melhor legibilidade
        self.root.clipboard_clear()
        self.root.clipboard_append(priority_text)
        self.root.update()
        messagebox.showinfo("OK!", "Prioridade(s) Gerada(s). Por favor colar no grupo do Meet")
     else:
        messagebox.showwarning("Nenhum item selecionado", "Por favor, selecione um ou mais itens na tabela.")
        #metodo que copia o resumo, funciona com uma ou mais linhas selecionadas
    def copy_resumo(self, event=None):
        selected_item = self.tree.selection()[0]
        try:
            chamado_value = self.tree.item(selected_item)['values'][7]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(chamado_value))
            self.root.update()
            messagebox.showinfo("OK!", "Resumo copiado.")
        except KeyError as e:
            messagebox.showerror("Erro!", f"A coluna '{e.args[0]}' não existe")
        #metodo que copia o resumo, funciona com uma ou mais linhas selecionadas
    def copy_desc(self, event=None):
        selected_item = self.tree.selection()[0]
        try:
            chamado_value = self.tree.item(selected_item)['values'][-1]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(chamado_value))
            self.root.update()
            messagebox.showinfo("OK!", "Descrição copiada.")
        except KeyError as e:
            messagebox.showerror("Erro!", f"A column '{e.args[0]}' does not exist in the DataFrame.")
        #metodo que copia o numero do problem, funciona com uma ou mais linhas selecionadas
    def copy_problema(self, event=None):
     selected_items = self.tree.selection()
     if selected_items:
        problema_values = [self.tree.item(item)['values'][3] for item in selected_items]
        problema_values = [value for value in problema_values if value]  
        if problema_values:
            self.root.clipboard_clear()
            self.root.clipboard_append('\n'.join(map(str, problema_values)))
            self.root.update()
            messagebox.showinfo("Ok!", "Problema(s) copiado(s). Pode pesquisar no TD Interativa.")
        else:
            messagebox.showwarning("Aviso!", "Nenhum problema atrelado aos chamados selecionados.")
     else:
        messagebox.showwarning("Nenhum item selecionado", "Por favor, selecione um ou mais itens na tabela.")
        #metodo que exporta para xls chamados selecionados na arvore
    def export_to_excel_selected(self, event=None):
        selected_items = self.tree.selection()
        data=[]
        for item in selected_items:
            values = self.tree.item(item)["values"]
            data.append(values)
        if len(data)>1:
            df = pd.DataFrame(data, columns=["DT_ABERTURA", "DT_SOLUÇÃO", "CHAMADO", "PROBLEMA", "GRUPO", "STATUS","TIPO", "RESUMO", "AGING_IN_DAYS", "LOCALIDADE", "SLA_VIOLADO", "DESCRICAO"])
            try:
                df.to_clipboard()
                messagebox.showinfo("Ok", f"{len(data)} Chamado(s) copiado(s) com sucesso, por favor colar no Google Planilhas!")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao exportar os dados: {str(e)}")
        else:
            messagebox.showwarning("Aviso", "Nenhum item selecionado para exportar.")
        #metodo que atualiza de forma dinamica, conforme a seleção os dados de SLA, RESUMO,  DESCRICAO e AGING
    def display_description(self, event=None):
     selected_item = self.tree.selection()[0]
     row_data = self.tree.item(selected_item, 'values')
     self.description_text.delete(1.0, tk.END)
     self.description_text.configure(font=self.roman_font)
     sla_status = "Dentro do SLA!" if row_data[10] == "Não Violado" else "SLA Violado!"
     color = "green" if row_data[10] == "Não Violado" else "red"
     bold = 'bold' if row_data[10] == "SLA Violado" else 'normal'
    
     self.description_text.insert(tk.END, sla_status, (f"{color}_text",))
     self.description_text.tag_config(f"{color}_text", foreground=color, font=(self.roman_font, 12, bold))
     self.description_text.insert(tk.INSERT, "\n\nResumo: ")
     self.description_text.insert(tk.INSERT, row_data[7])
     self.description_text.insert(tk.INSERT, "\n\nAging: ")
     self.description_text.insert(tk.INSERT, f"{row_data[8]} dias")
     self.description_text.insert(tk.INSERT, f"\n\n{row_data[-1]}")
        
        #metodo que exporta os dados estatisicos para a area de transferencia
    def handle_stats(self, event=None):
     self.export_stats()
    #metodo para normalizar o texto para o modelo de ML
    def preprocess_text(self,text):
     text = text.lower()
     text = re.sub(r'[^a-zA-Z0-9çáéíóúãõâêîôû\s]', '', text)
     text = re.sub(r'\s+', ' ', text).strip()
     return text
    #metodo autilixar para localizar resumos similares com uma melhor acertividade faz parte do pacote difflib
    def similar(self,a, b):
     return SequenceMatcher(None, a, b).ratio()
    #metodo para classificar e contar chamados baseado no campo resumo, o modelo faz a leitura e agrupa em palavras-chaves
    def ranking_top10(self, event=None):
    # Extrair os dados da treeview
     data = [self.tree.item(item)["values"] for item in self.tree.get_children()]
    
    # Criar o DataFrame com os dados extraídos
     columns = ["DT_ABERTURA", "DT_SOLUÇÃO", "CHAMADO", "PROBLEMA", "GRUPO", "STATUS", "TIPO", "RESUMO", "AGING_IN_DAYS", "LOCALIDADE", "SLA_VIOLADO", "DESCRICAO"]
     df = pd.DataFrame(data, columns=columns)
    
    # Preprocessar os textos na coluna 'RESUMO'
     def preprocess_resumo(text):
        # Remove termos irrelevantes como "GV"
        text = text.lower().replace("gv ", "")
        # Outros pré-processamentos podem ser aplicados aqui, como remoção de stopwords
        return text
    
     df['RESUMO'] = df['RESUMO'].apply(preprocess_resumo)
    
    # Agrupar os dados por 'GRUPO', 'LOCALIDADE' e 'RESUMO'
     df_grouped = df.groupby(['GRUPO', 'LOCALIDADE', 'RESUMO']).size().reset_index(name='Tickets')
    
    # Identificar temas similares (baseado em TF-IDF e similaridade de cosseno)
     def group_themes(resumos):
        vectorizer = TfidfVectorizer()
        vectors = vectorizer.fit_transform(resumos).toarray()
        cosine_matrix = cosine_similarity(vectors)
        return cosine_matrix
    
     df_grouped['Resumos_Similares'] = df_grouped.apply(lambda x: 'Divergência entre GPP e Gera' 
                                                       if 'divergencia' in x['RESUMO'] and 'gpp' in x['RESUMO'] and 'gera' in x['RESUMO'] 
                                                       else x['RESUMO'], axis=1)
    
    # Agrupar novamente pelo tema consolidado (Resumos_Similares) e somar os Tickets
     df_final = df_grouped.groupby(['GRUPO', 'LOCALIDADE', 'Resumos_Similares']).agg({'Tickets': 'sum'}).reset_index()
    
    # Ordenar o DataFrame pelos 'Tickets' em ordem decrescente
     df_final = df_final.sort_values(by='Tickets', ascending=False)
    
    # Copiar os dados para a área de transferência
     try:
        df_final.to_clipboard(sep="|", index=False)
        messagebox.showinfo("Ok", f"A IA analisou {len(df)} Chamados e Classificou da melhor forma possível, por favor colar os dados na Planilha Google ou Excel")
     except Exception as e:
        messagebox.showerror("Opa", f"Não foi possível copiar dados: {e}")
    
    

    #popups sobre o projeto e sobre o manual do usuario
    def help_popup(self,event=None):
     messagebox.showinfo("Que Fita!","Esse projeto está em status beta fechado, bugs podem, e irão aparecer, por favor reporte qualquer sintoma ou comportamento inesperado. Obrigado por testar!")

    def manual_popup(self,event=None):
     messagebox.showinfo("Manual","Esse projeto possui um guia de usuário que pode ser acessado em PDF, dentro da pasta do projeto e no drive da sustentação. Obrigado!")
   

if __name__ == "__main__":
    root = tk.Tk()
    app = BacklogViewer(root)
    app.center_window()
    root.mainloop()
