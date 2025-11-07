# -*- coding: utf-8 -*-
"""
Monitoramento de Ociosidade do Usuário com Interface Tkinter, registro em MySQL e gráficos/Excel

Necessário:
pip install pynput mysql-connector-python pandas matplotlib openpyxl

Configurar no XAMPP Banco:
CREATE DATABASE monitor_ociosidade;
USE monitor_ociosidade;
CREATE TABLE ociosidade_log (
    id INT PRIMARY KEY AUTO_INCREMENT,
    inicio_ocioso DATETIME,
    fim_ocioso DATETIME,
    duracao_segundos INT,
    usuario VARCHAR(64),
    host VARCHAR(64)
);
"""

import threading
import time
import os
import socket
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pynput import keyboard, mouse
import mysql.connector
import pandas as pd
import matplotlib.pyplot as plt

class DBHelper:
    """
    Classe responsável pela conexão e operações CRUD no banco MySQL.
    """
    def __init__(self, host='localhost', user='root', password='', database='monitor_ociosidade'):
        self.conn = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database,
            autocommit=True
        )
        self.cursor = self.conn.cursor()
        # Garante existência da tabela
        self._criar_tabela()

    def _criar_tabela(self):
        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS ociosidade_log (
            id INT PRIMARY KEY AUTO_INCREMENT,
            inicio_ocioso DATETIME,
            fim_ocioso DATETIME,
            duracao_segundos INT,
            usuario VARCHAR(64),
            host VARCHAR(64)
        )
        """)

    def inserir_ociosidade(self, inicio_ocioso, fim_ocioso, duracao, usuario, host):
        sql = """
        INSERT INTO ociosidade_log (inicio_ocioso, fim_ocioso, duracao_segundos, usuario, host)
        VALUES (%s, %s, %s, %s, %s)
        """
        self.cursor.execute(sql, (inicio_ocioso, fim_ocioso, duracao, usuario, host))

    def buscar_logs(self, dt_inicio=None, dt_fim=None):
        sql = "SELECT id, inicio_ocioso, fim_ocioso, duracao_segundos, usuario, host FROM ociosidade_log"
        conds = []
        vals = []
        if dt_inicio:
            conds.append("inicio_ocioso >= %s")
            vals.append(dt_inicio)
        if dt_fim:
            conds.append("fim_ocioso <= %s")
            vals.append(dt_fim)
        if conds:
            sql += " WHERE " + " AND ".join(conds)
        sql += " ORDER BY inicio_ocioso DESC"
        self.cursor.execute(sql, tuple(vals))
        return self.cursor.fetchall()

    def fechar(self):
        if self.cursor: self.cursor.close()
        if self.conn: self.conn.close()

class OciosidadeMonitor(threading.Thread):
    """
    Thread principal responsável por monitorar eventos de teclado e mouse.
    Detecta início e fim da ociosidade e salva eventos no banco de dados.
    """
    def __init__(self, tempo_ocioso, callback_on_evento_ocioso):
        super().__init__()
        self.tempo_ocioso = tempo_ocioso
        self.callback_on_evento_ocioso = callback_on_evento_ocioso
        self._stop_event = threading.Event()
        self._resetar_timer_atividade()
        self.db = DBHelper()
        self.usuario = os.getlogin() if hasattr(os, 'getlogin') else 'unknown'
        self.host = socket.gethostname()
        self.ocioso_ativo = False
        self.inicio_ocioso = None

    def run(self):
        # Inicia listeners de teclado e mouse em threads separadas
        listener_tec = keyboard.Listener(on_press=self.on_input_event)
        listener_mou = mouse.Listener(on_move=self.on_input_event,
                                      on_click=self.on_input_event,
                                      on_scroll=self.on_input_event)
        listener_tec.start()
        listener_mou.start()

        try:
            while not self._stop_event.is_set():
                agora = time.time()
                if not self.ocioso_ativo and (agora - self._ultimo_evento) > self.tempo_ocioso:
                    # Iniciando novo período de ociosidade
                    self.inicio_ocioso = datetime.datetime.now()
                    self.ocioso_ativo = True
                elif self.ocioso_ativo and (agora - self._ultimo_evento) <= self.tempo_ocioso:
                    # Terminando o período de ociosidade
                    fim_ocioso = datetime.datetime.now()
                    duracao = (fim_ocioso - self.inicio_ocioso).total_seconds()
                    # Evento de ociosidade detectado, executa callback e salva no banco
                    self.callback_on_evento_ocioso(self.inicio_ocioso, fim_ocioso, int(duracao))
                    self.db.inserir_ociosidade(self.inicio_ocioso, fim_ocioso, int(duracao), self.usuario, self.host)
                    self.ocioso_ativo = False
                time.sleep(1)
        finally:
            listener_tec.stop()
            listener_mou.stop()
            self.db.fechar()

    def on_input_event(self, *args, **kwargs):
        # É disparado por qualquer evento de teclado ou mouse
        self._resetar_timer_atividade()

    def _resetar_timer_atividade(self):
        self._ultimo_evento = time.time()

    def stop(self):
        self._stop_event.set()

class AppOciosidade(tk.Tk):
    """
    Classe principal da interface gráfica. Monta o dashboard, integra monitoramento e relatórios.
    """
    def __init__(self):
        super().__init__()
        self.title("Monitor de Ociosidade do Usuário")
        self.geometry("900x600")
        self.resizable(True, True)
        self.iconbitmap(default='') # Set icone aqui se desejar.

        # Parâmetros
        self.tempo_ocioso = tk.IntVar(value=5*60)  # tempo ocioso padrão: 5min
        self.db = DBHelper()
        self.monitor_thread = None
        self._montar_gui()
        self.protocol("WM_DELETE_WINDOW", self.fechar)

        # Inicia monitoramento em background
        self.iniciar_monitoramento()

    def _montar_gui(self):
        # Frame superior de controles
        frm_top = ttk.LabelFrame(self, text="Configuração e Controle")
        frm_top.pack(fill='x', padx=10, pady=5)

        ttk.Label(frm_top, text="Tempo de Ociosidade (segundos): ").pack(side='left', padx=(8,2))
        spin = ttk.Spinbox(frm_top, from_=30, to=60*60, increment=30, textvariable=self.tempo_ocioso, width=8)
        spin.pack(side='left')

        btn_reload = ttk.Button(frm_top, text="Atualizar Relatório", command=self.carregar_dados)
        btn_reload.pack(side='left', padx=8)

        btn_graf_dia = ttk.Button(frm_top, text="Gráfico Ociosidade por Dia", command=self.gerar_grafico_dia)
        btn_graf_dia.pack(side='left', padx=8)

        btn_graf_hora = ttk.Button(frm_top, text="Gráfico Ociosidade por Hora", command=self.gerar_grafico_hora)
        btn_graf_hora.pack(side='left', padx=8)

        btn_export = ttk.Button(frm_top, text="Exportar para Excel", command=self.exportar_excel)
        btn_export.pack(side='left', padx=8)

        # Frame meio/tabela de logs
        frm_table = ttk.Frame(self)
        frm_table.pack(fill='both', expand=1, padx=10, pady=5)

        colunas = ('ID', 'Início Ocioso', 'Fim Ocioso', 'Duração (s)', 'Usuário', 'Host')
        self.tree = ttk.Treeview(frm_table, columns=colunas, show='headings', selectmode='extended')
        for col in colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center', width=120)
        self.tree.pack(side='left', fill='both', expand=1)

        # Scrollbar para a treeview
        scroll = ttk.Scrollbar(frm_table, orient='vertical', command=self.tree.yview)
        self.tree.config(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        # Status bar
        self.status_var = tk.StringVar()
        status = ttk.Label(self, textvariable=self.status_var, relief='sunken', anchor='w')
        status.pack(fill='x', side='bottom')
        self._mensagem_status("Sistema inicializado.")
        self.carregar_dados()

    def _mensagem_status(self, msg):
        self.status_var.set(msg)

    def iniciar_monitoramento(self):
        # (Re)inicia a thread do monitorador
        if hasattr(self, 'monitor_thread') and self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.stop()
            self.monitor_thread.join(1)
        self.monitor_thread = OciosidadeMonitor(
            self.tempo_ocioso.get(), self.on_novo_ocioso)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        self._mensagem_status(f"Monitoramento iniciado (timeout: {self.tempo_ocioso.get()}s).")

    def on_novo_ocioso(self, inicio, fim, duracao):
        # Callback invocado toda vez que um novo período de ociosidade é detectado
        self._mensagem_status(
            f"Ociosidade detectada: início={inicio}, fim={fim}, duração={int(duracao)}s")
        self.carregar_dados(async_load=True)

    def carregar_dados(self, async_load=False):
        # Carrega e exibe os dados da tabela no widget
        def load():
            for i in self.tree.get_children():
                self.tree.delete(i)
            rows = self.db.buscar_logs()
            for row in rows:
                self.tree.insert('', 'end', values=row)
            self._mensagem_status(f"Registros carregados: {len(rows)}")
        if async_load:
            threading.Thread(target=load).start()
        else:
            load()

    def gerar_grafico_dia(self):
        # Gera gráfico de ociosidade agregada por dia
        rows = self.db.buscar_logs()
        if not rows:
            messagebox.showinfo("Informação", "Sem dados para gerar gráfico.")
            return
        df = pd.DataFrame(rows, columns=['id','inicio','fim','duracao','usuario','host'])
        df['dia'] = pd.to_datetime(df['inicio']).dt.date
        grupo = df.groupby('dia')['duracao'].sum().reset_index()
        plt.figure(figsize=(8,4))
        plt.bar(grupo['dia'].astype(str), grupo['duracao']/60)
        plt.title('Duração Total de Ociosidade por Dia (em minutos)')
        plt.xlabel('Data')
        plt.ylabel('Ociosidade (min)')
        plt.xticks(rotation=30)
        plt.tight_layout()
        plt.show()

    def gerar_grafico_hora(self):
        # Gera gráfico de ociosidade agregada por hora do dia
        rows = self.db.buscar_logs()
        if not rows:
            messagebox.showinfo("Informação", "Sem dados para gerar gráfico.")
            return
        df = pd.DataFrame(rows, columns=['id','inicio','fim','duracao','usuario','host'])
        df['hora'] = pd.to_datetime(df['inicio']).dt.hour
        grupo = df.groupby('hora')['duracao'].sum().reset_index()
        plt.figure(figsize=(8,4))
        plt.bar(grupo['hora'].astype(str), grupo['duracao']/60)
        plt.title('Duração Total de Ociosidade por Hora do Dia (em minutos)')
        plt.xlabel('Hora')
        plt.ylabel('Ociosidade (min)')
        plt.tight_layout()
        plt.show()

    def exportar_excel(self):
        # Exporta os registros exibidos na tabela para arquivo Excel
        rows = self.db.buscar_logs()
        if not rows:
            messagebox.showinfo("Informação", "Sem dados para exportar.")
            return
        df = pd.DataFrame(rows, columns=['ID','Início Ocioso','Fim Ocioso','Duração (s)','Usuário','Host'])
        filename = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel Files', '*.xlsx')],
            title='Salvar como')
        if filename:
            try:
                df.to_excel(filename, index=False)
                messagebox.showinfo("Exportação completa", f"Dados exportados para {filename}")
            except Exception as e:
                messagebox.showerror("Erro ao exportar", str(e))
        self._mensagem_status("Dados exportados para Excel.")

    def fechar(self):
        # Finaliza monitoramento e fecha a aplicação
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.stop()
        self.db.fechar()
        self.destroy()

if __name__ == '__main__':
    app = AppOciosidade()
    app.mainloop()