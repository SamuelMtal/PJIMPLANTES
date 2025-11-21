# -*- coding: utf-8 -*-
"""
Monitoramento de Ociosidade do Usuário com Interface Tkinter, registro em MySQL e gráficos/Excel
Agora com contador de ociosidade em tempo real.

Necessário:
pip install pynput mysql-connector-python pandas matplotlib openpyxl
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
    def __init__(self, host='localhost', user='root', password='', database='monitor_ociosidade'):
        self.conn = mysql.connector.connect(
            host=host, user=user, password=password, database=database, autocommit=True
        )
        self.cursor = self.conn.cursor()
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
        listener_tec = keyboard.Listener(on_press=self.on_input_event)
        listener_mou = mouse.Listener(
            on_move=self.on_input_event,
            on_click=self.on_input_event,
            on_scroll=self.on_input_event
        )
        listener_tec.start()
        listener_mou.start()

        try:
            while not self._stop_event.is_set():
                agora = time.time()

                # Início da ociosidade
                if not self.ocioso_ativo and (agora - self._ultimo_evento) > self.tempo_ocioso:
                    self.inicio_ocioso = datetime.datetime.now()
                    self.ocioso_ativo = True

                # Fim da ociosidade
                elif self.ocioso_ativo and (agora - self._ultimo_evento) <= self.tempo_ocioso:
                    fim_ocioso = datetime.datetime.now()
                    duracao = (fim_ocioso - self.inicio_ocioso).total_seconds()

                    self.callback_on_evento_ocioso(self.inicio_ocioso, fim_ocioso, int(duracao))
                    self.db.inserir_ociosidade(
                        self.inicio_ocioso, fim_ocioso, int(duracao),
                        self.usuario, self.host
                    )
                    self.ocioso_ativo = False

                time.sleep(1)
        finally:
            listener_tec.stop()
            listener_mou.stop()
            self.db.fechar()

    def on_input_event(self, *args, **kwargs):
        self._resetar_timer_atividade()

    def _resetar_timer_atividade(self):
        self._ultimo_evento = time.time()

    def stop(self):
        self._stop_event.set()


class AppOciosidade(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Monitor de Ociosidade do Usuário")
        self.geometry("900x650")
        self.resizable(True, True)

        self.tempo_ocioso = tk.IntVar(value=30)
        self.ocioso_atual = tk.StringVar(value="Ociosidade atual: 0s")

        self.db = DBHelper()
        self.monitor_thread = None

        self._montar_gui()
        self.protocol("WM_DELETE_WINDOW", self.fechar)

        # Inicia monitoramento
        self.iniciar_monitoramento()

        # Inicia contador de ociosidade em tempo real
        self.after(1000, self.atualizar_contador_ociosidade)

    def _montar_gui(self):
        frm_top = ttk.LabelFrame(self, text="Configuração e Controle")
        frm_top.pack(fill='x', padx=10, pady=5)

        ttk.Label(frm_top, text="Tempo Ocioso (segundos): ").pack(side='left', padx=(8, 2))
        spin = ttk.Spinbox(frm_top, from_=30, to=3600, increment=30,
                           textvariable=self.tempo_ocioso, width=8)
        spin.pack(side='left')

        ttk.Button(frm_top, text="Atualizar Relatório",
                   command=self.carregar_dados).pack(side='left', padx=8)

        ttk.Button(frm_top, text="Gráfico por Dia",
                   command=self.gerar_grafico_dia).pack(side='left', padx=8)

        ttk.Button(frm_top, text="Gráfico por Hora",
                   command=self.gerar_grafico_hora).pack(side='left', padx=8)

        ttk.Button(frm_top, text="Exportar Excel",
                   command=self.exportar_excel).pack(side='left', padx=8)

        # Contador de ociosidade
        lbl_ocioso = ttk.Label(self, textvariable=self.ocioso_atual, anchor='w')
        lbl_ocioso.pack(fill='x')

        frm_table = ttk.Frame(self)
        frm_table.pack(fill='both', expand=1, padx=10, pady=5)

        colunas = (
            'ID', 'Início da ociosidade', 'Fim da ociosidade',
            'Duração (s)', 'Usuário', 'Nome do Computador'
        )
        self.tree = ttk.Treeview(frm_table, columns=colunas, show='headings')
        for col in colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center', width=140)
        self.tree.pack(side='left', fill='both', expand=1)

        scroll = ttk.Scrollbar(frm_table, orient='vertical', command=self.tree.yview)
        self.tree.config(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')

        self.status_var = tk.StringVar()
        status = ttk.Label(self, textvariable=self.status_var, relief='sunken', anchor='w')
        status.pack(fill='x', side='bottom')

        self._mensagem_status("Sistema inicializado.")
        self.carregar_dados()

    def _mensagem_status(self, msg):
        self.status_var.set(msg)

    def iniciar_monitoramento(self):
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.stop()
            self.monitor_thread.join(1)

        self.monitor_thread = OciosidadeMonitor(
            self.tempo_ocioso.get(), self.on_novo_ocioso
        )
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        self._mensagem_status(f"Monitoramento iniciado (timeout: {self.tempo_ocioso.get()}s).")

    def on_novo_ocioso(self, inicio, fim, duracao):
        self._mensagem_status(
            f"Ociosidade detectada: início={inicio}, fim={fim}, duração={duracao}s"
        )
        self.carregar_dados(async_load=True)

    def carregar_dados(self, async_load=False):
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

    def atualizar_contador_ociosidade(self):
        if self.monitor_thread:
            agora = time.time()
            diff = agora - self.monitor_thread._ultimo_evento

            if diff < self.tempo_ocioso.get():
                self.ocioso_atual.set("Ociosidade atual: 0s")
            else:
                self.ocioso_atual.set(f"Ociosidade atual: {int(diff)}s")

        self.after(1000, self.atualizar_contador_ociosidade)

    def gerar_grafico_dia(self):
        rows = self.db.buscar_logs()
        if not rows:
            messagebox.showinfo("Informação", "Sem dados.")
            return

        df = pd.DataFrame(rows, columns=['id', 'inicio', 'fim', 'duracao', 'usuario', 'host'])
        df['dia'] = pd.to_datetime(df['inicio']).dt.date
        grupo = df.groupby('dia')['duracao'].sum().reset_index()

        plt.figure(figsize=(8, 4))
        plt.bar(grupo['dia'].astype(str), grupo['duracao'] / 60)
        plt.title('Ociosidade por Dia (minutos)')
        plt.xlabel('Data')
        plt.ylabel('Minutos ociosos')
        plt.xticks(rotation=30)
        plt.tight_layout()
        plt.show()

    def gerar_grafico_hora(self):
        rows = self.db.buscar_logs()
        if not rows:
            messagebox.showinfo("Informação", "Sem dados.")
            return

        df = pd.DataFrame(rows, columns=['id', 'inicio', 'fim', 'duracao', 'usuario', 'host'])
        df['hora'] = pd.to_datetime(df['inicio']).dt.hour
        grupo = df.groupby('hora')['duracao'].sum().reset_index()

        plt.figure(figsize=(8, 4))
        plt.bar(grupo['hora'], grupo['duracao'] / 60)
        plt.title('Ociosidade por Hora (minutos)')
        plt.xlabel('Hora')
        plt.ylabel('Minutos ociosos')
        plt.tight_layout()
        plt.show()

    def exportar_excel(self):
        rows = self.db.buscar_logs()
        if not rows:
            messagebox.showinfo("Informação", "Sem dados.")
            return

        df = pd.DataFrame(rows, columns=['ID', 'Início', 'Fim', 'Duração (s)', 'Usuário', 'Host'])

        filename = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel Files', '*.xlsx')],
            title='Salvar como'
        )

        if filename:
            try:
                df.to_excel(filename, index=False)
                messagebox.showinfo("Sucesso", "Arquivo exportado!")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

    def fechar(self):
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.stop()
        self.db.fechar()
        self.destroy()


if __name__ == '__main__':
    app = AppOciosidade()
    app.mainloop()
