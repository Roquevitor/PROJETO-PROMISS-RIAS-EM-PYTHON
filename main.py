from tkinter import *
from tkinter import ttk
from num2words import num2words
from reportlab.pdfgen import canvas
import datetime
from tkcalendar import DateEntry
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from tkinter import messagebox
from dateutil.relativedelta import relativedelta
import sys
import os
import locale
from tkinter import filedialog

pdfmetrics.registerFont(TTFont("Arial", "C:\\Windows\\Fonts\\arial.ttf"))
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    print("Locale pt_BR.UTF-8 não está disponível. Verifique as configurações do sistema.")



if getattr(sys, 'frozen', False): 
   
    pastaApp = os.path.dirname(sys.executable)
else:
     
    pastaApp = os.path.dirname(os.path.abspath(__file__))



class PDFGenerator:
    def __init__(self, pasta_app):
        self.pasta_app = pasta_app

    def remover_formatacao_reais(self, valor_str):
        """Remove 'R$', pontos de milhar e converte para float."""
        try:
            valor_str = valor_str.replace("R$", "").replace(".", "").replace(",", ".").strip()
            return float(valor_str)
        except ValueError:
            return 0.0
        
    def caminho_arquivo(self, nome_arquivo):
        if getattr(sys, 'frozen', False):  
            base_path = sys._MEIPASS
        else: 
            base_path = os.path.dirname(__file__)
        return os.path.join(base_path, nome_arquivo)

    
        
    def formatar_para_reais(event):
        """Formata o conteúdo do campo Entry para o formato brasileiro de moeda."""
        entry = event.widget
        valor = entry.get()

        try:
            
            valor = valor.replace(".", "").replace(",", ".")
            
            valor_float = float(valor)
            
            valor_formatado = f"R${valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            entry.delete(0, tk.END)
            entry.insert(0, valor_formatado)
        except ValueError:
            
            entry.delete(0, tk.END)

    def formatar_cpf_cnpj(event):
        """Formata o campo como CPF ou CNPJ baseado no tamanho do valor inserido."""
        entry = event.widget
        valor = entry.get()

        
        valor = valor.replace(".", "").replace("-", "").replace("/", "")

        if len(valor) == 11: 
            
            valor_formatado = f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
        elif len(valor) == 14:  
            
            valor_formatado = f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
        else:
            valor_formatado = valor 

        
        entry.delete(0, tk.END)
        entry.insert(0, valor_formatado)

    def formatar_para_maiusculas(event):
        """Formata o texto digitado no campo para letras maiúsculas."""
        entry = event.widget
        texto = entry.get()
        texto_maiusculo = texto.upper()
        entry.delete(0, tk.END)
        entry.insert(0, texto_maiusculo)
                
    def desenhar_texto_com_quebra(self, cnv, text, x_start, y_start, max_width):
        cnv.setFont("Times-Roman", 9)
        self.line_height = 15
        self.words = text.split()
        self.x = x_start
        self.y = y_start
        for word in self.words:
            self.word_width = cnv.stringWidth(word + " ", "Times-Roman", 9)
            if self.x + self.word_width > max_width:
                self.x = 153
                self.y -= self.line_height
            cnv.drawString(self.x, self.y, word)
            self.x += self.word_width

    def set_data(self, data):
        """Define os dados necessários para o PDF."""
        (
            self.divida_total, self.valor_pago, self.valor_parcela, self.data_selecionada,
            self.credor, self.cnpj_cpf_credor, self.emitente, self.cpf_emitente, self.endereco_emitente
        ) = data


    def criar_pdf(self, data):
        """Gera o arquivo PDF com base nos dados fornecidos."""
        try:
            pasta_promissoria = filedialog.askdirectory(title="Selecione a pasta de promissórias")
            
            
            if not os.path.exists(pasta_promissoria):
                os.makedirs(pasta_promissoria)

            
            

            
            (
                self.divida_total, self.valor_pago, self.valor_parcela, data_selecionada,
                self.credor, self.cnpj_cpf_credor, self.emitente, self.cpf_emitente, self.endereco_emitente
            ) = data

            # Cálculos
            divida_restante = self.divida_total - self.valor_pago
            parcelas_completas = divida_restante // self.valor_parcela
            residual = divida_restante % self.valor_parcela
            total_parcelas = int(parcelas_completas + (1 if residual > 0 else 0))

            caminho_pdf = os.path.join(pasta_promissoria, f"{self.emitente}.pdf")
            cnv = canvas.Canvas(caminho_pdf, pagesize=A4)

            
    
            n_t = 0
            data_obj = datetime.datetime.strptime(data_selecionada, "%d/%m/%Y")
            text_y = 740
            text_y = 1017
            image_y = 635 
            text_y_2 = 986
            text_y_3 = 960
            text_y_4 = 940
            text_y_5 = 911
            text_y_6 = 904
            text_y_7 = 888
            text_y_8 = 858
            print(total_parcelas)

            for i in range(total_parcelas + 1):
                


                if image_y < 2: 
                    cnv.showPage()
                    text_y = 740
                    text_y = 1017
                    image_y = 635 
                    text_y_2 = 986
                    text_y_3 = 960
                    text_y_4 = 940
                    text_y_5 = 911
                    text_y_6 = 904
                    text_y_7 = 888
                    text_y_8 = 858
                data_atual = data_obj + relativedelta(months=i)
                valor_parcela_atual = residual if i == total_parcelas - 1 and residual > 0 else self.valor_parcela
                data_obj = datetime.datetime.strptime(data_selecionada, "%d/%m/%Y")
                n_t += 1
                if i == total_parcelas:

                
                    cnv.drawImage(os.path.join(pastaApp, self.caminho_arquivo(r"C:\img\teste_1.jpg")), 20, image_y, width=540, height=200)
                    image_y -= 210
                    text_y -= 210
                    text_y_2 -= 210
                    text_y_4 -= 210
                    text_y_3 -=210
                    text_y_5 -= 210
                    text_y_6 -=210
                    text_y_7 -=210
                    text_y_8 -=210

                    
                    
                    agora = datetime.datetime.now()
                    

                    dia_h = agora.strftime("%d")
                    mes_h = agora.strftime("%B")
                    ano_h = agora.strftime("%Y")
                    
                    data_atual = data_obj
                    dia = data_atual.strftime("%d")
                    mes = data_atual.strftime("%B")
                    ano = data_atual.strftime("%Y")
                    
                    
                
                    
                    
                    cnv.setFont("Times-Bold", 10)
                    cnv.drawString(190, text_y, f"{n_t}/{total_parcelas}")
                    cnv.drawString(480, text_y, f"{self.valor_pago:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    cnv.setFont("Times-Roman", 10)
                    cnv.drawString(258, text_y, f"         {dia_h}                                           {ano_h}".encode('latin1').decode('utf-8'))
                    cnv.drawString(260, text_y, f"                       {mes_h}               ".encode('latin1').decode('utf-8').upper())
                    cnv.setFont("Times-Roman", 9)
                    cnv.drawString(163, text_y_3, f"{self.credor}")
                    cnv.drawString(465, text_y_3, f"{self.cnpj_cpf_credor}")
                    divida_total = str(num2words(self.valor_pago, lang='pt'))
                    cnv.drawString(285, text_y_4, f"{divida_total} REAIS".encode('latin1').decode('utf-8').upper())
                    cnv.setFont("Times-Bold", 9)
                    cnv.drawString(198, text_y_6, f"{self.emitente}")
                    cnv.setFont("Times-Roman", 9)
                    mes_h = agora.strftime("%m")
                
                    cnv.drawString(481, text_y_6, f"     {dia_h}   {mes_h}   {ano_h}")
                    cnv.drawString(197, text_y_7, f"{self.cpf_emitente}")
                    texto = self.endereco_emitente
                    PDFGenerator.desenhar_texto_com_quebra(self, cnv, text=texto, x_start=385, y_start=text_y_7-3, max_width=555)
                else:
                
                    cnv.drawImage(os.path.join(pastaApp, self.caminho_arquivo(r"C:\img\teste_1.jpg")), 20, image_y, width=540, height=200)
                    image_y -= 210  
                    text_y -= 210
                    text_y_2 -= 210
                    text_y_4 -= 210
                    text_y_3 -=210
                    text_y_5 -= 210
                    text_y_6 -=210
                    text_y_7 -=210
                    text_y_8 -=210

                    
                    agora = datetime.datetime.now()
                    dia_h = agora.strftime("%d")
                    mes_h = agora.strftime("%m")
                    ano_h = agora.strftime("%Y")
                    

                    data_atual = data_obj + relativedelta(months=i)
                    dia = data_atual.strftime("%d")
                    mes = data_atual.strftime("%B")
                    ano = data_atual.strftime("%Y")
                    
                    
                
                    
                    
                    cnv.setFont("Times-Bold", 10)
                    cnv.drawString(190, text_y, f"{n_t}/{total_parcelas}")
                    cnv.drawString(480, text_y, f"{valor_parcela_atual:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    cnv.setFont("Times-Roman", 10)
                    cnv.drawString(258, text_y, f"         {dia}                                           {ano}".encode('latin1').decode('utf-8'))
                    cnv.drawString(260, text_y, f"                        {mes}               ".encode('latin1').decode('utf-8').upper())
                    cnv.setFont("Times-Roman", 9)
                    cnv.drawString(163, text_y_3, f"{self.credor}")
                    cnv.drawString(465, text_y_3, f"{self.cnpj_cpf_credor}")
                    valor_parcela_extenso = str(num2words(valor_parcela_atual, lang='pt'))
                    cnv.drawString(285, text_y_4, f"{valor_parcela_extenso} REAIS".encode('latin1').decode('utf-8').upper())
                    cnv.setFont("Times-Bold", 9)
                    cnv.drawString(198, text_y_6, f"{self.emitente}")
                    cnv.setFont("Times-Roman", 9)
                    mes_1 = data_obj.strftime("%m")
                    dia_1 = data_obj.strftime("%d")
                    ano_1 = data_obj.strftime("%Y")
                    cnv.drawString(481, text_y_6, f"     {dia_h}   {mes_h}   {ano_h}")
                    cnv.drawString(197, text_y_7, f"{self.cpf_emitente}")
                    texto = self.endereco_emitente
                    PDFGenerator.desenhar_texto_com_quebra(self, cnv, text=texto, x_start=385, y_start=text_y_7-3, max_width=555)
                
            cnv.save()
            messagebox.showinfo("PDF Gerado", f"PDF salvo em: {caminho_pdf}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF: {e}")




co = "#D9D9D9"



import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry

import customtkinter


COR_FUNDO = "#7E7979" 
COR_BORDA = "#CCCCCC"  



class App:

    def __init__(self, root):

        self.root = root
        self.root.title("PROMISSÓRIAS")
        self.root.geometry("900x520")
        self.root.resizable(False, False)
        self.root.iconbitmap(r"C:\img\icon.ico")
        self.pdf_generator = PDFGenerator(pastaApp)
        self.setup_ui()

    def setup_ui(self):
        """Configura os elementos da interface."""

  
        customtkinter.set_appearance_mode("System")
        customtkinter.set_default_color_theme("dark-blue")




        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

 
        frame =  customtkinter.CTkFrame(self.root, width=905, height=530, fg_color="#CDCBCB")
        frame.place(x=0, y=0)
        frame1 =  customtkinter.CTkFrame(self.root, width=230, fg_color="#E6E6E6", corner_radius=8, bg_color="#CDCBCB", border_width=1, border_color="#CDCBCB")
        frame1.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        frame2 =  customtkinter.CTkFrame(self.root, width=230, fg_color="#E6E6E6", corner_radius=8, bg_color="#CDCBCB", border_width=1, border_color="#CDCBCB")
        frame2.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        frame3 =  customtkinter.CTkFrame(self.root, width=230, fg_color="#E6E6E6", corner_radius=8, bg_color="#CDCBCB", border_width=1, border_color="#CDCBCB")
        frame3.grid(row=1, column=0, sticky="nsew", padx=10, pady=20)
        frame4 =  customtkinter.CTkFrame(self.root, width=230, fg_color="#E6E6E6", corner_radius=8, bg_color="#CDCBCB", border_width=1, border_color="#CDCBCB")
        frame4.grid(row=1, column=1, sticky="nsew", padx=10, pady=20)

       
        label_cab = customtkinter.CTkLabel(
            frame1, 
            text="INFORMAÇÕES DA PROMISSÓRIA", 
            font=("Helvetica", 12, "bold"),
            fg_color="#444444",
            bg_color="#E6E6E6",
            width=35,
            pady=5,
            corner_radius=4
        )
        label_cab.pack(fill="x")

 
        container_frame1 = customtkinter.CTkFrame(frame1, fg_color="#E6E6E6", corner_radius=8)
        container_frame1.pack(padx=10, pady=10, fill="both", expand=True)

        label_divida = customtkinter.CTkLabel(container_frame1, text="VALOR TOTAL:", text_color="#3E3D3D",font=("Arial", 12, "bold"), fg_color="#E6E6E6")
        label_divida.grid(row=0, column=0, sticky="w", pady=10, padx=2)

        self.input_divida = customtkinter.CTkEntry(container_frame1, width=180, font=("Helvetica", 10, "bold"), fg_color="white", text_color="black", placeholder_text="R$", placeholder_text_color="#B5ACAC", border_width=1)
        self.input_divida.grid(row=0, column=0, pady=10, padx=105)


        label_pago = customtkinter.CTkLabel(container_frame1, text="VALOR PAGO:", font=("Arial", 12, "bold"), text_color="#3E3D3D", fg_color="#E6E6E6")
        label_pago.grid(row=1, column=0, sticky="w", pady=5, padx=2)

        self.input_pago = customtkinter.CTkEntry(container_frame1, width=180, font=("Helvetica", 10, "bold"), fg_color="white", text_color="black", placeholder_text="R$", placeholder_text_color="#B5ACAC", border_width=1,)
        self.input_pago.grid(row=1, column=0, pady=10, padx=105)

        label_valor_parcela = customtkinter.CTkLabel(container_frame1, text="PARCELAS:", font=("Arial", 12, "bold"), text_color="#3E3D3D", fg_color="#E6E6E6")
        label_valor_parcela.grid(row=2, column=0, sticky="w", pady=5, padx=2)

        self.input_valor_parcela = customtkinter.CTkEntry(container_frame1, width=180, font=("Helvetica", 10, "bold"), fg_color="white", text_color="black", placeholder_text="R$", placeholder_text_color="#B5ACAC", border_width=1)
        self.input_valor_parcela.grid(row=2, column=0, pady=10, padx=105)


        self.input_divida.bind("<FocusOut>", PDFGenerator.formatar_para_reais)
        self.input_pago.bind("<FocusOut>", PDFGenerator.formatar_para_reais)
        self.input_valor_parcela.bind("<FocusOut>", PDFGenerator.formatar_para_reais)

        label_data = customtkinter.CTkLabel(container_frame1, text="DATA VENCIMENTO:", font=("Arial", 10, "bold"), text_color="#3E3D3D", fg_color="#E6E6E6")
        label_data.grid(row=3, column=0, sticky="w", pady=10, padx=2)
        

        self.entry_data = DateEntry(
        container_frame1, width=25, background='darkblue',
        foreground='white', borderwidth=2, date_pattern="dd/MM/yyyy"
    )
        self.entry_data.grid(row=3, column=0, pady=10, padx=100)

   

        label_cab1 = customtkinter.CTkLabel(
        frame2, 
        text="INFORMAÇÕES DO EMITENTE", 
        font=("Helvetica", 12, "bold"),
        fg_color="#444444",
        bg_color="#E6E6E6",
        width=35,
        pady=5,
        corner_radius=4
    )
        label_cab1.pack(fill="x")

        container_frame2 = customtkinter.CTkFrame(frame2, fg_color="#E6E6E6", corner_radius=8)
        container_frame2.pack(padx=10, pady=20, fill="both", expand=True)

        label_emitente = customtkinter.CTkLabel(container_frame2, text="EMITENTE:",  font=("Arial", 12, "bold"), bg_color="#E6E6E6", text_color="#3E3D3D")
        label_emitente.grid(row=0, column=0, sticky="w", pady=10, padx=2)

        self.input_emitente = customtkinter.CTkEntry(container_frame2, width=280, font=("Helvetica", 10, "bold"), placeholder_text="Digite o nome do emitente", placeholder_text_color="#B5ACAC", border_width=1,  fg_color="white", text_color="black")
        self.input_emitente.grid(row=0, column=0, pady=10, padx=95)
        self.input_emitente.bind("<KeyRelease>", PDFGenerator.formatar_para_maiusculas)

        label_cpf_emitente = customtkinter.CTkLabel(container_frame2, text="CPF/CNPJ:",  font=("Arial", 12, "bold"), bg_color="#E6E6E6", text_color="#3E3D3D")
        label_cpf_emitente.grid(row=1, column=0, sticky="w", pady=5, padx=2)

        self.input_cpf_emitente = customtkinter.CTkEntry(container_frame2, width=280, font=("Helvetica", 10, "bold"), placeholder_text="000.000.000-00", placeholder_text_color="#B5ACAC", border_width=1,  fg_color="white", text_color="black")
        self.input_cpf_emitente.grid(row=1, column=0, pady=10, padx=95)
        
       

        label_endereco_emitente = customtkinter.CTkLabel(container_frame2, text="ENDEREÇO:",  font=("Arial", 12, "bold"), width=20,bg_color="#E6E6E6", text_color="#3E3D3D")
        label_endereco_emitente.grid(row=2, column=0, sticky="w", pady=5, padx=2)


        self.input_endereco_emitente = customtkinter.CTkEntry(container_frame2, width=280, font=("Helvetica", 10, "bold"), placeholder_text="Digite o endereço do emitente",  border_width=1, placeholder_text_color="#B5ACAC", fg_color="white", text_color="black")
        self.input_endereco_emitente.grid(row=2, column=0, pady=10, padx=95)
        self.input_endereco_emitente.bind("<KeyRelease>", PDFGenerator.formatar_para_maiusculas)




        label_cab2 = customtkinter.CTkLabel(
            frame3, 
            text="INFORMAÇÕES DO CREDOR", 
            font=("Helvetica", 12, "bold"),
            bg_color="#E6E6E6",
            fg_color="#444444",
            width=35,
            pady=5,
            corner_radius=4
        )
        label_cab2.pack(fill="x")


        container_frame3 = customtkinter.CTkFrame(frame3, fg_color="#E6E6E6", corner_radius=8)
        container_frame3.pack(padx=10, pady=20, fill="both", expand=True)

        label_credor = customtkinter.CTkLabel(container_frame3, text="CREDOR:", font=("Arial", 12, "bold"), bg_color="#E6E6E6", text_color="#3E3D3D")
        label_credor.grid(row=0, column=0, sticky="w", pady=15, padx=2)

        self.entry_credor = customtkinter.CTkEntry(container_frame3, width=280, font=("Helvetica", 10, "bold"), placeholder_text="Digite o nome do credor", placeholder_text_color="#B5ACAC", border_width=1, fg_color="white", text_color="black")
        self.entry_credor.grid(row=0, column=0, pady=15, padx=85)
        self.entry_credor.bind("<KeyRelease>", PDFGenerator.formatar_para_maiusculas)


        label_credor_cpf_cnpj = customtkinter.CTkLabel(container_frame3, text="CPF/CNPJ:",font=("Arial", 12, "bold"), bg_color="#E6E6E6", text_color="#3E3D3D")
        label_credor_cpf_cnpj.grid(row=1, column=0, sticky="w", pady=15, padx=2)


        self.entry_credor_cpf_cnpj = customtkinter.CTkEntry(container_frame3, width=280, font=("Helvetica", 10, "bold"), placeholder_text="000.000.000-00", placeholder_text_color="#B5ACAC", border_width=1, fg_color="white", text_color="black")
        self.entry_credor_cpf_cnpj.grid(row=1, column=0, pady=15, padx=85)
        self.input_cpf_emitente.bind("<FocusOut>", PDFGenerator.formatar_cpf_cnpj)
        self.entry_credor_cpf_cnpj.bind("<FocusOut>", PDFGenerator.formatar_cpf_cnpj)




 
        label_cab3 = customtkinter.CTkLabel(
            frame4, 
            text="AÇÕES", 
            font=("Helvetica", 12, "bold"),
            bg_color="#E6E6E6",
            fg_color="#444444",
            width=35,
            pady=5,
            corner_radius=4
        )
        label_cab3.pack(fill="x")

        container_frame4 = customtkinter.CTkFrame(frame4, corner_radius=8, fg_color="#E6E6E6")
        container_frame4.pack(padx=10, pady=20, fill="both", expand=True)

        btn_criarPdf = customtkinter.CTkButton(container_frame4, width=280, text="Gerar Promissórias", font=("Arial", 12, "bold"), command=self.gerar_pdf, fg_color="green", hover_color="#2A5111" )
        btn_criarPdf.grid(row=0, column=0, pady=10, padx=75, sticky="w")

        btn_limpar_campos = customtkinter.CTkButton(container_frame4, width=280, text="Limpar Campos", font=("Arial", 12, "bold"), command=self.limpar_campos, fg_color="#ECB10E", hover_color="#A17804")
        btn_limpar_campos.grid(row=1, column=0, pady=5, padx=75, sticky="w")


        btn_sair = customtkinter.CTkButton(container_frame4, width=280, text="Sair", font=("Arial", 12, "bold"), command=root.quit, fg_color="#E80C0C", hover_color="#7B0D0D")
        btn_sair.grid(row=2, column=0, pady=5, padx=75, sticky="w")
        
        

        


        rodape = customtkinter.CTkLabel(
        self.root,
        text="Application developed by Roque Vitor — All rights reserved, Version 1.1 • Instagram: @vitor.goncalvess1",
        font=("Arial", 10, "italic", "bold"),
        fg_color="#CDCBCB", 
        bg_color="#CDCBCB",
        text_color="#3E3D3D",
        compound="left",
        pady = -100,
        padx =195,
        height= 5,
        corner_radius=4
    )
        rodape.place(x=1, y=505) 

    def limpar_campos(self):
       
        self.input_divida.delete(0, "end")
        self.input_pago.delete(0, "end")
        self.input_valor_parcela.delete(0, "end")
        self.entry_credor.delete(0, "end")
        self.entry_credor_cpf_cnpj.delete(0, "end")
        self.input_emitente.delete(0, "end")
        self.input_cpf_emitente.delete(0, "end")
        self.input_endereco_emitente.delete(0, "end")

    def gerar_pdf(self):
        """Obtém os dados do formulário e gera o PDF."""
        try:
            
            pdf_generator = PDFGenerator(pastaApp)
            
           
            self.divida_total = self.pdf_generator.remover_formatacao_reais(self.input_divida.get())
            self.valor_pago = self.pdf_generator.remover_formatacao_reais(self.input_pago.get())
            self.valor_parcela = self.pdf_generator.remover_formatacao_reais(self.input_valor_parcela.get())
            self.credor = self.entry_credor.get()
            self.cnpj_cpf_credor = self.entry_credor_cpf_cnpj.get()
            self.cpf_emitente = self.input_cpf_emitente.get()
            self.emitente = self.input_emitente.get()
            self.endereco_emitente = self.input_endereco_emitente.get()
            self.data_selecionada = self.entry_data.get()

            
    

            
            data = (
                self.divida_total, self.valor_pago, self.valor_parcela, self.data_selecionada,
                self.credor, self.cnpj_cpf_credor, self.emitente, self.cpf_emitente, self.endereco_emitente
            )

           
            pdf_generator.criar_pdf(data)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar dados: {e}")
            print(e)





if __name__ == "__main__":
    root = Tk()
    app = App(root)
    root.mainloop()