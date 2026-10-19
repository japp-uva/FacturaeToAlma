# SPDX-License-Identifier: GPL-3.0-only
# FacturaeToAlma 4.0

import os
import sys
import json
import queue
import tempfile
import threading
import re
import copy
import html
import logging
import traceback
import unicodedata
import webbrowser
import subprocess  # nosec B404
import platform
import shutil
from dataclasses import dataclass, field
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk, UnidentifiedImageError


import defusedxml.ElementTree as ET
import defusedxml.minidom
from defusedxml.common import DefusedXmlException

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

APP_NAME = "FacturaeToAlma"
APP_VERSION = "4.0"
APP_TITLE = f"Universidad de Valladolid. Biblioteca -> {APP_NAME} {APP_VERSION}"
LOG_FILE = "facturae_alma.log"
PROJECT_REPOSITORY_URL = "https://github.com/japp-uva/FacturaeToAlma"
PROJECT_GUIDE_URL = "https://biblioguias.uva.es/facturae_to_alma"
UI_BG = "#f0f0f0"
TITLE_MATCH_MIN_SCORE = 0.78
TITLE_MATCH_MIN_MARGIN = 0.08
LINE_TYPE_OPTIONS = ["REGULAR", "OVERHEAD", "OTHER", "SHIPMENT", "DISCOUNT", "INSURANCE", "ADDITIONAL_CHARGES"]
EDITABLE_HEADER_FIELDS = {"currency", "vendor_account", "note", "inclusive"}
EDITABLE_LINE_FIELDS = {
    "line_type", "po_line", "vendor_ref", "issn", "isbn", "title",
    "reporting_code", "secondary_reporting_code", "tertiary_reporting_code",
    "fourth_reporting_code", "fifth_reporting_code", "start_date", "end_date",
    "note", "fund_1", "fund_2",
}
READ_ONLY_HEADER_FIELDS = {"invoice_number", "date", "total_amount", "vat_amount"}
READ_ONLY_LINE_FIELDS = {"invoice", "line", "price", "quantity", "vat_amount", "vat_percentage"}
DANGEROUS_EXCEL_PREFIXES = ("=", "+", "-", "@", "\t", "\r", "\n")
MAX_LOGO_HEIGHT = 100

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")

VENDOR_MAP = {"DEFAULT": "PROV_GENERICO"}

DEFAULT_HINV_HEADERS = [
    "HINV", "Invoice Number", "Date", "Total amount", "Currency", "Vendor Account", "Payment method", "Note",
    "Use Pro Rata", "Shipment amount", "Overhead amount", "Insurance amount", "Discount amount", "VAT Percentage",
    "VAT Amount", "Inclusive", "Expended From Fund", "PrePaid", "Payment status", "Voucher number", "Voucher date",
    "Voucher Amount", "Invoice Reference Number", "VAT In Invoice Line Level", "Report TAX"
]
DEFAULT_HIL_HEADERS = [
    "HIL", "Line type", "Line Number", "PO Line", "Vendor Ref Number", "ISSN", "ISBN", "Title", "Price", "Quantity",
    "Reporting Code", "Secondary Reporting Code", "Tertiary Reporting Code", "Fourth Reporting Code", "Fifth Reporting Code",
    "Start subs date", "End subs date", "Note", "Fund and percent", "Fund and percent", "VAT Amount", "VAT Percentage"
]

STOPWORDS = {
    "el", "la", "los", "las", "un", "una", "unos", "unas", "a", "ante", "bajo", "con", "contra", "de", "desde", "en", "entre", "hacia", "hasta", "para", "por", "segun", "según", "sin", "sobre", "tras", "y", "e", "ni", "que", "pero", "aunque", "sino", "porque", "si", "sí", "me", "te", "se", "nos", "mi", "mí", "tu", "tú", "su", "este", "esta", "esto",
    "an", "the", "in", "on", "at", "by", "for", "with", "about", "from", "to", "of", "under", "over", "and", "but", "or", "so", "because", "although", "yet", "i", "you", "he", "she", "it", "we", "they", "this", "that", "there",
    "le", "les", "l", "une", "des", "à", "dans", "sur", "avec", "pour", "par", "sans", "chez", "et", "ou", "mais", "donc", "car", "je", "il", "elle", "nous", "vous", "ce",
    "der", "die", "das", "dem", "den", "ein", "eine", "einer", "auf", "zu", "mit", "von", "fur", "für", "bei", "durch", "aus", "vor", "und", "oder", "aber", "denn", "weil", "obwohl", "sondern", "ich", "du", "er", "sie", "es", "wir", "ihr", "sich",
    "il", "lo", "gli", "di", "da", "su", "per", "tra", "fra", "ed", "ma", "oppure", "perche", "perché", "anche", "ti", "ci", "io", "lui", "lei", "noi", "voi",
    "els", "amb", "sense", "sota", "perque", "perquè", "doncs", "jo", "ell", "ella", "nosaltres", "vosaltres", "em", "et",
    "o", "os", "as", "unha", "baixo", "segundo", "sen", "senon", "senón", "mais", "nin", "eu", "ela", "nós", "vós", "vos",
    "ak", "ari", "ariak", "bat", "batzuk", "baino", "gabe", "gora", "behera", "arte", "bila", "kontra", "zehar", "aurka", "eta", "edo", "baina", "denez", "beraz", "orduan", "baita", "ezta", "zein", "hi", "hura", "gu", "zu", "zuek", "haiek", "hau", "hori",
    "um", "uma", "umas", "com", "em", "sem", "sob", "trás", "ate", "até", "mas", "nem", "embora", "contudo", "pois", "ele", "isto", "esse", "essa", "isso"
}
STOPWORD_PHRASES = {"per a", "des de", "cap a", "fins a", "encara que", "ainda que", "aínda que"}

COLOR_BG = "#f8fafc"
COLOR_CARD = "#ffffff"
COLOR_BORDER = "#cbd5e1"
COLOR_TEXT_PRIMARY = "#0f172a"
COLOR_TEXT_MUTED = "#64748b"
COLOR_PRIMARY = "#1e40af"
COLOR_PRIMARY_HOVER = "#1d4ed8"
COLOR_PRIMARY_TEXT = "#ffffff"
COLOR_TABLE_HEADER = "#e2e8f0"
COLOR_ROW_ALT = "#f1f5f9"

ROLE_DESCRIPTIONS = {
    "01": "Oficina Contable",
    "02": "Órgano Gestor",
    "03": "Unidad Tramitadora",
    "04": "Órgano Proponente"
}

def clean_str(val):
    if val is None:
        return ""
    text = str(val).replace("\xa0", " ").strip()
    return "" if text.lower() == "nan" else text

def format_currency(val, currency="€"):
    if val is None or val == "":
        return f"0,00 {currency}"
    try:
        dec = Decimal(str(val))
        return f"{dec:,.2f} {currency}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (InvalidOperation, ValueError):
        return f"{val} {currency}"

def format_date(val):
    text = clean_str(val)
    if not text:
        return ""
    text = text[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            pass
    return text

class InvoiceDataExtractor:
    """Extrae datos normalizados del XML de Facturae 3.2 / 3.2.1 / 3.2.2."""

    def __init__(self, file_path):
        self.file_path = file_path
        self.raw_xml_text = ""
        self.root = None

    def load(self):
        encodings = ["utf-8-sig", "utf-8", "iso-8859-1", "latin-1"]
        last_error = None
        for enc in encodings:
            try:
                with open(self.file_path, "r", encoding=enc) as f:
                    content = f.read().strip()
                xml_start = content.find("<")
                if xml_start > 0:
                    content = content[xml_start:]
                self.root = ET.fromstring(content)
                self.raw_xml_text = content
                return
            except DefusedXmlException as e:
                raise ValueError(f"Estructura XML bloqueada por seguridad (Bandit / DefusedXML): {e}")
            except Exception as e:
                last_error = e
        raise ValueError(f"No se pudo decodificar el archivo XML/XSIG: {last_error}")

    def _t(self, elem, tag, default=""):
        if elem is None:
            return default
        found = elem.find(f".//{{*}}{tag}")
        return clean_str(found.text) if (found is not None and found.text) else default

    def _tp(self, elem, path, default=""):
        if elem is None:
            return default
        xpath = "./" + "/".join(f"{{*}}{p}" for p in path.split("/"))
        found = elem.find(xpath)
        return clean_str(found.text) if (found is not None and found.text) else default

    def _get_party_name(self, party):
        if party is None:
            return ""
        legal = party.find(".//{*}LegalEntity")
        indiv = party.find(".//{*}Individual")
        for c in [
            self._t(legal, "CorporateName"),
            self._t(indiv, "CorporateName"),
            self._t(legal, "TradeName"),
            self._t(indiv, "TradeName")
        ]:
            if c:
                return c
        parts = [self._t(indiv, "Name"), self._t(indiv, "FirstSurname"), self._t(indiv, "SecondSurname")]
        full = " ".join(p for p in parts if p).strip()
        return full if full else ""

    def _get_address(self, party):
        if party is None:
            return ""
        addr = party.find(".//{*}AddressInSpain")
        if addr is None:
            addr = party.find(".//{*}OverseasAddress")
        if addr is None:
            return ""
        street = self._t(addr, "Address")
        pc = self._t(addr, "PostCode") or self._t(addr, "PostCodeAndTown")
        town = self._t(addr, "Town")
        prov = self._t(addr, "Province")
        country = self._t(addr, "CountryCode", "ESP")
        return f"{street}, {pc} {town} ({prov}) {country}".strip(", ")

    def extract(self):
        if self.root is None:
            self.load()

        inv = self.root.find(".//{*}Invoice")
        if inv is None:
            raise ValueError("No se encontró ningún elemento <Invoice> en el comprobante.")

        seller = self.root.find(".//{*}SellerParty")
        buyer = self.root.find(".//{*}BuyerParty")

        # Emisor
        seller_nif = self._t(seller, "TaxIdentificationNumber")
        seller_name = self._get_party_name(seller)
        seller_addr = self._get_address(seller)
        seller_tel = self._tp(seller, "LegalEntity/ContactDetails/Telephone") or self._tp(seller, "Individual/ContactDetails/Telephone")
        seller_email = self._tp(seller, "LegalEntity/ContactDetails/ElectronicMail") or self._tp(seller, "Individual/ContactDetails/ElectronicMail")

        # Receptor
        buyer_nif = self._t(buyer, "TaxIdentificationNumber")
        buyer_name = self._get_party_name(buyer)
        buyer_addr = self._get_address(buyer)

        # Centros Administrativos DIR3
        dir3_list = []
        for c in buyer.findall(".//{*}AdministrativeCentre"):
            role_code = self._t(c, "RoleTypeCode")
            role_label = ROLE_DESCRIPTIONS.get(role_code, f"Rol {role_code}")
            code = self._t(c, "CentreCode")
            desc = self._t(c, "CentreDescription") or self._t(c, "Name")
            dir3_list.append({"role": role_label, "code": code, "desc": desc})

        # Datos Generales de Cabecera
        header = inv.find(".//{*}InvoiceHeader")
        issue = inv.find(".//{*}InvoiceIssueData")
        totals = inv.find(".//{*}InvoiceTotals")

        invoice_number = self._t(header, "InvoiceNumber")
        series = self._t(header, "InvoiceSeriesCode")
        issue_date = format_date(self._t(issue, "IssueDate"))
        operation_date = format_date(self._t(issue, "OperationDate"))
        currency = self._t(issue, "InvoiceCurrencyCode", "EUR")
        curr_symbol = "€" if currency == "EUR" else currency

        # Totales
        gross_amount = self._t(totals, "TotalGrossAmountBeforeTaxes") or self._t(totals, "TotalGrossAmount")
        total_taxes = self._t(totals, "TotalTaxOutputs")
        total_retention = self._t(totals, "TotalTaxesWithheld", "0.00")
        invoice_total = self._t(totals, "InvoiceTotal")
        total_payable = self._t(totals, "TotalOutstandingAmount") or invoice_total

        # Líneas de Factura (Items)
        lines = []
        for idx, line in enumerate(inv.findall(".//{*}InvoiceLine"), start=1):
            desc = self._t(line, "ItemDescription")
            qty = self._t(line, "Quantity", "1")
            price = self._t(line, "UnitPriceWithoutTax", "0.00")
            line_gross = self._t(line, "GrossAmount") or self._t(line, "TotalCost", "0.00")
            code = self._t(line, "ArticleCode")
            tax_rate = self._tp(line, "TaxesOutputs/Tax/TaxRate", "4.00")
            tax_amount = self._tp(line, "TaxesOutputs/Tax/TaxAmount/TotalAmount", "0.00")
            lines.append({
                "line": idx,
                "desc": desc,
                "code": code,
                "qty": qty,
                "price": price,
                "gross": line_gross,
                "tax_rate": tax_rate,
                "tax_amount": tax_amount
            })

        # Datos de Pago
        iban = self._tp(inv, "PaymentDetails/Installment/AccountToBeCredited/IBAN")
        payment_means = self._tp(inv, "PaymentDetails/Installment/PaymentMeans")
        due_date = format_date(self._tp(inv, "PaymentDetails/Installment/InstallmentDueDate"))

        return {
            "seller_nif": seller_nif,
            "seller_name": seller_name,
            "seller_addr": seller_addr,
            "seller_tel": seller_tel,
            "seller_email": seller_email,
            "buyer_nif": buyer_nif,
            "buyer_name": buyer_name,
            "buyer_addr": buyer_addr,
            "dir3_list": dir3_list,
            "invoice_number": invoice_number,
            "series": series,
            "issue_date": issue_date,
            "operation_date": operation_date,
            "currency_symbol": curr_symbol,
            "gross_amount": gross_amount,
            "total_taxes": total_taxes,
            "total_retention": total_retention,
            "invoice_total": invoice_total,
            "total_payable": total_payable,
            "lines": lines,
            "iban": iban,
            "payment_means": payment_means,
            "due_date": due_date,
            "raw_xml": self.raw_xml_text
        }

class FacturaeViewerTab(ttk.Frame):
    """
    Componente modular que se puede usar en una ventana aislada o
    incrustarse directamente como pestaña en FacturaeToAlma.
    """

    def __init__(self, parent, *args, show_loader=False, **kwargs):
        self.show_loader = show_loader
        super().__init__(parent, style="App.TFrame", *args, **kwargs)
        self.current_data = None
        self.pdf_canvas = None
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1 if self.show_loader else 0, weight=1)

        # Barra superior de selección
        self.top_bar = ttk.Frame(self, style="App.TFrame", padding=(0, 0, 0, 8))
        top_bar = self.top_bar
        if self.show_loader:
            top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.columnconfigure(1, weight=1)

        ttk.Label(top_bar, text="Factura electrónica:", font=("Segoe UI", 9, "bold"), background=COLOR_BG).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.file_var = tk.StringVar()
        ttk.Entry(top_bar, textvariable=self.file_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(top_bar, text="Cargar factura...", command=self.load_file_dialog, style="Primary.TButton").grid(row=0, column=2)

        # Sub-Notebook: Se sitúa en primer lugar la Vista Factura (PDF) para que prevalezca
        self.viewer_notebook = ttk.Notebook(self)
        self.viewer_notebook.grid(row=1 if self.show_loader else 0, column=0, sticky="nsew")

        self.tab_pdf_view = ttk.Frame(self.viewer_notebook, style="App.TFrame")
        self.tab_xml_view = ttk.Frame(self.viewer_notebook, style="App.TFrame")

        self.viewer_notebook.add(self.tab_pdf_view, text="  Vista Factura (Estilo PDF)  ")
        self.viewer_notebook.add(self.tab_xml_view, text="  Código XML (FACe)  ")

        self._build_pdf_view_ui()
        self._build_xml_view_ui()

    def _build_pdf_view_ui(self):
        container = ttk.Frame(self.tab_pdf_view, style="App.TFrame")
        container.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        canvas = tk.Canvas(container, bg=COLOR_BG, highlightthickness=0)
        self.pdf_canvas = canvas
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_pdf_frame = ttk.Frame(canvas, style="App.TFrame")

        self.scrollable_pdf_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_pdf_frame, anchor="nw")

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas.bind("<Enter>", self._bind_pdf_mousewheel)
        canvas.bind("<Leave>", self._unbind_pdf_mousewheel)
        self.scrollable_pdf_frame.bind("<Enter>", self._bind_pdf_mousewheel)
        self.scrollable_pdf_frame.bind("<Leave>", self._unbind_pdf_mousewheel)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Placeholder inicial
        self.placeholder_label = ttk.Label(
            self.scrollable_pdf_frame,
            text="Seleccione un archivo de factura electrónica (.xsig, .xml, .txt) para visualizar sus datos.",
            font=("Segoe UI", 10, "italic"),
            foreground=COLOR_TEXT_MUTED,
            background=COLOR_BG
        )
        self.placeholder_label.pack(pady=40)

    def _build_xml_view_ui(self):
        container = ttk.Frame(self.tab_xml_view, style="App.TFrame")
        container.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # Marco de texto con Scrollbars
        txt_frame = ttk.Frame(container)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        self.xml_text = tk.Text(
            txt_frame,
            wrap="none",
            font=("Consolas", 10),
            background="#ffffff",
            foreground=COLOR_TEXT_PRIMARY,
            padx=10,
            pady=10,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER
        )
        x_scroll = ttk.Scrollbar(txt_frame, orient="horizontal", command=self.xml_text.xview)
        y_scroll = ttk.Scrollbar(txt_frame, orient="vertical", command=self.xml_text.yview)

        self.xml_text.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)

        self.xml_text.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        txt_frame.rowconfigure(0, weight=1)
        txt_frame.columnconfigure(0, weight=1)

        # Configurar colores de etiquetas XML
        self.xml_text.tag_configure("tag", foreground="#0369a1", font=("Consolas", 10, "bold"))
        self.xml_text.tag_configure("attr", foreground="#b45309")
        self.xml_text.tag_configure("val", foreground="#15803d")
        self.xml_text.tag_configure("comment", foreground="#94a3b8", font=("Consolas", 10, "italic"))

    def _bind_pdf_mousewheel(self, _event=None):
        self.bind_all("<MouseWheel>", self._on_pdf_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_pdf_mousewheel, add="+")
        self.bind_all("<Button-5>", self._on_pdf_mousewheel, add="+")

    def _unbind_pdf_mousewheel(self, _event=None):
        self.unbind_all("<MouseWheel>")
        self.unbind_all("<Button-4>")
        self.unbind_all("<Button-5>")

    def _on_pdf_mousewheel(self, event):
        if not self.pdf_canvas:
            return
        if getattr(event, "num", None) == 4:
            units = -3
        elif getattr(event, "num", None) == 5:
            units = 3
        else:
            delta = getattr(event, "delta", 0)
            units = -int(delta / 120) if delta else 0
            if units == 0 and delta:
                units = -1 if delta > 0 else 1
        self.pdf_canvas.yview_scroll(units, "units")
        return "break"

    def load_file_dialog(self):
        path = filedialog.askopenfilename(
            title="Seleccionar factura electrónica FACe",
            filetypes=[
                ("Facturas compatibles", "*.xsig *.xml *.txt"),
                ("Archivos XSIG", "*.xsig"),
                ("Archivos XML", "*.xml"),
                ("Archivos TXT", "*.txt"),
                ("Todos los archivos", "*.*")
            ]
        )
        if path:
            self.load_file(path)

    def load_file(self, path):
        try:
            extractor = InvoiceDataExtractor(path)
            data = extractor.extract()
            self.file_var.set(path)
            self.current_data = data
            self._render_pdf_view(data)
            self._render_xml_view(data["raw_xml"])
        except Exception as e:
            messagebox.showerror("Error al leer factura", f"No se pudo cargar la factura seleccionada:\n{e}")

    def _render_pdf_view(self, data):
        for widget in self.scrollable_pdf_frame.winfo_children():
            widget.destroy()

        # Lienzo "Hoja de Papel / Tarjeta"
        sheet = tk.Frame(self.scrollable_pdf_frame, bg=COLOR_CARD, padx=24, pady=20, highlightbackground=COLOR_BORDER, highlightthickness=1)
        sheet.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        sheet.columnconfigure(0, weight=1)
        sheet.columnconfigure(1, weight=1)

        # 1. Cabecera principal: Factura & Emisión
        hdr_box = tk.Frame(sheet, bg=COLOR_CARD)
        hdr_box.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        hdr_box.columnconfigure(0, weight=1)
        hdr_box.columnconfigure(1, weight=1)

        title_lbl = tk.Label(hdr_box, text=f"FACTURA: {data['invoice_number']}", font=("Segoe UI", 16, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD)
        title_lbl.grid(row=0, column=0, sticky="w")

        meta_text = f"Serie: {data['series'] or '-'}   |   Fecha de Emisión: {data['issue_date']}   |   Fecha de Operación: {data['operation_date'] or data['issue_date']}"
        meta_lbl = tk.Label(hdr_box, text=meta_text, font=("Segoe UI", 9), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD)
        meta_lbl.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # 2. Bloques Emisor / Receptor
        emisor_card = tk.LabelFrame(sheet, text=" DATOS DEL EMISOR ", font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=8, highlightbackground=COLOR_BORDER)
        emisor_card.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 12))

        receptor_card = tk.LabelFrame(sheet, text=" DATOS DEL RECEPTOR ", font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=8, highlightbackground=COLOR_BORDER)
        receptor_card.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(0, 12))

        # Rellenar Emisor
        tk.Label(emisor_card, text=data["seller_name"], font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w")
        tk.Label(emisor_card, text=f"CIF/NIF: {data['seller_nif']}", font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))
        tk.Label(emisor_card, text=f"Dirección: {data['seller_addr']}", font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w", pady=(2, 0))
        if data["seller_tel"] or data["seller_email"]:
            contact = f"Contacto: {data['seller_tel']}  {data['seller_email']}".strip()
            tk.Label(emisor_card, text=contact, font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))

        # Rellenar Receptor
        tk.Label(receptor_card, text=data["buyer_name"], font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w")
        tk.Label(receptor_card, text=f"CIF/NIF: {data['buyer_nif']}", font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))
        tk.Label(receptor_card, text=f"Dirección: {data['buyer_addr']}", font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w", pady=(2, 0))

        # 3. Bloque DIR3 (Centros Administrativos FACe)
        if data["dir3_list"]:
            dir3_card = tk.LabelFrame(sheet, text=" CENTROS ADMINISTRATIVOS FACe (DIR3) ", font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=6)
            dir3_card.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 14))
            for item in data["dir3_list"]:
                line_text = f"• {item['role']}: [{item['code']}] {item['desc']}"
                tk.Label(dir3_card, text=line_text, font=("Segoe UI", 8), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD, anchor="w").pack(fill=tk.X, pady=1)

        # 4. Tabla de Líneas / Conceptos
        items_frame = tk.LabelFrame(sheet, text=" DETALLE ECONÓMICO (LÍNEAS) ", font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=8, pady=8)
        items_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0, 14))
        items_frame.columnconfigure(0, weight=1)

        cols = ("line", "code", "desc", "qty", "price", "gross", "tax")
        tree = ttk.Treeview(items_frame, columns=cols, show="headings", height=min(max(len(data["lines"]), 4), 14))
        tree.heading("line", text="#")
        tree.heading("code", text="Artículo / ISBN")
        tree.heading("desc", text="Descripción / Título")
        tree.heading("qty", text="Cant.")
        tree.heading("price", text="Precio Unit.")
        tree.heading("gross", text="Importe Bruto")
        tree.heading("tax", text="IVA Repercutido")

        tree.column("line", width=35, anchor="center")
        tree.column("code", width=140, anchor="w")
        tree.column("desc", width=390, anchor="w")
        tree.column("qty", width=55, anchor="e")
        tree.column("price", width=85, anchor="e")
        tree.column("gross", width=95, anchor="e")
        tree.column("tax", width=110, anchor="e")

        tree.tag_configure("even", background="#ffffff")
        tree.tag_configure("odd", background=COLOR_ROW_ALT)

        for i, line in enumerate(data["lines"]):
            tag = "even" if i % 2 == 0 else "odd"
            tree.insert("", "end", values=(
                line["line"],
                line["code"] or "-",
                line["desc"],
                line["qty"],
                format_currency(line["price"], data["currency_symbol"]),
                format_currency(line["gross"], data["currency_symbol"]),
                f"{line['tax_rate']}% ({format_currency(line['tax_amount'], data['currency_symbol'])})"
            ), tags=(tag,))

        tree.pack(fill=tk.BOTH, expand=True)

        # 5. Sección Inferior: Datos de Pago y Resumen de Totales
        bottom_frame = tk.Frame(sheet, bg=COLOR_CARD)
        bottom_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)

        # Forma de pago & Vencimiento
        pay_box = tk.LabelFrame(bottom_frame, text=" DATOS DE PAGO ", font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=10)
        pay_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        tk.Label(pay_box, text=f"Medio de Pago: Transferencia (Código {data['payment_means'] or '04'})", font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w")
        tk.Label(pay_box, text=f"Fecha Vencimiento: {data['due_date'] or data['issue_date']}", font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))
        tk.Label(pay_box, text=f"Cuenta / IBAN: {data['iban'] or 'No especificado'}", font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(4, 0))

        # Totales (similar al cuadro de resumen de los PDF adjuntos)
        totals_box = tk.Frame(bottom_frame, bg=COLOR_ROW_ALT, padx=14, pady=10, highlightbackground=COLOR_BORDER, highlightthickness=1)
        totals_box.grid(row=0, column=1, sticky="nsew")
        totals_box.columnconfigure(0, weight=1)
        totals_box.columnconfigure(1, weight=1)

        def add_tot_row(label, val, row, is_grand=False):
            fg = COLOR_PRIMARY if is_grand else COLOR_TEXT_PRIMARY
            font = ("Segoe UI", 11, "bold") if is_grand else ("Segoe UI", 9)
            tk.Label(totals_box, text=label, font=font, fg=fg, bg=COLOR_ROW_ALT).grid(row=row, column=0, sticky="w", pady=2)
            tk.Label(totals_box, text=val, font=font, fg=fg, bg=COLOR_ROW_ALT).grid(row=row, column=1, sticky="e", pady=2)

        add_tot_row("Base Imponible:", format_currency(data["gross_amount"], data["currency_symbol"]), 0)
        add_tot_row("Impuestos Repercutidos:", format_currency(data["total_taxes"], data["currency_symbol"]), 1)
        if Decimal(data["total_retention"] or "0.00") > 0:
            add_tot_row("Impuestos Retenidos:", format_currency(data["total_retention"], data["currency_symbol"]), 2)
        ttk.Separator(totals_box, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=4)
        add_tot_row("TOTAL FACTURA:", format_currency(data["invoice_total"], data["currency_symbol"]), 4, is_grand=True)
        add_tot_row("Total a Pagar:", format_currency(data["total_payable"], data["currency_symbol"]), 5)

    def _render_xml_view(self, raw_xml):
        self.xml_text.config(state="normal")
        self.xml_text.delete("1.0", tk.END)

        # Formateo indentado seguro con minidom de defusedxml
        try:
            dom = defusedxml.minidom.parseString(raw_xml.encode("utf-8"))
            pretty_xml = dom.toprettyxml(indent="  ")
            lines = [line for line in pretty_xml.splitlines() if line.strip()]
            formatted_text = "\n".join(lines)
        except Exception:
            formatted_text = raw_xml

        self.xml_text.insert(tk.END, formatted_text)

        # Resaltado de sintaxis XML (Tags, Atributos y Comentarios)
        content = self.xml_text.get("1.0", tk.END)
        for match in re.finditer(r"(</?[\w\:\-]+)(\s|>|/)", content):
            start = f"1.0 + {match.start(1)} chars"
            end = f"1.0 + {match.end(1)} chars"
            self.xml_text.tag_add("tag", start, end)

        for match in re.finditer(r"<!--[\s\S]*?-->", content):
            start = f"1.0 + {match.start()} chars"
            end = f"1.0 + {match.end()} chars"
            self.xml_text.tag_add("comment", start, end)

        self.xml_text.config(state="disabled")



@dataclass
class AlmaPOLRecord:
    po_line: str
    title: str = None
    reporting_code: str = None
    secondary_reporting_code: str = None
    tertiary_reporting_code: str = None
    fourth_reporting_code: str = None
    fifth_reporting_code: str = None
    start_subs_date: datetime = None
    isbn_raw: str = None
    issn_raw: str = None
    quantity_for_pricing: Decimal = None
    po_price: Decimal = None
    funds: list = field(default_factory=list)
    normalized_title_words: list = field(default_factory=list)
    normalized_title_text: str = ""

@dataclass
class ValidationIssue:
    issue_type: str
    invoice_number: str
    invoice_line_number: int
    isbn: str
    issn: str
    title: str
    candidates: list
    message: str
    po_line: str = None
    invoice_quantity: Decimal = None
    alma_quantity_for_pricing: Decimal = None

@dataclass
class InvoiceLineData:
    line_number: int
    po_line: str = None
    line_type: str = "REGULAR"
    vendor_ref: str = None
    issn: str = None
    isbn: str = None
    title: str = None
    price: Decimal = Decimal("0.00")
    quantity: Decimal = Decimal("1.00")
    start_date: datetime = None
    end_date: datetime = None
    note: str = None
    vat_amount: Decimal = None
    vat_percentage: Decimal = None
    reporting_code: str = None
    secondary_reporting_code: str = None
    tertiary_reporting_code: str = None
    fourth_reporting_code: str = None
    fifth_reporting_code: str = None
    alma_quantity_for_pricing: Decimal = None
    funds: list = field(default_factory=list)
    match_status: str = None
    match_message: str = None
    match_candidates: list = field(default_factory=list)

@dataclass
class InvoiceData:
    invoice_number: str
    date: datetime
    total_amount: Decimal
    currency: str
    vendor_account: str
    note: str
    seller_name: str = None
    seller_tax_id: str = None
    source_file: str = None
    vat_amount: Decimal = None
    inclusive: bool = False
    lines: list = field(default_factory=list)

def sanitize_excel_value(value):
    if isinstance(value, str) and value.startswith(DANGEROUS_EXCEL_PREFIXES):
        return "'" + value
    return value

def clean_cell(value):
    if value is None:
        return None
    text = str(value).replace("\xa0", " ").strip()
    if not text or text.lower() == "nan":
        return None
    return text

def alma_empty_to_none(value):
    value = clean_cell(value)
    return None if value in (None, "", "-1") else value

def safe_decimal(value, default=Decimal("0.00")):
    if value is None:
        return default
    raw = str(value).strip().replace(" ", "")
    if not raw:
        return default
    if "." in raw and "," in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw and "." not in raw:
        raw = raw.replace(",", ".")
    try:
        return Decimal(raw)
    except InvalidOperation:
        logging.warning(f"No se pudo convertir a Decimal: {value!r}")
        return default

def decimal_or_none(value):
    if value is None or not str(value).strip():
        return None
    return safe_decimal(value)

def decimal_to_float_or_none(value):
    return None if value is None else float(value)

def parse_date_for_excel(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    text = clean_cell(value)
    if not text:
        return None
    text = text[:10]
    for fmt in ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    logging.warning(f"No se pudo convertir la fecha: {value!r}")
    return None

def normalize_identifier(value):
    if not value:
        return ""
    return re.sub(r"[^0-9X]", "", str(value).replace("\xa0", " ").upper())

def split_identifiers(value):
    value = clean_cell(value)
    if not value:
        return []
    return [norm for norm in (normalize_identifier(part) for part in value.split(";")) if norm]

def is_valid_isbn10(code):
    code = normalize_identifier(code)
    if not re.fullmatch(r"\d{9}[\dX]", code):
        return False
    total = sum((10 - i) * (10 if c == "X" else int(c)) for i, c in enumerate(code))
    return total % 11 == 0

def is_valid_isbn13(code):
    code = normalize_identifier(code)
    if not re.fullmatch(r"\d{13}", code):
        return False
    total = sum(int(c) if i % 2 == 0 else int(c) * 3 for i, c in enumerate(code[:12]))
    check = (10 - (total % 10)) % 10
    return check == int(code[-1])

def is_valid_issn(code):
    code = normalize_identifier(code)
    if not re.fullmatch(r"\d{7}[\dX]", code):
        return False
    total = sum(int(code[i]) * (8 - i) for i in range(7))
    check = (11 - (total % 11)) % 11
    expected = "X" if check == 10 else str(check)
    return code[-1] == expected

def looks_like_issn(code):
    return bool(re.fullmatch(r"\d{7}[\dX]", normalize_identifier(code)))

def looks_like_isbn(code):
    return bool(re.fullmatch(r"\d{10}|\d{13}", normalize_identifier(code)))

def extract_identifiers_from_text(text):
    if not text:
        return None, None, text
    original_text = str(text)
    isbn = None
    issn = None
    matched = []
    candidate_pattern = re.compile(r"[0-9Xx][0-9Xx\-\s]{6,25}[0-9Xx]")
    for match in candidate_pattern.finditer(original_text):
        fragment = match.group(0)
        normalized = normalize_identifier(fragment)
        if not isbn and len(normalized) == 13 and is_valid_isbn13(normalized):
            isbn = normalized
            matched.append(fragment)
            continue
        if not isbn and len(normalized) == 10 and is_valid_isbn10(normalized):
            isbn = normalized
            matched.append(fragment)
            continue
        if not issn and len(normalized) == 8 and is_valid_issn(normalized):
            issn = normalized
            matched.append(fragment)
            continue
    cleaned = original_text
    for fragment in matched:
        cleaned = cleaned.replace(fragment, " ")
    cleaned = re.sub(r"\s*/\s*$", "", cleaned)
    cleaned = re.sub(r"\s*[-–—:;,/]\s*$", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return isbn, issn, cleaned

def classify_article_code(code):
    if not code:
        return None, None, None
    isbn, issn, _ = extract_identifiers_from_text(code)
    if isbn:
        return None, isbn, None
    if issn:
        return issn, None, None
    if looks_like_issn(code):
        return code, None, None
    if looks_like_isbn(code):
        return None, code, None
    return None, None, code

def normalize_header(header):
    return "" if header is None else " ".join(str(header).strip().split())

def get_vendor_account(nif, seller_name=None):
    nif = nif.strip() if nif else ""
    if nif and nif in VENDOR_MAP:
        return VENDOR_MAP[nif]
    if seller_name:
        return seller_name.strip()
    return VENDOR_MAP["DEFAULT"]

def build_validation_output_path(output_path):
    base, _ = os.path.splitext(output_path)
    return f"{base}_validacion.xlsx"

def remove_accents(text):
    if not text:
        return ""
    normalized = unicodedata.normalize("NFD", text)
    return "".join(char for char in normalized if unicodedata.category(char) != "Mn")

def normalize_title_for_matching(title):
    if not title:
        return []
    text = remove_accents(html.unescape(str(title)).lower()).replace("’", "'").replace("´", "'")
    for phrase in STOPWORD_PHRASES:
        normalized_phrase = re.sub(r"[^a-z0-9]+", " ", remove_accents(phrase.lower())).strip()
        if normalized_phrase:
            text = re.sub(rf"\b{re.escape(normalized_phrase)}\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return [w for w in text.split() if len(w) > 1 and w not in STOPWORDS]

GENERIC_CONTAINMENT_WORDS = {
    "access", "acceso", "base", "bases", "collection", "coleccion",
    "database", "datos", "digital", "electronic", "electronica",
    "online", "package", "paquete", "platform", "plataforma",
    "service", "servicio", "subscription", "suscripcion",
}


def contains_word_sequence(container_words, candidate_words):
    """Indica si candidate_words aparece como secuencia completa en container_words."""
    if not container_words or not candidate_words:
        return False
    candidate_length = len(candidate_words)
    if candidate_length > len(container_words):
        return False
    return any(
        container_words[index:index + candidate_length] == candidate_words
        for index in range(len(container_words) - candidate_length + 1)
    )


def is_safe_title_fragment(words):
    """Evita emparejar automaticamente fragmentos demasiado cortos o genericos."""
    if not words:
        return False
    if len(words) == 1:
        word = words[0]
        return len(word) >= 5 and word not in GENERIC_CONTAINMENT_WORDS
    significant = [word for word in words if word not in GENERIC_CONTAINMENT_WORDS]
    return bool(significant) and sum(len(word) for word in significant) >= 5


def title_containment_score(invoice_words, alma_words):
    """Puntua una inclusion segura entre secuencias de palabras normalizadas."""
    if invoice_words == alma_words and invoice_words:
        return 1.0

    shorter, longer = (
        (invoice_words, alma_words)
        if len(invoice_words) <= len(alma_words)
        else (alma_words, invoice_words)
    )
    if not is_safe_title_fragment(shorter):
        return 0.0
    if not contains_word_sequence(longer, shorter):
        return 0.0

    # Una secuencia mas larga y especifica recibe una puntuacion ligeramente mayor.
    coverage = len(shorter) / len(longer)
    return 0.94 + min(0.05, coverage * 0.05)


def extract_significant_numbers(words):
    """Extrae numeros que suelen distinguir curso, volumen o nivel."""
    values = set()
    roman = {"i": "1", "ii": "2", "iii": "3", "iv": "4", "v": "5", "vi": "6"}
    for word in words:
        if word.isdigit():
            # Los anos/editoriales se conservan, pero pesan menos en el diagnostico.
            values.add(word)
        elif word in roman:
            values.add(roman[word])
        else:
            match = re.fullmatch(r"(\d+)(?:ep|o|a)?", word)
            if match:
                values.add(match.group(1))
    return values


def diagnostic_title_score(invoice_title, record):
    """Puntuacion explicable para ordenar candidatas, nunca determinante."""
    invoice_words = normalize_title_for_matching(invoice_title)
    alma_words = record.normalized_title_words
    if not invoice_words or not alma_words:
        return 0.0, []
    invoice_set, alma_set = set(invoice_words), set(alma_words)
    common = invoice_set & alma_set
    union = invoice_set | alma_set
    jaccard = len(common) / len(union) if union else 0.0
    containment = len(common) / min(len(invoice_set), len(alma_set))
    sequence = SequenceMatcher(None, " ".join(invoice_words), " ".join(alma_words)).ratio()
    fuzzy_hits = 0
    for word in invoice_set:
        if len(word) < 5:
            continue
        if any(SequenceMatcher(None, word, other).ratio() >= 0.84 for other in alma_set if len(other) >= 5):
            fuzzy_hits += 1
    fuzzy = fuzzy_hits / max(1, min(len(invoice_set), len(alma_set)))
    score = 0.38 * containment + 0.27 * sequence + 0.20 * jaccard + 0.15 * fuzzy
    reasons = [f"lexico={score:.3f}"]
    invoice_numbers = extract_significant_numbers(invoice_words)
    alma_numbers = extract_significant_numbers(alma_words)
    if invoice_numbers and alma_numbers:
        if invoice_numbers & alma_numbers:
            score += 0.08
            reasons.append("numero/volumen compatible")
        elif any(len(n) <= 2 for n in invoice_numbers | alma_numbers):
            score -= 0.28
            reasons.append("numero/curso incompatible")
    return max(0.0, min(1.0, score)), reasons


def title_similarity_score(invoice_title, record):
    invoice_words = normalize_title_for_matching(invoice_title)
    alma_words = record.normalized_title_words
    if not invoice_words or not alma_words:
        return 0.0
    invoice_set = set(invoice_words)
    alma_set = set(alma_words)
    common = invoice_set.intersection(alma_set)
    word_score = len(common) / len(invoice_set)
    sequence_score = SequenceMatcher(None, " ".join(invoice_words), record.normalized_title_text).ratio()
    return word_score * 0.75 + sequence_score * 0.25

def normalize_provider_for_comparison(value):
    if not value:
        return ""
    text = remove_accents(str(value).lower())
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def get_vendor_batch_key(invoice):
    if invoice.seller_tax_id:
        return normalize_identifier(invoice.seller_tax_id)
    if invoice.vendor_account:
        return normalize_provider_for_comparison(invoice.vendor_account)
    return ""

def build_conversion_summary(invoices, validation_issues):
    summary = {"total_lines": 0, "match_isbn": 0, "match_issn": 0, "match_title": 0, "duplicate_isbn": 0, "duplicate_issn": 0, "duplicate_title": 0, "ambiguous_title": 0, "no_match": 0, "validation_issues": len(validation_issues), "lines_without_po_line": 0}
    for invoice in invoices:
        for line in invoice.lines:
            summary["total_lines"] += 1
            status = line.match_status or ""
            if status == "MATCH_ISBN": summary["match_isbn"] += 1
            elif status == "MATCH_ISSN": summary["match_issn"] += 1
            elif status == "MATCH_TITLE": summary["match_title"] += 1
            elif status == "DUPLICATE_ISBN": summary["duplicate_isbn"] += 1
            elif status == "DUPLICATE_ISSN": summary["duplicate_issn"] += 1
            elif status == "DUPLICATE_TITLE": summary["duplicate_title"] += 1
            elif status == "AMBIGUOUS_TITLE": summary["ambiguous_title"] += 1
            elif status == "NO_MATCH": summary["no_match"] += 1
            if not line.po_line:
                summary["lines_without_po_line"] += 1
    return summary

class AlmaPOLineLookup:
    REQUIRED_COLUMNS = ["PO Line", "Reporting Code", "Secondary Reporting Code", "Tertiary Reporting Code", "Start subs date", "Title", "ISSN", "ISBN", "Quantity for Pricing"]
    def __init__(self, alma_excel_path):
        self.alma_excel_path = alma_excel_path
        self.records = []
        self.isbn_index = {}
        self.issn_index = {}
        self.po_line_index = {}
    def load(self):
        try:
            wb = load_workbook(self.alma_excel_path, data_only=True)
            ws = wb.active
        except Exception as e:
            logging.exception("Error cargando Excel de Alma")
            raise ValueError(f"No se pudo leer el Excel de Alma con PO Lines: {e}")
        headers = self._read_headers(ws)
        self._validate_headers(headers)
        col_map = {}
        col_map_multi = {}
        for idx, header in enumerate(headers, start=1):
            normalized = normalize_header(header)
            if normalized not in col_map:
                col_map[normalized] = idx
            col_map_multi.setdefault(normalized, []).append(idx)
        fund_columns = col_map_multi.get("Fund and percent", [])
        price_header = next((header for header in (
            "PO Line Price", "Net Price", "Price", "Amount", "Line Amount", "Encumbrance"
        ) if header in col_map), None)
        for row_idx in range(2, ws.max_row + 1):
            po_line = clean_cell(ws.cell(row=row_idx, column=col_map["PO Line"]).value)
            if not po_line:
                continue
            title = clean_cell(ws.cell(row=row_idx, column=col_map["Title"]).value)
            title_words = normalize_title_for_matching(title)
            funds = []
            for fund_col in fund_columns:
                fund_value = alma_empty_to_none(ws.cell(row=row_idx, column=fund_col).value)
                if fund_value:
                    funds.append(fund_value)
            rec = AlmaPOLRecord(
                po_line=po_line,
                reporting_code=alma_empty_to_none(ws.cell(row=row_idx, column=col_map["Reporting Code"]).value),
                secondary_reporting_code=alma_empty_to_none(ws.cell(row=row_idx, column=col_map["Secondary Reporting Code"]).value),
                tertiary_reporting_code=alma_empty_to_none(ws.cell(row=row_idx, column=col_map["Tertiary Reporting Code"]).value),
                fourth_reporting_code=alma_empty_to_none(ws.cell(row=row_idx, column=col_map["Fourth Reporting Code"]).value) if "Fourth Reporting Code" in col_map else None,
                fifth_reporting_code=alma_empty_to_none(ws.cell(row=row_idx, column=col_map["Fifth Reporting Code"]).value) if "Fifth Reporting Code" in col_map else None,
                start_subs_date=parse_date_for_excel(ws.cell(row=row_idx, column=col_map["Start subs date"]).value),
                title=title,
                issn_raw=clean_cell(ws.cell(row=row_idx, column=col_map["ISSN"]).value),
                quantity_for_pricing=decimal_or_none(ws.cell(row=row_idx, column=col_map["Quantity for Pricing"]).value),
                po_price=decimal_or_none(ws.cell(row=row_idx, column=col_map[price_header]).value) if price_header else None,
                isbn_raw=clean_cell(ws.cell(row=row_idx, column=col_map["ISBN"]).value),
                funds=funds,
                normalized_title_words=title_words,
                normalized_title_text=" ".join(title_words)
            )
            self.records.append(rec)
            self.po_line_index[rec.po_line] = rec
            for isbn in split_identifiers(rec.isbn_raw):
                self.isbn_index.setdefault(isbn, []).append(rec)
            for issn in split_identifiers(rec.issn_raw):
                self.issn_index.setdefault(issn, []).append(rec)
        logging.info(f"Excel Alma cargado: {len(self.records)} registros, {len(self.isbn_index)} ISBN indexados, {len(self.issn_index)} ISSN indexados")
    def _read_headers(self, ws):
        return [clean_cell(cell.value) or "" for cell in ws[1]]
    def _validate_headers(self, headers):
        normalized = {normalize_header(h) for h in headers}
        missing = [c for c in self.REQUIRED_COLUMNS if c not in normalized]
        if missing:
            raise ValueError("El Excel de Alma no contiene las columnas esperadas. Faltan: " + ", ".join(missing))
    def find_by_po_line(self, po_line):
        return self.po_line_index.get(clean_cell(po_line))

    def find_by_isbn(self, isbn):
        return self.isbn_index.get(normalize_identifier(isbn), []) if normalize_identifier(isbn) else []
    def find_by_issn(self, issn):
        return self.issn_index.get(normalize_identifier(issn), []) if normalize_identifier(issn) else []
    def rank_title_candidates(self, title, limit=8):
        ranked = []
        for record in self.records:
            score, reasons = diagnostic_title_score(title, record)
            if score >= 0.20:
                ranked.append((record, score, reasons))
        ranked.sort(key=lambda item: item[1], reverse=True)
        return ranked[:limit]

    def find_by_title(self, title):
        if not title:
            return {"status": "NO_MATCH", "candidates": [], "scores": [], "method": None}

        invoice_words = normalize_title_for_matching(title)
        if not invoice_words:
            return {"status": "NO_MATCH", "candidates": [], "scores": [], "method": None}

        # 1. Coincidencia exacta tras normalizar mayusculas, tildes y puntuacion.
        exact = self._unique_by_po_line([
            record for record in self.records
            if record.normalized_title_words == invoice_words
        ])
        if exact:
            return self._title_result(exact, [1.0] * len(exact), "EXACT")

        # 2. Inclusion segura por secuencias completas de palabras.
        # Evita falsos positivos como 'art' dentro de 'artificial'.
        contained = []
        for record in self.records:
            alma_words = record.normalized_title_words
            if not alma_words:
                continue
            score = title_containment_score(invoice_words, alma_words)
            if score > 0:
                contained.append((record, score))

        if contained:
            contained.sort(
                key=lambda item: (
                    item[1],
                    len(item[0].normalized_title_words),
                    len(item[0].normalized_title_text),
                ),
                reverse=True,
            )
            unique_records, unique_scores = self._unique_scored_by_po_line(contained)
            return self._title_result(unique_records, unique_scores, "CONTAINMENT")

        # 3. Ultimo recurso: algoritmo de similitud historico.
        scored = []
        for record in self.records:
            if record.normalized_title_words:
                score = title_similarity_score(title, record)
                if score >= TITLE_MATCH_MIN_SCORE:
                    scored.append((record, score))
        scored.sort(key=lambda item: item[1], reverse=True)
        if not scored:
            return {"status": "NO_MATCH", "candidates": [], "scores": [], "method": None}

        unique_records, unique_scores = self._unique_scored_by_po_line(scored)
        if len(unique_records) == 1:
            return {"status": "MATCH", "candidates": unique_records, "scores": unique_scores, "method": "SIMILARITY"}

        best_score = unique_scores[0]
        if best_score - unique_scores[1] >= TITLE_MATCH_MIN_MARGIN:
            return {"status": "MATCH", "candidates": [unique_records[0]], "scores": [best_score], "method": "SIMILARITY"}

        ambiguous = [
            (record, score)
            for record, score in zip(unique_records, unique_scores)
            if best_score - score < TITLE_MATCH_MIN_MARGIN
        ]
        return {
            "status": "AMBIGUOUS",
            "candidates": [record for record, _ in ambiguous],
            "scores": [score for _, score in ambiguous],
            "method": "SIMILARITY",
        }

    @staticmethod
    def _unique_by_po_line(records):
        unique = {}
        for record in records:
            unique.setdefault(record.po_line, record)
        return list(unique.values())

    @staticmethod
    def _unique_scored_by_po_line(scored_records):
        unique = {}
        for record, score in scored_records:
            current = unique.get(record.po_line)
            if current is None or score > current[1]:
                unique[record.po_line] = (record, score)
        ordered = sorted(unique.values(), key=lambda item: item[1], reverse=True)
        return [record for record, _ in ordered], [score for _, score in ordered]

    @staticmethod
    def _title_result(records, scores, method):
        if len(records) == 1:
            return {"status": "MATCH", "candidates": records, "scores": scores, "method": method}
        return {"status": "AMBIGUOUS", "candidates": records, "scores": scores, "method": method}

class FacturaeParser:
    def __init__(self, invoice_path):
        self.invoice_path = invoice_path
        self.tree = None
        self.root = None
    def load(self):
        try:
            self.tree = ET.parse(self.invoice_path)
            self.root = self.tree.getroot()
            logging.info(f"Archivo de factura cargado correctamente: {self.invoice_path}")
            return
        except DefusedXmlException as e:
            logging.exception("XML bloqueado por seguridad")
            raise ValueError("El archivo XML/XSIG contiene estructuras XML no permitidas por seguridad y no se procesara. " + str(e))
        except Exception as first_error:
            logging.warning(f"No se pudo parsear directamente. Se intentara lectura alternativa. Error inicial: {first_error}")
        encodings = ["utf-8-sig", "utf-8", "iso-8859-1", "latin-1"]
        last_error = None
        for enc in encodings:
            try:
                with open(self.invoice_path, "r", encoding=enc) as f:
                    content = f.read().strip()
                xml_start = content.find("<")
                if xml_start > 0:
                    content = content[xml_start:]
                self.root = ET.fromstring(content)
                self.tree = ET.ElementTree(self.root)
                logging.info(f"Archivo de factura cargado mediante lectura alternativa con codificacion {enc}: {self.invoice_path}")
                return
            except DefusedXmlException as e:
                logging.exception("XML bloqueado por seguridad")
                raise ValueError("El archivo XML/XSIG contiene estructuras XML no permitidas por seguridad y no se procesara. " + str(e))
            except Exception as e:
                last_error = e
        logging.exception("Error cargando archivo de factura")
        raise ValueError(f"No se pudo leer o parsear el archivo de factura. Debe ser XML valido. Ultimo error: {last_error}")
    def find_text(self, element, tag, default=""):
        if element is None:
            return default
        found = element.find(f".//{{*}}{tag}")
        return found.text.strip() if found is not None and found.text else default
    def find_direct_text(self, element, tag, default=""):
        if element is None:
            return default
        found = element.find(f"./{{*}}{tag}")
        return found.text.strip() if found is not None and found.text else default
    def find_path_text(self, element, path, default=""):
        if element is None:
            return default
        xpath = "./" + "/".join(f"{{*}}{part}" for part in path.split("/"))
        found = element.find(xpath)
        return found.text.strip() if found is not None and found.text else default
    def get_seller_party_name(self, seller_party):
        if seller_party is None:
            return None
        legal = seller_party.find(".//{*}LegalEntity")
        individual = seller_party.find(".//{*}Individual")
        for candidate in [self.find_direct_text(legal, "CorporateName"), self.find_direct_text(individual, "CorporateName"), self.find_direct_text(legal, "TradeName"), self.find_direct_text(individual, "TradeName")]:
            if candidate:
                return candidate
        parts = [self.find_direct_text(individual, "Name"), self.find_direct_text(individual, "FirstSurname"), self.find_direct_text(individual, "SecondSurname")]
        full = " ".join(p.strip() for p in parts if p and p.strip())
        return full if full else None
    def parse(self):
        if self.root is None:
            self.load()
        invoices_xml = self.root.findall(".//{*}Invoice")
        logging.info(f"Numero de nodos Invoice detectados: {len(invoices_xml)}")
        if not invoices_xml:
            raise ValueError("No se encontraron elementos <Invoice> en el archivo.")
        return [self._parse_invoice(inv) for inv in invoices_xml]
    def _parse_invoice(self, inv):
        seller_party = self.root.find(".//{*}SellerParty")
        seller_tax_id = self.find_text(seller_party, "TaxIdentificationNumber", "")
        seller_name = self.get_seller_party_name(seller_party)
        vendor_account = get_vendor_account(seller_tax_id, seller_name)
        invoice_data = InvoiceData(
            invoice_number=self.find_text(inv, "InvoiceNumber"),
            date=parse_date_for_excel(self.find_text(inv, "IssueDate")),
            total_amount=safe_decimal(self.find_text(inv, "InvoiceTotal")),
            currency=self.find_text(inv, "InvoiceCurrencyCode", "EUR"),
            vendor_account=vendor_account,
            note=self.find_text(inv, "InvoiceDescription") or None,
            seller_name=seller_name,
            seller_tax_id=seller_tax_id,
            source_file=self.invoice_path,
            vat_amount=decimal_or_none(self.find_path_text(inv, "TaxesOutputs/Tax/TaxAmount/TotalAmount")),
            inclusive=False,
            lines=[]
        )
        lines_xml = inv.findall(".//{*}InvoiceLine")
        logging.info(f"Factura {invoice_data.invoice_number}: {len(lines_xml)} lineas")
        for idx, line_xml in enumerate(lines_xml, start=1):
            invoice_data.lines.append(self._parse_line(line_xml, idx))
        return invoice_data
    def _parse_line(self, line, idx):
        seq_str = self.find_text(line, "SequenceNumber", str(idx))
        try:
            seq_num = int(seq_str)
        except ValueError:
            seq_num = idx
        raw_title = self.find_text(line, "ItemDescription") or None
        desc_isbn, desc_issn, cleaned_title = extract_identifiers_from_text(raw_title)
        title = cleaned_title or raw_title
        quantity = safe_decimal(self.find_text(line, "Quantity", "1"), Decimal("1.00"))
        gross_amount_text = self.find_path_text(line, "GrossAmount")
        unit_price_text = self.find_text(line, "UnitPriceWithoutTax")
        price = safe_decimal(gross_amount_text, Decimal("0.00")) if gross_amount_text else safe_decimal(unit_price_text, Decimal("0.00"))
        issn, isbn, vendor_ref = classify_article_code(self.find_text(line, "ArticleCode") or None)
        if not isbn and desc_isbn:
            isbn = desc_isbn
        if not issn and desc_issn:
            issn = desc_issn
        return InvoiceLineData(
            line_number=seq_num,
            po_line=None,
            vendor_ref=vendor_ref,
            issn=issn,
            isbn=isbn,
            title=title,
            price=price,
            quantity=quantity,
            vat_amount=decimal_or_none(self.find_path_text(line, "TaxesOutputs/Tax/TaxAmount/TotalAmount")),
            vat_percentage=decimal_or_none(self.find_path_text(line, "TaxesOutputs/Tax/TaxRate"))
        )

class AlmaLineEnricher:
    def __init__(self, lookup, manual_assignments=None):
        self.lookup = lookup
        self.manual_assignments = manual_assignments or {}
        self.validation_issues = []
    def enrich(self, invoices):
        for invoice in invoices:
            for line in invoice.lines:
                self._enrich_line(invoice, line)
        return self.validation_issues
    def _clear_alma_fields(self, line):
        line.po_line = None
        line.reporting_code = None
        line.secondary_reporting_code = None
        line.tertiary_reporting_code = None
        line.fourth_reporting_code = None
        line.fifth_reporting_code = None
        line.funds = []
    def _enrich_line(self, invoice, line):
        source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
        manual_po_line = self.manual_assignments.get((source_key, line.line_number))
        if manual_po_line:
            manual_record = self.lookup.find_by_po_line(manual_po_line)
            if manual_record:
                self._resolve_candidates(invoice, line, [manual_record], "MANUAL")
                self.validation_issues.append(ValidationIssue(
                    "ASIGNACION MANUAL", invoice.invoice_number, line.line_number,
                    line.isbn, line.issn, line.title, [manual_po_line],
                    "PO Line asignada manualmente durante esta sesion; el XML original no ha sido modificado.",
                    po_line=manual_po_line,
                    invoice_quantity=line.quantity,
                    alma_quantity_for_pricing=manual_record.quantity_for_pricing,
                ))
                return
        if line.isbn:
            c = self.lookup.find_by_isbn(line.isbn)
            if c:
                self._resolve_candidates(invoice, line, c, "ISBN")
                return
        if line.issn:
            c = self.lookup.find_by_issn(line.issn)
            if c:
                self._resolve_candidates(invoice, line, c, "ISSN")
                return
        if line.title:
            result = self.lookup.find_by_title(line.title)
            if result["status"] == "MATCH":
                self._resolve_candidates(invoice, line, result["candidates"], "TITLE")
                return
            if result["status"] == "AMBIGUOUS":
                candidate_po_lines = [r.po_line for r in result["candidates"]]
                score_text = "; ".join(f"{r.po_line}: {s:.3f}" for r, s in zip(result["candidates"], result["scores"]))
                title_method = result.get("method") or "SIMILARITY"
                self._clear_alma_fields(line)
                line.match_status = "AMBIGUOUS_TITLE"
                line.match_message = "Varias coincidencias posibles por titulo"
                line.match_candidates = candidate_po_lines
                self.validation_issues.append(ValidationIssue(
                    "TITULO AMBIGUO",
                    invoice.invoice_number,
                    line.line_number,
                    line.isbn,
                    line.issn,
                    line.title,
                    candidate_po_lines,
                    f"Varias PO Lines candidatas por titulo ({title_method}). Puntuaciones: {score_text}",
                ))
                return
        self._clear_alma_fields(line)
        line.match_status = "NO_MATCH"
        line.match_message = "Sin coincidencia en Alma por ISBN, ISSN ni titulo"
        self.validation_issues.append(ValidationIssue("SIN COINCIDENCIA", invoice.invoice_number, line.line_number, line.isbn, line.issn, line.title, [], "No se encontro ninguna PO Line por ISBN, ISSN ni titulo."))
    def _resolve_candidates(self, invoice, line, candidates, match_type):
        unique = {c.po_line: c for c in candidates}
        unique_candidates = list(unique.values())
        if len(unique_candidates) == 1:
            r = unique_candidates[0]
            line.po_line = r.po_line
            line.reporting_code = r.reporting_code
            line.secondary_reporting_code = r.secondary_reporting_code
            line.tertiary_reporting_code = r.tertiary_reporting_code
            line.fourth_reporting_code = r.fourth_reporting_code
            line.fifth_reporting_code = r.fifth_reporting_code
            line.alma_quantity_for_pricing = r.quantity_for_pricing
            line.funds = list(r.funds or [])
            if r.title:
                line.title = r.title
            if r.start_subs_date:
                line.start_date = r.start_subs_date
            line.match_status = f"MATCH_{match_type}"
            line.match_message = f"Coincidencia unica por {match_type}"
            line.match_candidates = [r.po_line]

            if (
                r.quantity_for_pricing is not None
                and line.quantity is not None
                and line.quantity != r.quantity_for_pricing
            ):
                self.validation_issues.append(
                    ValidationIssue(
                        issue_type="REVISAR CANTIDAD",
                        invoice_number=invoice.invoice_number,
                        invoice_line_number=line.line_number,
                        isbn=line.isbn,
                        issn=line.issn,
                        title=line.title,
                        candidates=[r.po_line],
                        message=(
                            f"La factura indica Quantity={line.quantity}, mientras que "
                            f"Alma indica Quantity for Pricing={r.quantity_for_pricing}. "
                            "Puede tratarse de una facturacion parcial o de una entrega final "
                            "que complete una linea facturada previamente. "
                            "Revisar manualmente antes de cargar en Alma."
                        ),
                        po_line=r.po_line,
                        invoice_quantity=line.quantity,
                        alma_quantity_for_pricing=r.quantity_for_pricing,
                    )
                )
            return
        candidate_po_lines = sorted(r.po_line for r in unique_candidates)
        self._clear_alma_fields(line)
        line.match_status = f"DUPLICATE_{match_type}"
        line.match_message = f"Duplicado por {match_type}. Revision manual necesaria."
        line.match_candidates = candidate_po_lines
        self.validation_issues.append(ValidationIssue(f"DUPLICADO {match_type}", invoice.invoice_number, line.line_number, line.isbn, line.issn, line.title, candidate_po_lines, "El identificador tiene varias PO Lines candidatas."))

class AlmaExcelExporter:
    DATE_FORMAT = "mm/dd/yyyy"
    def __init__(self, output_path, template_path=None):
        self.output_path = output_path
        self.template_path = template_path
        self.workbook = None
        self.worksheet = None
        self.hinv_headers = []
        self.hil_headers = []
        self.hinv_style = None
        self.inv_style = None
        self.hil_style = None
        self.il_style = None
        self.column_widths = {}
    def create_or_load_workbook(self):
        if self.template_path and os.path.exists(self.template_path):
            self._load_from_template()
        else:
            self._create_without_template()
    def _load_from_template(self):
        wb = load_workbook(self.template_path)
        ws = wb.active
        hinv_row = self._find_marker_row(ws, "HINV")
        hil_row = self._find_marker_row(ws, "HIL")
        if not hinv_row or not hil_row:
            raise ValueError("La plantilla seleccionada no contiene filas HINV y HIL reconocibles.")
        inv_row = hinv_row + 1
        il_row = hil_row + 1
        self.hinv_headers = self._read_row_values(ws, hinv_row)
        self.hil_headers = self._read_row_values(ws, hil_row)
        max_cols = max(len(self.hinv_headers), len(self.hil_headers), 25)
        self.hinv_style = self._capture_row_style(ws, hinv_row, max_cols)
        self.inv_style = self._capture_row_style(ws, inv_row, max_cols)
        self.hil_style = self._capture_row_style(ws, hil_row, max_cols)
        self.il_style = self._capture_row_style(ws, il_row, max_cols)
        self.column_widths = {cl: ws.column_dimensions[cl].width for cl in ws.column_dimensions}
        self.workbook = wb
        self.worksheet = ws
        self._delete_other_sheets()
        self._clear_sheet_fully(ws)
        for cl, width in self.column_widths.items():
            if width:
                self.worksheet.column_dimensions[cl].width = width
    def _create_without_template(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "Sheet1"
        self.hinv_headers = DEFAULT_HINV_HEADERS
        self.hil_headers = DEFAULT_HIL_HEADERS
    def _find_marker_row(self, ws, marker):
        for r in range(1, ws.max_row + 1):
            if clean_cell(ws.cell(r, 1).value) == marker:
                return r
        return None
    def _read_row_values(self, ws, row_idx):
        last_col = 0
        for c in range(1, ws.max_column + 1):
            if clean_cell(ws.cell(row_idx, c).value) is not None:
                last_col = c
        return [clean_cell(ws.cell(row_idx, c).value) or "" for c in range(1, last_col + 1)] if last_col else []
    def _capture_row_style(self, ws, row_idx, max_cols):
        styles = []
        for c in range(1, max_cols + 1):
            cell = ws.cell(row_idx, c)
            styles.append({"style": copy.copy(cell._style) if cell.has_style else None, "number_format": cell.number_format, "font": copy.copy(cell.font), "fill": copy.copy(cell.fill), "border": copy.copy(cell.border), "alignment": copy.copy(cell.alignment), "protection": copy.copy(cell.protection)})
        return {"height": ws.row_dimensions[row_idx].height, "cells": styles}
    def _apply_row_style(self, row_idx, row_style, max_cols):
        if not row_style:
            return
        self.worksheet.row_dimensions[row_idx].height = row_style["height"]
        for c in range(1, min(max_cols, len(row_style["cells"])) + 1):
            info = row_style["cells"][c - 1]
            cell = self.worksheet.cell(row_idx, c)
            if info["style"] is not None:
                cell._style = copy.copy(info["style"])
            cell.number_format = info["number_format"]
            cell.font = copy.copy(info["font"])
            cell.fill = copy.copy(info["fill"])
            cell.border = copy.copy(info["border"])
            cell.alignment = copy.copy(info["alignment"])
            cell.protection = copy.copy(info["protection"])
    def _delete_other_sheets(self):
        active = self.worksheet
        for ws in list(self.workbook.worksheets):
            if ws != active:
                self.workbook.remove(ws)
    def _clear_sheet_fully(self, ws):
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
    def _write_row(self, row_idx, values, row_style):
        self._apply_row_style(row_idx, row_style, len(values))
        for c, value in enumerate(values, start=1):
            cell = self.worksheet.cell(row_idx, c)
            value = sanitize_excel_value(value)
            cell.value = value
            if isinstance(value, datetime):
                cell.number_format = self.DATE_FORMAT
    def _find_header_col(self, headers, target_header):
        target = normalize_header(target_header)
        for idx, header in enumerate(headers, start=1):
            if normalize_header(header) == target:
                return idx
        return None
    def _apply_line_type_dropdown(self, start_row, end_row):
        if end_row < start_row:
            return
        col = self._find_header_col(self.hil_headers, "Line type")
        if not col:
            return
        validation = DataValidation(type="list", formula1='"' + ",".join(LINE_TYPE_OPTIONS) + '"', allow_blank=False)
        validation.error = "Selecciona un tipo de linea valido para Alma."
        validation.errorTitle = "Tipo de linea no valido"
        validation.prompt = "Selecciona un tipo de linea."
        validation.promptTitle = "Line type"
        self.worksheet.add_data_validation(validation)
        validation.add(f"{get_column_letter(col)}{start_row}:{get_column_letter(col)}{end_row}")
    def write_invoices(self, invoices):
        if self.workbook is None or self.worksheet is None:
            self.create_or_load_workbook()
        row = 1
        for invoice in invoices:
            self._write_row(row, self.hinv_headers, self.hinv_style); row += 1
            self._write_row(row, self._build_invoice_row_from_headers(invoice, self.hinv_headers), self.inv_style); row += 1
            self._write_row(row, self.hil_headers, self.hil_style); row += 1
            first_il_row = row
            for line in invoice.lines:
                self._write_row(row, self._build_line_row_from_headers(line, self.hil_headers), self.il_style)
                row += 1
            self._apply_line_type_dropdown(first_il_row, row - 1)
    def _build_invoice_row_from_headers(self, invoice, headers):
        values = {"HINV": "INV", "Invoice Number": invoice.invoice_number, "Date": invoice.date, "Total amount": float(invoice.total_amount), "Currency": invoice.currency, "Vendor Account": invoice.vendor_account, "Payment method": None, "Note": invoice.note, "Use Pro Rata": None, "Shipment amount": None, "Overhead amount": None, "Insurance amount": None, "Discount amount": None, "VAT Percentage": None, "VAT Amount": decimal_to_float_or_none(invoice.vat_amount), "Inclusive": invoice.inclusive, "Expended From Fund": None, "PrePaid": None, "Payment status": None, "Voucher number": None, "Voucher date": None, "Voucher Amount": None, "Invoice Reference Number": None, "VAT In Invoice Line Level": "YES", "Report TAX": None, "VAT code": None}
        return [values.get(normalize_header(h), None) for h in headers]
    def _build_line_row_from_headers(self, line, headers):
        values = {"HIL": "IL", "Line type": line.line_type or "REGULAR", "Line Number": line.line_number, "PO Line": line.po_line, "Vendor Ref Number": line.vendor_ref, "ISSN": line.issn, "ISBN": line.isbn, "Title": line.title, "Price": float(line.price), "Price Note": None, "Quantity": float(line.quantity), "Reporting Code": line.reporting_code, "Secondary Reporting Code": line.secondary_reporting_code, "Tertiary Reporting Code": line.tertiary_reporting_code, "Fourth Reporting Code": line.fourth_reporting_code, "Fifth Reporting Code": line.fifth_reporting_code, "VAT code": None, "Start subs date": line.start_date, "End subs date": line.end_date, "Additional Info": None, "Note": line.note, "VAT Amount": decimal_to_float_or_none(line.vat_amount), "VAT Percentage": decimal_to_float_or_none(line.vat_percentage)}
        row = []
        fund_index = 0
        for h in headers:
            nh = normalize_header(h)
            if nh == "Fund and percent":
                row.append(line.funds[fund_index] if fund_index < len(line.funds) else None)
                fund_index += 1
            else:
                row.append(values.get(nh, None))
        return row
    def save(self):
        try:
            self.workbook.save(self.output_path)
            logging.info(f"Excel Alma guardado correctamente: {self.output_path}")
        except PermissionError:
            logging.exception("No se puede guardar el Excel porque esta abierto o bloqueado")
            raise ValueError(f"No se pudo guardar el archivo '{self.output_path}'. Puede que este abierto en Excel.")
        except Exception as e:
            logging.exception("Error guardando Excel")
            raise ValueError(f"Error al guardar el Excel: {e}")

class ValidationReportExporter:
    def __init__(self, output_path):
        self.output_path = output_path
    def save(self, issues):
        if not issues:
            return None
        wb = Workbook()
        ws = wb.active
        ws.title = "Validacion"
        ws.append([
            "Tipo aviso", "Factura", "Linea factura", "PO Line",
            "ISBN XSIG/XML", "ISSN XSIG/XML", "Titulo",
            "Cantidad factura", "Quantity for Pricing Alma",
            "PO Lines candidatas", "Mensaje"
        ])
        for issue in issues:
            ws.append([sanitize_excel_value(v) for v in [
                issue.issue_type,
                issue.invoice_number,
                issue.invoice_line_number,
                issue.po_line,
                issue.isbn,
                issue.issn,
                issue.title,
                decimal_to_float_or_none(issue.invoice_quantity),
                decimal_to_float_or_none(issue.alma_quantity_for_pricing),
                "; ".join(issue.candidates),
                issue.message,
            ]])
        wb.save(self.output_path)
        logging.info(f"Informe de validacion guardado: {self.output_path}")
        return self.output_path

def parse_invoice_paths(invoice_paths, invoice_number_overrides=None):
    if isinstance(invoice_paths, str):
        invoice_paths = [invoice_paths]
    invoices = []
    invoice_number_overrides = invoice_number_overrides or {}
    for p in invoice_paths:
        parsed = FacturaeParser(p).parse()
        override = invoice_number_overrides.get(os.path.normcase(os.path.abspath(p)))
        if override:
            for invoice in parsed:
                invoice.invoice_number = override
        invoices.extend(parsed)
    return invoices

def validate_single_vendor_batch(invoices):
    keys = {}
    for inv in invoices:
        key = get_vendor_batch_key(inv)
        label = inv.vendor_account or inv.seller_name or "Proveedor no identificado"
        keys.setdefault(key, set()).add(label)
    meaningful_keys = {k for k in keys if k}
    if len(meaningful_keys) <= 1:
        return
    details = []
    for labels in keys.values():
        for label in sorted(labels):
            details.append(f"- {label}")
    raise ValueError("El lote contiene facturas de varios proveedores. Selecciona unicamente facturas del mismo proveedor.\n\nProveedores detectados:\n" + "\n".join(details))

def _normalized_excel_header(value):
    return normalize_header(clean_cell(value) or "")


def apply_saved_edits_to_excel(output_path, invoices, line_overrides, invoice_header_overrides, invoice_number_overrides):
    """Apply saved edits to the generated workbook and return verification targets."""
    line_overrides = line_overrides or {}
    invoice_header_overrides = invoice_header_overrides or {}
    invoice_number_overrides = invoice_number_overrides or {}
    expected = []
    if not (line_overrides or invoice_header_overrides or invoice_number_overrides):
        return expected

    workbook = load_workbook(output_path)
    worksheet = workbook.active
    marker_rows = [
        row for row in range(1, worksheet.max_row + 1)
        if _normalized_excel_header(worksheet.cell(row=row, column=1).value) == "HINV"
    ]
    if len(marker_rows) < len(invoices):
        raise ValueError(
            f"No se pudieron localizar todos los bloques HINV del Excel final: "
            f"{len(marker_rows)} encontrados para {len(invoices)} facturas."
        )

    header_map = {
        "invoice_number": "Invoice Number", "date": "Date", "total_amount": "Total amount",
        "currency": "Currency", "vendor_account": "Vendor Account", "note": "Note",
        "vat_amount": "VAT Amount", "inclusive": "Inclusive",
    }
    line_map = {
        "line_type": "Line type", "po_line": "PO Line", "vendor_ref": "Vendor Ref Number",
        "issn": "ISSN", "isbn": "ISBN", "title": "Title", "price": "Price",
        "quantity": "Quantity", "reporting_code": "Reporting Code",
        "secondary_reporting_code": "Secondary Reporting Code",
        "tertiary_reporting_code": "Tertiary Reporting Code",
        "fourth_reporting_code": "Fourth Reporting Code", "fifth_reporting_code": "Fifth Reporting Code",
        "start_date": "Start subs date", "end_date": "End subs date", "note": "Note",
        "vat_amount": "VAT Amount", "vat_percentage": "VAT Percentage",
    }

    def header_columns(row_index):
        result = {}
        for column in range(1, worksheet.max_column + 1):
            normalized = _normalized_excel_header(worksheet.cell(row=row_index, column=column).value)
            if normalized:
                result.setdefault(normalized, []).append(column)
        return result

    for invoice_index, invoice in enumerate(invoices):
        hinv_row = marker_rows[invoice_index]
        block_end = marker_rows[invoice_index + 1] - 1 if invoice_index + 1 < len(marker_rows) else worksheet.max_row
        inv_row = hinv_row + 1
        source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
        hinv_columns = header_columns(hinv_row)
        header_changes = dict(invoice_header_overrides.get(source_key, {}))
        if source_key in invoice_number_overrides:
            header_changes["invoice_number"] = invoice_number_overrides[source_key]
        for field, value in header_changes.items():
            normalized = normalize_header(header_map.get(field, ""))
            columns = hinv_columns.get(normalized, [])
            if not columns:
                raise ValueError(f"No existe la columna de cabecera necesaria para guardar '{field}'.")
            cell = worksheet.cell(row=inv_row, column=columns[0])
            cell.value = sanitize_excel_value(value)
            expected.append((worksheet.title, inv_row, columns[0], cell.value, f"cabecera {field}"))

        hil_row = next((
            row for row in range(inv_row + 1, block_end + 1)
            if _normalized_excel_header(worksheet.cell(row=row, column=1).value) == "HIL"
        ), None)
        if hil_row is None:
            raise ValueError(f"No se encontro el bloque HIL de la factura {invoice.invoice_number}.")
        hil_columns = header_columns(hil_row)
        il_rows = [
            row for row in range(hil_row + 1, block_end + 1)
            if _normalized_excel_header(worksheet.cell(row=row, column=1).value) == "IL"
        ]
        line_number_cols = hil_columns.get(normalize_header("Line Number"), [])
        by_number = {}
        if line_number_cols:
            for row in il_rows:
                try:
                    by_number[int(worksheet.cell(row=row, column=line_number_cols[0]).value)] = row
                except (TypeError, ValueError):
                    pass
        by_order = {line.line_number: row for line, row in zip(invoice.lines, il_rows)}

        for (override_source, line_number), changes in line_overrides.items():
            if override_source != source_key:
                continue
            row = by_number.get(line_number) or by_order.get(line_number)
            if row is None:
                raise ValueError(f"No se encontro la linea {line_number} de la factura {invoice.invoice_number} en el Excel final.")
            for field, value in changes.items():
                if field in {"fund_1", "fund_2"}:
                    columns = hil_columns.get(normalize_header("Fund and percent"), [])
                    occurrence = int(field.split("_", 1)[1]) - 1
                    if occurrence >= len(columns):
                        raise ValueError(f"No existe la columna necesaria para guardar '{field}'.")
                    column = columns[occurrence]
                else:
                    normalized = normalize_header(line_map.get(field, ""))
                    columns = hil_columns.get(normalized, [])
                    if not columns:
                        raise ValueError(f"No existe la columna necesaria para guardar '{field}'.")
                    column = columns[0]
                if field in {"price", "quantity", "vat_amount", "vat_percentage"} and value not in (None, ""):
                    value = float(safe_decimal(value))
                elif field in {"start_date", "end_date"} and value:
                    value = parse_date_for_excel(value)
                cell = worksheet.cell(row=row, column=column)
                cell.value = sanitize_excel_value(value)
                expected.append((worksheet.title, row, column, cell.value, f"linea {line_number} {field}"))

    audit_name = "Cambios FacturaeToAlma"
    if audit_name in workbook.sheetnames:
        del workbook[audit_name]
    audit = workbook.create_sheet(audit_name)
    audit.append(["Tipo", "Factura/Fichero", "Linea", "Campo", "Valor guardado"])
    for invoice in invoices:
        source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
        if source_key in invoice_number_overrides:
            audit.append(["Cabecera", os.path.basename(invoice.source_file or ""), "", "Invoice Number", invoice_number_overrides[source_key]])
        for field, value in invoice_header_overrides.get(source_key, {}).items():
            audit.append(["Cabecera", os.path.basename(invoice.source_file or ""), "", field, value])
        for (override_source, line_number), changes in line_overrides.items():
            if override_source != source_key:
                continue
            for field, value in changes.items():
                audit.append(["Linea", os.path.basename(invoice.source_file or ""), line_number, field, value])
    audit.freeze_panes = "A2"
    workbook.save(output_path)
    return expected


def verify_saved_edits(output_path, expected):
    workbook = load_workbook(output_path, read_only=True, data_only=False)
    failures = []
    try:
        for sheet_name, row, column, expected_value, label in expected:
            actual = workbook[sheet_name].cell(row=row, column=column).value
            if isinstance(expected_value, datetime) and isinstance(actual, datetime):
                matches = expected_value == actual
            else:
                matches = str(actual if actual is not None else "") == str(expected_value if expected_value is not None else "")
            if not matches:
                failures.append(f"{label}: esperado={expected_value!r}, obtenido={actual!r}")
    finally:
        workbook.close()
    if failures:
        raise ValueError("El Excel generado no conserva todos los cambios:\n" + "\n".join(failures[:20]))


def convert_facturae_to_alma_excel(invoice_paths, alma_pol_excel_path, excel_output_path, template_path=None, invoice_number_overrides=None, manual_assignments=None, line_overrides=None, invoice_header_overrides=None):
    invoices = parse_invoice_paths(invoice_paths, invoice_number_overrides)
    if not invoices:
        raise ValueError("No se encontro ninguna factura para procesar.")
    validate_single_vendor_batch(invoices)
    lookup = AlmaPOLineLookup(alma_pol_excel_path)
    lookup.load()
    enricher = AlmaLineEnricher(lookup, manual_assignments)
    validation_issues = enricher.enrich(invoices)
    line_overrides = line_overrides or {}
    invoice_header_overrides = invoice_header_overrides or {}
    for invoice in invoices:
        source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
        for field_name, value in invoice_header_overrides.get(source_key, {}).items():
            if field_name == "total_amount":
                value = safe_decimal(value, invoice.total_amount)
            elif field_name == "vat_amount":
                value = safe_decimal(value, invoice.vat_amount)
            elif field_name == "date":
                value = parse_date_for_excel(value) or invoice.date
            elif field_name == "inclusive":
                value = str(value).strip().lower() in {"true", "1", "yes", "si", "sí"}
            if hasattr(invoice, field_name):
                setattr(invoice, field_name, value)
    editable_fields = set(EDITABLE_LINE_FIELDS)

    for invoice in invoices:
        source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
        for line in invoice.lines:
            changes = line_overrides.get((source_key, line.line_number), {})
            for field_name, value in changes.items():
                if field_name not in editable_fields:
                    continue
                if field_name in {"price", "quantity", "vat_amount", "vat_percentage"}:
                    value = safe_decimal(value, getattr(line, field_name, Decimal("0.00")))
                elif field_name in {"start_date", "end_date"}:
                    value = parse_date_for_excel(value)
                elif field_name == "line_type":
                    value = value if value in LINE_TYPE_OPTIONS else "REGULAR"
                if field_name.startswith("fund_"):
                    fund_index = int(field_name.split("_", 1)[1]) - 1
                    while len(line.funds) <= fund_index:
                        line.funds.append(None)
                    line.funds[fund_index] = value or None
                else:
                    setattr(line, field_name, value or None)
    output_directory = os.path.dirname(os.path.abspath(excel_output_path)) or os.getcwd()
    os.makedirs(output_directory, exist_ok=True)
    temp_handle, temporary_output = tempfile.mkstemp(prefix="FacturaeToAlma_", suffix=".xlsx", dir=output_directory)
    os.close(temp_handle)
    try:
        os.remove(temporary_output)
    except FileNotFoundError:
        pass
    exporter = AlmaExcelExporter(output_path=temporary_output, template_path=template_path)
    exporter.create_or_load_workbook()
    exporter.write_invoices(invoices)
    exporter.save()
    expected_edits = apply_saved_edits_to_excel(
        temporary_output, invoices, line_overrides, invoice_header_overrides, invoice_number_overrides
    )
    requested_edits = len(invoice_number_overrides or {}) + sum(len(v) for v in (invoice_header_overrides or {}).values()) + sum(len(v) for v in (line_overrides or {}).values())
    if requested_edits and len(expected_edits) != requested_edits:
        raise ValueError(
            f"No se pudieron transferir todos los cambios al Excel: {requested_edits} solicitados, {len(expected_edits)} escritos."
        )
    verify_saved_edits(temporary_output, expected_edits)
    os.replace(temporary_output, excel_output_path)
    logging.info("Excel final verificado y reemplazado atomicamente: %s", excel_output_path)
    validation_path = None
    if validation_issues:
        validation_path = build_validation_output_path(excel_output_path)
        ValidationReportExporter(validation_path).save(validation_issues)
    summary = build_conversion_summary(invoices, validation_issues)
    summary["verified_editor_changes"] = len(expected_edits)
    return len(invoices), len(validation_issues), validation_path, summary

class FacturaeApp:
    DEFAULT_INSTITUTION = "Universidad de Valladolid. Biblioteca"

    def __init__(self, root):
        self.root = root
        self.config_path = self._get_config_path()
        self.preferences = self._load_preferences()
        self.conversion_mode = tk.StringVar(value="single")
        self.invoice_file = tk.StringVar()
        self.alma_pol_file = tk.StringVar()
        self.template_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.status_text = tk.StringVar(value="Estado: pendiente de seleccionar archivos.")
        self.institution_var = tk.StringVar(value=self.preferences.get("institution", self.DEFAULT_INSTITUTION))
        self.custom_invoice_dir = tk.StringVar(value=self.preferences.get("custom_invoice_dir", ""))
        self.custom_alma_dir = tk.StringVar(value=self.preferences.get("custom_alma_dir", ""))
        self.custom_template_dir = tk.StringVar(value=self.preferences.get("custom_template_dir", ""))
        self.custom_output_dir = tk.StringVar(value=self.preferences.get("custom_output_dir", ""))
        self.logo_path_var = tk.StringVar(value=self.preferences.get("logo_path", ""))
        self.logo_photo = None
        self.invoice_files = []
        self.invoice_metadata = {}
        self.invoice_number_overrides = {}
        self.manual_po_assignments = {}
        self.line_overrides = {}
        self.invoice_header_overrides = {}
        self.edit_original_values = {}
        self.edit_header_original_values = {}
        self.edit_invoice_by_label = {}
        self.edit_combo_editor = None
        self.edit_combo_context = None
        self.header_cell_editor = None
        self.header_editor_context = None
        self.edit_draft_signature = None
        self.edit_changes_saved = False
        self.saved_line_overrides = {}
        self.saved_header_overrides = {}
        self.saved_number_overrides = {}
        self.saved_manual_assignments = {}
        self.viewer_path_by_label = {}
        self.context_widget = None
        self.worker_queue = queue.Queue()
        self.worker_thread = None
        remembered_template = self.preferences.get("last_template_path", "")
        if remembered_template and os.path.isfile(remembered_template):
            self.template_file.set(remembered_template)
        self.root.geometry("1180x940")
        self.root.minsize(1000, 820)
        self.root.resizable(True, True)
        self.root.configure(bg=UI_BG)
        self._setup_style()
        self._build_ui()
        self._build_global_context_menu()
        self._update_title()
        self._load_logo_from_preferences(show_warning=False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _application_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _get_config_path(self):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        folder = os.path.join(base, APP_NAME)
        try:
            os.makedirs(folder, exist_ok=True)
            return os.path.join(folder, "config.json")
        except OSError:
            return os.path.join(self._application_dir(), "facturae_to_alma_config.json")

    def _default_preferences(self):
        return {
            "institution": self.DEFAULT_INSTITUTION,
            "custom_invoice_dir": "",
            "custom_alma_dir": "",
            "custom_template_dir": "",
            "custom_output_dir": "",
            "logo_path": "",
            "last_invoice_dir": "",
            "last_alma_dir": "",
            "last_template_dir": "",
            "last_output_dir": "",
            "last_template_path": "",
        }

    def _load_preferences(self):
        data = self._default_preferences()
        try:
            if os.path.isfile(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    data.update(loaded)
        except (OSError, json.JSONDecodeError) as e:
            logging.warning(f"No se pudo leer la configuracion: {e}")
        return data

    def _save_preferences_file(self):
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.preferences, f, ensure_ascii=False, indent=2)
        except OSError as e:
            logging.exception("No se pudo guardar la configuracion")
            raise ValueError(f"No se pudo guardar la configuracion: {e}")

    def _valid_dir(self, value):
        return value if value and os.path.isdir(value) else ""

    def _initial_dir(self, custom_key, last_key, fallback=None):
        custom = self._valid_dir(self.preferences.get(custom_key, ""))
        if custom:
            return custom
        last = self._valid_dir(self.preferences.get(last_key, ""))
        if last:
            return last
        if fallback and os.path.isdir(fallback):
            return fallback
        return self._application_dir()

    def _remember_dir(self, key, path):
        folder = path if os.path.isdir(path) else os.path.dirname(path)
        if folder and os.path.isdir(folder):
            self.preferences[key] = folder
            try:
                self._save_preferences_file()
            except ValueError:
                logging.warning("No se pudo persistir el ultimo directorio")

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError as exc:
            logging.warning("No se pudo aplicar el tema grafico 'clam': %s", exc)
        style.configure("App.TFrame", background=UI_BG)
        style.configure("App.TLabel", background=UI_BG)
        style.configure("App.TLabelframe", background=UI_BG)
        style.configure("App.TLabelframe.Label", background=UI_BG)
        style.configure("App.TRadiobutton", background=UI_BG)
        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=(20, 12))
        style.configure("HelpTitle.TLabel", background=UI_BG, font=("Segoe UI", 12, "bold"))
        style.configure("Treeview", rowheight=25)

    def _update_title(self, *_):
        institution = self.institution_var.get().strip() or self.DEFAULT_INSTITUTION
        title = f"{institution} - {APP_NAME} {APP_VERSION}"
        self.root.title(title)
        if hasattr(self, "title_label"):
            self.title_label.config(text=title)

    def _build_ui(self):
        shell = ttk.Frame(self.root, padding=16, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        self.header_frame = ttk.Frame(shell, style="App.TFrame")
        self.header_frame.pack(fill=tk.X, pady=(0, 12))
        self.logo_label = ttk.Label(self.header_frame, style="App.TLabel")
        self.logo_label.pack(anchor="w", pady=(0, 8))
        self.title_label = ttk.Label(
            self.header_frame,
            text="",
            font=("Segoe UI", 15, "bold"),
            style="App.TLabel",
        )
        self.title_label.pack(fill=tk.X, anchor="w")
        self.notebook = ttk.Notebook(shell)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.convert_tab = ttk.Frame(self.notebook, padding=14, style="App.TFrame")
        self.visualize_tab = ttk.Frame(self.notebook, padding=12, style="App.TFrame")
        self.edit_tab = ttk.Frame(self.notebook, padding=12, style="App.TFrame")
        self.customize_tab = ttk.Frame(self.notebook, padding=18, style="App.TFrame")
        self.help_tab = ttk.Frame(self.notebook, padding=18, style="App.TFrame")
        self.notebook.add(self.convert_tab, text="Convertir")
        self.notebook.add(self.visualize_tab, text="Visualizar")
        self.notebook.add(self.edit_tab, text="Editar")
        self.notebook.add(self.customize_tab, text="Personalizar")
        self.notebook.add(self.help_tab, text="Ayuda")
        self._build_convert_tab()
        self._build_visualize_tab()
        self._build_edit_tab()
        self._build_customize_tab()
        self._build_help_tab()

    def _build_convert_tab(self):
        tab = self.convert_tab
        tab.columnconfigure(0, weight=1)
        mode = ttk.LabelFrame(tab, text="1. Modo de conversion", padding=10, style="App.TLabelframe")
        mode.grid(row=0, column=0, sticky="ew")
        ttk.Radiobutton(mode, text="Factura individual", variable=self.conversion_mode, value="single", command=self.on_mode_change, style="App.TRadiobutton").grid(row=0, column=0, padx=(0, 24), sticky="w")
        ttk.Radiobutton(mode, text="Lote de facturas del mismo proveedor y biblioteca", variable=self.conversion_mode, value="batch", command=self.on_mode_change, style="App.TRadiobutton").grid(row=0, column=1, sticky="w")

        input_frame = ttk.LabelFrame(tab, text="2. Archivos de entrada", padding=10, style="App.TLabelframe")
        input_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        input_frame.columnconfigure(0, weight=1)
        self.invoice_label = ttk.Label(input_frame, text="Factura electronica XSIG / XML / TXT *", style="App.TLabel")
        self.invoice_label.grid(row=0, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.invoice_file, state="readonly").grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(2, 8))
        self.invoice_button = ttk.Button(input_frame, text="Seleccionar factura...", command=self.select_invoice_or_batch)
        self.invoice_button.grid(row=1, column=1, sticky="ew", pady=(2, 8))
        ttk.Label(input_frame, text="Fichero Excel de Alma Analytics con PO Lines *", style="App.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.alma_pol_file).grid(row=3, column=0, sticky="ew", padx=(0, 10), pady=(2, 8))
        ttk.Button(input_frame, text="Seleccionar fichero...", command=self.select_alma_pol).grid(row=3, column=1, sticky="ew", pady=(2, 8))
        ttk.Label(input_frame, text="Plantilla Alma .xlsx *", style="App.TLabel").grid(row=4, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.template_file).grid(row=5, column=0, sticky="ew", padx=(0, 10), pady=(2, 0))
        ttk.Button(input_frame, text="Seleccionar plantilla...", command=self.select_template).grid(row=5, column=1, sticky="ew", pady=(2, 0))

        self.batch_frame = ttk.LabelFrame(tab, text="Facturas del lote", padding=8, style="App.TLabelframe")
        self.batch_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        self.batch_frame.columnconfigure(0, weight=1)
        self.batch_frame.rowconfigure(0, weight=1)
        cols = ("invoice", "vendor", "file")
        self.batch_tree = ttk.Treeview(self.batch_frame, columns=cols, show="headings", height=10, selectmode="extended")
        self.batch_tree.heading("invoice", text="Numero de factura")
        self.batch_tree.heading("vendor", text="Proveedor")
        self.batch_tree.heading("file", text="Fichero")
        self.batch_tree.column("invoice", width=150, anchor="w")
        self.batch_tree.column("vendor", width=310, anchor="w")
        self.batch_tree.column("file", width=520, anchor="w")
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(self.batch_frame, orient="vertical", command=self.batch_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.batch_tree.configure(yscrollcommand=scroll.set)
        batch_buttons = ttk.Frame(self.batch_frame, style="App.TFrame")
        batch_buttons.grid(row=1, column=0, columnspan=2, sticky="w", pady=(7, 0))
        ttk.Button(batch_buttons, text="Anadir facturas", command=self.add_batch_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(batch_buttons, text="Quitar seleccionadas", command=self.remove_selected_batch_files).pack(side=tk.LEFT, padx=6)
        ttk.Button(batch_buttons, text="Ordenar por numero", command=self.sort_batch_files).pack(side=tk.LEFT, padx=6)
        ttk.Button(batch_buttons, text="Vaciar lote", command=self.clear_batch_files).pack(side=tk.LEFT, padx=6)

        output = ttk.LabelFrame(tab, text="3. Archivo de salida", padding=10, style="App.TLabelframe")
        output.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        output.columnconfigure(0, weight=1)
        ttk.Entry(output, textvariable=self.output_file).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(output, text="Elegir destino...", command=self.select_output).grid(row=0, column=1)

        actions = ttk.LabelFrame(tab, text="4. Acciones", padding=10, style="App.TLabelframe")
        actions.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=2)
        for c in (1, 2, 3): actions.columnconfigure(c, weight=1)
        self.convert_button = ttk.Button(actions, text="Convertir a formato Alma", command=self.run_conversion, style="Primary.TButton")
        self.convert_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(actions, text="Limpiar datos", command=self.clear_fields).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="Abrir carpeta de salida", command=self.open_output_folder).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(actions, text="Salir", command=self._on_close).grid(row=0, column=3, sticky="ew", padx=(4, 0))
        ttk.Label(tab, textvariable=self.status_text, style="App.TLabel", font=("Segoe UI", 9, "italic")).grid(row=5, column=0, sticky="w", pady=(10, 0))
        tab.rowconfigure(3, weight=1)
        self.on_mode_change(initial=True)

    def _build_visualize_tab(self):
        tab = self.visualize_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        controls = ttk.Frame(tab, style="App.TFrame")
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        controls.columnconfigure(1, weight=1)
        ttk.Label(controls, text="Factura:", style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=(0, 8))
        self.viewer_invoice_var = tk.StringVar()
        self.viewer_selector = ttk.Combobox(controls, textvariable=self.viewer_invoice_var, state="readonly")
        self.viewer_selector.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.viewer_selector.bind("<<ComboboxSelected>>", self._on_viewer_invoice_selected)
        self.invoice_viewer = FacturaeViewerTab(tab, show_loader=False)
        self.invoice_viewer.grid(row=1, column=0, sticky="nsew")
        ttk.Label(tab, text="Vista de consulta de la factura original. Los cambios del Excel de Alma se realizan en la pestaña Editar.", style="App.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))

    def _refresh_viewer_selector(self, select_path=None):
        labels = []
        self.viewer_path_by_label = {}
        for path in self.invoice_files:
            meta = self.invoice_metadata.get(path, {})
            number = meta.get("number", "(sin numero)")
            vendor = meta.get("vendor", "")
            label = f"{number} | {vendor} | {os.path.basename(path)}"
            labels.append(label)
            self.viewer_path_by_label[label] = path
        self.viewer_selector["values"] = labels
        if not labels:
            self.viewer_invoice_var.set("")
            return
        target = next((label for label, path in self.viewer_path_by_label.items() if path == select_path), labels[0])
        self.viewer_invoice_var.set(target)
        self._load_active_viewer_invoice()

    def _active_viewer_path(self):
        return self.viewer_path_by_label.get(self.viewer_invoice_var.get())

    def _on_viewer_invoice_selected(self, _event=None):
        self._load_active_viewer_invoice()

    def _load_active_viewer_invoice(self):
        path = self._active_viewer_path()
        if not path:
            return
        self.invoice_viewer.load_file(path)
        if self.invoice_viewer.current_data:
            effective = self.invoice_metadata.get(path, {}).get("number")
            if effective:
                self.invoice_viewer.current_data["invoice_number"] = effective
            self.invoice_viewer._render_pdf_view(self.invoice_viewer.current_data)

    def edit_active_invoice_number(self):
        path = self._active_viewer_path()
        if not path:
            messagebox.showinfo("Visualizar", "No hay ninguna factura activa.")
            return
        original = self.invoice_metadata.get(path, {}).get("original_number") or self.invoice_metadata.get(path, {}).get("number", "")
        dialog = tk.Toplevel(self.root)
        dialog.title("Corregir numero de factura en memoria")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(1, weight=1)
        ttk.Label(dialog, text="Es probable que el numero pertenezca a una secuencia que se reinicia periodicamente.\nPuedes anadir temporalmente un prefijo o sufijo para desambiguar.\nEl XML original no se modificara.", wraplength=520, justify="left").grid(row=0, column=0, columnspan=2, padx=16, pady=12, sticky="w")
        ttk.Label(dialog, text=f"Numero original: {original}").grid(row=1, column=0, columnspan=2, padx=16, sticky="w")
        current_year = str(datetime.now().year)
        prefix = tk.StringVar(value=f"[{current_year}]")
        suffix = tk.StringVar()
        preview = tk.StringVar(value=f"[{current_year}]{original}")
        def update(*_): preview.set(f"{prefix.get()}{original}{suffix.get()}")
        prefix.trace_add("write", update); suffix.trace_add("write", update)
        ttk.Label(dialog, text="Prefijo").grid(row=2, column=0, padx=16, pady=6, sticky="w")
        ttk.Entry(dialog, textvariable=prefix).grid(row=2, column=1, padx=16, pady=6, sticky="ew")
        ttk.Label(dialog, text="Sufijo").grid(row=3, column=0, padx=16, pady=6, sticky="w")
        ttk.Entry(dialog, textvariable=suffix).grid(row=3, column=1, padx=16, pady=6, sticky="ew")
        ttk.Label(dialog, textvariable=preview, font=("Segoe UI", 10, "bold")).grid(row=4, column=0, columnspan=2, padx=16, pady=10)
        def accept():
            value = preview.get().strip()
            if not value: return
            key = os.path.normcase(os.path.abspath(path))
            self.invoice_number_overrides[key] = value
            self.invoice_metadata[path]["number"] = value
            self._refresh_batch_tree(); self._refresh_viewer_selector(select_path=path)
            dialog.destroy()
        buttons = ttk.Frame(dialog); buttons.grid(row=5, column=0, columnspan=2, pady=(0, 14))
        ttk.Button(buttons, text="Aplicar en esta sesion", command=accept).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons, text="Cancelar", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def reset_active_invoice_number(self):
        path = self._active_viewer_path()
        if not path: return False
        original = self.invoice_metadata.get(path, {}).get("original_number")
        if original:
            self.invoice_metadata[path]["number"] = original
        self.invoice_number_overrides.pop(os.path.normcase(os.path.abspath(path)), None)
        self._refresh_batch_tree(); self._refresh_viewer_selector(select_path=path)

    def _build_edit_tab(self):
        tab = self.edit_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)

        toolbar = ttk.Frame(tab, style="App.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        toolbar.columnconfigure(1, weight=1)
        ttk.Label(toolbar, text="Factura:", style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=(0, 8))
        self.edit_invoice_var = tk.StringVar()
        self.edit_invoice_selector = ttk.Combobox(toolbar, textvariable=self.edit_invoice_var, state="readonly")
        self.edit_invoice_selector.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.edit_invoice_selector.bind("<<ComboboxSelected>>", lambda _event: self._refresh_edit_tree())

        number_box = ttk.LabelFrame(tab, text="Numero de factura en el Excel de Alma", padding=8, style="App.TLabelframe")
        number_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        number_box.columnconfigure(5, weight=1)
        ttk.Label(number_box, text="Original:", style="App.TLabel").grid(row=0, column=0, padx=(0, 4))
        self.edit_original_number_var = tk.StringVar()
        ttk.Label(number_box, textvariable=self.edit_original_number_var, style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, padx=(0, 12))
        self.number_mode_var = tk.StringVar(value="original")
        ttk.Radiobutton(number_box, text="Original", variable=self.number_mode_var, value="original", command=self._update_number_preview).grid(row=0, column=2)
        ttk.Radiobutton(number_box, text="Sufijo", variable=self.number_mode_var, value="suffix", command=self._update_number_preview).grid(row=0, column=3)
        ttk.Radiobutton(number_box, text="Prefijo", variable=self.number_mode_var, value="prefix", command=self._update_number_preview).grid(row=0, column=4)
        self.number_affix_var = tk.StringVar(value=f"[{datetime.now().year}]")
        self.number_affix_var.trace_add("write", self._update_number_preview)
        self.number_affix_entry = ttk.Entry(number_box, textvariable=self.number_affix_var, width=14)
        self.number_affix_entry.grid(row=0, column=5, padx=8)
        self.number_preview_var = tk.StringVar()
        ttk.Label(number_box, textvariable=self.number_preview_var, style="App.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky="w")
        ttk.Button(number_box, text="Aplicar numero", command=self.apply_edit_invoice_number).grid(row=0, column=7, padx=(8, 0))

        header_frame = ttk.LabelFrame(tab, text="Cabecera de factura (INV)", padding=6, style="App.TLabelframe")
        header_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        header_frame.columnconfigure(0, weight=1)
        self.header_columns = ("invoice_number", "date", "total_amount", "currency", "vendor_account", "note", "vat_amount", "inclusive")
        header_titles = {
            "invoice_number": "Invoice Number", "date": "Date", "total_amount": "Total amount",
            "currency": "Currency", "vendor_account": "Vendor Account", "note": "Note",
            "vat_amount": "VAT Amount", "inclusive": "Inclusive",
        }
        self.header_tree = ttk.Treeview(header_frame, columns=self.header_columns, show="headings", height=1, selectmode="browse")
        for col in self.header_columns:
            self.header_tree.heading(col, text=header_titles[col])
            self.header_tree.column(col, width=140 if col != "note" else 260, minwidth=75, anchor="w")
        self.header_tree.grid(row=0, column=0, sticky="ew")
        self.header_tree.bind("<Double-1>", self._begin_edit_header_cell)
        self.header_tree.bind("<Button-3>", self._show_header_context_menu)
        self.header_context_menu = tk.Menu(self.header_tree, tearoff=0)
        self.header_context_menu.add_command(label="Editar celda", command=self._edit_selected_header_cell)
        self.header_context_menu.add_command(label="Restablecer celda", command=self._restore_selected_header_cell)

        grid_frame = ttk.Frame(tab, style="App.TFrame")
        grid_frame.grid(row=3, column=0, sticky="nsew")
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)
        self.edit_columns = (
            "invoice", "line", "line_type", "po_line", "vendor_ref", "issn", "isbn", "title",
            "price", "quantity", "reporting_code", "secondary_reporting_code", "tertiary_reporting_code",
            "fourth_reporting_code", "fifth_reporting_code", "start_date", "end_date", "note",
            "fund_1", "fund_2", "vat_amount", "vat_percentage",
        )
        headings = {
            "invoice": "Factura", "line": "Linea", "line_type": "Line type", "po_line": "PO Line",
            "vendor_ref": "Vendor Ref Number", "issn": "ISSN", "isbn": "ISBN", "title": "Title",
            "price": "Price", "quantity": "Quantity", "reporting_code": "Reporting Code",
            "secondary_reporting_code": "Secondary Reporting Code", "tertiary_reporting_code": "Tertiary Reporting Code",
            "fourth_reporting_code": "Fourth Reporting Code", "fifth_reporting_code": "Fifth Reporting Code",
            "start_date": "Start subs date", "end_date": "End subs date", "note": "Note",
            "fund_1": "Fund and percent 1", "fund_2": "Fund and percent 2",
            "vat_amount": "VAT Amount", "vat_percentage": "VAT Percentage",
        }
        self.edit_tree = ttk.Treeview(grid_frame, columns=self.edit_columns, show="headings", selectmode="extended")
        for col in self.edit_columns:
            self.edit_tree.heading(col, text=headings[col])
            width = 80 if col in {"invoice", "line", "price", "quantity"} else 130
            if col == "title": width = 320
            self.edit_tree.column(col, width=width, minwidth=55, anchor="w")
        self.edit_tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.edit_tree.yview)
        xscroll = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.edit_tree.xview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.edit_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.edit_tree.bind("<Double-1>", self._begin_edit_cell)
        self.edit_tree.bind("<MouseWheel>", self._on_edit_mousewheel)
        self.edit_tree.bind("<Button-4>", self._on_edit_mousewheel)
        self.edit_tree.bind("<Button-5>", self._on_edit_mousewheel)
        self.edit_tree.bind("<Button-3>", self._show_edit_context_menu)
        self.edit_context_menu = tk.Menu(self.edit_tree, tearoff=0)
        self.edit_context_menu.add_command(label="Editar celda", command=self._edit_selected_cell)
        self.edit_context_menu.add_command(label="Vaciar celda", command=self._clear_selected_edit_cell)
        self.edit_context_menu.add_command(label="Restablecer celda", command=self._restore_selected_edit_cell)
        self.edit_context_menu.add_separator()
        self.edit_context_menu.add_command(label="Sugerir valor por contexto", command=self._suggest_selected_context_value)
        self.edit_context_menu.add_command(label="Aplicar valor a lineas seleccionadas", command=self._bulk_apply_selected_value)
        self.edit_context_menu.add_command(label="Rellenar blancos seleccionados", command=self._fill_selected_blanks)
        self.edit_context_menu.add_command(label="Rellenar blancos inferiores", command=self._fill_blanks_below)
        edit_bottom = ttk.Frame(tab, style="App.TFrame")
        edit_bottom.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        edit_bottom.columnconfigure(0, weight=1)
        ttk.Label(edit_bottom, text="Doble clic para editar. Pulsa Guardar cambios antes de convertir.", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(edit_bottom, text="Guardar cambios y actualizar Excel", command=self.save_edit_changes, style="Primary.TButton").grid(row=0, column=1, padx=(12, 6))
        ttk.Button(edit_bottom, text="Restablecer factura", command=self.reset_edit_invoice).grid(row=0, column=2)

    def _on_edit_mousewheel(self, event):
        if getattr(event, "num", None) == 4: units = -3
        elif getattr(event, "num", None) == 5: units = 3
        else:
            delta = getattr(event, "delta", 0)
            units = -int(delta / 120) if delta else 0
            if units == 0 and delta: units = -1 if delta > 0 else 1
        self.edit_tree.yview_scroll(units, "units")
        return "break"

    def _refresh_edit_selector(self, select_path=None):
        labels = []
        self.edit_invoice_by_label = {}
        for path in self.invoice_files:
            meta = self.invoice_metadata.get(path, {})
            label = f"{meta.get('number', '(sin numero)')} | {meta.get('vendor', '')} | {os.path.basename(path)}"
            labels.append(label); self.edit_invoice_by_label[label] = path
        self.edit_invoice_selector["values"] = labels
        if not labels:
            self.edit_invoice_var.set(""); self.edit_original_number_var.set(""); self.number_preview_var.set("")
            self._refresh_edit_tree(); return
        target = next((label for label, path in self.edit_invoice_by_label.items() if path == select_path), labels[0])
        self.edit_invoice_var.set(target)
        path = self.edit_invoice_by_label[target]
        original = self.invoice_metadata.get(path, {}).get("original_number", "")
        self.edit_original_number_var.set(original)
        self._update_number_preview(); self._refresh_edit_tree()

    def _active_edit_path(self):
        return self.edit_invoice_by_label.get(self.edit_invoice_var.get())

    def _update_number_preview(self, *_):
        original = self.edit_original_number_var.get()
        affix = self.number_affix_var.get()
        mode = self.number_mode_var.get()
        if mode == "suffix":
            value = f"{original}{affix}"
        elif mode == "prefix":
            value = f"{affix}{original}"
        else:
            value = original
        self.number_preview_var.set(value)
        self.edit_changes_saved = False
        if hasattr(self, "number_affix_entry"):
            self.number_affix_entry.configure(state="normal" if mode in {"suffix", "prefix"} else "disabled")

    def apply_edit_invoice_number(self):
        path = self._active_edit_path()
        if not path: return
        value = self.number_preview_var.get().strip()
        if not value: return
        key = os.path.normcase(os.path.abspath(path))
        original = self.invoice_metadata[path].get("original_number", value)
        if self.number_mode_var.get() == "original" or value == original:
            self.invoice_number_overrides.pop(key, None)
            value = original
        else:
            self.invoice_number_overrides[key] = value
        self.invoice_metadata[path]["number"] = value
        self._refresh_batch_tree(); self._refresh_viewer_selector(select_path=path); self._refresh_edit_selector(select_path=path)

    def _current_edit_signature(self):
        invoice_signature = tuple(os.path.normcase(os.path.abspath(path)) for path in self.invoice_files)
        return (
            invoice_signature,
            os.path.normcase(os.path.abspath(self.alma_pol_file.get().strip())) if self.alma_pol_file.get().strip() else "",
            os.path.normcase(os.path.abspath(self.template_file.get().strip())) if self.template_file.get().strip() else "",
        )

    def _invalidate_edit_draft(self, clear_overrides=True):
        self.edit_draft_signature = None
        self.edit_original_values.clear()
        self.edit_header_original_values.clear()
        if clear_overrides:
            self.line_overrides.clear()
            self.manual_po_assignments.clear()
            self.invoice_number_overrides.clear()
            self.invoice_header_overrides.clear()
            self.saved_line_overrides.clear()
            self.saved_manual_assignments.clear()
            self.saved_number_overrides.clear()
            self.saved_header_overrides.clear()
            self.edit_changes_saved = False
        if hasattr(self, "edit_tree"):
            for item in self.edit_tree.get_children():
                self.edit_tree.delete(item)
        if hasattr(self, "edit_invoice_selector"):
            self._refresh_edit_selector(select_path=self.invoice_files[0] if self.invoice_files else None)

    def _prune_edit_state_to_current_invoices(self):
        valid_sources = {os.path.normcase(os.path.abspath(path)) for path in self.invoice_files}
        self.line_overrides = {key: value for key, value in self.line_overrides.items() if key[0] in valid_sources}
        self.manual_po_assignments = {key: value for key, value in self.manual_po_assignments.items() if key[0] in valid_sources}
        self.invoice_number_overrides = {key: value for key, value in self.invoice_number_overrides.items() if key in valid_sources}
        self.invoice_header_overrides = {key: value for key, value in self.invoice_header_overrides.items() if key in valid_sources}

    def prepare_edit_draft(self, silent=False):
        try:
            invoice_paths, alma_path, _output, _template = self._validate_inputs()
            self._prune_edit_state_to_current_invoices()
            invoices = parse_invoice_paths(invoice_paths, self.invoice_number_overrides)
            validate_single_vendor_batch(invoices)
            lookup = AlmaPOLineLookup(alma_path)
            lookup.load()
            AlmaLineEnricher(lookup, self.manual_po_assignments).enrich(invoices)

            refreshed_originals = {}
            for invoice in invoices:
                source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
                for line in invoice.lines:
                    key = (source_key, line.line_number)
                    refreshed_originals[key] = self._editable_line_values(invoice, line)

            valid_keys = set(refreshed_originals)
            self.line_overrides = {key: value for key, value in self.line_overrides.items() if key in valid_keys}
            self.manual_po_assignments = {key: value for key, value in self.manual_po_assignments.items() if key in valid_keys}
            self.edit_original_values = refreshed_originals
            self.edit_header_original_values = {
                os.path.normcase(os.path.abspath(invoice.source_file or "")): self._editable_header_values(invoice)
                for invoice in invoices
            }
            valid_sources = set(self.edit_header_original_values)
            self.invoice_header_overrides = {key: value for key, value in self.invoice_header_overrides.items() if key in valid_sources}
            self.edit_draft_signature = self._current_edit_signature()

            selected_path = self._active_edit_path()
            if selected_path not in self.invoice_files:
                selected_path = self.invoice_files[0] if self.invoice_files else None
            self._refresh_edit_selector(select_path=selected_path)
            self.status_text.set("Estado: borrador editable del Excel de Alma preparado.")
            if not silent:
                messagebox.showinfo(
                    "Editar",
                    "Se ha reconstruido el borrador del Excel final de Alma con la factura y el informe de Analytics actualmente seleccionados."
                )
            return True
        except (ValueError, FileNotFoundError) as exc:
            self.edit_draft_signature = None
            self.edit_original_values.clear()
            self.edit_header_original_values.clear()
            self._refresh_edit_tree()
            if not silent:
                messagebox.showerror("No se pudo preparar el borrador", str(exc))
            return False

    def _editable_header_values(self, invoice):
        return {
            "invoice_number": invoice.invoice_number or "", "date": invoice.date.strftime("%m/%d/%Y") if invoice.date else "",
            "total_amount": str(invoice.total_amount), "currency": invoice.currency or "", "vendor_account": invoice.vendor_account or "",
            "note": invoice.note or "", "vat_amount": "" if invoice.vat_amount is None else str(invoice.vat_amount),
            "inclusive": str(bool(invoice.inclusive)),
        }

    def _editable_line_values(self, invoice, line):
        return {
            "invoice": invoice.invoice_number, "line": line.line_number, "line_type": line.line_type or "REGULAR",
            "po_line": line.po_line or "", "vendor_ref": line.vendor_ref or "", "issn": line.issn or "", "isbn": line.isbn or "",
            "title": line.title or "", "price": str(line.price), "quantity": str(line.quantity),
            "reporting_code": line.reporting_code or "", "secondary_reporting_code": line.secondary_reporting_code or "",
            "tertiary_reporting_code": line.tertiary_reporting_code or "", "fourth_reporting_code": line.fourth_reporting_code or "",
            "fifth_reporting_code": line.fifth_reporting_code or "",
            "start_date": line.start_date.strftime("%Y-%m-%d") if line.start_date else "",
            "end_date": line.end_date.strftime("%Y-%m-%d") if line.end_date else "", "note": line.note or "",
            "fund_1": line.funds[0] if len(line.funds) > 0 and line.funds[0] else "",
            "fund_2": line.funds[1] if len(line.funds) > 1 and line.funds[1] else "",
            "vat_amount": "" if line.vat_amount is None else str(line.vat_amount),
            "vat_percentage": "" if line.vat_percentage is None else str(line.vat_percentage),
        }

    def _refresh_edit_tree(self):
        if not hasattr(self, "edit_tree"): return
        for item in self.edit_tree.get_children(): self.edit_tree.delete(item)
        path = self._active_edit_path()
        if not path: return
        source_key = os.path.normcase(os.path.abspath(path))
        for item in self.header_tree.get_children(): self.header_tree.delete(item)
        header_values = dict(self.edit_header_original_values.get(source_key, {}))
        header_values.update(self.invoice_header_overrides.get(source_key, {}))
        if header_values:
            self.header_tree.insert("", "end", iid="header", values=[header_values.get(col, "") for col in self.header_columns])
        for key, original in sorted(self.edit_original_values.items(), key=lambda item: item[0][1]):
            if key[0] != source_key: continue
            values = dict(original); values.update(self.line_overrides.get(key, {}))
            iid = f"{key[1]}"
            self.edit_tree.insert("", "end", iid=iid, values=[values.get(col, "") for col in self.edit_columns])

    def _begin_edit_cell(self, event):
        row_id = self.edit_tree.identify_row(event.y); col_id = self.edit_tree.identify_column(event.x)
        if not row_id or not col_id: return
        col_index = int(col_id[1:]) - 1; column = self.edit_columns[col_index]
        if column in READ_ONLY_LINE_FIELDS:
            messagebox.showinfo("Campo protegido", "Este campo procede de la factura original y no es editable en FacturaeToAlma.")
            return
        self._open_tree_cell_editor(row_id, column, col_id)

    def _open_tree_cell_editor(self, row_id, column, col_id=None):
        if col_id is None: col_id = f"#{self.edit_columns.index(column)+1}"
        bbox = self.edit_tree.bbox(row_id, col_id)
        if not bbox: return
        x, y, width, height = bbox
        current = self.edit_tree.set(row_id, column)
        values = None
        if column == "line_type": values = LINE_TYPE_OPTIONS
        elif column in {"reporting_code", "secondary_reporting_code", "tertiary_reporting_code", "fourth_reporting_code", "fifth_reporting_code", "fund_1", "fund_2"}:
            values = self._context_suggestions(column)
        if values:
            editor = ttk.Combobox(self.edit_tree, values=values, state="normal")
            editor.set(current)
        else:
            editor = ttk.Entry(self.edit_tree); editor.insert(0, current)
        editor.place(x=x, y=y, width=width, height=height); editor.focus_set(); editor.select_range(0, tk.END)
        self.edit_combo_editor = editor
        self.edit_combo_context = (row_id, column)
        def commit(_event=None):
            self._commit_active_line_editor()
        def cancel(_event=None):
            if self.edit_combo_editor is not None and self.edit_combo_editor.winfo_exists():
                self.edit_combo_editor.destroy()
            self.edit_combo_editor = None
            self.edit_combo_context = None
        def commit_tab(_event=None):
            if self._commit_active_line_editor():
                self.root.after_idle(lambda: self._focus_next_editable_cell(row_id, column))
            return "break"
        editor.bind("<Return>", commit)
        editor.bind("<Tab>", commit_tab)
        editor.bind("<Escape>", cancel)

    def _commit_active_line_editor(self):
        editor = self.edit_combo_editor
        context = self.edit_combo_context
        if editor is None or context is None:
            return True
        self.edit_combo_editor = None
        self.edit_combo_context = None
        value = editor.get()
        row_id, column = context
        success = self._set_edit_value(row_id, column, value)
        if editor.winfo_exists():
            editor.destroy()
        return bool(success)

    def _focus_next_editable_cell(self, row_id, column):
        editable = [name for name in self.edit_columns if name in EDITABLE_LINE_FIELDS]
        if column not in editable:
            return
        rows = list(self.edit_tree.get_children())
        column_index = editable.index(column)
        row_index = rows.index(row_id)
        if column_index + 1 < len(editable):
            next_row, next_column = row_id, editable[column_index + 1]
        elif row_index + 1 < len(rows):
            next_row, next_column = rows[row_index + 1], editable[0]
        else:
            return
        self.edit_tree.selection_set(next_row)
        self.edit_tree.focus(next_row)
        self.edit_tree.see(next_row)
        self._open_tree_cell_editor(next_row, next_column)

    def _focus_next_header_cell(self, column):
        editable = [name for name in self.header_columns if name in EDITABLE_HEADER_FIELDS]
        if column not in editable:
            return
        index = editable.index(column)
        if index + 1 < len(editable):
            self._open_header_cell_editor(editable[index + 1])

    def _set_edit_value(self, row_id, column, value, update_widget=True):
        path = self._active_edit_path()
        if not path: return
        key = (os.path.normcase(os.path.abspath(path)), int(row_id))
        if column == "line_type" and value not in LINE_TYPE_OPTIONS:
            messagebox.showerror("Valor no valido", "Selecciona un Line type permitido por Alma."); return False
        if column in {"price", "quantity", "vat_amount", "vat_percentage"}:
            try: Decimal(str(value).replace(",", "."))
            except InvalidOperation:
                messagebox.showerror("Valor no valido", f"{column} debe ser numerico."); return False
        clean_value = sanitize_excel_value(value.strip()) if isinstance(value, str) else value
        self.line_overrides.setdefault(key, {})[column] = clean_value
        if update_widget:
            self.edit_tree.set(row_id, column, clean_value)
        if column == "po_line":
            self.manual_po_assignments[key] = clean_value if clean_value else None
        self.edit_changes_saved = False
        return True

    def _context_suggestions(self, column):
        counts = {}
        for original in self.edit_original_values.values():
            value = original.get(column)
            if value: counts[value] = counts.get(value, 0) + 1
        for changes in self.line_overrides.values():
            value = changes.get(column)
            if value: counts[value] = counts.get(value, 0) + 1
        return [value for value, _count in sorted(counts.items(), key=lambda item: (-item[1], str(item[0])))]

    def _reporting_code_suggestions(self, column):
        return self._context_suggestions(column)

    def _show_edit_context_menu(self, event):
        row_id = self.edit_tree.identify_row(event.y); col_id = self.edit_tree.identify_column(event.x)
        if row_id and row_id not in self.edit_tree.selection(): self.edit_tree.selection_set(row_id)
        self.edit_context_column = self.edit_columns[int(col_id[1:])-1] if col_id else None
        try: self.edit_context_menu.tk_popup(event.x_root, event.y_root)
        finally: self.edit_context_menu.grab_release()

    def _edit_selected_cell(self):
        selected = self.edit_tree.selection()
        if selected and getattr(self, "edit_context_column", None): self._open_tree_cell_editor(selected[0], self.edit_context_column)

    def _clear_selected_edit_cell(self):
        column = getattr(self, "edit_context_column", None)
        if column in READ_ONLY_LINE_FIELDS or column is None: return
        for row_id in self.edit_tree.selection(): self._set_edit_value(row_id, column, "")

    def _restore_selected_edit_cell(self):
        column = getattr(self, "edit_context_column", None); path = self._active_edit_path()
        if not column or not path: return
        source_key = os.path.normcase(os.path.abspath(path))
        for row_id in self.edit_tree.selection():
            key = (source_key, int(row_id)); original = self.edit_original_values.get(key, {}).get(column, "")
            self.line_overrides.get(key, {}).pop(column, None); self.edit_tree.set(row_id, column, original)

    def _suggest_selected_context_value(self):
        column = getattr(self, "edit_context_column", None)
        if not column or column in {"invoice", "line", "price", "quantity", "vat_amount", "vat_percentage"}: return
        suggestions = self._context_suggestions(column)
        if suggestions:
            for row_id in self.edit_tree.selection(): self._set_edit_value(row_id, column, suggestions[0])
        else: messagebox.showinfo("Sin sugerencias", "No hay valores de contexto para sugerir en esta columna.")

    def _prompt_bulk_value(self, column, initial=""):
        dialog = tk.Toplevel(self.root); dialog.title(f"Asignar {column}"); dialog.transient(self.root); dialog.grab_set()
        value_var = tk.StringVar(value=initial)
        suggestions = LINE_TYPE_OPTIONS if column == "line_type" else self._context_suggestions(column)
        ttk.Label(dialog, text=f"Valor para {column}:").pack(anchor="w", padx=14, pady=(12, 4))
        editor = ttk.Combobox(dialog, textvariable=value_var, values=suggestions, state="normal", width=55)
        editor.pack(fill=tk.X, padx=14, pady=4); editor.focus_set()
        result = {"value": None}
        def accept(): result["value"] = value_var.get(); dialog.destroy()
        buttons=ttk.Frame(dialog); buttons.pack(pady=(6,12)); ttk.Button(buttons,text="Aplicar",command=accept).pack(side=tk.LEFT,padx=4); ttk.Button(buttons,text="Cancelar",command=dialog.destroy).pack(side=tk.LEFT,padx=4)
        dialog.bind("<Return>", lambda _e: accept()); self.root.wait_window(dialog)
        return result["value"]

    def _bulk_apply_selected_value(self):
        column = getattr(self, "edit_context_column", None); selected = self.edit_tree.selection()
        if not selected or column in READ_ONLY_LINE_FIELDS or column is None: return
        value = self._prompt_bulk_value(column, self.edit_tree.set(selected[0], column))
        if value is None: return
        for row_id in selected: self._set_edit_value(row_id, column, value, update_widget=False)
        self._refresh_edit_tree()

    def _fill_selected_blanks(self):
        column = getattr(self, "edit_context_column", None); selected = self.edit_tree.selection()
        if not selected or column in READ_ONLY_LINE_FIELDS or column is None: return
        source = next((self.edit_tree.set(row, column) for row in selected if self.edit_tree.set(row, column)), "")
        if not source: source = self._prompt_bulk_value(column, "")
        if source is None or source == "": return
        for row_id in selected:
            if not self.edit_tree.set(row_id, column): self._set_edit_value(row_id, column, source, update_widget=False)
        self._refresh_edit_tree()

    def _fill_blanks_below(self):
        column = getattr(self, "edit_context_column", None); selected = self.edit_tree.selection()
        if not selected or column in READ_ONLY_LINE_FIELDS or column is None: return
        children = list(self.edit_tree.get_children()); anchor = selected[0]; start_index = children.index(anchor)
        source = self.edit_tree.set(anchor, column)
        if not source: source = self._prompt_bulk_value(column, "")
        if source is None or source == "": return
        for row_id in children[start_index + 1:]:
            if not self.edit_tree.set(row_id, column): self._set_edit_value(row_id, column, source, update_widget=False)
        self._refresh_edit_tree()

    def _begin_edit_header_cell(self, event):
        row_id=self.header_tree.identify_row(event.y); col_id=self.header_tree.identify_column(event.x)
        if not row_id or not col_id: return
        column=self.header_columns[int(col_id[1:])-1]
        if column in READ_ONLY_HEADER_FIELDS:
            messagebox.showinfo("Campo protegido", "El número, la fecha y los importes de cabecera no se editan directamente. El número solo admite prefijo o sufijo.")
            return
        self._open_header_cell_editor(column, col_id)

    def _open_header_cell_editor(self, column, col_id=None):
        if col_id is None: col_id=f"#{self.header_columns.index(column)+1}"
        bbox=self.header_tree.bbox("header", col_id)
        if not bbox: return
        x,y,w,h=bbox; current=self.header_tree.set("header", column)
        if column == "inclusive":
            editor=ttk.Combobox(self.header_tree, values=["False","True"], state="readonly"); editor.set(current)
        else:
            editor=ttk.Entry(self.header_tree); editor.insert(0,current)
        editor.place(x=x,y=y,width=w,height=h); editor.focus_set(); editor.select_range(0,tk.END)
        self.header_cell_editor = editor
        self.header_editor_context = column
        def commit(_event=None): self._commit_active_header_editor()
        def cancel(_event=None):
            if self.header_cell_editor is not None and self.header_cell_editor.winfo_exists(): self.header_cell_editor.destroy()
            self.header_cell_editor = None; self.header_editor_context = None
        def commit_tab(_event=None):
            if self._commit_active_header_editor():
                self.root.after_idle(lambda: self._focus_next_header_cell(column))
            return "break"
        editor.bind("<Return>",commit)
        editor.bind("<Tab>",commit_tab)
        editor.bind("<Escape>",cancel)

    def _commit_active_header_editor(self):
        editor = self.header_cell_editor
        column = self.header_editor_context
        if editor is None or column is None:
            return True
        self.header_cell_editor = None
        self.header_editor_context = None
        value = editor.get()
        success = self._set_header_value(column, value)
        if editor.winfo_exists():
            editor.destroy()
        return bool(success)

    def _flush_editors(self):
        if not self._commit_active_line_editor(): return False
        if not self._commit_active_header_editor(): return False
        self.root.update_idletasks()
        return True

    def _set_header_value(self, column, value):
        path=self._active_edit_path()
        if not path: return False
        source_key=os.path.normcase(os.path.abspath(path))
        if column in {"total_amount","vat_amount"} and value:
            try: Decimal(str(value).replace(",","."))
            except InvalidOperation: messagebox.showerror("Valor no valido",f"{column} debe ser numerico."); return False
        if column == "date" and value and parse_date_for_excel(value) is None:
            messagebox.showerror("Fecha no valida","Usa una fecha valida, por ejemplo 2026-10-19."); return False
        if column == "invoice_number":
            self.invoice_metadata[path]["number"] = value; self.invoice_number_overrides[source_key] = value
        else:
            self.invoice_header_overrides.setdefault(source_key,{})[column]=sanitize_excel_value(value.strip()) if isinstance(value,str) else value
        self.header_tree.set("header",column,value)
        self.edit_changes_saved = False
        return True

    def _show_header_context_menu(self,event):
        row=self.header_tree.identify_row(event.y); col=self.header_tree.identify_column(event.x)
        if row: self.header_tree.selection_set(row)
        self.header_context_column=self.header_columns[int(col[1:])-1] if col else None
        try:self.header_context_menu.tk_popup(event.x_root,event.y_root)
        finally:self.header_context_menu.grab_release()

    def _edit_selected_header_cell(self):
        column = getattr(self, "header_context_column", None)
        if column in READ_ONLY_HEADER_FIELDS:
            messagebox.showinfo("Campo protegido", "Este campo no es editable.")
            return
        if column:
            self._open_header_cell_editor(column)

    def _restore_selected_header_cell(self):
        path=self._active_edit_path(); column=getattr(self,"header_context_column",None)
        if not path or not column:return
        source_key=os.path.normcase(os.path.abspath(path)); original=self.edit_header_original_values.get(source_key,{}).get(column,"")
        self.invoice_header_overrides.get(source_key,{}).pop(column,None); self.header_tree.set("header",column,original)

    def _harvest_visible_editor(self):
        """Read every displayed editable cell into the model before saving."""
        if not self._flush_editors():
            return False
        path = self._active_edit_path()
        if not path:
            return False
        source_key = os.path.normcase(os.path.abspath(path))

        # Header: only explicitly allowed non-financial fields.
        if self.header_tree.exists("header"):
            for column in EDITABLE_HEADER_FIELDS:
                current = self.header_tree.set("header", column)
                original = self.edit_header_original_values.get(source_key, {}).get(column, "")
                if str(current) != str(original):
                    self.invoice_header_overrides.setdefault(source_key, {})[column] = sanitize_excel_value(current)
                else:
                    self.invoice_header_overrides.get(source_key, {}).pop(column, None)

        # Lines: every visible editable value is compared with its original value.
        for row_id in self.edit_tree.get_children():
            key = (source_key, int(row_id))
            original_values = self.edit_original_values.get(key, {})
            for column in EDITABLE_LINE_FIELDS:
                current = self.edit_tree.set(row_id, column)
                original = original_values.get(column, "")
                if str(current) != str(original):
                    self.line_overrides.setdefault(key, {})[column] = sanitize_excel_value(current)
                else:
                    self.line_overrides.get(key, {}).pop(column, None)
            if not self.line_overrides.get(key):
                self.line_overrides.pop(key, None)
            po_line = self.edit_tree.set(row_id, "po_line")
            if str(po_line) != str(original_values.get("po_line", "")):
                self.manual_po_assignments[key] = po_line or None
            else:
                self.manual_po_assignments.pop(key, None)

        # Number: only Original / Prefix / Suffix controls are authoritative.
        original_number = self.invoice_metadata.get(path, {}).get("original_number", "")
        mode = self.number_mode_var.get()
        affix = self.number_affix_var.get()
        if mode == "prefix":
            effective_number = f"{affix}{original_number}"
        elif mode == "suffix":
            effective_number = f"{original_number}{affix}"
        else:
            effective_number = original_number
        if effective_number != original_number:
            self.invoice_number_overrides[source_key] = effective_number
        else:
            self.invoice_number_overrides.pop(source_key, None)
        self.invoice_metadata[path]["number"] = effective_number
        return True

    def save_edit_changes(self):
        path = self._active_edit_path()
        if not path:
            messagebox.showinfo("Editar", "No hay ninguna factura activa en el editor.")
            return
        if not self._harvest_visible_editor():
            return
        if not messagebox.askyesno(
            "Guardar cambios",
            "¿Quieres guardar todos los cambios y actualizar ahora el Excel de salida?\n\n"
            "Se sobrescribirá el Excel indicado en Convertir. El XML original no se modificará.",
        ):
            return
        # Immutable snapshot. Conversion never reads the mutable Treeview state.
        self.saved_line_overrides = copy.deepcopy(self.line_overrides)
        self.saved_header_overrides = copy.deepcopy(self.invoice_header_overrides)
        self.saved_number_overrides = copy.deepcopy(self.invoice_number_overrides)
        self.saved_manual_assignments = copy.deepcopy(self.manual_po_assignments)
        total_saved_fields = sum(len(values) for values in self.saved_line_overrides.values()) + sum(len(values) for values in self.saved_header_overrides.values()) + len(self.saved_number_overrides)
        if total_saved_fields == 0:
            messagebox.showwarning("Sin cambios", "No se ha detectado ningun cambio respecto al borrador original.")
            self.edit_changes_saved = False
            return
        self.edit_draft_signature = self._current_edit_signature()
        self.edit_changes_saved = True
        saved_line_fields = sum(len(values) for values in self.saved_line_overrides.values())
        saved_header_fields = sum(len(values) for values in self.saved_header_overrides.values())
        saved_number_fields = len(self.saved_number_overrides)
        logging.info(
            "Editor guardado: %d campos de linea, %d campos de cabecera, %d numeros, %d POL manuales",
            saved_line_fields, saved_header_fields, len(self.saved_number_overrides), len(self.saved_manual_assignments),
        )
        self.status_text.set(
            f"Estado: cambios guardados ({saved_number_fields} números, {saved_header_fields} campos de cabecera, "
            f"{saved_line_fields} campos de línea)."
        )
        self.status_text.set(
            f"Estado: guardados {saved_number_fields} cambios de número, {saved_header_fields} de cabecera y "
            f"{saved_line_fields} de línea; actualizando el Excel de salida..."
        )
        # A Save action must persist to disk immediately, not merely to memory.
        self.root.after_idle(self.run_conversion)

    def reset_edit_invoice(self):
        path = self._active_edit_path()
        if not path: return
        if not messagebox.askyesno("Restablecer factura", "¿Quieres descartar todos los cambios de esta factura?"):
            return
        source_key = os.path.normcase(os.path.abspath(path))
        self.line_overrides = {key: value for key, value in self.line_overrides.items() if key[0] != source_key}
        self.manual_po_assignments = {key: value for key, value in self.manual_po_assignments.items() if key[0] != source_key}
        self.invoice_header_overrides.pop(source_key, None)
        self.saved_line_overrides = {key: value for key, value in self.saved_line_overrides.items() if key[0] != source_key}
        self.saved_manual_assignments = {key: value for key, value in self.saved_manual_assignments.items() if key[0] != source_key}
        self.saved_header_overrides.pop(source_key, None)
        self.saved_number_overrides.pop(source_key, None)
        self.reset_active_invoice_number()
        self.saved_line_overrides.clear(); self.saved_header_overrides.clear()
        self.saved_number_overrides.clear(); self.saved_manual_assignments.clear()
        self.edit_changes_saved = False
        self.number_mode_var.set("original")
        self._update_number_preview()
        self._refresh_edit_tree()

    def _directory_row(self, parent, row, label, variable, command):
        ttk.Label(parent, text=label, style="App.TLabel").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=10, pady=6)
        ttk.Button(parent, text="Seleccionar carpeta...", command=command).grid(row=row, column=2, pady=6)

    def _build_customize_tab(self):
        tab = self.customize_tab
        tab.columnconfigure(0, weight=1)
        frame = ttk.LabelFrame(tab, text="Preferencias", padding=16, style="App.TLabelframe")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Institucion", style="App.TLabel").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.institution_var).grid(row=0, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=6)
        self._directory_row(frame, 1, "Directorio de facturas FACe", self.custom_invoice_dir, lambda: self._choose_preference_dir(self.custom_invoice_dir))
        self._directory_row(frame, 2, "Directorio del fichero Excel de Alma", self.custom_alma_dir, lambda: self._choose_preference_dir(self.custom_alma_dir))
        self._directory_row(frame, 3, "Directorio de plantilla Alma", self.custom_template_dir, lambda: self._choose_preference_dir(self.custom_template_dir))
        self._directory_row(frame, 4, "Directorio de Excel finales", self.custom_output_dir, lambda: self._choose_preference_dir(self.custom_output_dir))
        ttk.Label(frame, text="Logo institucional", style="App.TLabel").grid(row=5, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.logo_path_var).grid(row=5, column=1, sticky="ew", padx=10, pady=6)
        logo_buttons = ttk.Frame(frame, style="App.TFrame")
        logo_buttons.grid(row=5, column=2, sticky="e", pady=6)
        ttk.Button(logo_buttons, text="Seleccionar logo...", command=self.select_logo).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(logo_buttons, text="Eliminar logo", command=self.remove_logo).pack(side=tk.LEFT)
        self.logo_status_label = ttk.Label(frame, text="PNG, JPG o JPEG. Altura maxima: 100 px.", style="App.TLabel")
        self.logo_status_label.grid(row=6, column=1, columnspan=2, sticky="w", padx=10)
        note = ("Todos los campos son opcionales. Si una carpeta personalizada queda vacia, se usa la ultima carpeta valida; "
                "en la primera ejecucion se usa la carpeta de la aplicacion. La salida se guarda junto a la factura, salvo que se defina un directorio de salida.")
        ttk.Label(frame, text=note, wraplength=880, justify="left", style="App.TLabel").grid(row=7, column=0, columnspan=3, sticky="w", pady=(14, 8))
        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.grid(row=8, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(buttons, text="Guardar preferencias", command=self.save_preferences).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Restablecer valores", command=self.reset_preferences).pack(side=tk.LEFT)

    def _build_help_tab(self):
        tab = self.help_tab
        tab.columnconfigure(0, weight=1)
        ttk.Label(tab, text=f"{APP_NAME} {APP_VERSION}", style="HelpTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))
        help_text = (
            "Factura individual: selecciona una factura, el informe de Alma Analytics, la plantilla y el destino.\n\n"
            "Modo lote: todas las facturas deben pertenecer al mismo proveedor y biblioteca. La lista permite anadir, quitar, ordenar y detectar duplicados.\n\n"
            "Informe de validacion: se genera cuando faltan PO Lines, hay coincidencias ambiguas, duplicados o diferencias entre Quantity y Quantity for Pricing. "
            "Una diferencia de cantidad requiere revision manual y no determina por si sola si la facturacion es parcial o total.\n\n"
            "Limitaciones confirmadas por Ex Libris: la plantilla Excel no permite indicar Line Exclusive ni declarar explicitamente facturacion parcial o completa.\n\n"
            "Licencia: GPL-3.0-only."
        )
        ttk.Label(tab, text=help_text, wraplength=940, justify="left", style="App.TLabel").grid(row=1, column=0, sticky="nw")
        links = ttk.LabelFrame(tab, text="Documentacion y soporte", padding=14, style="App.TLabelframe")
        links.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        ttk.Button(links, text="Abrir repositorio de GitHub", command=lambda: self._open_url(PROJECT_REPOSITORY_URL)).pack(anchor="w", pady=4)
        ttk.Button(links, text="Abrir Biblioguia / LibGuides", command=lambda: self._open_url(PROJECT_GUIDE_URL)).pack(anchor="w", pady=4)
        ttk.Label(links, text=PROJECT_REPOSITORY_URL, style="App.TLabel").pack(anchor="w", pady=(10, 2))
        ttk.Label(links, text=PROJECT_GUIDE_URL, style="App.TLabel").pack(anchor="w")

    def select_logo(self):
        current = self.logo_path_var.get().strip()
        initial = os.path.dirname(current) if current and os.path.isfile(current) else self._application_dir()
        path = filedialog.askopenfilename(
            title="Seleccionar logo institucional",
            initialdir=initial,
            filetypes=[
                ("Imagenes compatibles", "*.png *.jpg *.jpeg"),
                ("Imagen PNG", "*.png"),
                ("Imagen JPEG", "*.jpg *.jpeg"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if not path:
            return
        self.logo_path_var.set(path)
        if self._apply_logo(path, show_warning=True):
            self.logo_status_label.config(text=f"Logo cargado: {os.path.basename(path)}")

    def remove_logo(self):
        self.logo_path_var.set("")
        self.logo_photo = None
        self.logo_label.config(image="")
        self.logo_status_label.config(text="Sin logo. PNG, JPG o JPEG; altura maxima: 100 px.")

    def _load_logo_from_preferences(self, show_warning=False):
        path = self.logo_path_var.get().strip()
        if not path:
            self.logo_photo = None
            if hasattr(self, "logo_label"):
                self.logo_label.config(image="")
            return False
        return self._apply_logo(path, show_warning=show_warning)

    def _apply_logo(self, path, show_warning=False):
        if not hasattr(self, "logo_label"):
            return False
        if not os.path.isfile(path):
            self.logo_photo = None
            self.logo_label.config(image="")
            if hasattr(self, "logo_status_label"):
                self.logo_status_label.config(text="El logo configurado no esta disponible.")
            if show_warning:
                messagebox.showwarning("Logo no disponible", f"No se encontro el archivo de logo:\n{path}")
            return False
        try:
            with Image.open(path) as opened:
                image = opened.copy()
            if image.mode not in ("RGB", "RGBA", "L", "LA"):
                image = image.convert("RGBA" if "transparency" in image.info else "RGB")
            if image.height > MAX_LOGO_HEIGHT:
                ratio = MAX_LOGO_HEIGHT / image.height
                new_width = max(1, round(image.width * ratio))
                image = image.resize((new_width, MAX_LOGO_HEIGHT), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(image)
            self.logo_label.config(image=self.logo_photo)
            return True
        except (UnidentifiedImageError, OSError, ValueError) as e:
            logging.warning(f"No se pudo cargar el logo {path}: {e}")
            self.logo_photo = None
            self.logo_label.config(image="")
            if hasattr(self, "logo_status_label"):
                self.logo_status_label.config(text="El archivo seleccionado no es una imagen valida.")
            if show_warning:
                messagebox.showwarning("Logo no valido", f"No se pudo cargar el logo:\n{path}\n\n{e}")
            return False

    def _choose_preference_dir(self, variable):
        initial = self._valid_dir(variable.get()) or self._application_dir()
        folder = filedialog.askdirectory(initialdir=initial, title="Seleccionar carpeta")
        if folder:
            variable.set(folder)

    def save_preferences(self):
        institution = self.institution_var.get().strip() or self.DEFAULT_INSTITUTION
        self.institution_var.set(institution)
        for value, label in [
            (self.custom_invoice_dir.get().strip(), "Directorio de facturas FACe"),
            (self.custom_alma_dir.get().strip(), "Directorio de Alma"),
            (self.custom_template_dir.get().strip(), "Directorio de plantilla"),
            (self.custom_output_dir.get().strip(), "Directorio de salida"),
        ]:
            if value and not os.path.isdir(value):
                messagebox.showerror("Preferencias no validas", f"{label} no es una carpeta valida:\n{value}")
                return
        logo_path = self.logo_path_var.get().strip()
        if logo_path and not self._apply_logo(logo_path, show_warning=True):
            return
        self.preferences.update({
            "institution": institution,
            "custom_invoice_dir": self.custom_invoice_dir.get().strip(),
            "custom_alma_dir": self.custom_alma_dir.get().strip(),
            "custom_template_dir": self.custom_template_dir.get().strip(),
            "custom_output_dir": self.custom_output_dir.get().strip(),
            "logo_path": logo_path,
        })
        try:
            self._save_preferences_file()
            self._update_title()
            self._load_logo_from_preferences(show_warning=False)
            messagebox.showinfo("Preferencias", "Las preferencias se han guardado correctamente.")
        except ValueError as e:
            messagebox.showerror("Error al guardar", str(e))

    def reset_preferences(self):
        if not messagebox.askyesno("Restablecer preferencias", "¿Quieres restablecer la institucion y todos los directorios personalizados?"):
            return
        defaults = self._default_preferences()
        self.preferences = defaults
        self.institution_var.set(defaults["institution"])
        self.custom_invoice_dir.set("")
        self.custom_alma_dir.set("")
        self.custom_template_dir.set("")
        self.custom_output_dir.set("")
        self.logo_path_var.set("")
        self.logo_photo = None
        self.logo_label.config(image="")
        self.logo_status_label.config(text="Sin logo. PNG, JPG o JPEG; altura maxima: 100 px.")
        self.template_file.set("")
        try:
            self._save_preferences_file()
        except ValueError as e:
            messagebox.showerror("Error al restablecer", str(e))
            return
        self._update_title()
        messagebox.showinfo("Preferencias", "Se han restablecido los valores iniciales.")

    def _build_global_context_menu(self):
        self.global_context_menu = tk.Menu(self.root, tearoff=0)
        self.root.bind_all("<Button-3>", self._show_global_context_menu, add="+")

    def _show_global_context_menu(self, event):
        widget = event.widget
        menu = self.global_context_menu
        menu.delete(0, tk.END)
        self.context_widget = widget
        editable = isinstance(widget, (tk.Entry, ttk.Entry, tk.Text, ttk.Combobox))
        if editable:
            menu.add_command(label="Cortar", command=lambda: widget.event_generate("<<Cut>>"))
            menu.add_command(label="Copiar", command=lambda: widget.event_generate("<<Copy>>"))
            menu.add_command(label="Pegar", command=lambda: widget.event_generate("<<Paste>>"))
            menu.add_separator()
            menu.add_command(label="Seleccionar todo", command=lambda: self._select_all_widget(widget))
        else:
            menu.add_command(label="Copiar texto", command=lambda: self._copy_widget_text(widget))
        if hasattr(self, "visualize_tab"):
            menu.add_separator()
            menu.add_command(label="Ir a Visualizar", command=lambda: self.notebook.select(self.visualize_tab))
            menu.add_command(label="Ir a Editar", command=lambda: self.notebook.select(self.edit_tab))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def _select_all_widget(self, widget):
        if isinstance(widget, tk.Text):
            widget.tag_add(tk.SEL, "1.0", tk.END)
        else:
            try:
                widget.selection_range(0, tk.END)
                widget.icursor(tk.END)
            except tk.TclError:
                pass

    def _copy_widget_text(self, widget):
        try:
            text = widget.cget("text")
        except (tk.TclError, AttributeError):
            text = ""
        if text:
            self.root.clipboard_clear(); self.root.clipboard_append(str(text))

    def _open_url(self, url):
        try:
            webbrowser.open_new(url)
        except Exception as e:
            messagebox.showerror("Error al abrir enlace", f"No se pudo abrir el enlace:\n{url}\n\n{e}")

    def _open_folder(self, folder_path):
        folder_path = os.path.realpath(os.path.abspath(folder_path))
        if not os.path.isdir(folder_path):
            messagebox.showwarning("Aviso", "No se encontro la carpeta de salida.")
            return
        try:
            system_name = platform.system()
            if system_name == "Windows":
                os.startfile(folder_path)  # nosec B606
                return

            command_name = "open" if system_name == "Darwin" else "xdg-open"
            opener_path = shutil.which(command_name)
            if not opener_path or not os.path.isabs(opener_path):
                raise FileNotFoundError(
                    f"No se encontro una ruta absoluta para la utilidad '{command_name}'."
                )

            subprocess.run(  # nosec B603
                [opener_path, folder_path],
                check=False,
                shell=False,
            )
        except (OSError, subprocess.SubprocessError) as e:
            messagebox.showerror("Error al abrir carpeta", f"No se pudo abrir la carpeta:\n{folder_path}\n\n{e}")

    def open_output_folder(self):
        output_path = self.output_file.get().strip()
        folder = os.path.dirname(output_path) if output_path else ""
        if not folder:
            messagebox.showwarning("Aviso", "No hay una carpeta de salida definida.")
            return
        self._open_folder(folder)

    def on_mode_change(self, initial=False):
        mode = self.conversion_mode.get()
        if not initial:
            self.invoice_files = []
            self.invoice_metadata = {}
            self.invoice_file.set("")
            self.output_file.set("")
            self._refresh_batch_tree()
            self._invalidate_edit_draft(clear_overrides=True)
        if mode == "batch":
            self.invoice_label.config(text="Facturas XSIG / XML / TXT del mismo proveedor y biblioteca *")
            self.invoice_button.config(text="Seleccionar facturas...")
            self.batch_frame.grid()
            self.status_text.set("Estado: modo lote. Selecciona facturas del mismo proveedor y biblioteca.")
        else:
            self.invoice_label.config(text="Factura electronica XSIG / XML / TXT *")
            self.invoice_button.config(text="Seleccionar factura...")
            self.batch_frame.grid_remove()
            self.status_text.set("Estado: modo factura individual.")

    def select_invoice_or_batch(self):
        if self.conversion_mode.get() == "batch":
            self.add_batch_files(replace=True)
        else:
            self.select_invoice_single()

    def _invoice_filetypes(self):
        return [("Facturas Universitas XXI / Facturae", "*.xsig *.xml *.txt"), ("Archivos XSIG", "*.xsig"), ("Archivos XML", "*.xml"), ("Archivos TXT", "*.txt"), ("Todos los archivos", "*.*")]

    def select_invoice_single(self):
        initial = self._initial_dir("custom_invoice_dir", "last_invoice_dir")
        path = filedialog.askopenfilename(title="Seleccionar factura XSIG / XML / TXT", initialdir=initial, filetypes=self._invoice_filetypes())
        if not path:
            return
        try:
            metadata = self._read_invoice_metadata(path)
        except Exception as e:
            messagebox.showerror("Factura no valida", str(e))
            return
        self._invalidate_edit_draft(clear_overrides=True)
        self.invoice_files = [path]
        self.invoice_metadata = {path: metadata}
        self.invoice_file.set(path)
        self._remember_dir("last_invoice_dir", path)
        self._set_default_output(path, batch=False)
        self.status_text.set(f"Estado: factura {metadata['number']} seleccionada. Proveedor: {metadata['vendor']}.")
        self._refresh_viewer_selector(select_path=path)
        self._refresh_edit_selector(select_path=path)

    def add_batch_files(self, replace=False):
        initial = self._initial_dir("custom_invoice_dir", "last_invoice_dir")
        paths = filedialog.askopenfilenames(title="Seleccionar facturas del mismo proveedor y biblioteca", initialdir=initial, filetypes=self._invoice_filetypes())
        if not paths:
            return
        if replace:
            self._invalidate_edit_draft(clear_overrides=True)
            self.invoice_files = []
            self.invoice_metadata = {}
        else:
            self.edit_draft_signature = None
        existing_norm = {os.path.normcase(os.path.abspath(p)) for p in self.invoice_files}
        existing_numbers = {m["number"] for m in self.invoice_metadata.values() if m.get("number")}
        added = 0
        warnings = []
        for path in paths:
            norm = os.path.normcase(os.path.abspath(path))
            if norm in existing_norm:
                warnings.append(f"Ruta repetida: {os.path.basename(path)}")
                continue
            try:
                metadata = self._read_invoice_metadata(path)
            except Exception as e:
                warnings.append(f"No se pudo leer {os.path.basename(path)}: {e}")
                continue
            number = metadata.get("number")
            if number and number in existing_numbers:
                warnings.append(f"Numero de factura duplicado {number}: {os.path.basename(path)}")
                continue
            self.invoice_files.append(path)
            self.invoice_metadata[path] = metadata
            existing_norm.add(norm)
            if number:
                existing_numbers.add(number)
            added += 1
        self._refresh_batch_tree()
        self._refresh_viewer_selector(select_path=self.invoice_files[0] if self.invoice_files else None)
        self._refresh_edit_selector(select_path=self.invoice_files[0] if self.invoice_files else None)
        if self.invoice_files:
            self.invoice_file.set(f"{len(self.invoice_files)} facturas seleccionadas")
            self._remember_dir("last_invoice_dir", self.invoice_files[0])
            self._set_default_output(self.invoice_files[0], batch=True)
        self.status_text.set(f"Estado: {len(self.invoice_files)} facturas en el lote.")
        if warnings:
            messagebox.showwarning("Facturas no incorporadas", "\n".join(warnings))
        elif added == 0:
            messagebox.showinfo("Lote", "No se incorporaron facturas nuevas.")

    def select_invoice_batch(self):
        self.add_batch_files(replace=True)

    def _read_invoice_metadata(self, path):
        parsed = FacturaeParser(path).parse()
        if len(parsed) != 1:
            raise ValueError(f"El fichero debe contener una sola factura y contiene {len(parsed)}.")
        inv = parsed[0]
        number = inv.invoice_number or "(sin numero)"
        return {"number": number, "original_number": number, "vendor": inv.vendor_account or inv.seller_name or "Proveedor no identificado"}

    def _refresh_batch_tree(self):
        if not hasattr(self, "batch_tree"):
            return
        for item in self.batch_tree.get_children():
            self.batch_tree.delete(item)
        for path in self.invoice_files:
            meta = self.invoice_metadata.get(path, {})
            self.batch_tree.insert("", "end", iid=path, values=(meta.get("number", ""), meta.get("vendor", ""), os.path.basename(path)))

    def remove_selected_batch_files(self):
        selected = list(self.batch_tree.selection())
        if not selected:
            messagebox.showinfo("Lote", "Selecciona al menos una factura de la lista.")
            return
        selected_set = set(selected)
        self.invoice_files = [p for p in self.invoice_files if p not in selected_set]
        for p in selected:
            self.invoice_metadata.pop(p, None)
        self.edit_draft_signature = None
        self.edit_original_values.clear()
        self._prune_edit_state_to_current_invoices()
        self._refresh_batch_tree()
        self._refresh_viewer_selector(select_path=self.invoice_files[0] if self.invoice_files else None)
        self._refresh_edit_selector(select_path=self.invoice_files[0] if self.invoice_files else None)
        self.invoice_file.set(f"{len(self.invoice_files)} facturas seleccionadas" if self.invoice_files else "")
        self.status_text.set(f"Estado: {len(self.invoice_files)} facturas en el lote.")

    def clear_batch_files(self):
        if self.invoice_files and not messagebox.askyesno("Vaciar lote", "¿Quieres eliminar todas las facturas del lote?"):
            return
        self.invoice_files = []
        self.invoice_metadata = {}
        self.invoice_file.set("")
        self.output_file.set("")
        self._refresh_batch_tree()
        self._invalidate_edit_draft(clear_overrides=True)
        self._refresh_viewer_selector(); self._refresh_edit_selector()
        self.status_text.set("Estado: lote vacio.")

    def sort_batch_files(self):
        self.invoice_files.sort(key=lambda p: (self.invoice_metadata.get(p, {}).get("number", "").casefold(), os.path.basename(p).casefold()))
        self._refresh_batch_tree()
        self.status_text.set("Estado: lote ordenado por numero de factura.")

    def _set_default_output(self, invoice_path, batch):
        custom_output = self._valid_dir(self.preferences.get("custom_output_dir", ""))
        invoice_folder = os.path.dirname(invoice_path)
        output_folder = custom_output or invoice_folder or self._initial_dir("custom_output_dir", "last_output_dir")
        base = os.path.basename(os.path.splitext(invoice_path)[0])
        if batch:
            output_name = f"Lote_{base}_Alma.xlsx"
        else:
            output_name = f"{base}_Alma.xlsx"
        self.output_file.set(os.path.join(output_folder, output_name))

    def select_alma_pol(self):
        fallback = os.path.dirname(self.invoice_files[0]) if self.invoice_files else None
        initial = self._initial_dir("custom_alma_dir", "last_alma_dir", fallback)
        path = filedialog.askopenfilename(title="Seleccionar fichero Excel de Alma Analytics", initialdir=initial, filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
        if path:
            self.alma_pol_file.set(path)
            self._remember_dir("last_alma_dir", path)
            self._invalidate_edit_draft(clear_overrides=False)
            self.status_text.set("Estado: fichero Excel de Alma Analytics seleccionado; actualiza el borrador en Editar.")

    def select_template(self):
        fallback = os.path.dirname(self.alma_pol_file.get()) if self.alma_pol_file.get() else None
        initial = self._initial_dir("custom_template_dir", "last_template_dir", fallback)
        path = filedialog.askopenfilename(title="Seleccionar plantilla Alma .xlsx", initialdir=initial, filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
        if path:
            self.template_file.set(path)
            self.preferences["last_template_path"] = path
            self._remember_dir("last_template_dir", path)
            self._invalidate_edit_draft(clear_overrides=False)
            self.status_text.set("Estado: plantilla Alma seleccionada; actualiza el borrador en Editar.")

    def select_output(self):
        fallback = os.path.dirname(self.invoice_files[0]) if self.invoice_files else None
        initial = self._initial_dir("custom_output_dir", "last_output_dir", fallback)
        initialfile = os.path.basename(self.output_file.get()) if self.output_file.get() else "FacturaeToAlma.xlsx"
        path = filedialog.asksaveasfilename(title="Definir archivo Excel de salida", initialdir=initial, initialfile=initialfile, defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
        if path:
            self.output_file.set(path)
            self._remember_dir("last_output_dir", path)
            self.status_text.set("Estado: archivo de salida definido.")

    def clear_fields(self):
        self.invoice_files = []
        self.invoice_metadata = {}
        self.invoice_file.set("")
        self.output_file.set("")
        self._refresh_batch_tree()
        self._refresh_viewer_selector()
        self._refresh_edit_selector()
        self.invoice_number_overrides.clear()
        self.manual_po_assignments.clear()
        self.line_overrides.clear(); self.edit_original_values.clear(); self.edit_header_original_values.clear(); self.invoice_header_overrides.clear(); self.edit_draft_signature = None
        self.number_mode_var.set("original")
        if self.alma_pol_file.get().strip() and messagebox.askyesno("Limpiar fichero Excel de Alma", "¿Quieres limpiar tambien el fichero Excel de Alma Analytics?\n\nLa plantilla Alma se conservara."):
            self.alma_pol_file.set("")
        self.status_text.set("Estado: datos limpiados. La plantilla Alma se conserva.")

    def _validate_inputs(self):
        if not self.invoice_files:
            raise ValueError("Debes seleccionar una factura o un lote de facturas.")
        allowed = {".xsig", ".xml", ".txt"}
        for path in self.invoice_files:
            if os.path.splitext(path)[1].lower() not in allowed:
                raise ValueError(f"Extension de factura no valida: {path}")
            if not os.path.isfile(path):
                raise FileNotFoundError(f"No existe la factura: {path}")
        alma = self.alma_pol_file.get().strip()
        template = self.template_file.get().strip()
        output = self.output_file.get().strip()
        if not alma:
            raise ValueError("Debes seleccionar el fichero Excel de Alma Analytics.")
        if not template:
            raise ValueError("Debes seleccionar la plantilla Alma.")
        if not output:
            raise ValueError("Debes definir el Excel de salida.")
        if not os.path.isfile(alma):
            raise FileNotFoundError(f"No existe el fichero Excel de Alma Analytics: {alma}")
        if not os.path.isfile(template):
            raise FileNotFoundError(f"No existe la plantilla Alma: {template}")
        output_norm = os.path.normcase(os.path.abspath(output))
        protected_paths = {
            os.path.normcase(os.path.abspath(alma)),
            os.path.normcase(os.path.abspath(template)),
            *(os.path.normcase(os.path.abspath(path)) for path in self.invoice_files),
        }
        if output_norm in protected_paths:
            raise ValueError("El Excel de salida no puede coincidir con la factura, el Analytics ni la plantilla. Elige otro nombre o carpeta.")
        duplicate_numbers = {}
        for path in self.invoice_files:
            number = self.invoice_metadata.get(path, {}).get("number")
            if number:
                duplicate_numbers.setdefault(number, []).append(path)
        repeated = {n: paths for n, paths in duplicate_numbers.items() if len(paths) > 1}
        if repeated:
            raise ValueError("Hay numeros de factura duplicados en el lote: " + ", ".join(sorted(repeated)))
        return list(self.invoice_files), alma, output, template

    def run_conversion(self):
        live_changes = bool(self.line_overrides or self.invoice_header_overrides or self.invoice_number_overrides or self.edit_combo_editor or self.header_cell_editor)
        if live_changes and not self.edit_changes_saved:
            messagebox.showwarning(
                "Cambios sin guardar",
                "Hay cambios en Editar que no se han guardado. Pulsa Guardar cambios antes de convertir."
            )
            self.notebook.select(self.edit_tab)
            return
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Conversion en curso", "Ya hay una conversion en curso.")
            return
        try:
            inputs = self._validate_inputs()
        except (ValueError, FileNotFoundError) as e:
            messagebox.showerror("Datos de entrada no validos", str(e))
            self.status_text.set("Estado: revisa los datos de entrada.")
            return
        if self.edit_draft_signature != self._current_edit_signature():
            if self.edit_changes_saved:
                messagebox.showerror(
                    "Borrador desactualizado",
                    "La factura, Analytics o plantilla han cambiado después de guardar. Vuelve a Editar, revisa y guarda de nuevo."
                )
                self.notebook.select(self.edit_tab)
                return
            if not self.prepare_edit_draft(silent=True):
                messagebox.showerror("Borrador no disponible", "No se pudo preparar el borrador editable antes de convertir.")
                return
        self.convert_button.config(state="disabled")
        self.status_text.set("Estado: convirtiendo en segundo plano...")
        self.worker_thread = threading.Thread(target=self._conversion_worker, args=inputs + (
            copy.deepcopy(self.saved_number_overrides), copy.deepcopy(self.saved_manual_assignments),
            copy.deepcopy(self.saved_line_overrides), copy.deepcopy(self.saved_header_overrides)
        ), daemon=True)
        self.worker_thread.start()
        self.root.after(100, self._poll_worker)

    def _conversion_worker(self, invoice_paths, alma_path, output_path, template_path, invoice_number_overrides, manual_assignments, line_overrides, invoice_header_overrides):
        try:
            result = convert_facturae_to_alma_excel(invoice_paths, alma_path, output_path, template_path, invoice_number_overrides, manual_assignments, line_overrides, invoice_header_overrides)
            self.worker_queue.put(("success", result, output_path))
        except Exception as e:
            logging.exception("Error durante la conversion")
            self.worker_queue.put(("error", e, traceback.format_exc()))

    def _poll_worker(self):
        try:
            kind, payload, extra = self.worker_queue.get_nowait()
        except queue.Empty:
            if self.worker_thread and self.worker_thread.is_alive():
                self.root.after(100, self._poll_worker)
            return
        self.convert_button.config(state="normal")
        if kind == "success":
            self._handle_conversion_success(payload, extra)
        else:
            self._handle_conversion_error(payload, extra)

    def _build_summary_message(self, summary):
        return (
            "\n\nResumen de lineas:\n"
            f"- Lineas procesadas: {summary['total_lines']}\n"
            f"- Coincidencias por ISBN: {summary['match_isbn']}\n"
            f"- Coincidencias por ISSN: {summary['match_issn']}\n"
            f"- Coincidencias por titulo: {summary['match_title']}\n"
            f"- Duplicados ISBN: {summary['duplicate_isbn']}\n"
            f"- Duplicados ISSN: {summary['duplicate_issn']}\n"
            f"- Titulos ambiguos: {summary['ambiguous_title']}\n"
            f"- Sin coincidencia: {summary['no_match']}\n"
            f"- Lineas sin PO Line: {summary['lines_without_po_line']}"
        )

    def _handle_conversion_success(self, result, output_path):
        total_invoices, total_issues, validation_path, summary = result
        heading = "CONVERSION COMPLETA SIN INCIDENCIAS" if not validation_path else "INCIDENCIAS EN LA CONVERSION"
        message = (
            f"{heading}\n\n"
            f"Facturas procesadas: {total_invoices}\n"
            f"Avisos de validacion: {total_issues}\n\n"
            f"Excel Alma generado:\n{output_path}"
        ) + self._build_summary_message(summary)
        if validation_path:
            message += f"\n\nInforme de validacion:\n{validation_path}"
        if summary.get("lines_without_po_line", 0) > 0:
            message += "\n\nAtencion: hay lineas sin PO Line. Revisa el informe antes de cargar en Alma."
        self._remember_dir("last_output_dir", output_path)
        self.status_text.set("Estado: conversion completada con incidencias." if validation_path else "Estado: conversion completada sin incidencias.")
        if validation_path:
            messagebox.showwarning("Incidencias en la conversion", message)
        else:
            messagebox.showinfo("Conversion completada", message)

    def _handle_conversion_error(self, error, trace):
        if isinstance(error, DefusedXmlException):
            title = "XML bloqueado por seguridad"
        elif isinstance(error, PermissionError):
            title = "Fichero bloqueado o sin permisos"
        elif isinstance(error, FileNotFoundError):
            title = "Fichero no encontrado"
        elif isinstance(error, ValueError):
            title = "Datos o formato no validos"
        else:
            title = "Error durante la conversion"
        logging.error(trace)
        self.status_text.set("Estado: error durante la conversion.")
        messagebox.showerror(title, f"{error}\n\nConsulta el archivo de log '{LOG_FILE}' para mas detalles.")

    def _on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("Conversion en curso", "Hay una conversion en curso. ¿Quieres cerrar la aplicacion de todos modos?"):
                return
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = FacturaeApp(root)
    root.mainloop()
