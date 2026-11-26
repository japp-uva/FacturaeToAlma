# SPDX-License-Identifier: GPL-3.0-only
# FacturaeToAlma 5.1

import os
import sys
import json
import math
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
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

APP_NAME = "FacturaeToAlma"
APP_VERSION = "5.1"
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
    "lo", "gli", "di", "da", "per", "tra", "fra", "ed", "ma", "oppure", "perche", "perché", "anche", "ti", "ci", "io", "lui", "lei", "noi", "voi",
    "els", "amb", "sense", "sota", "perque", "perquè", "doncs", "jo", "ell", "ella", "nosaltres", "vosaltres", "em",
    "o", "os", "as", "unha", "baixo", "segundo", "sen", "senon", "senón", "nin", "eu", "ela", "nós", "vós", "vos",
    "ak", "ari", "ariak", "bat", "batzuk", "baino", "gabe", "gora", "behera", "arte", "bila", "kontra", "zehar", "aurka", "eta", "edo", "baina", "denez", "beraz", "orduan", "baita", "ezta", "zein", "hi", "hura", "gu", "zuek", "haiek", "hau", "hori",
    "um", "uma", "umas", "com", "sem", "sob", "trás", "ate", "até", "mas", "nem", "embora", "contudo", "pois", "ele", "isto", "esse", "essa", "isso"
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
                raise ValueError(f"Estructura XML bloqueada por seguridad (Bandit / DefusedXML): {e}") from e
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

        seller_nif = self._t(seller, "TaxIdentificationNumber")
        seller_name = self._get_party_name(seller)
        seller_addr = self._get_address(seller)
        seller_tel = self._tp(seller, "LegalEntity/ContactDetails/Telephone") or self._tp(seller, "Individual/ContactDetails/Telephone")
        seller_email = self._tp(seller, "LegalEntity/ContactDetails/ElectronicMail") or self._tp(seller, "Individual/ContactDetails/ElectronicMail")

        buyer_nif = self._t(buyer, "TaxIdentificationNumber")
        buyer_name = self._get_party_name(buyer)
        buyer_addr = self._get_address(buyer)

        dir3_list = []
        if buyer is not None:
            for c in buyer.findall(".//{*}AdministrativeCentre"):
                role_code = self._t(c, "RoleTypeCode")
                role_label = localize_role_label(role_code)
                code = self._t(c, "CentreCode")
                desc = self._t(c, "CentreDescription") or self._t(c, "Name")
                dir3_list.append({"role_code": role_code, "role": role_label, "code": code, "desc": desc})

        header = inv.find(".//{*}InvoiceHeader")
        issue = inv.find(".//{*}InvoiceIssueData")
        totals = inv.find(".//{*}InvoiceTotals")

        invoice_number = self._t(header, "InvoiceNumber")
        series = self._t(header, "InvoiceSeriesCode")
        issue_date = format_date(self._t(issue, "IssueDate"))
        operation_date = format_date(self._t(issue, "OperationDate"))
        currency = self._t(issue, "InvoiceCurrencyCode", "EUR")
        curr_symbol = "€" if currency == "EUR" else currency

        # Totales (lógica exacta 4.0)
        gross_amount = self._t(totals, "TotalGrossAmountBeforeTaxes") or self._t(totals, "TotalGrossAmount")
        total_taxes = self._t(totals, "TotalTaxOutputs")
        total_retention = self._t(totals, "TotalTaxesWithheld", "0.00")
        invoice_total = self._t(totals, "InvoiceTotal")
        total_payable = self._t(totals, "TotalOutstandingAmount") or invoice_total

        # Líneas de Factura (Items - lógica exacta 4.0)
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
        super().__init__(parent, *args, style="App.TFrame", **kwargs)
        self.current_data = None
        self.pdf_canvas = None
        self.placeholder_label = None
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1 if self.show_loader else 0, weight=1)

        self.top_bar = ttk.Frame(self, style="App.TFrame", padding=(0, 0, 0, 8))
        top_bar = self.top_bar
        if self.show_loader:
            top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.columnconfigure(1, weight=1)

        self.elec_inv_label = ttk.Label(top_bar, text=tr("vw_elec_inv_lbl"), font=("Segoe UI", 9, "bold"), background=COLOR_BG)
        self.elec_inv_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.file_var = tk.StringVar()
        ttk.Entry(top_bar, textvariable=self.file_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.load_inv_btn = ttk.Button(top_bar, text=tr("vw_load_btn"), command=self.load_file_dialog, style="Primary.TButton")
        self.load_inv_btn.grid(row=0, column=2)

        self.viewer_notebook = ttk.Notebook(self)
        self.viewer_notebook.grid(row=1 if self.show_loader else 0, column=0, sticky="nsew")

        self.tab_pdf_view = ttk.Frame(self.viewer_notebook, style="App.TFrame")
        self.tab_xml_view = ttk.Frame(self.viewer_notebook, style="App.TFrame")

        self.viewer_notebook.add(self.tab_pdf_view, text=tr("vw_tab_pdf"))
        self.viewer_notebook.add(self.tab_xml_view, text=tr("vw_tab_xml"))

        self._build_pdf_view_ui()
        self._build_xml_view_ui()

    def refresh_language(self):
        if hasattr(self, "elec_inv_label"):
            self.elec_inv_label.config(text=tr("vw_elec_inv_lbl"))
        if hasattr(self, "load_inv_btn"):
            self.load_inv_btn.config(text=tr("vw_load_btn"))
        if hasattr(self, "viewer_notebook"):
            self.viewer_notebook.tab(self.tab_pdf_view, text=tr("vw_tab_pdf"))
            self.viewer_notebook.tab(self.tab_xml_view, text=tr("vw_tab_xml"))
        if self.current_data is not None:
            self._render_pdf_view(self.current_data)
        elif self.placeholder_label is not None and self.placeholder_label.winfo_exists():
            self.placeholder_label.config(text=tr("vw_placeholder"))

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

        self._show_placeholder()

    def _show_placeholder(self):
        for widget in self.scrollable_pdf_frame.winfo_children():
            widget.destroy()
        self.placeholder_label = ttk.Label(
            self.scrollable_pdf_frame,
            text=tr("vw_placeholder"),
            font=("Segoe UI", 10, "italic"),
            foreground=COLOR_TEXT_MUTED,
            background=COLOR_BG
        )
        self.placeholder_label.pack(pady=40)

    def clear_view(self):
        self.current_data = None
        self.file_var.set("")
        self._show_placeholder()
        self.xml_text.config(state="normal")
        self.xml_text.delete("1.0", tk.END)
        self.xml_text.config(state="disabled")

    def _build_xml_view_ui(self):
        container = ttk.Frame(self.tab_xml_view, style="App.TFrame")
        container.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

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
            title=tr("fd_sel_inv_face"),
            filetypes=[
                (tr("fd_ft_compat_inv"), "*.xsig *.xml *.txt"),
                (tr("fd_ft_xsig"), "*.xsig"),
                (tr("fd_ft_xml"), "*.xml"),
                (tr("fd_ft_txt"), "*.txt"),
                (tr("fd_ft_all"), "*.*")
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
            messagebox.showerror(tr("mb_err_read_inv_title"), tr("mb_err_read_inv_msg", err=e))

    def _render_pdf_view(self, data):
        for widget in self.scrollable_pdf_frame.winfo_children():
            widget.destroy()
        self.placeholder_label = None

        sheet = tk.Frame(self.scrollable_pdf_frame, bg=COLOR_CARD, padx=24, pady=20, highlightbackground=COLOR_BORDER, highlightthickness=1)
        sheet.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        sheet.columnconfigure(0, weight=1)
        sheet.columnconfigure(1, weight=1)

        hdr_box = tk.Frame(sheet, bg=COLOR_CARD)
        hdr_box.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        hdr_box.columnconfigure(0, weight=1)
        hdr_box.columnconfigure(1, weight=1)

        title_lbl = tk.Label(hdr_box, text=tr("vw_hdr_invoice", number=data["invoice_number"]), font=("Segoe UI", 16, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD)
        title_lbl.grid(row=0, column=0, sticky="w")

        meta_text = tr(
            "vw_hdr_meta",
            series=data["series"] or "-",
            issue=data["issue_date"],
            op=data["operation_date"] or data["issue_date"],
        )
        meta_lbl = tk.Label(hdr_box, text=meta_text, font=("Segoe UI", 9), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD)
        meta_lbl.grid(row=1, column=0, sticky="w", pady=(2, 0))

        emisor_card = tk.LabelFrame(sheet, text=tr("vw_card_seller"), font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=8, highlightbackground=COLOR_BORDER)
        emisor_card.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 12))

        receptor_card = tk.LabelFrame(sheet, text=tr("vw_card_buyer"), font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=8, highlightbackground=COLOR_BORDER)
        receptor_card.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(0, 12))

        tk.Label(emisor_card, text=data["seller_name"], font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w")
        tk.Label(emisor_card, text=tr("vw_nif", nif=data["seller_nif"]), font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))
        tk.Label(emisor_card, text=tr("vw_addr", addr=data["seller_addr"]), font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w", pady=(2, 0))
        if data["seller_tel"] or data["seller_email"]:
            contact_val = f"{data['seller_tel']}  {data['seller_email']}".strip()
            tk.Label(emisor_card, text=tr("vw_contact", contact=contact_val), font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))

        tk.Label(receptor_card, text=data["buyer_name"], font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w")
        tk.Label(receptor_card, text=tr("vw_nif", nif=data["buyer_nif"]), font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))
        tk.Label(receptor_card, text=tr("vw_addr", addr=data["buyer_addr"]), font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD, wraplength=420, justify="left").pack(anchor="w", pady=(2, 0))

        if data["dir3_list"]:
            dir3_card = tk.LabelFrame(sheet, text=tr("vw_card_dir3"), font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=6)
            dir3_card.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 14))
            for item in data["dir3_list"]:
                role_txt = localize_role_label(item.get("role_code", "")) if item.get("role_code") is not None else item["role"]
                line_text = f"• {role_txt}: [{item['code']}] {item['desc']}"
                tk.Label(dir3_card, text=line_text, font=("Segoe UI", 8), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD, anchor="w").pack(fill=tk.X, pady=1)

        items_frame = tk.LabelFrame(sheet, text=tr("vw_card_lines"), font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=8, pady=8)
        items_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0, 14))
        items_frame.columnconfigure(0, weight=1)

        cols = ("line", "code", "desc", "qty", "price", "gross", "tax")
        tree = ttk.Treeview(items_frame, columns=cols, show="headings", height=min(max(len(data["lines"]), 4), 14))
        tree.heading("line", text="#")
        tree.heading("code", text=tr("vw_col_code"))
        tree.heading("desc", text=tr("vw_col_desc"))
        tree.heading("qty", text=tr("vw_col_qty"))
        tree.heading("price", text=tr("vw_col_price"))
        tree.heading("gross", text=tr("vw_col_gross"))
        tree.heading("tax", text=tr("vw_col_tax"))

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

        bottom_frame = tk.Frame(sheet, bg=COLOR_CARD)
        bottom_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)

        pay_box = tk.LabelFrame(bottom_frame, text=tr("vw_card_pay"), font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD, padx=12, pady=10)
        pay_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        tk.Label(pay_box, text=tr("vw_pay_means", code=data["payment_means"] or "04"), font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w")
        tk.Label(pay_box, text=tr("vw_pay_due", date=data["due_date"] or data["issue_date"]), font=("Segoe UI", 9), fg=COLOR_TEXT_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(2, 0))
        tk.Label(pay_box, text=tr("vw_pay_iban", iban=data["iban"] or tr("vw_not_specified")), font=("Segoe UI", 9, "bold"), fg=COLOR_PRIMARY, bg=COLOR_CARD).pack(anchor="w", pady=(4, 0))

        totals_box = tk.Frame(bottom_frame, bg=COLOR_ROW_ALT, padx=14, pady=10, highlightbackground=COLOR_BORDER, highlightthickness=1)
        totals_box.grid(row=0, column=1, sticky="nsew")
        totals_box.columnconfigure(0, weight=1)
        totals_box.columnconfigure(1, weight=1)

        def add_tot_row(label, val, row, is_grand=False):
            fg = COLOR_PRIMARY if is_grand else COLOR_TEXT_PRIMARY
            font = ("Segoe UI", 11, "bold") if is_grand else ("Segoe UI", 9)
            tk.Label(totals_box, text=label, font=font, fg=fg, bg=COLOR_ROW_ALT).grid(row=row, column=0, sticky="w", pady=2)
            tk.Label(totals_box, text=val, font=font, fg=fg, bg=COLOR_ROW_ALT).grid(row=row, column=1, sticky="e", pady=2)

        add_tot_row(tr("vw_tot_base"), format_currency(data["gross_amount"], data["currency_symbol"]), 0)
        add_tot_row(tr("vw_tot_taxes"), format_currency(data["total_taxes"], data["currency_symbol"]), 1)
        if Decimal(data["total_retention"] or "0.00") > 0:
            add_tot_row(tr("vw_tot_ret"), format_currency(data["total_retention"], data["currency_symbol"]), 2)
        ttk.Separator(totals_box, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=4)
        add_tot_row(tr("vw_tot_grand"), format_currency(data["invoice_total"], data["currency_symbol"]), 4, is_grand=True)
        add_tot_row(tr("vw_tot_payable"), format_currency(data["total_payable"], data["currency_symbol"]), 5)

    def _render_xml_view(self, raw_xml):
        self.xml_text.config(state="normal")
        self.xml_text.delete("1.0", tk.END)

        try:
            dom = defusedxml.minidom.parseString(raw_xml.encode("utf-8"))
            pretty_xml = dom.toprettyxml(indent="  ")
            lines = [line for line in pretty_xml.splitlines() if line.strip()]
            formatted_text = "\n".join(lines)
        except Exception:
            formatted_text = raw_xml

        self.xml_text.insert(tk.END, formatted_text)

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
    main_title_words: list = field(default_factory=list)

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
    xml_po_line: str = None

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
    parts = re.split(r"[;,/\s]+", str(value))
    return [norm for norm in (normalize_identifier(part) for part in parts) if norm]

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

def isbn10_to_isbn13(isbn10):
    code = normalize_identifier(isbn10)
    if len(code) != 10 or not is_valid_isbn10(code):
        return None
    stem = "978" + code[:9]
    total = sum(int(c) if i % 2 == 0 else int(c) * 3 for i, c in enumerate(stem))
    check = (10 - (total % 10)) % 10
    return f"{stem}{check}"

def isbn13_to_isbn10(isbn13):
    code = normalize_identifier(isbn13)
    if len(code) != 13 or not code.startswith("978") or not is_valid_isbn13(code):
        return None
    stem = code[3:12]
    total = sum((10 - i) * int(c) for i, c in enumerate(stem))
    rem = (11 - (total % 11)) % 11
    check = "X" if rem == 10 else str(rem)
    return f"{stem}{check}"

def expand_isbn_variants(isbn):
    code = normalize_identifier(isbn)
    if not code:
        return []
    variants = [code]
    if len(code) == 10 and is_valid_isbn10(code):
        v13 = isbn10_to_isbn13(code)
        if v13 and v13 not in variants:
            variants.append(v13)
    elif len(code) == 13 and is_valid_isbn13(code):
        v10 = isbn13_to_isbn10(code)
        if v10 and v10 not in variants:
            variants.append(v10)
    return variants

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

def extract_po_line_from_text(*texts):
    pattern = re.compile(r"\b(POL-\d+)\b", re.IGNORECASE)
    for text in texts:
        if not text:
            continue
        match = pattern.search(str(text))
        if match:
            return match.group(1).upper()
    return None

SHIPMENT_KEYWORDS = (
    "gastos de envio", "gastos envio", "portes", "gastos de expedicion",
    "gastos expedicion", "transporte y envio", "shipping charges",
    "shipping and handling", "delivery charges", "freight charges", "postage",
)
DISCOUNT_KEYWORDS = (
    "descuento comercial", "bonificacion", "descuento pronto pago",
)

def detect_special_line_type(title, price=None, isbn=None, issn=None):
    if isbn or issn:
        return "REGULAR"
    norm = remove_accents(str(title or "").lower()).strip()
    norm_clean = re.sub(r"[^a-z0-9\s]+", " ", norm)
    norm_clean = re.sub(r"\s+", " ", norm_clean).strip()
    if norm_clean and any(kw in norm_clean for kw in SHIPMENT_KEYWORDS) and len(norm_clean) <= 55:
        return "SHIPMENT"
    if price is not None and price < Decimal("0.00"):
        return "DISCOUNT"
    if norm_clean and any(kw in norm_clean for kw in DISCOUNT_KEYWORDS) and len(norm_clean) <= 55:
        return "DISCOUNT"
    return "REGULAR"

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

def normalize_tax_id(value):
    if not value:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).strip().upper())

SUPPORTED_LANGUAGES = (('es', 'Español / Castellano (es)'),
 ('ca', 'Català (ca)'),
 ('va', 'Valencià (va)'),
 ('gl', 'Galego (gl)'),
 ('eu', 'Euskara (eu)'),
 ('oc', 'Aranés / Occitan (oc)'))
LANGUAGE_CODE_TO_LABEL = dict(SUPPORTED_LANGUAGES)
LANGUAGE_LABEL_TO_CODE = {label: code for code, label in SUPPORTED_LANGUAGES}

STATUS_FILTER_CODES = ("all", "problems", "unmatched", "ok", "edited")
STATUS_FILTER_KEYS = {
    "all": "ed_flt_st_all",
    "problems": "ed_flt_st_prob",
    "unmatched": "ed_flt_st_unm",
    "ok": "ed_flt_st_ok",
    "edited": "ed_flt_st_edit",
}

FILTER_OP_CODES = ("contains", "exact", "not_equal", "blank", "not_blank")
FILTER_OP_KEYS = {
    "contains": "ed_flt_op_contains",
    "exact": "ed_flt_op_exact",
    "not_equal": "ed_flt_op_neq",
    "blank": "ed_flt_op_blank",
    "not_blank": "ed_flt_op_nblank",
}

VENDOR_IMPORT_MODES = ("incremental", "only_new", "full")
VENDOR_IMPORT_MODE_KEYS = {
    "incremental": "cust_vend_mode_inc",
    "only_new": "cust_vend_mode_new",
    "full": "cust_vend_mode_full",
}

VENDOR_EXCEL_HEADERS = ("NIF / CIF", "Vendor Account")

I18N = {'ca': {'afdlg_apply_next': 'Aplicar i continuar filtrant una altra columna',
        'afdlg_blank_chk': '(Buides / Sense valor)',
        'afdlg_check_all': 'Marcar tots',
        'afdlg_col_lbl': 'Columna a filtrar:',
        'afdlg_list_frame': 'Valors presents a la columna (marca els que vols mostrar)',
        'afdlg_only_blank': 'Només buides',
        'afdlg_only_val': 'Només amb valor',
        'afdlg_remove_col': "Treure filtre d'aquesta columna",
        'afdlg_search_lbl': 'Cercar text / valor:',
        'afdlg_title': 'Filtre autofiltre per columnes (estil Excel)',
        'afdlg_uncheck_all': 'Desmarcar tots',
        'bulkdlg_lbl': 'Valor per a {col}:',
        'bulkdlg_title': 'Assignar {col}',
        'conv_actions_frame': '4. Accions',
        'conv_alma_btn': 'Seleccionar fitxer...',
        'conv_alma_lbl': "Fitxer Excel d'Alma Analytics amb PO Lines *",
        'conv_batch_add': 'Afegir factures',
        'conv_batch_clear': 'Buidar lot',
        'conv_batch_count': '{count} factures seleccionades',
        'conv_batch_frame': 'Factures del lot',
        'conv_batch_hint': '(Doble clic en una factura per obrir-la a Editar)',
        'conv_batch_remove': 'Treure seleccionades',
        'conv_batch_sort': 'Ordenar per número',
        'conv_clear_btn': 'Netejar dades',
        'conv_col_file': 'Fitxer',
        'conv_col_invoice': 'Número de factura',
        'conv_col_vendor': 'Proveïdor',
        'conv_exit_btn': 'Sortir',
        'conv_inputs_frame': "2. Fitxers d'entrada",
        'conv_inv_batch_btn': 'Seleccionar factures...',
        'conv_inv_batch_lbl': 'Factures XSIG / XML / TXT del mateix proveïdor i biblioteca *',
        'conv_inv_single_btn': 'Seleccionar factura...',
        'conv_inv_single_lbl': 'Factura electrònica XSIG / XML / TXT *',
        'conv_mode_batch': 'Lot de factures del mateix proveïdor i biblioteca',
        'conv_mode_frame': '1. Mode de conversió',
        'conv_mode_single': 'Factura individual',
        'conv_ok_heading': 'CONVERSIÓ COMPLETA SENSE INCIDÈNCIES',
        'conv_open_out_btn': 'Obrir carpeta de sortida',
        'conv_output_btn': 'Triar destinació...',
        'conv_output_frame': '3. Fitxer de sortida',
        'conv_result_body': '{heading}\n'
                            '\n'
                            'Factures processades: {total_inv}\n'
                            'Avisos de validació: {total_iss}\n'
                            '\n'
                            'Excel Alma generat:\n'
                            '{out}',
        'conv_run_btn': 'Convertir a format Alma',
        'conv_tpl_btn': 'Seleccionar plantilla...',
        'conv_tpl_lbl': 'Plantilla Alma .xlsx *',
        'conv_unmatched_warn': '\n'
                               '\n'
                               "Atenció: hi ha línies sense PO Line. Revisa l'informe abans de carregar a "
                               'Alma.',
        'conv_val_report': '\n\nInforme de validació:\n{path}',
        'conv_warn_heading': 'INCIDÈNCIES EN LA CONVERSIÓ',
        'ctx_bulk_apply': 'Aplicar valor a les línies seleccionades...',
        'ctx_clear_cell': 'Buidar cel·la (Supr)',
        'ctx_edit_cell': 'Editar cel·la',
        'ctx_fill_blanks_down': 'Emplenar blancs inferiors (■)',
        'ctx_fill_blanks_sel': 'Emplenar blancs seleccionats',
        'ctx_fill_down': 'Emplenar cap avall (Ctrl+D)',
        'ctx_flt_cell_val': "Filtrar pel valor d'aquesta cel·la",
        'ctx_flt_clear_all': 'Treure tots els filtres',
        'ctx_flt_col_excel': 'Filtrar aquesta columna (estil Excel)...',
        'ctx_redo': 'Refer (Ctrl+Y)',
        'ctx_restore_cell': 'Restablir cel·la',
        'ctx_restore_orig': 'Restablir cel·la original',
        'ctx_search_pol': "Cercar PO Line al catàleg d'Alma...",
        'ctx_suggest_ctx': 'Suggerir valor pel context',
        'ctx_undo': 'Desfer (Ctrl+Z)',
        'cust_choose_dir_btn': 'Seleccionar carpeta...',
        'cust_dir_alma': "Directori del fitxer Excel d'Alma",
        'cust_dir_inv': 'Directori de factures FACe',
        'cust_dir_out': "Directori d'Excel finals",
        'cust_dir_tpl': 'Directori de plantilla Alma',
        'cust_inst_lbl': 'Institució',
        'cust_lang_lbl': 'Idioma de la interfície',
        'cust_logo_del_btn': 'Eliminar logotip',
        'cust_logo_hint': 'PNG, JPG o JPEG. Alçada màxima: 100 px.',
        'cust_logo_invalid': 'El fitxer seleccionat no és una imatge vàlida.',
        'cust_logo_lbl': 'Logotip institucional',
        'cust_logo_loaded': 'Logotip carregat: {name}',
        'cust_logo_none': 'Sense logotip. PNG, JPG o JPEG; alçada màxima: 100 px.',
        'cust_logo_sel_btn': 'Seleccionar logotip...',
        'cust_logo_unavail': 'El logotip configurat no està disponible.',
        'cust_pref_frame': 'Preferències',
        'cust_pref_note': 'Tots els camps són opcionals. Si una carpeta personalitzada queda buida, es fa '
                          "servir l'última carpeta vàlida; en la primera execució s'utilitza la carpeta de "
                          "l'aplicació. La sortida es desa al costat de la factura, llevat que es defineixi "
                          'un directori de sortida.',
        'cust_reset_pref_btn': 'Restablir valors',
        'cust_save_pref_btn': 'Desar preferències',
        'cust_vend_acc_lbl': 'Vendor Account a Alma:',
        'cust_vend_add_btn': 'Afegir / Actualitzar',
        'cust_vend_clr_btn': 'Buidar llista',
        'cust_vend_col_acc': 'Vendor Account a Alma',
        'cust_vend_col_nif': 'NIF / CIF normalitzat',
        'cust_vend_del_btn': 'Eliminar seleccionat',
        'cust_vend_exp_btn': 'Exportar Excel...',
        'cust_vend_frame': 'Proveïdors recordats automàticament (NIF / CIF → Vendor Account a Alma)',
        'cust_vend_imp_btn': 'Importar Excel...',
        'cust_vend_imp_mode_lbl': "Mode d'importació Excel:",
        'cust_vend_mode_full': 'Càrrega completa (reemplaçar llista actual)',
        'cust_vend_mode_inc': 'Incremental / acumulatiu (afegir i actualitzar)',
        'cust_vend_mode_new': 'Només afegir nous (no sobreescriure)',
        'cust_vend_nif_lbl': 'NIF / CIF:',
        'dlg_accept': "D'acord",
        'dlg_apply': 'Aplicar',
        'dlg_cancel': 'Cancel·lar',
        'ed_apply_num_btn': 'Aplicar número',
        'ed_bottom_hint': 'Escriu directament o fes doble clic per editar | Arrossega ■ per emplenar | '
                          'Ctrl+F cercar | Clic a capçaleres ▼ per filtrar.',
        'ed_col_invoice': 'Factura',
        'ed_col_line': 'Línia',
        'ed_diag_edited': '✎ Línia {line}: modificada manualment en aquesta sessió.',
        'ed_diag_init': "Selecciona una línia per veure'n el diagnòstic de creuament i validació.",
        'ed_diag_manual': "✓ Línia {line}: PO Line '{pol}' assignada manualment.",
        'ed_diag_matched': '✓ Línia {line}: {msg}.',
        'ed_diag_no_inv': 'Selecciona una factura per revisar-ne les línies i els avisos.',
        'ed_diag_ready': '✓ Línia {line}: llesta per exportar.',
        'ed_diag_recon': '⚠ Quadrament factura: {msg}',
        'ed_diag_unmatched': "⚠ Línia {line}: sense PO Line assignada (fes servir 'Cercar PO Line a Alma...' "
                             'o escriu-la a la cel·la).',
        'ed_diag_warning': '⚠ Línia {line}: {notes}',
        'ed_flt_active_prefix': 'Filtres actius -> ',
        'ed_flt_add_btn': '+ Afegir filtre',
        'ed_flt_clear_btn': 'Netejar filtres',
        'ed_flt_col_lbl': 'Columna:',
        'ed_flt_contains_q': " conté '{q}'",
        'ed_flt_excel_btn': 'Filtre Excel...',
        'ed_flt_frame': 'Filtrar línies de factura (per estat i per una o diverses columnes estil Excel)',
        'ed_flt_in_list': '{col} a [{items}]',
        'ed_flt_none': 'Sense filtres actius (fes clic en qualsevol capçalera ▼ per filtrar per una o més '
                       'columnes estil Excel).',
        'ed_flt_op_blank': 'Buides',
        'ed_flt_op_contains': 'Conté / Igual',
        'ed_flt_op_exact': 'Igual exacte',
        'ed_flt_op_nblank': 'No buides',
        'ed_flt_op_neq': 'Diferent de',
        'ed_flt_st_all': 'Totes les línies',
        'ed_flt_st_edit': 'Només línies modificades manualment',
        'ed_flt_st_ok': 'Només línies correctes (amb PO Line i sense avisos)',
        'ed_flt_st_prob': 'Només línies amb problemes (sense PO Line o amb avisos)',
        'ed_flt_st_unm': 'Només línies sense PO Line (pendents)',
        'ed_flt_state_lbl': 'Estat:',
        'ed_flt_tag': 'Filtre',
        'ed_flt_val_blank': '(Buides)',
        'ed_flt_val_lbl': 'Valor:',
        'ed_flt_val_nblank': '(No buides)',
        'ed_flt_values_cnt': '{col} ({count} valors)',
        'ed_hdr_frame': 'Capçalera de factura (INV)',
        'ed_load_draft_btn': 'Carregar esborrany...',
        'ed_mode_orig': 'Original',
        'ed_mode_prefix': 'Prefix',
        'ed_mode_suffix': 'Sufix',
        'ed_num_frame': "Número de factura a l'Excel d'Alma",
        'ed_orig_lbl': 'Original:',
        'ed_refresh_btn': 'Actualitzar esborrany',
        'ed_reset_inv_btn': 'Restablir factura',
        'ed_save_draft_btn': 'Desar esborrany...',
        'ed_save_excel_btn': 'Desar canvis i actualitzar Excel',
        'ed_search_pol_btn': 'Cercar PO Line a Alma...',
        'ed_stats_empty': 'Mostrant: 0/0 | Correctes: 0 | Amb problemes: 0',
        'ed_stats_fmt': 'Mostrant: {shown}/{total} | Correctes: {ok} | Amb problemes: {prob} (Sense POL: '
                        '{unm})',
        'ed_tb_bulk': 'Assignar valor a seleccionades...',
        'ed_tb_clear': 'Buidar cel·la (Supr)',
        'ed_tb_copy': 'Copiar (Ctrl+C)',
        'ed_tb_fill_blanks': '⬇ Emplenar blancs (■)',
        'ed_tb_fill_down': '⬇ Emplenar avall (Ctrl+D)',
        'ed_tb_paste': 'Enganxar (Ctrl+V)',
        'ed_tb_redo': '↪ Refer (Ctrl+Y)',
        'ed_tb_restore': 'Restablir cel·la',
        'ed_tb_search_lbl': 'Cercar (Ctrl+F):',
        'ed_tb_undo': '↩ Desfer (Ctrl+Z)',
        'ed_tot_init': 'Import visible: 0,00 €',
        'ed_tot_sel': '   |   Selecció ({count}): {base}',
        'ed_tot_vis': 'Visible ({count}): Base {base} | IVA {vat}',
        'err_conv_title': 'Error durant la conversió',
        'err_fnf_title': 'Fitxer no trobat',
        'err_log_hint': "{err}\n\nConsulta el fitxer de registre '{log}' per a més detalls.",
        'err_perm_title': 'Fitxer bloquejat o sense permisos',
        'err_val_title': 'Dades o format no vàlids',
        'err_xml_sec_title': 'XML bloquejat per seguretat',
        'fd_exp_vendors': 'Exportar proveïdors (NIF / CIF i Vendor Account) a Excel',
        'fd_ft_all': 'Tots els fitxers',
        'fd_ft_compat_inv': 'Factures compatibles',
        'fd_ft_draft': 'Esborrany FacturaeToAlma',
        'fd_ft_excel': 'Fitxers Excel',
        'fd_ft_img_compat': 'Imatges compatibles',
        'fd_ft_jpg': 'Imatge JPEG',
        'fd_ft_png': 'Imatge PNG',
        'fd_ft_txt': 'Fitxers TXT',
        'fd_ft_xml': 'Fitxers XML',
        'fd_ft_xsig': 'Fitxers XSIG',
        'fd_imp_vendors': "Importar proveïdors (NIF / CIF i Vendor Account) des d'Excel",
        'fd_load_draft': 'Carregar esborrany de sessió FacturaeToAlma',
        'fd_save_draft': 'Desar esborrany de sessió FacturaeToAlma',
        'fd_sel_alma': "Seleccionar fitxer Excel d'Alma Analytics",
        'fd_sel_folder': 'Seleccionar carpeta',
        'fd_sel_inv_batch': 'Seleccionar factures del mateix proveïdor i biblioteca',
        'fd_sel_inv_face': 'Seleccionar factura electrònica FACe',
        'fd_sel_inv_single': 'Seleccionar factura XSIG / XML / TXT',
        'fd_sel_logo': 'Seleccionar logotip institucional',
        'fd_sel_out': 'Definir fitxer Excel de sortida',
        'fd_sel_tpl': 'Seleccionar plantilla Alma .xlsx',
        'gctx_copy': 'Copiar',
        'gctx_copy_txt': 'Copiar text',
        'gctx_cut': 'Retallar',
        'gctx_go_edit': 'Anar a Editar',
        'gctx_go_vis': 'Anar a Visualitzar',
        'gctx_paste': 'Enganxar',
        'gctx_sel_all': 'Seleccionar-ho tot',
        'hlp_body_5_0': "Factura individual: selecciona una factura, l'informe d'Alma Analytics, la "
                        'plantilla i la destinació.\n'
                        '\n'
                        'Mode lot: totes les factures han de pertànyer al mateix proveïdor i biblioteca. La '
                        'llista permet afegir, treure, ordenar i obrir cada factura amb doble clic a '
                        'Editar.\n'
                        '\n'
                        'Edició i filtratge: revisa la capçalera (INV) i les línies (IL), cerca PO Lines al '
                        "catàleg d'Alma Analytics, aplica autofiltres estil Excel o desa/carrega esborranys "
                        '(.ftalma.json).\n'
                        '\n'
                        "Proveïdors i idiomes: a Personalitzar pots canviar l'idioma oficial/cooficial (es, "
                        'ca, va, gl, eu, oc) i gestionar, importar o exportar en Excel la taula de '
                        'proveïdors (NIF / CIF i Vendor Account) en mode incremental, només nous o càrrega '
                        'completa.\n'
                        '\n'
                        'Informe de validació: es genera quan falten PO Lines, hi ha coincidències ambigües, '
                        'duplicats, desquadraments comptables o diferències entre Quantity i Quantity for '
                        'Pricing.\n'
                        '\n'
                        'Limitacions confirmades per Ex Libris: la plantilla Excel no permet indicar Line '
                        'Exclusive ni declarar explícitament facturació parcial o completa.\n'
                        '\n'
                        'Llicència: GPL-3.0-only.',
        'hlp_body_5_1': "Factura individual: selecciona una factura, l'informe d'Alma Analytics, la "
                        'plantilla i la destinació.\n'
                        '\n'
                        'Mode lot: totes les factures han de pertànyer al mateix proveïdor i biblioteca. La '
                        'llista permet afegir, treure, ordenar i obrir cada factura amb doble clic a '
                        'Editar.\n'
                        '\n'
                        "Imports i IVA compatibles amb Línia exclusiva (lògica exacta 4.0): l'extracció i "
                        "l'exportació d'imports i IVA reprodueixen exactament el comportament de la versió "
                        '4.0 perquè el canvi manual a «Línia exclusiva» a Alma quadri sempre.\n'
                        '\n'
                        'Proveïdors i idiomes: a Personalitzar pots seleccionar qualsevol de les llengües '
                        "oficials i cooficials d'Espanya (es, ca, va, gl, eu, oc) i importar o exportar en "
                        'Excel la taula de proveïdors (només columnes NIF / CIF i Vendor Account) en mode '
                        'incremental/acumulatiu, només afegir nous o càrrega completa.\n'
                        '\n'
                        'Informe de validació: es genera quan falten PO Lines, hi ha coincidències ambigües, '
                        'duplicats o diferències entre Quantity i Quantity for Pricing.\n'
                        '\n'
                        'Limitacions confirmades per Ex Libris: la plantilla Excel no permet indicar Line '
                        'Exclusive ni declarar explícitament facturació parcial o completa.\n'
                        '\n'
                        'Llicència: GPL-3.0-only.',
        'hlp_guide_btn': 'Obrir Biblioguia / LibGuides',
        'hlp_links_frame': 'Documentació i suport',
        'hlp_repo_btn': 'Obrir repositori de GitHub',
        'match_msg_ambig': 'Diverses coincidències possibles per títol',
        'match_msg_dup': 'Duplicat per {mtype}. Cal revisió manual.',
        'match_msg_manual': 'PO Line assignada manualment',
        'match_msg_none': 'Sense coincidència a Alma per ISBN, ISSN ni títol',
        'match_msg_special': 'Línia de tipus {ltype} (no requereix PO Line bibliogràfica)',
        'match_msg_unique': 'Coincidència única per {mtype}',
        'mb_batch_clr_ask': 'Vols eliminar totes les factures del lot?',
        'mb_batch_clr_title': 'Buidar lot',
        'mb_batch_dup_num': 'Número de factura duplicat {num}: {name}',
        'mb_batch_none_added': "No s'han incorporat factures noves.",
        'mb_batch_read_err': "No s'ha pogut llegir {name}: {err}",
        'mb_batch_rep_path': 'Ruta repetida: {name}',
        'mb_batch_sel_one': 'Selecciona almenys una factura de la llista.',
        'mb_batch_title': 'Lot',
        'mb_batch_warn_title': 'Factures no incorporades',
        'mb_clr_alma_ask': "Vols netejar també el fitxer Excel d'Alma Analytics?\n"
                           '\n'
                           'La plantilla Alma es conservarà.',
        'mb_clr_alma_title': "Netejar fitxer Excel d'Alma",
        'mb_conv_close_ask': "Hi ha una conversió en curs. Vols tancar l'aplicació igualment?",
        'mb_conv_ok_title': 'Conversió completada',
        'mb_conv_running_msg': 'Ja hi ha una conversió en curs.',
        'mb_conv_running_title': 'Conversió en curs',
        'mb_conv_warn_title': 'Incidències en la conversió',
        'mb_draft_loaded_msg': "S'han restaurat les factures i els canvis de l'esborrany.{warn}",
        'mb_draft_loaded_title': 'Esborrany carregat',
        'mb_draft_missing_warn': '\n'
                                 '\n'
                                 "Avís: no s'han trobat algunes factures mogudes o eliminades:\n"
                                 '- {files}',
        'mb_draft_outdated_msg': 'La factura, Analytics o la plantilla han canviat després de desar. Torna a '
                                 'Editar, revisa-ho i desa de nou.',
        'mb_draft_outdated_title': 'Esborrany desactualitzat',
        'mb_draft_rebuilt_alma': "S'ha reconstruït l'esborrany de l'Excel final d'Alma amb la factura i "
                                 "l'informe d'Analytics seleccionats actualment.",
        'mb_draft_rebuilt_no_alma': "S'han carregat les línies de la factura a l'editor.\n"
                                    '\n'
                                    "Si selecciones també el fitxer Excel d'Alma Analytics, es creuaran "
                                    'automàticament les PO Lines, els Reporting Codes i els fons.',
        'mb_draft_saved_msg': "S'ha desat el progrés actual a:\n"
                              '{path}\n'
                              '\n'
                              "Podràs reprendre l'edició més tard amb 'Carregar esborrany...'.",
        'mb_draft_saved_title': 'Esborrany desat',
        'mb_draft_unavail_msg': "No s'ha pogut preparar l'esborrany editable abans de convertir.",
        'mb_draft_unavail_title': 'Esborrany no disponible',
        'mb_empty_src_msg': 'La cel·la activa està buida. Escriu o selecciona una cel·la amb valor abans '
                            "d'emplenar els blancs.",
        'mb_empty_src_title': "Cel·la d'origen buida",
        'mb_err_folder_msg': "No s'ha pogut obrir la carpeta:\n{path}\n\n{err}",
        'mb_err_folder_title': 'Error en obrir la carpeta',
        'mb_err_load_draft': "Error en carregar l'esborrany",
        'mb_err_load_draft_msg': "No s'ha pogut carregar l'esborrany:\n{err}",
        'mb_err_prep_draft': "No s'ha pogut preparar l'esborrany",
        'mb_err_read_inv_msg': "No s'ha pogut carregar la factura seleccionada:\n{err}",
        'mb_err_read_inv_title': 'Error en llegir la factura',
        'mb_err_reset_title': 'Error en restablir',
        'mb_err_save_draft': "Error en desar l'esborrany",
        'mb_err_save_draft_msg': "No s'ha pogut desar l'esborrany:\n{err}",
        'mb_err_save_title': 'Error en desar',
        'mb_err_url_msg': "No s'ha pogut obrir l'enllaç:\n{url}\n\n{err}",
        'mb_err_url_title': "Error en obrir l'enllaç",
        'mb_flt_cols_no_inv': 'Selecciona primer una factura a la pestanya Editar.',
        'mb_flt_cols_title': 'Filtrar columnes',
        'mb_inv_invalid_title': 'Factura no vàlida',
        'mb_invalid_date_msg': 'Fes servir una data vàlida, per exemple 2026-10-19.',
        'mb_invalid_date_title': 'Data no vàlida',
        'mb_invalid_inputs_title': "Dades d'entrada no vàlides",
        'mb_invalid_ltype': 'Selecciona un Line type permès per Alma.',
        'mb_invalid_num': '{col} ha de ser numèric.',
        'mb_invalid_val_title': 'Valor no vàlid',
        'mb_logo_invalid_msg': "No s'ha pogut carregar el logotip:\n{path}\n\n{err}",
        'mb_logo_invalid_title': 'Logotip no vàlid',
        'mb_logo_unavail_msg': "No s'ha trobat el fitxer de logotip:\n{path}",
        'mb_logo_unavail_title': 'Logotip no disponible',
        'mb_no_active_inv': 'No hi ha cap factura activa.',
        'mb_no_edit_inv': "No hi ha cap factura activa a l'editor.",
        'mb_no_out_dir_def': 'No hi ha cap carpeta de sortida definida.',
        'mb_no_out_dir_found': "No s'ha trobat la carpeta de sortida.",
        'mb_no_sugg_msg': 'No hi ha valors de context per suggerir en aquesta columna.',
        'mb_no_sugg_title': 'Sense suggeriments',
        'mb_pol_ask_alma': 'Aquest cercador consulta les línies de comanda (PO Lines) dins del fitxer Excel '
                           "exportat d'Alma Analytics.\n"
                           '\n'
                           "Encara no has seleccionat cap fitxer Excel d'Alma Analytics.\n"
                           'Vols seleccionar-lo ara?',
        'mb_pol_no_inv': 'Selecciona primer una factura (o lot de factures) a la pestanya Convertir.',
        'mb_pol_no_lines': "No hi ha línies visibles a l'editor.",
        'mb_pol_sel_row': 'Selecciona una fila de la llista.',
        'mb_pol_title': 'Cercar PO Line a Alma',
        'mb_pref_invalid_dir': '{label} no és una carpeta vàlida:\n{val}',
        'mb_pref_invalid_title': 'Preferències no vàlides',
        'mb_pref_reset_ask': "Vols restablir la institució, l'idioma i tots els directoris personalitzats?",
        'mb_pref_reset_ok': "S'han restablert els valors inicials.",
        'mb_pref_reset_title': 'Restablir preferències',
        'mb_pref_saved_msg': "Les preferències s'han desat correctament.",
        'mb_prot_drag_msg': 'Aquesta columna està protegida i no es pot emplenar arrossegant.',
        'mb_prot_field_title': 'Camp protegit',
        'mb_prot_generic_msg': 'Aquest camp no és editable.',
        'mb_prot_hdr_msg': "El número, la data i els imports de capçalera no s'editen directament. El número "
                           'només admet prefix o sufix.',
        'mb_prot_line_msg': 'Aquest camp prové de la factura original i no és editable a FacturaeToAlma.',
        'mb_reset_inv_ask': "Vols descartar tots els canvis d'aquesta factura?",
        'mb_reset_inv_title': 'Restablir factura',
        'mb_save_changes_ask': "Vols desar tots els canvis i actualitzar ara l'Excel de sortida?\n"
                               '\n'
                               "Es sobreescriurà l'Excel indicat a Convertir. L'XML original no es "
                               'modificarà.',
        'mb_save_changes_title': 'Desar canvis',
        'mb_save_draft_empty': 'No hi ha cap factura carregada per desar en un esborrany.',
        'mb_save_draft_title': 'Desar esborrany',
        'mb_unsaved_ed_ask': 'Hi ha canvis fets a Editar (número de factura, capçalera o línies) que encara '
                             "no s'havien desat.\n"
                             '\n'
                             'Vols aplicar i desar aquests canvis per convertir ara?',
        'mb_unsaved_ed_title': 'Canvis pendents a Editar',
        'mb_vend_clr_ask': 'Vols eliminar totes les associacions NIF → Vendor Account desades?',
        'mb_vend_clr_title': 'Buidar proveïdors recordats',
        'mb_vend_exp_err_title': 'Error en exportar proveïdors',
        'mb_vend_exp_ok_msg': "S'han exportat {count} proveïdors (columnes NIF / CIF i Vendor Account) al "
                              'fitxer:\n'
                              '{path}',
        'mb_vend_exp_ok_title': 'Exportació de proveïdors completada',
        'mb_vend_full_ask': 'El mode «Càrrega completa» reemplaçarà els {count} proveïdors desats actualment '
                            'pels del fitxer Excel seleccionat.\n'
                            '\n'
                            'Vols continuar?',
        'mb_vend_imp_err_title': 'Error en importar proveïdors',
        'mb_vend_imp_ok_msg': 'Mode aplicat: {mode}\n'
                              '\n'
                              '- Nous afegits: {added}\n'
                              '- Existents actualitzats: {updated}\n'
                              '- Sense canvis / omesos: {skipped}\n'
                              '- Files incompletes ignorades: {invalid}\n'
                              '\n'
                              'Total de proveïdors desats: {total}',
        'mb_vend_imp_ok_title': 'Importació de proveïdors completada',
        'mb_vend_missing': "Indica tant el NIF/CIF com el codi Vendor Account d'Alma.",
        'mb_vend_sel_del': 'Selecciona una fila de la taula per eliminar-la.',
        'mb_vend_title': 'Proveïdors recordats',
        'mb_warn_title': 'Avís',
        'numdlg_apply': 'Aplicar en aquesta sessió',
        'numdlg_desc': 'És probable que el número pertanyi a una seqüència que es reinicia periòdicament.\n'
                       'Pots afegir temporalment un prefix o sufix per desambiguar-lo.\n'
                       "L'XML original no es modificarà.",
        'numdlg_orig': 'Número original: {original}',
        'numdlg_title': 'Corregir número de factura en memòria',
        'poldlg_all_btn': 'Veure tots',
        'poldlg_assign_full': "Assignar PO Line + codis i fons d'Analytics",
        'poldlg_assign_only': 'Assignar només PO Line',
        'poldlg_col_score': 'Afinitat',
        'poldlg_dblclick_hint': 'El doble clic assigna únicament la PO Line (sense alterar Reporting Codes '
                                'ni fons).',
        'poldlg_inv_title': 'Títol a la factura: {title}',
        'poldlg_line_frame': 'Línia de factura #{line}',
        'poldlg_search_btn': 'Cercar',
        'poldlg_search_lbl': 'Cercar a Analytics (Títol / ISBN / ISSN / PO Line):',
        'poldlg_summary': 'ISBN: {isbn}   |   ISSN: {issn}   |   Preu: {price}   |   PO Line actual: {pol}',
        'poldlg_title': "Cercar PO Line al catàleg d'Alma - Línia {line}",
        'poldlg_unassigned': '(sense assignar)',
        'st_alma_selected': "Estat: fitxer Excel d'Alma Analytics seleccionat.",
        'st_batch_count': 'Estat: {count} factures al lot.',
        'st_batch_empty': 'Estat: lot buit.',
        'st_batch_sorted': 'Estat: lot ordenat per número de factura.',
        'st_check_inputs': "Estat: revisa les dades d'entrada.",
        'st_cleared': 'Estat: dades netejades. La plantilla Alma es conserva.',
        'st_conv_err': 'Estat: error durant la conversió.',
        'st_conv_ok': 'Estat: conversió completada sense incidències.',
        'st_conv_warn': 'Estat: conversió completada amb incidències.',
        'st_converting': 'Estat: convertint en segon pla...',
        'st_copied': "Estat: copiat el valor '{val}' de {col}.",
        'st_draft_loaded': "Estat: esborrany '{name}' carregat correctament.",
        'st_draft_ready_alma': "Estat: esborrany editable de l'Excel d'Alma preparat (creuat amb Alma "
                               'Analytics).',
        'st_draft_ready_no_alma': "Estat: factura carregada a Editar (selecciona l'Excel d'Alma Analytics "
                                  'per creuar PO Lines).',
        'st_draft_saved': "Estat: esborrany desat a '{name}'.",
        'st_fill_blanks': "Estat: emplenades {count} cel·les en blanc inferiors amb '{val}'.",
        'st_fill_drag': "Estat: emplenades {count} cel·les a '{col}' amb '{val}'.",
        'st_inv_selected': 'Estat: factura {number} seleccionada. Proveïdor: {vendor}.',
        'st_lang_changed': 'Estat: idioma canviat a {lang}.',
        'st_mode_batch': 'Estat: mode lot. Selecciona factures del mateix proveïdor i biblioteca.',
        'st_mode_single': 'Estat: mode factura individual.',
        'st_no_redo': 'Estat: no hi ha accions per refer.',
        'st_no_undo': 'Estat: no hi ha accions per desfer.',
        'st_num_updated': "Estat: número de factura actualitzat a l'esborrany a '{number}'.",
        'st_open_in_edit': "Estat: oberta la factura '{name}' a Editar.",
        'st_out_selected': 'Estat: fitxer de sortida definit.',
        'st_pending': 'Estat: pendent de seleccionar fitxers.',
        'st_redone': "Estat: refeta l'acció '{action}'.",
        'st_saving_edits': 'Estat: desats {n_num} canvis de número, {n_hdr} de capçalera i {n_line} de '
                           "línia; actualitzant l'Excel de sortida...",
        'st_tpl_selected': 'Estat: plantilla Alma seleccionada.',
        'st_undone': "Estat: desfeta l'acció '{action}'.",
        'st_vendor_cleared': 'Estat: llista de proveïdors recordats buidada.',
        'st_vendor_deleted': "Estat: eliminat el proveïdor recordat per al NIF '{nif}'.",
        'st_vendor_exported': "Estat: exportats {count} proveïdors a '{name}'.",
        'st_vendor_imported': "Estat: proveïdors importats des de '{name}' (total: {total}).",
        'st_vendor_saved': 'Estat: proveïdor desat ({nif} → {vendor}).',
        'sum_body': '\n'
                    '\n'
                    'Resum de línies:\n'
                    '- Línies processades: {total_lines}\n'
                    '- Coincidències per ISBN: {match_isbn}\n'
                    '- Coincidències per ISSN: {match_issn}\n'
                    '- Coincidències per títol: {match_title}\n'
                    '{extra}- Duplicats ISBN: {duplicate_isbn}\n'
                    '- Duplicats ISSN: {duplicate_issn}\n'
                    '- Títols ambigus: {ambiguous_title}\n'
                    '- Sense coincidència: {no_match}\n'
                    '- Línies sense PO Line: {lines_without_po_line}',
        'sum_pol_man': '- Assignacions manuals de PO Line: {n}\n',
        'sum_pol_xml': "- Coincidències per PO Line a l'XML: {n}\n",
        'tab_convert': 'Convertir',
        'tab_customize': 'Personalitzar',
        'tab_edit': 'Editar',
        'tab_help': 'Ajuda',
        'tab_visualize': 'Visualitzar',
        'val_bad_draft': "El fitxer no té un format d'esborrany vàlid de FacturaeToAlma.",
        'val_bad_ext': 'Extensió de factura no vàlida: {path}',
        'val_draft_no_files': "Cap de les factures referenciades a l'esborrany existeix ja en la seva ruta "
                              'original.',
        'val_dup_numbers': 'Hi ha números de factura duplicats al lot: {nums}',
        'val_multi_vendor': 'El lot conté factures de diversos proveïdors. Selecciona únicament factures del '
                            'mateix proveïdor.\n'
                            '\n'
                            'Proveïdors detectats:\n'
                            '{details}',
        'val_need_alma': "Has de seleccionar el fitxer Excel d'Alma Analytics.",
        'val_need_inv': 'Has de seleccionar una factura o un lot de factures.',
        'val_need_inv_edit': 'Has de seleccionar primer una factura o un lot de factures a la pestanya '
                             'Convertir.',
        'val_need_out': "Has de definir l'Excel de sortida.",
        'val_need_tpl': 'Has de seleccionar la plantilla Alma.',
        'val_no_alma_file': "No existeix el fitxer Excel d'Alma Analytics: {path}",
        'val_no_inv_file': 'No existeix la factura: {path}',
        'val_no_tpl_file': 'No existeix la plantilla Alma: {path}',
        'val_out_collision': "L'Excel de sortida no pot coincidir amb la factura, l'Analytics ni la "
                             'plantilla. Tria un altre nom o carpeta.',
        'val_single_inv_only': 'El fitxer ha de contenir una sola factura i en conté {count}.',
        'val_vend_empty_xlsx': 'El fitxer Excel està buit.',
        'val_vend_po_cols': "El fitxer conté columnes de línia de comanda ({cols}). L'Excel de proveïdors ha "
                            'de tenir únicament les 2 columnes «NIF / CIF» i «Vendor Account».',
        'val_vend_two_cols': "L'Excel de proveïdors ha de contenir exactament les 2 columnes «NIF / CIF» i "
                             '«Vendor Account» (sense camps de línia de comanda).',
        'vis_footer_note': "Vista de consulta de la factura original. Els canvis de l'Excel d'Alma es fan a "
                           'la pestanya Editar.',
        'vis_inv_lbl': 'Factura:',
        'vw_addr': 'Adreça: {addr}',
        'vw_card_buyer': ' DADES DEL RECEPTOR ',
        'vw_card_dir3': ' CENTRES ADMINISTRATIUS FACe (DIR3) ',
        'vw_card_lines': ' DETALL ECONÒMIC (LÍNIES) ',
        'vw_card_pay': ' DADES DE PAGAMENT ',
        'vw_card_seller': " DADES DE L'EMISSOR ",
        'vw_col_code': 'Article / ISBN',
        'vw_col_desc': 'Descripció / Títol',
        'vw_col_gross': 'Import Brut',
        'vw_col_price': 'Preu Unit.',
        'vw_col_qty': 'Quant.',
        'vw_col_tax': 'IVA Repercutit',
        'vw_contact': 'Contacte: {contact}',
        'vw_elec_inv_lbl': 'Factura electrònica:',
        'vw_hdr_invoice': 'FACTURA: {number}',
        'vw_hdr_meta': "Sèrie: {series}   |   Data d'emissió: {issue}   |   Data d'operació: {op}",
        'vw_load_btn': 'Carregar factura...',
        'vw_nif': 'CIF/NIF: {nif}',
        'vw_no_number': '(sense número)',
        'vw_not_specified': 'No especificat',
        'vw_pay_due': 'Data de venciment: {date}',
        'vw_pay_iban': 'Compte / IBAN: {iban}',
        'vw_pay_means': 'Mitjà de pagament: Transferència (Codi {code})',
        'vw_placeholder': 'Seleccioneu un fitxer de factura electrònica (.xsig, .xml, .txt) per '
                          'visualitzar-ne les dades.',
        'vw_role_01': 'Oficina Comptable',
        'vw_role_02': 'Òrgan Gestor',
        'vw_role_03': 'Unitat Tramitadora',
        'vw_role_04': 'Òrgan Proponent',
        'vw_role_other': 'Rol {code}',
        'vw_tab_pdf': '  Vista Factura (Estil PDF)  ',
        'vw_tab_xml': '  Codi XML (FACe)  ',
        'vw_tot_base': 'Base Imposable:',
        'vw_tot_grand': 'TOTAL FACTURA:',
        'vw_tot_payable': 'Total a Pagar:',
        'vw_tot_ret': 'Impostos Retinguts:',
        'vw_tot_taxes': 'Impostos Repercutits:',
        'vw_unknown_vendor': 'Proveïdor no identificat'},
 'es': {'afdlg_apply_next': 'Aplicar y seguir filtrando otra columna',
        'afdlg_blank_chk': '(Vacías / Sin valor)',
        'afdlg_check_all': 'Marcar todos',
        'afdlg_col_lbl': 'Columna a filtrar:',
        'afdlg_list_frame': 'Valores presentes en la columna (marca los que deseas mostrar)',
        'afdlg_only_blank': 'Solo vacías',
        'afdlg_only_val': 'Solo con valor',
        'afdlg_remove_col': 'Quitar filtro de esta columna',
        'afdlg_search_lbl': 'Buscar texto / valor:',
        'afdlg_title': 'Filtro autofiltro por columnas (estilo Excel)',
        'afdlg_uncheck_all': 'Desmarcar todos',
        'bulkdlg_lbl': 'Valor para {col}:',
        'bulkdlg_title': 'Asignar {col}',
        'conv_actions_frame': '4. Acciones',
        'conv_alma_btn': 'Seleccionar fichero...',
        'conv_alma_lbl': 'Fichero Excel de Alma Analytics con PO Lines *',
        'conv_batch_add': 'Añadir facturas',
        'conv_batch_clear': 'Vaciar lote',
        'conv_batch_count': '{count} facturas seleccionadas',
        'conv_batch_frame': 'Facturas del lote',
        'conv_batch_hint': '(Doble clic en una factura para abrirla en Editar)',
        'conv_batch_remove': 'Quitar seleccionadas',
        'conv_batch_sort': 'Ordenar por número',
        'conv_clear_btn': 'Limpiar datos',
        'conv_col_file': 'Fichero',
        'conv_col_invoice': 'Número de factura',
        'conv_col_vendor': 'Proveedor',
        'conv_exit_btn': 'Salir',
        'conv_inputs_frame': '2. Archivos de entrada',
        'conv_inv_batch_btn': 'Seleccionar facturas...',
        'conv_inv_batch_lbl': 'Facturas XSIG / XML / TXT del mismo proveedor y biblioteca *',
        'conv_inv_single_btn': 'Seleccionar factura...',
        'conv_inv_single_lbl': 'Factura electrónica XSIG / XML / TXT *',
        'conv_mode_batch': 'Lote de facturas del mismo proveedor y biblioteca',
        'conv_mode_frame': '1. Modo de conversión',
        'conv_mode_single': 'Factura individual',
        'conv_ok_heading': 'CONVERSIÓN COMPLETA SIN INCIDENCIAS',
        'conv_open_out_btn': 'Abrir carpeta de salida',
        'conv_output_btn': 'Elegir destino...',
        'conv_output_frame': '3. Archivo de salida',
        'conv_result_body': '{heading}\n'
                            '\n'
                            'Facturas procesadas: {total_inv}\n'
                            'Avisos de validación: {total_iss}\n'
                            '\n'
                            'Excel Alma generado:\n'
                            '{out}',
        'conv_run_btn': 'Convertir a formato Alma',
        'conv_tpl_btn': 'Seleccionar plantilla...',
        'conv_tpl_lbl': 'Plantilla Alma .xlsx *',
        'conv_unmatched_warn': '\n'
                               '\n'
                               'Atención: hay líneas sin PO Line. Revisa el informe antes de cargar en Alma.',
        'conv_val_report': '\n\nInforme de validación:\n{path}',
        'conv_warn_heading': 'INCIDENCIAS EN LA CONVERSIÓN',
        'ctx_bulk_apply': 'Aplicar valor a líneas seleccionadas...',
        'ctx_clear_cell': 'Vaciar celda (Supr)',
        'ctx_edit_cell': 'Editar celda',
        'ctx_fill_blanks_down': 'Rellenar blancos inferiores (■)',
        'ctx_fill_blanks_sel': 'Rellenar blancos seleccionados',
        'ctx_fill_down': 'Rellenar hacia abajo (Ctrl+D)',
        'ctx_flt_cell_val': 'Filtrar por el valor de esta celda',
        'ctx_flt_clear_all': 'Quitar todos los filtros',
        'ctx_flt_col_excel': 'Filtrar esta columna (estilo Excel)...',
        'ctx_redo': 'Rehacer (Ctrl+Y)',
        'ctx_restore_cell': 'Restablecer celda',
        'ctx_restore_orig': 'Restablecer celda original',
        'ctx_search_pol': 'Buscar PO Line en catálogo Alma...',
        'ctx_suggest_ctx': 'Sugerir valor por contexto',
        'ctx_undo': 'Deshacer (Ctrl+Z)',
        'cust_choose_dir_btn': 'Seleccionar carpeta...',
        'cust_dir_alma': 'Directorio del fichero Excel de Alma',
        'cust_dir_inv': 'Directorio de facturas FACe',
        'cust_dir_out': 'Directorio de Excel finales',
        'cust_dir_tpl': 'Directorio de plantilla Alma',
        'cust_inst_lbl': 'Institución',
        'cust_lang_lbl': 'Idioma de la interfaz',
        'cust_logo_del_btn': 'Eliminar logo',
        'cust_logo_hint': 'PNG, JPG o JPEG. Altura máxima: 100 px.',
        'cust_logo_invalid': 'El archivo seleccionado no es una imagen válida.',
        'cust_logo_lbl': 'Logo institucional',
        'cust_logo_loaded': 'Logo cargado: {name}',
        'cust_logo_none': 'Sin logo. PNG, JPG o JPEG; altura máxima: 100 px.',
        'cust_logo_sel_btn': 'Seleccionar logo...',
        'cust_logo_unavail': 'El logo configurado no está disponible.',
        'cust_pref_frame': 'Preferencias',
        'cust_pref_note': 'Todos los campos son opcionales. Si una carpeta personalizada queda vacía, se usa '
                          'la última carpeta válida; en la primera ejecución se usa la carpeta de la '
                          'aplicación. La salida se guarda junto a la factura, salvo que se defina un '
                          'directorio de salida.',
        'cust_reset_pref_btn': 'Restablecer valores',
        'cust_save_pref_btn': 'Guardar preferencias',
        'cust_vend_acc_lbl': 'Vendor Account en Alma:',
        'cust_vend_add_btn': 'Añadir / Actualizar',
        'cust_vend_clr_btn': 'Vaciar lista',
        'cust_vend_col_acc': 'Vendor Account en Alma',
        'cust_vend_col_nif': 'NIF / CIF normalizado',
        'cust_vend_del_btn': 'Eliminar seleccionado',
        'cust_vend_exp_btn': 'Exportar Excel...',
        'cust_vend_frame': 'Proveedores recordados automáticamente (NIF / CIF → Vendor Account en Alma)',
        'cust_vend_imp_btn': 'Importar Excel...',
        'cust_vend_imp_mode_lbl': 'Modo de importación Excel:',
        'cust_vend_mode_full': 'Carga completa (reemplazar lista actual)',
        'cust_vend_mode_inc': 'Incremental / acumulativo (añadir y actualizar)',
        'cust_vend_mode_new': 'Solo añadir nuevos (no sobrescribir)',
        'cust_vend_nif_lbl': 'NIF / CIF:',
        'dlg_accept': 'Aceptar',
        'dlg_apply': 'Aplicar',
        'dlg_cancel': 'Cancelar',
        'ed_apply_num_btn': 'Aplicar número',
        'ed_bottom_hint': 'Escribe directo o doble clic para editar | Arrastra ■ para rellenar | Ctrl+F '
                          'buscar | Clic en cabeceras ▼ para filtrar.',
        'ed_col_invoice': 'Factura',
        'ed_col_line': 'Línea',
        'ed_diag_edited': '✎ Línea {line}: modificada manualmente en esta sesión.',
        'ed_diag_init': 'Selecciona una línea para ver su diagnóstico de cruce y validación.',
        'ed_diag_manual': "✓ Línea {line}: PO Line '{pol}' asignada manualmente.",
        'ed_diag_matched': '✓ Línea {line}: {msg}.',
        'ed_diag_no_inv': 'Selecciona una factura para revisar sus líneas y avisos.',
        'ed_diag_ready': '✓ Línea {line}: lista para exportar.',
        'ed_diag_recon': '⚠ Cuadre factura: {msg}',
        'ed_diag_unmatched': "⚠ Línea {line}: sin PO Line asignada (usa 'Buscar PO Line en Alma...' o "
                             'escríbela en la celda).',
        'ed_diag_warning': '⚠ Línea {line}: {notes}',
        'ed_flt_active_prefix': 'Filtros activos -> ',
        'ed_flt_add_btn': '+ Añadir filtro',
        'ed_flt_clear_btn': 'Limpiar filtros',
        'ed_flt_col_lbl': 'Columna:',
        'ed_flt_contains_q': " contiene '{q}'",
        'ed_flt_excel_btn': 'Filtro Excel...',
        'ed_flt_frame': 'Filtrar líneas de factura (por estado y por una o varias columnas estilo Excel)',
        'ed_flt_in_list': '{col} en [{items}]',
        'ed_flt_none': 'Sin filtros activos (clic en cualquier cabecera ▼ para filtrar por una o más '
                       'columnas estilo Excel).',
        'ed_flt_op_blank': 'Vacías',
        'ed_flt_op_contains': 'Contiene / Igual',
        'ed_flt_op_exact': 'Igual exacto',
        'ed_flt_op_nblank': 'No vacías',
        'ed_flt_op_neq': 'Distinto de',
        'ed_flt_st_all': 'Todas las líneas',
        'ed_flt_st_edit': 'Solo líneas modificadas manualmente',
        'ed_flt_st_ok': 'Solo líneas correctas (con PO Line y sin avisos)',
        'ed_flt_st_prob': 'Solo líneas con problemas (sin PO Line o con avisos)',
        'ed_flt_st_unm': 'Solo líneas sin PO Line (pendientes)',
        'ed_flt_state_lbl': 'Estado:',
        'ed_flt_tag': 'Filtro',
        'ed_flt_val_blank': '(Vacías)',
        'ed_flt_val_lbl': 'Valor:',
        'ed_flt_val_nblank': '(No vacías)',
        'ed_flt_values_cnt': '{col} ({count} valores)',
        'ed_hdr_frame': 'Cabecera de factura (INV)',
        'ed_load_draft_btn': 'Cargar borrador...',
        'ed_mode_orig': 'Original',
        'ed_mode_prefix': 'Prefijo',
        'ed_mode_suffix': 'Sufijo',
        'ed_num_frame': 'Número de factura en el Excel de Alma',
        'ed_orig_lbl': 'Original:',
        'ed_refresh_btn': 'Actualizar borrador',
        'ed_reset_inv_btn': 'Restablecer factura',
        'ed_save_draft_btn': 'Guardar borrador...',
        'ed_save_excel_btn': 'Guardar cambios y actualizar Excel',
        'ed_search_pol_btn': 'Buscar PO Line en Alma...',
        'ed_stats_empty': 'Mostrando: 0/0 | Correctas: 0 | Con problemas: 0',
        'ed_stats_fmt': 'Mostrando: {shown}/{total} | Correctas: {ok} | Con problemas: {prob} (Sin POL: '
                        '{unm})',
        'ed_tb_bulk': 'Asignar valor a seleccionadas...',
        'ed_tb_clear': 'Vaciar celda (Supr)',
        'ed_tb_copy': 'Copiar (Ctrl+C)',
        'ed_tb_fill_blanks': '⬇ Rellenar blancos (■)',
        'ed_tb_fill_down': '⬇ Rellenar abajo (Ctrl+D)',
        'ed_tb_paste': 'Pegar (Ctrl+V)',
        'ed_tb_redo': '↪ Rehacer (Ctrl+Y)',
        'ed_tb_restore': 'Restablecer celda',
        'ed_tb_search_lbl': 'Buscar (Ctrl+F):',
        'ed_tb_undo': '↩ Deshacer (Ctrl+Z)',
        'ed_tot_init': 'Importe visible: 0,00 €',
        'ed_tot_sel': '   |   Selección ({count}): {base}',
        'ed_tot_vis': 'Visible ({count}): Base {base} | IVA {vat}',
        'err_conv_title': 'Error durante la conversión',
        'err_fnf_title': 'Fichero no encontrado',
        'err_log_hint': "{err}\n\nConsulta el archivo de log '{log}' para más detalles.",
        'err_perm_title': 'Fichero bloqueado o sin permisos',
        'err_val_title': 'Datos o formato no válidos',
        'err_xml_sec_title': 'XML bloqueado por seguridad',
        'fd_exp_vendors': 'Exportar proveedores (NIF / CIF y Vendor Account) a Excel',
        'fd_ft_all': 'Todos los archivos',
        'fd_ft_compat_inv': 'Facturas compatibles',
        'fd_ft_draft': 'Borrador FacturaeToAlma',
        'fd_ft_excel': 'Archivos Excel',
        'fd_ft_img_compat': 'Imágenes compatibles',
        'fd_ft_jpg': 'Imagen JPEG',
        'fd_ft_png': 'Imagen PNG',
        'fd_ft_txt': 'Archivos TXT',
        'fd_ft_xml': 'Archivos XML',
        'fd_ft_xsig': 'Archivos XSIG',
        'fd_imp_vendors': 'Importar proveedores (NIF / CIF y Vendor Account) desde Excel',
        'fd_load_draft': 'Cargar borrador de sesión FacturaeToAlma',
        'fd_save_draft': 'Guardar borrador de sesión FacturaeToAlma',
        'fd_sel_alma': 'Seleccionar fichero Excel de Alma Analytics',
        'fd_sel_folder': 'Seleccionar carpeta',
        'fd_sel_inv_batch': 'Seleccionar facturas del mismo proveedor y biblioteca',
        'fd_sel_inv_face': 'Seleccionar factura electrónica FACe',
        'fd_sel_inv_single': 'Seleccionar factura XSIG / XML / TXT',
        'fd_sel_logo': 'Seleccionar logo institucional',
        'fd_sel_out': 'Definir archivo Excel de salida',
        'fd_sel_tpl': 'Seleccionar plantilla Alma .xlsx',
        'gctx_copy': 'Copiar',
        'gctx_copy_txt': 'Copiar texto',
        'gctx_cut': 'Cortar',
        'gctx_go_edit': 'Ir a Editar',
        'gctx_go_vis': 'Ir a Visualizar',
        'gctx_paste': 'Pegar',
        'gctx_sel_all': 'Seleccionar todo',
        'hlp_body_5_0': 'Factura individual: selecciona una factura, el informe de Alma Analytics, la '
                        'plantilla y el destino.\n'
                        '\n'
                        'Modo lote: todas las facturas deben pertenecer al mismo proveedor y biblioteca. La '
                        'lista permite añadir, quitar, ordenar y abrir cada factura con doble clic en '
                        'Editar.\n'
                        '\n'
                        'Edición y filtrado: revisa cabecera (INV) y líneas (IL), busca PO Lines en el '
                        'catálogo de Alma Analytics, aplica autofiltros estilo Excel o guarda/carga '
                        'borradores (.ftalma.json).\n'
                        '\n'
                        'Proveedores e idiomas: en Personalizar puedes cambiar el idioma oficial/cooficial '
                        '(es, ca, va, gl, eu, oc) y gestionar, importar o exportar en Excel la tabla de '
                        'proveedores (NIF / CIF y Vendor Account) en modo incremental, solo nuevos o carga '
                        'completa.\n'
                        '\n'
                        'Informe de validación: se genera cuando faltan PO Lines, hay coincidencias '
                        'ambiguas, duplicados, descuadres contables o diferencias entre Quantity y Quantity '
                        'for Pricing.\n'
                        '\n'
                        'Limitaciones confirmadas por Ex Libris: la plantilla Excel no permite indicar Line '
                        'Exclusive ni declarar explícitamente facturación parcial o completa.\n'
                        '\n'
                        'Licencia: GPL-3.0-only.',
        'hlp_body_5_1': 'Factura individual: selecciona una factura, el informe de Alma Analytics, la '
                        'plantilla y el destino.\n'
                        '\n'
                        'Modo lote: todas las facturas deben pertenecer al mismo proveedor y biblioteca. La '
                        'lista permite añadir, quitar, ordenar y abrir cada factura con doble clic en '
                        'Editar.\n'
                        '\n'
                        'Importes e IVA compatibles con Línea exclusiva (lógica exacta 4.0): la extracción y '
                        'exportación de importes e IVA reproduce exactamente el comportamiento de la versión '
                        '4.0 para que el cambio manual a «Línea exclusiva» en Alma cuadre siempre.\n'
                        '\n'
                        'Proveedores e idiomas: en Personalizar puedes seleccionar cualquiera de las lenguas '
                        'oficiales y cooficiales de España (es, ca, va, gl, eu, oc) e importar o exportar en '
                        'Excel la tabla de proveedores (solo columnas NIF / CIF y Vendor Account) en modo '
                        'incremental/acumulativo, solo añadir nuevos o carga completa.\n'
                        '\n'
                        'Informe de validación: se genera cuando faltan PO Lines, hay coincidencias '
                        'ambiguas, duplicados o diferencias entre Quantity y Quantity for Pricing.\n'
                        '\n'
                        'Limitaciones confirmadas por Ex Libris: la plantilla Excel no permite indicar Line '
                        'Exclusive ni declarar explícitamente facturación parcial o completa.\n'
                        '\n'
                        'Licencia: GPL-3.0-only.',
        'hlp_guide_btn': 'Abrir Biblioguía / LibGuides',
        'hlp_links_frame': 'Documentación y soporte',
        'hlp_repo_btn': 'Abrir repositorio de GitHub',
        'match_msg_ambig': 'Varias coincidencias posibles por título',
        'match_msg_dup': 'Duplicado por {mtype}. Revisión manual necesaria.',
        'match_msg_manual': 'PO Line asignada manualmente',
        'match_msg_none': 'Sin coincidencia en Alma por ISBN, ISSN ni título',
        'match_msg_special': 'Línea de tipo {ltype} (no requiere PO Line bibliográfica)',
        'match_msg_unique': 'Coincidencia única por {mtype}',
        'mb_batch_clr_ask': '¿Quieres eliminar todas las facturas del lote?',
        'mb_batch_clr_title': 'Vaciar lote',
        'mb_batch_dup_num': 'Número de factura duplicado {num}: {name}',
        'mb_batch_none_added': 'No se incorporaron facturas nuevas.',
        'mb_batch_read_err': 'No se pudo leer {name}: {err}',
        'mb_batch_rep_path': 'Ruta repetida: {name}',
        'mb_batch_sel_one': 'Selecciona al menos una factura de la lista.',
        'mb_batch_title': 'Lote',
        'mb_batch_warn_title': 'Facturas no incorporadas',
        'mb_clr_alma_ask': '¿Quieres limpiar también el fichero Excel de Alma Analytics?\n'
                           '\n'
                           'La plantilla Alma se conservará.',
        'mb_clr_alma_title': 'Limpiar fichero Excel de Alma',
        'mb_conv_close_ask': 'Hay una conversión en curso. ¿Quieres cerrar la aplicación de todos modos?',
        'mb_conv_ok_title': 'Conversión completada',
        'mb_conv_running_msg': 'Ya hay una conversión en curso.',
        'mb_conv_running_title': 'Conversión en curso',
        'mb_conv_warn_title': 'Incidencias en la conversión',
        'mb_draft_loaded_msg': 'Se han restaurado las facturas y cambios del borrador.{warn}',
        'mb_draft_loaded_title': 'Borrador cargado',
        'mb_draft_missing_warn': '\n'
                                 '\n'
                                 'Aviso: no se encontraron algunas facturas movidas o eliminadas:\n'
                                 '- {files}',
        'mb_draft_outdated_msg': 'La factura, Analytics o plantilla han cambiado después de guardar. Vuelve '
                                 'a Editar, revisa y guarda de nuevo.',
        'mb_draft_outdated_title': 'Borrador desactualizado',
        'mb_draft_rebuilt_alma': 'Se ha reconstruido el borrador del Excel final de Alma con la factura y el '
                                 'informe de Analytics actualmente seleccionados.',
        'mb_draft_rebuilt_no_alma': 'Se han cargado las líneas de la factura en el editor.\n'
                                    '\n'
                                    'Si seleccionas también el fichero Excel de Alma Analytics, se cruzarán '
                                    'automáticamente las PO Lines, Reporting Codes y fondos.',
        'mb_draft_saved_msg': 'Se ha guardado el progreso actual en:\n'
                              '{path}\n'
                              '\n'
                              "Podrás reanudar la edición más tarde con 'Cargar borrador...'.",
        'mb_draft_saved_title': 'Borrador guardado',
        'mb_draft_unavail_msg': 'No se pudo preparar el borrador editable antes de convertir.',
        'mb_draft_unavail_title': 'Borrador no disponible',
        'mb_empty_src_msg': 'La celda activa está vacía. Escribe o selecciona una celda con valor antes de '
                            'rellenar los blancos.',
        'mb_empty_src_title': 'Celda origen vacía',
        'mb_err_folder_msg': 'No se pudo abrir la carpeta:\n{path}\n\n{err}',
        'mb_err_folder_title': 'Error al abrir carpeta',
        'mb_err_load_draft': 'Error al cargar borrador',
        'mb_err_load_draft_msg': 'No se pudo cargar el borrador:\n{err}',
        'mb_err_prep_draft': 'No se pudo preparar el borrador',
        'mb_err_read_inv_msg': 'No se pudo cargar la factura seleccionada:\n{err}',
        'mb_err_read_inv_title': 'Error al leer factura',
        'mb_err_reset_title': 'Error al restablecer',
        'mb_err_save_draft': 'Error al guardar borrador',
        'mb_err_save_draft_msg': 'No se pudo guardar el borrador:\n{err}',
        'mb_err_save_title': 'Error al guardar',
        'mb_err_url_msg': 'No se pudo abrir el enlace:\n{url}\n\n{err}',
        'mb_err_url_title': 'Error al abrir enlace',
        'mb_flt_cols_no_inv': 'Selecciona primero una factura en la pestaña Editar.',
        'mb_flt_cols_title': 'Filtrar columnas',
        'mb_inv_invalid_title': 'Factura no válida',
        'mb_invalid_date_msg': 'Usa una fecha válida, por ejemplo 2026-10-19.',
        'mb_invalid_date_title': 'Fecha no válida',
        'mb_invalid_inputs_title': 'Datos de entrada no válidos',
        'mb_invalid_ltype': 'Selecciona un Line type permitido por Alma.',
        'mb_invalid_num': '{col} debe ser numérico.',
        'mb_invalid_val_title': 'Valor no válido',
        'mb_logo_invalid_msg': 'No se pudo cargar el logo:\n{path}\n\n{err}',
        'mb_logo_invalid_title': 'Logo no válido',
        'mb_logo_unavail_msg': 'No se encontró el archivo de logo:\n{path}',
        'mb_logo_unavail_title': 'Logo no disponible',
        'mb_no_active_inv': 'No hay ninguna factura activa.',
        'mb_no_edit_inv': 'No hay ninguna factura activa en el editor.',
        'mb_no_out_dir_def': 'No hay una carpeta de salida definida.',
        'mb_no_out_dir_found': 'No se encontró la carpeta de salida.',
        'mb_no_sugg_msg': 'No hay valores de contexto para sugerir en esta columna.',
        'mb_no_sugg_title': 'Sin sugerencias',
        'mb_pol_ask_alma': 'Este buscador consulta las líneas de pedido (PO Lines) dentro del fichero Excel '
                           'exportado de Alma Analytics.\n'
                           '\n'
                           'Todavía no has seleccionado ningún fichero Excel de Alma Analytics.\n'
                           '¿Quieres seleccionarlo ahora?',
        'mb_pol_no_inv': 'Selecciona primero una factura (o lote de facturas) en la pestaña Convertir.',
        'mb_pol_no_lines': 'No hay líneas visibles en el editor.',
        'mb_pol_sel_row': 'Selecciona una fila de la lista.',
        'mb_pol_title': 'Buscar PO Line en Alma',
        'mb_pref_invalid_dir': '{label} no es una carpeta válida:\n{val}',
        'mb_pref_invalid_title': 'Preferencias no válidas',
        'mb_pref_reset_ask': '¿Quieres restablecer la institución, el idioma y todos los directorios '
                             'personalizados?',
        'mb_pref_reset_ok': 'Se han restablecido los valores iniciales.',
        'mb_pref_reset_title': 'Restablecer preferencias',
        'mb_pref_saved_msg': 'Las preferencias se han guardado correctamente.',
        'mb_prot_drag_msg': 'Esta columna está protegida y no puede rellenarse arrastrando.',
        'mb_prot_field_title': 'Campo protegido',
        'mb_prot_generic_msg': 'Este campo no es editable.',
        'mb_prot_hdr_msg': 'El número, la fecha y los importes de cabecera no se editan directamente. El '
                           'número solo admite prefijo o sufijo.',
        'mb_prot_line_msg': 'Este campo procede de la factura original y no es editable en FacturaeToAlma.',
        'mb_reset_inv_ask': '¿Quieres descartar todos los cambios de esta factura?',
        'mb_reset_inv_title': 'Restablecer factura',
        'mb_save_changes_ask': '¿Quieres guardar todos los cambios y actualizar ahora el Excel de salida?\n'
                               '\n'
                               'Se sobrescribirá el Excel indicado en Convertir. El XML original no se '
                               'modificará.',
        'mb_save_changes_title': 'Guardar cambios',
        'mb_save_draft_empty': 'No hay ninguna factura cargada para guardar en un borrador.',
        'mb_save_draft_title': 'Guardar borrador',
        'mb_unsaved_ed_ask': 'Hay cambios realizados en Editar (número de factura, cabecera o líneas) que '
                             'aún no se habían guardado.\n'
                             '\n'
                             '¿Quieres aplicar y guardar estos cambios para convertir ahora?',
        'mb_unsaved_ed_title': 'Cambios pendientes en Editar',
        'mb_vend_clr_ask': '¿Quieres eliminar todas las asociaciones NIF → Vendor Account guardadas?',
        'mb_vend_clr_title': 'Vaciar proveedores recordados',
        'mb_vend_exp_err_title': 'Error al exportar proveedores',
        'mb_vend_exp_ok_msg': 'Se han exportado {count} proveedores (columnas NIF / CIF y Vendor Account) al '
                              'archivo:\n'
                              '{path}',
        'mb_vend_exp_ok_title': 'Exportación de proveedores completada',
        'mb_vend_full_ask': 'El modo «Carga completa» reemplazará los {count} proveedores guardados '
                            'actualmente por los del archivo Excel seleccionado.\n'
                            '\n'
                            '¿Deseas continuar?',
        'mb_vend_imp_err_title': 'Error al importar proveedores',
        'mb_vend_imp_ok_msg': 'Modo aplicado: {mode}\n'
                              '\n'
                              '- Nuevos añadidos: {added}\n'
                              '- Existentes actualizados: {updated}\n'
                              '- Sin cambios / omitidos: {skipped}\n'
                              '- Filas incompletas ignoradas: {invalid}\n'
                              '\n'
                              'Total de proveedores guardados: {total}',
        'mb_vend_imp_ok_title': 'Importación de proveedores completada',
        'mb_vend_missing': 'Indica tanto el NIF/CIF como el código Vendor Account de Alma.',
        'mb_vend_sel_del': 'Selecciona una fila de la tabla para eliminarla.',
        'mb_vend_title': 'Proveedores recordados',
        'mb_warn_title': 'Aviso',
        'numdlg_apply': 'Aplicar en esta sesión',
        'numdlg_desc': 'Es probable que el número pertenezca a una secuencia que se reinicia '
                       'periódicamente.\n'
                       'Puedes añadir temporalmente un prefijo o sufijo para desambiguar.\n'
                       'El XML original no se modificará.',
        'numdlg_orig': 'Número original: {original}',
        'numdlg_title': 'Corregir número de factura en memoria',
        'poldlg_all_btn': 'Ver todos',
        'poldlg_assign_full': 'Asignar PO Line + códigos y fondos de Analytics',
        'poldlg_assign_only': 'Asignar solo PO Line',
        'poldlg_col_score': 'Afinidad',
        'poldlg_dblclick_hint': 'Doble clic asigna únicamente la PO Line (sin alterar Reporting Codes ni '
                                'fondos).',
        'poldlg_inv_title': 'Título en factura: {title}',
        'poldlg_line_frame': 'Línea de factura #{line}',
        'poldlg_search_btn': 'Buscar',
        'poldlg_search_lbl': 'Buscar en Analytics (Título / ISBN / ISSN / PO Line):',
        'poldlg_summary': 'ISBN: {isbn}   |   ISSN: {issn}   |   Precio: {price}   |   PO Line actual: {pol}',
        'poldlg_title': 'Buscar PO Line en catálogo Alma - Línea {line}',
        'poldlg_unassigned': '(sin asignar)',
        'st_alma_selected': 'Estado: fichero Excel de Alma Analytics seleccionado.',
        'st_batch_count': 'Estado: {count} facturas en el lote.',
        'st_batch_empty': 'Estado: lote vacío.',
        'st_batch_sorted': 'Estado: lote ordenado por número de factura.',
        'st_check_inputs': 'Estado: revisa los datos de entrada.',
        'st_cleared': 'Estado: datos limpiados. La plantilla Alma se conserva.',
        'st_conv_err': 'Estado: error durante la conversión.',
        'st_conv_ok': 'Estado: conversión completada sin incidencias.',
        'st_conv_warn': 'Estado: conversión completada con incidencias.',
        'st_converting': 'Estado: convirtiendo en segundo plano...',
        'st_copied': "Estado: copiado valor '{val}' de {col}.",
        'st_draft_loaded': "Estado: borrador '{name}' cargado correctamente.",
        'st_draft_ready_alma': 'Estado: borrador editable del Excel de Alma preparado (cruzado con Alma '
                               'Analytics).',
        'st_draft_ready_no_alma': 'Estado: factura cargada en Editar (selecciona el Excel de Alma Analytics '
                                  'para cruzar PO Lines).',
        'st_draft_saved': "Estado: borrador guardado en '{name}'.",
        'st_fill_blanks': "Estado: rellenadas {count} celdas en blanco inferiores con '{val}'.",
        'st_fill_drag': "Estado: rellenadas {count} celdas en '{col}' con '{val}'.",
        'st_inv_selected': 'Estado: factura {number} seleccionada. Proveedor: {vendor}.',
        'st_lang_changed': 'Estado: idioma cambiado a {lang}.',
        'st_mode_batch': 'Estado: modo lote. Selecciona facturas del mismo proveedor y biblioteca.',
        'st_mode_single': 'Estado: modo factura individual.',
        'st_no_redo': 'Estado: no hay acciones para rehacer.',
        'st_no_undo': 'Estado: no hay acciones para deshacer.',
        'st_num_updated': "Estado: número de factura actualizado en el borrador a '{number}'.",
        'st_open_in_edit': "Estado: abierta factura '{name}' en Editar.",
        'st_out_selected': 'Estado: archivo de salida definido.',
        'st_pending': 'Estado: pendiente de seleccionar archivos.',
        'st_redone': "Estado: rehecha acción '{action}'.",
        'st_saving_edits': 'Estado: guardados {n_num} cambios de número, {n_hdr} de cabecera y {n_line} de '
                           'línea; actualizando el Excel de salida...',
        'st_tpl_selected': 'Estado: plantilla Alma seleccionada.',
        'st_undone': "Estado: deshecha acción '{action}'.",
        'st_vendor_cleared': 'Estado: lista de proveedores recordados vaciada.',
        'st_vendor_deleted': "Estado: eliminado proveedor recordado para NIF '{nif}'.",
        'st_vendor_exported': "Estado: exportados {count} proveedores a '{name}'.",
        'st_vendor_imported': "Estado: proveedores importados desde '{name}' (total: {total}).",
        'st_vendor_saved': 'Estado: proveedor guardado ({nif} → {vendor}).',
        'sum_body': '\n'
                    '\n'
                    'Resumen de líneas:\n'
                    '- Líneas procesadas: {total_lines}\n'
                    '- Coincidencias por ISBN: {match_isbn}\n'
                    '- Coincidencias por ISSN: {match_issn}\n'
                    '- Coincidencias por título: {match_title}\n'
                    '{extra}- Duplicados ISBN: {duplicate_isbn}\n'
                    '- Duplicados ISSN: {duplicate_issn}\n'
                    '- Títulos ambiguos: {ambiguous_title}\n'
                    '- Sin coincidencia: {no_match}\n'
                    '- Líneas sin PO Line: {lines_without_po_line}',
        'sum_pol_man': '- Asignaciones manuales de PO Line: {n}\n',
        'sum_pol_xml': '- Coincidencias por PO Line en XML: {n}\n',
        'tab_convert': 'Convertir',
        'tab_customize': 'Personalizar',
        'tab_edit': 'Editar',
        'tab_help': 'Ayuda',
        'tab_visualize': 'Visualizar',
        'val_bad_draft': 'El fichero no tiene un formato de borrador válido de FacturaeToAlma.',
        'val_bad_ext': 'Extensión de factura no válida: {path}',
        'val_draft_no_files': 'Ninguna de las facturas referenciadas en el borrador existe ya en su ruta '
                              'original.',
        'val_dup_numbers': 'Hay números de factura duplicados en el lote: {nums}',
        'val_multi_vendor': 'El lote contiene facturas de varios proveedores. Selecciona únicamente facturas '
                            'del mismo proveedor.\n'
                            '\n'
                            'Proveedores detectados:\n'
                            '{details}',
        'val_need_alma': 'Debes seleccionar el fichero Excel de Alma Analytics.',
        'val_need_inv': 'Debes seleccionar una factura o un lote de facturas.',
        'val_need_inv_edit': 'Debes seleccionar primero una factura o un lote de facturas en la pestaña '
                             'Convertir.',
        'val_need_out': 'Debes definir el Excel de salida.',
        'val_need_tpl': 'Debes seleccionar la plantilla Alma.',
        'val_no_alma_file': 'No existe el fichero Excel de Alma Analytics: {path}',
        'val_no_inv_file': 'No existe la factura: {path}',
        'val_no_tpl_file': 'No existe la plantilla Alma: {path}',
        'val_out_collision': 'El Excel de salida no puede coincidir con la factura, el Analytics ni la '
                             'plantilla. Elige otro nombre o carpeta.',
        'val_single_inv_only': 'El fichero debe contener una sola factura y contiene {count}.',
        'val_vend_empty_xlsx': 'El archivo Excel está vacío.',
        'val_vend_po_cols': 'El archivo contiene columnas de línea de pedido ({cols}). El Excel de '
                            'proveedores debe tener únicamente las 2 columnas «NIF / CIF» y «Vendor '
                            'Account».',
        'val_vend_two_cols': 'El Excel de proveedores debe contener exactamente las 2 columnas «NIF / CIF» y '
                             '«Vendor Account» (sin campos de línea de pedido).',
        'vis_footer_note': 'Vista de consulta de la factura original. Los cambios del Excel de Alma se '
                           'realizan en la pestaña Editar.',
        'vis_inv_lbl': 'Factura:',
        'vw_addr': 'Dirección: {addr}',
        'vw_card_buyer': ' DATOS DEL RECEPTOR ',
        'vw_card_dir3': ' CENTROS ADMINISTRATIVOS FACe (DIR3) ',
        'vw_card_lines': ' DETALLE ECONÓMICO (LÍNEAS) ',
        'vw_card_pay': ' DATOS DE PAGO ',
        'vw_card_seller': ' DATOS DEL EMISOR ',
        'vw_col_code': 'Artículo / ISBN',
        'vw_col_desc': 'Descripción / Título',
        'vw_col_gross': 'Importe Bruto',
        'vw_col_price': 'Precio Unit.',
        'vw_col_qty': 'Cant.',
        'vw_col_tax': 'IVA Repercutido',
        'vw_contact': 'Contacto: {contact}',
        'vw_elec_inv_lbl': 'Factura electrónica:',
        'vw_hdr_invoice': 'FACTURA: {number}',
        'vw_hdr_meta': 'Serie: {series}   |   Fecha de Emisión: {issue}   |   Fecha de Operación: {op}',
        'vw_load_btn': 'Cargar factura...',
        'vw_nif': 'CIF/NIF: {nif}',
        'vw_no_number': '(sin número)',
        'vw_not_specified': 'No especificado',
        'vw_pay_due': 'Fecha Vencimiento: {date}',
        'vw_pay_iban': 'Cuenta / IBAN: {iban}',
        'vw_pay_means': 'Medio de Pago: Transferencia (Código {code})',
        'vw_placeholder': 'Seleccione un archivo de factura electrónica (.xsig, .xml, .txt) para visualizar '
                          'sus datos.',
        'vw_role_01': 'Oficina Contable',
        'vw_role_02': 'Órgano Gestor',
        'vw_role_03': 'Unidad Tramitadora',
        'vw_role_04': 'Órgano Proponente',
        'vw_role_other': 'Rol {code}',
        'vw_tab_pdf': '  Vista Factura (Estilo PDF)  ',
        'vw_tab_xml': '  Código XML (FACe)  ',
        'vw_tot_base': 'Base Imponible:',
        'vw_tot_grand': 'TOTAL FACTURA:',
        'vw_tot_payable': 'Total a Pagar:',
        'vw_tot_ret': 'Impuestos Retenidos:',
        'vw_tot_taxes': 'Impuestos Repercutidos:',
        'vw_unknown_vendor': 'Proveedor no identificado'},
 'eu': {'afdlg_apply_next': 'Aplikatu eta jarraitu beste zutabe bat iragazten',
        'afdlg_blank_chk': '(Hutsik / Baliorik gabe)',
        'afdlg_check_all': 'Markatu denak',
        'afdlg_col_lbl': 'Iragazi beharreko zutabea:',
        'afdlg_list_frame': 'Zutabean dauden balioak (markatu erakutsi nahi dituzunak)',
        'afdlg_only_blank': 'Hutsak soilik',
        'afdlg_only_val': 'Baliodunak soilik',
        'afdlg_remove_col': 'Kendu zutabe honen iragazkia',
        'afdlg_search_lbl': 'Bilatu testua / balioa:',
        'afdlg_title': 'Zutabeen autoiragazkia (Excel estiloa)',
        'afdlg_uncheck_all': 'Desmarkatu denak',
        'bulkdlg_lbl': '{col} zutaberako balioa:',
        'bulkdlg_title': 'Esleitu {col}',
        'conv_actions_frame': '4. Ekintzak',
        'conv_alma_btn': 'Hautatu fitxategia...',
        'conv_alma_lbl': 'PO Line-ak dituen Alma Analytics-eko Excel fitxategia *',
        'conv_batch_add': 'Gehitu fakturak',
        'conv_batch_clear': 'Hustu sorta',
        'conv_batch_count': '{count} faktura hautatuta',
        'conv_batch_frame': 'Sortako fakturak',
        'conv_batch_hint': '(Klik bikoitza faktura batean Editatu fitxan irekitzeko)',
        'conv_batch_remove': 'Kendu hautatutakoak',
        'conv_batch_sort': 'Ordenatu zenbakiz',
        'conv_clear_btn': 'Garbitu datuak',
        'conv_col_file': 'Fitxategia',
        'conv_col_invoice': 'Faktura-zenbakia',
        'conv_col_vendor': 'Hornitzailea',
        'conv_exit_btn': 'Irten',
        'conv_inputs_frame': '2. Sarrerako fitxategiak',
        'conv_inv_batch_btn': 'Hautatu fakturak...',
        'conv_inv_batch_lbl': 'Hornitzaile eta liburutegi bereko XSIG / XML / TXT fakturak *',
        'conv_inv_single_btn': 'Hautatu faktura...',
        'conv_inv_single_lbl': 'XSIG / XML / TXT faktura elektronikoa *',
        'conv_mode_batch': 'Hornitzaile eta liburutegi bereko fakturen sorta',
        'conv_mode_frame': '1. Bihurketa modua',
        'conv_mode_single': 'Faktura bakarra',
        'conv_ok_heading': 'BIHURKETA OSOA GORABEHERARIK GABE',
        'conv_open_out_btn': 'Ireki irteerako karpeta',
        'conv_output_btn': 'Aukeratu helmuga...',
        'conv_output_frame': '3. Irteerako fitxategia',
        'conv_result_body': '{heading}\n'
                            '\n'
                            'Prozesatutako fakturak: {total_inv}\n'
                            'Baliozkotze-oharrak: {total_iss}\n'
                            '\n'
                            'Sortutako Alma Excel-a:\n'
                            '{out}',
        'conv_run_btn': 'Bihurtu Alma formatura',
        'conv_tpl_btn': 'Hautatu txantiloia...',
        'conv_tpl_lbl': 'Alma .xlsx txantiloia *',
        'conv_unmatched_warn': '\n'
                               '\n'
                               'Adi: PO Line gabeko lerroak daude. Berrikusi txostena Alman kargatu '
                               'aurretik.',
        'conv_val_report': '\n\nBaliozkotze-txostena:\n{path}',
        'conv_warn_heading': 'GORABEHERAK BIHURKETAN',
        'ctx_bulk_apply': 'Aplikatu balioa hautatutako lerroei...',
        'ctx_clear_cell': 'Hustu gelaxka (Ezab)',
        'ctx_edit_cell': 'Editatu gelaxka',
        'ctx_fill_blanks_down': 'Bete beheko hutsuneak (■)',
        'ctx_fill_blanks_sel': 'Bete hautatutako hutsuneak',
        'ctx_fill_down': 'Bete beherantz (Ctrl+D)',
        'ctx_flt_cell_val': 'Iragazi gelaxka honen balioaren arabera',
        'ctx_flt_clear_all': 'Kendu iragazki guztiak',
        'ctx_flt_col_excel': 'Iragazi zutabe hau (Excel estiloa)...',
        'ctx_redo': 'Berregin (Ctrl+Y)',
        'ctx_restore_cell': 'Berrezarri gelaxka',
        'ctx_restore_orig': 'Berrezarri jatorrizko gelaxka',
        'ctx_search_pol': 'Bilatu PO Line Almako katalogoan...',
        'ctx_suggest_ctx': 'Iradoki balioa testuinguruaren arabera',
        'ctx_undo': 'Desegin (Ctrl+Z)',
        'cust_choose_dir_btn': 'Hautatu karpeta...',
        'cust_dir_alma': 'Almako Excel fitxategiaren direktorioa',
        'cust_dir_inv': 'FACe fakturen direktorioa',
        'cust_dir_out': 'Azken Excel fitxategien direktorioa',
        'cust_dir_tpl': 'Alma txantiloiaren direktorioa',
        'cust_inst_lbl': 'Erakundea',
        'cust_lang_lbl': 'Interfazearen hizkuntza',
        'cust_logo_del_btn': 'Ezabatu logotipoa',
        'cust_logo_hint': 'PNG, JPG edo JPEG. Gehieneko altuera: 100 px.',
        'cust_logo_invalid': 'Hautatutako fitxategia ez da baliozko irudi bat.',
        'cust_logo_lbl': 'Erakundearen logotipoa',
        'cust_logo_loaded': 'Logotipoa kargatuta: {name}',
        'cust_logo_none': 'Logotiporik gabe. PNG, JPG edo JPEG; gehieneko altuera: 100 px.',
        'cust_logo_sel_btn': 'Hautatu logotipoa...',
        'cust_logo_unavail': 'Konfiguratutako logotipoa ez dago erabilgarri.',
        'cust_pref_frame': 'Hobespenak',
        'cust_pref_note': 'Eremu guztiak aukerakoak dira. Karpeta pertsonalizatu bat hutsik geratzen bada, '
                          'azken baliozko karpeta erabiliko da; lehen exekuzioan aplikazioaren karpeta '
                          'erabiliko da. Irteera fakturaren ondoan gordetzen da, irteera-direktorio bat '
                          'zehaztu ezean.',
        'cust_reset_pref_btn': 'Berrezarri balioak',
        'cust_save_pref_btn': 'Gorde hobespenak',
        'cust_vend_acc_lbl': 'Vendor Account Alman:',
        'cust_vend_add_btn': 'Gehitu / Eguneratu',
        'cust_vend_clr_btn': 'Hustu zerrenda',
        'cust_vend_col_acc': 'Vendor Account Alman',
        'cust_vend_col_nif': 'NIF / CIF normalizatua',
        'cust_vend_del_btn': 'Ezabatu hautatutakoa',
        'cust_vend_exp_btn': 'Esportatu Excel...',
        'cust_vend_frame': 'Automatikoki gogoratutako hornitzaileak (IFZ / IFK → Vendor Account Alman)',
        'cust_vend_imp_btn': 'Inportatu Excel...',
        'cust_vend_imp_mode_lbl': 'Excel inportazio-modua:',
        'cust_vend_mode_full': 'Karga osoa (ordezkatu uneko zerrenda)',
        'cust_vend_mode_inc': 'Inkrementala / metagarria (gehitu eta eguneratu)',
        'cust_vend_mode_new': 'Berriak soilik gehitu (ez gainidatzi)',
        'cust_vend_nif_lbl': 'NIF / CIF:',
        'dlg_accept': 'Onartu',
        'dlg_apply': 'Aplikatu',
        'dlg_cancel': 'Utzi',
        'ed_apply_num_btn': 'Aplikatu zenbakia',
        'ed_bottom_hint': 'Idatzi zuzenean edo klik bikoitza editatzeko | Arrastatu ■ betetzeko | Ctrl+F '
                          'bilatu | Klik ▼ goiburuetan iragazteko.',
        'ed_col_invoice': 'Faktura',
        'ed_col_line': 'Lerroa',
        'ed_diag_edited': '✎ {line}. lerroa: saio honetan eskuz aldatuta.',
        'ed_diag_init': 'Hautatu lerro bat gurutzaketa- eta baliozkotze-diagnostikoa ikusteko.',
        'ed_diag_manual': "✓ {line}. lerroa: '{pol}' PO Line eskuz esleituta.",
        'ed_diag_matched': '✓ {line}. lerroa: {msg}.',
        'ed_diag_no_inv': 'Hautatu faktura bat bere lerroak eta oharrak berrikusteko.',
        'ed_diag_ready': '✓ {line}. lerroa: esportatzeko prest.',
        'ed_diag_recon': '⚠ Fakturaren koadratzea: {msg}',
        'ed_diag_unmatched': "⚠ {line}. lerroa: PO Line esleitu gabe (erabili 'Bilatu PO Line Alman...' edo "
                             'idatzi gelaxkan).',
        'ed_diag_warning': '⚠ {line}. lerroa: {notes}',
        'ed_flt_active_prefix': 'Iragazki aktiboak -> ',
        'ed_flt_add_btn': '+ Gehitu iragazkia',
        'ed_flt_clear_btn': 'Garbitu iragazkiak',
        'ed_flt_col_lbl': 'Zutabea:',
        'ed_flt_contains_q': " '{q}' dauka",
        'ed_flt_excel_btn': 'Excel iragazkia...',
        'ed_flt_frame': 'Iragazi faktura-lerroak (egoeraren eta Excel estiloko zutabe baten edo gehiagoren '
                        'arabera)',
        'ed_flt_in_list': '{col} [{items}]-n',
        'ed_flt_none': 'Iragazki aktiborik ez (egin klik edozein ▼ goiburutan Excel estiloko zutabeen '
                       'arabera iragazteko).',
        'ed_flt_op_blank': 'Hutsik',
        'ed_flt_op_contains': 'Badauka / Berdin',
        'ed_flt_op_exact': 'Berdin zehatza',
        'ed_flt_op_nblank': 'Ez hutsik',
        'ed_flt_op_neq': 'Ezberdina',
        'ed_flt_st_all': 'Lerro guztiak',
        'ed_flt_st_edit': 'Eskuz aldatutako lerroak soilik',
        'ed_flt_st_ok': 'Lerro zuzenak soilik (PO Line-arekin eta oharrik gabe)',
        'ed_flt_st_prob': 'Arazoak dituzten lerroak soilik (PO Line gabe edo oharrekin)',
        'ed_flt_st_unm': 'PO Line gabeko lerroak soilik (zain)',
        'ed_flt_state_lbl': 'Egoera:',
        'ed_flt_tag': 'Iragazkia',
        'ed_flt_val_blank': '(Hutsik)',
        'ed_flt_val_lbl': 'Balioa:',
        'ed_flt_val_nblank': '(Ez hutsik)',
        'ed_flt_values_cnt': '{col} ({count} balio)',
        'ed_hdr_frame': 'Fakturaren goiburua (INV)',
        'ed_load_draft_btn': 'Kargatu zirriborroa...',
        'ed_mode_orig': 'Jatorrizkoa',
        'ed_mode_prefix': 'Aurrizkia',
        'ed_mode_suffix': 'Atzizkia',
        'ed_num_frame': 'Faktura-zenbakia Almako Excel-ean',
        'ed_orig_lbl': 'Jatorrizkoa:',
        'ed_refresh_btn': 'Eguneratu zirriborroa',
        'ed_reset_inv_btn': 'Berrezarri faktura',
        'ed_save_draft_btn': 'Gorde zirriborroa...',
        'ed_save_excel_btn': 'Gorde aldaketak eta eguneratu Excel',
        'ed_search_pol_btn': 'Bilatu PO Line Alman...',
        'ed_stats_empty': 'Erakusten: 0/0 | Zuzenak: 0 | Arazoekin: 0',
        'ed_stats_fmt': 'Erakusten: {shown}/{total} | Zuzenak: {ok} | Arazoekin: {prob} (POL gabe: {unm})',
        'ed_tb_bulk': 'Esleitu balioa hautatutakoei...',
        'ed_tb_clear': 'Hustu gelaxka (Ezab)',
        'ed_tb_copy': 'Kopiatu (Ctrl+C)',
        'ed_tb_fill_blanks': '⬇ Bete hutsuneak (■)',
        'ed_tb_fill_down': '⬇ Bete behera (Ctrl+D)',
        'ed_tb_paste': 'Itsatsi (Ctrl+V)',
        'ed_tb_redo': '↪ Berregin (Ctrl+Y)',
        'ed_tb_restore': 'Berrezarri gelaxka',
        'ed_tb_search_lbl': 'Bilatu (Ctrl+F):',
        'ed_tb_undo': '↩ Desegin (Ctrl+Z)',
        'ed_tot_init': 'Agertzen den zenbatekoa: 0,00 €',
        'ed_tot_sel': '   |   Hautapena ({count}): {base}',
        'ed_tot_vis': 'Ikusgai ({count}): Oinarria {base} | BEZa {vat}',
        'err_conv_title': 'Errorea bihurketan zehar',
        'err_fnf_title': 'Fitxategia ez da aurkitu',
        'err_log_hint': "{err}\n\nKontsultatu '{log}' erregistro-fitxategia xehetasun gehiagorako.",
        'err_perm_title': 'Fitxategia blokeatuta edo baimenik gabe',
        'err_val_title': 'Datu edo formatu baliogabeak',
        'err_xml_sec_title': 'XMLa segurtasunez blokeatuta',
        'fd_exp_vendors': 'Esportatu hornitzaileak (NIF / CIF eta Vendor Account) Excel-era',
        'fd_ft_all': 'Fitxategi guztiak',
        'fd_ft_compat_inv': 'Faktura bateragarriak',
        'fd_ft_draft': 'FacturaeToAlma zirriborroa',
        'fd_ft_excel': 'Excel fitxategiak',
        'fd_ft_img_compat': 'Irudi bateragarriak',
        'fd_ft_jpg': 'JPEG irudia',
        'fd_ft_png': 'PNG irudia',
        'fd_ft_txt': 'TXT fitxategiak',
        'fd_ft_xml': 'XML fitxategiak',
        'fd_ft_xsig': 'XSIG fitxategiak',
        'fd_imp_vendors': 'Inportatu hornitzaileak (NIF / CIF eta Vendor Account) Excel-etik',
        'fd_load_draft': 'Kargatu FacturaeToAlma saioko zirriborroa',
        'fd_save_draft': 'Gorde FacturaeToAlma saioko zirriborroa',
        'fd_sel_alma': 'Hautatu Alma Analytics-eko Excel fitxategia',
        'fd_sel_folder': 'Hautatu karpeta',
        'fd_sel_inv_batch': 'Hautatu hornitzaile eta liburutegi bereko fakturak',
        'fd_sel_inv_face': 'Hautatu FACe faktura elektronikoa',
        'fd_sel_inv_single': 'Hautatu XSIG / XML / TXT faktura',
        'fd_sel_logo': 'Hautatu erakundearen logotipoa',
        'fd_sel_out': 'Zehaztu irteerako Excel fitxategia',
        'fd_sel_tpl': 'Hautatu Alma .xlsx txantiloia',
        'gctx_copy': 'Kopiatu',
        'gctx_copy_txt': 'Kopiatu testua',
        'gctx_cut': 'Ebaki',
        'gctx_go_edit': 'Joan Editatu fitxara',
        'gctx_go_vis': 'Joan Bistaratu fitxara',
        'gctx_paste': 'Itsatsi',
        'gctx_sel_all': 'Hautatu dena',
        'hlp_body_5_0': 'Faktura bakarra: hautatu faktura bat, Alma Analytics txostena, txantiloia eta '
                        'helmuga.\n'
                        '\n'
                        'Sorta modua: faktura guztiek hornitzaile eta liburutegi berekokoak izan behar dute. '
                        'Zerrendak fakturak gehitu, kendu, ordenatu eta klik bikoitzarekin Editatu fitxan '
                        'irekitzeko aukera ematen du.\n'
                        '\n'
                        'Edizioa eta iragazketa: berrikusi goiburua (INV) eta lerroak (IL), bilatu PO '
                        'Line-ak Alma Analytics katalogoan, aplikatu Excel estiloko autoiragazkiak edo '
                        'gorde/kargatu zirriborroak (.ftalma.json).\n'
                        '\n'
                        'Hornitzaileak eta hizkuntzak: Pertsonalizatu fitxan Espainiako hizkuntza ofizial '
                        'eta koofizialak (es, ca, va, gl, eu, oc) hauta ditzakezu, baita hornitzaileen taula '
                        '(NIF / CIF eta Vendor Account) Excel bidez inportatu edo esportatu ere (modu '
                        'inkrementalean, berriak soilik edo karga osoan).\n'
                        '\n'
                        'Baliozkotze-txostena: PO Line-ak falta direnean, bat-etortze anbiguoak, bikoiztuak, '
                        'kontabilitate-desorekak edo Quantity eta Quantity for Pricing arteko aldeak '
                        'daudenean sortzen da.\n'
                        '\n'
                        'Ex Librisek baieztatutako mugak: Excel txantiloiak ez du uzten Line Exclusive '
                        'adierazten ezta fakturazio partziala edo osoa esplizituki zehazten ere.\n'
                        '\n'
                        'Lizentzia: GPL-3.0-only.',
        'hlp_body_5_1': 'Faktura bakarra: hautatu faktura bat, Alma Analytics txostena, txantiloia eta '
                        'helmuga.\n'
                        '\n'
                        'Sorta modua: faktura guztiek hornitzaile eta liburutegi berekokoak izan behar dute. '
                        'Zerrendak fakturak gehitu, kendu, ordenatu eta klik bikoitzarekin Editatu fitxan '
                        'irekitzeko aukera ematen du.\n'
                        '\n'
                        'Lerro esklusiboarekin bateragarriak diren zenbatekoak eta BEZa (4.0 logika '
                        'zehatza): zenbatekoen eta BEZaren erauzketak eta esportazioak 4.0 bertsioaren '
                        'portaera zehatza errepikatzen dute, Alman «Lerro esklusiboa» aukerara aldatzean '
                        'beti koadratu dadin.\n'
                        '\n'
                        'Hornitzaileak eta hizkuntzak: Pertsonalizatu fitxan Espainiako hizkuntza ofizial '
                        'eta koofizialak (es, ca, va, gl, eu, oc) hauta ditzakezu, eta hornitzaileen taula '
                        '(NIF / CIF eta Vendor Account zutabeak soilik) Excel bidez inportatu edo esportatu '
                        '(inkrementala/metagarria, berriak soilik gehitu edo karga osoa).\n'
                        '\n'
                        'Baliozkotze-txostena: PO Line-ak falta direnean, bat-etortze anbiguoak, bikoiztuak '
                        'edo Quantity eta Quantity for Pricing arteko aldeak daudenean sortzen da.\n'
                        '\n'
                        'Ex Librisek baieztatutako mugak: Excel txantiloiak ez du uzten Line Exclusive '
                        'adierazten ezta fakturazio partziala edo osoa esplizituki zehazten ere.\n'
                        '\n'
                        'Lizentzia: GPL-3.0-only.',
        'hlp_guide_btn': 'Ireki Biblioguia / LibGuides',
        'hlp_links_frame': 'Dokumentazioa eta laguntza',
        'hlp_repo_btn': 'Ireki GitHub biltegia',
        'match_msg_ambig': 'Izenburu bidezko hainbat bat-etortze posible',
        'match_msg_dup': '{mtype} bidez bikoiztua. Eskuzko berrikuspena behar da.',
        'match_msg_manual': 'PO Line eskuz esleituta',
        'match_msg_none': 'Bat-etortzerik ez Alman ISBN, ISSN edo izenburu bidez',
        'match_msg_special': '{ltype} motako lerroa (ez du PO Line bibliografikorik behar)',
        'match_msg_unique': 'Bat-etortze bakarra {mtype} bidez',
        'mb_batch_clr_ask': 'Sortako faktura guztiak ezabatu nahi dituzu?',
        'mb_batch_clr_title': 'Hustu sorta',
        'mb_batch_dup_num': '{num} faktura-zenbaki bikoiztua: {name}',
        'mb_batch_none_added': 'Ez da faktura berririk gehitu.',
        'mb_batch_read_err': 'Ezin izan da {name} irakurri: {err}',
        'mb_batch_rep_path': 'Bide errepikatua: {name}',
        'mb_batch_sel_one': 'Hautatu zerrendako faktura bat gutxienez.',
        'mb_batch_title': 'Sorta',
        'mb_batch_warn_title': 'Sartu gabeko fakturak',
        'mb_clr_alma_ask': 'Alma Analytics-eko Excel fitxategia ere garbitu nahi duzu?\n'
                           '\n'
                           'Alma txantiloia mantenduko da.',
        'mb_clr_alma_title': 'Garbitu Almako Excel fitxategia',
        'mb_conv_close_ask': 'Bihurketa bat abian da. Aplikazioa itxi nahi duzu hala ere?',
        'mb_conv_ok_title': 'Bihurketa osatuta',
        'mb_conv_running_msg': 'Dagoeneko bihurketa bat abian da.',
        'mb_conv_running_title': 'Bihurketa abian',
        'mb_conv_warn_title': 'Gorabeherak bihurketan',
        'mb_draft_loaded_msg': 'Zirriborroko fakturak eta aldaketak leheneratu dira.{warn}',
        'mb_draft_loaded_title': 'Zirriborroa kargatuta',
        'mb_draft_missing_warn': '\n'
                                 '\n'
                                 'Oharra: lekuz aldatutako edo ezabatutako faktura batzuk ez dira aurkitu:\n'
                                 '- {files}',
        'mb_draft_outdated_msg': 'Faktura, Analytics edo txantiloia gorde ondoren aldatu dira. Itzuli '
                                 'Editatu fitxara, berrikusi eta gorde berriro.',
        'mb_draft_outdated_title': 'Zirriborro zaharkitua',
        'mb_draft_rebuilt_alma': 'Almako azken Excel-aren zirriborroa berreraiki da hautatutako fakturarekin '
                                 'eta Analytics txostenarekin.',
        'mb_draft_rebuilt_no_alma': 'Fakturaren lerroak editorean kargatu dira.\n'
                                    '\n'
                                    'Alma Analytics-eko Excel fitxategia ere hautatzen baduzu, PO Line-ak, '
                                    'Reporting Code-ak eta funtsak automatikoki gurutzatuko dira.',
        'mb_draft_saved_msg': 'Uneko aurrerapena hemen gorde da:\n'
                              '{path}\n'
                              '\n'
                              "Geroago edizioari berrekin ahal izango diozu 'Kargatu zirriborroa...' "
                              'erabiliz.',
        'mb_draft_saved_title': 'Zirriborroa gordeta',
        'mb_draft_unavail_msg': 'Ezin izan da zirriborro editagarria prestatu bihurtu aurretik.',
        'mb_draft_unavail_title': 'Zirriborroa ez dago erabilgarri',
        'mb_empty_src_msg': 'Gelaxka aktiboa hutsik dago. Idatzi edo hautatu balioa duen gelaxka bat '
                            'hutsuneak bete aurretik.',
        'mb_empty_src_title': 'Jatorrizko gelaxka hutsik',
        'mb_err_folder_msg': 'Ezin izan da karpeta ireki:\n{path}\n\n{err}',
        'mb_err_folder_title': 'Errorea karpeta irekitzean',
        'mb_err_load_draft': 'Errorea zirriborroa kargatzean',
        'mb_err_load_draft_msg': 'Ezin izan da zirriborroa kargatu:\n{err}',
        'mb_err_prep_draft': 'Ezin izan da zirriborroa prestatu',
        'mb_err_read_inv_msg': 'Ezin izan da hautatutako faktura kargatu:\n{err}',
        'mb_err_read_inv_title': 'Errorea faktura irakurtzean',
        'mb_err_reset_title': 'Errorea berrezartzean',
        'mb_err_save_draft': 'Errorea zirriborroa gordetzean',
        'mb_err_save_draft_msg': 'Ezin izan da zirriborroa gorde:\n{err}',
        'mb_err_save_title': 'Errorea gordetzean',
        'mb_err_url_msg': 'Ezin izan da esteka ireki:\n{url}\n\n{err}',
        'mb_err_url_title': 'Errorea esteka irekitzean',
        'mb_flt_cols_no_inv': 'Hautatu lehenik faktura bat Editatu fitxan.',
        'mb_flt_cols_title': 'Iragazi zutabeak',
        'mb_inv_invalid_title': 'Faktura baliogabea',
        'mb_invalid_date_msg': 'Erabili baliozko data bat, adibidez 2026-10-19.',
        'mb_invalid_date_title': 'Data baliogabea',
        'mb_invalid_inputs_title': 'Sarrerako datu baliogabeak',
        'mb_invalid_ltype': 'Hautatu Almak onartutako Line type bat.',
        'mb_invalid_num': '{col} zenbakizkoa izan behar da.',
        'mb_invalid_val_title': 'Balio baliogabea',
        'mb_logo_invalid_msg': 'Ezin izan da logotipoa kargatu:\n{path}\n\n{err}',
        'mb_logo_invalid_title': 'Logotipo baliogabea',
        'mb_logo_unavail_msg': 'Ez da logotipoaren fitxategia aurkitu:\n{path}',
        'mb_logo_unavail_title': 'Logotipoa ez dago erabilgarri',
        'mb_no_active_inv': 'Ez dago faktura aktiborik.',
        'mb_no_edit_inv': 'Ez dago faktura aktiborik editorean.',
        'mb_no_out_dir_def': 'Ez dago irteerako karpetarik zehaztuta.',
        'mb_no_out_dir_found': 'Ez da irteerako karpeta aurkitu.',
        'mb_no_sugg_msg': 'Ez dago testuinguruko baliorik zutabe honetan iradokitzeko.',
        'mb_no_sugg_title': 'Iradokizunik gabe',
        'mb_pol_ask_alma': 'Bilatzaile honek Alma Analytics-etik esportatutako Excel fitxategiko '
                           'eskaera-lerroak (PO Lines) kontsultatzen ditu.\n'
                           '\n'
                           'Oraindik ez duzu Alma Analytics-eko Excel fitxategirik hautatu.\n'
                           'Orain hautatu nahi duzu?',
        'mb_pol_no_inv': 'Hautatu lehenik faktura bat (edo faktura-sorta bat) Bihurtu fitxan.',
        'mb_pol_no_lines': 'Ez dago lerro ikusgairik editorean.',
        'mb_pol_sel_row': 'Hautatu zerrendako errenkada bat.',
        'mb_pol_title': 'Bilatu PO Line Alman',
        'mb_pref_invalid_dir': '{label} ez da baliozko karpeta bat:\n{val}',
        'mb_pref_invalid_title': 'Hobespen baliogabeak',
        'mb_pref_reset_ask': 'Erakundea, hizkuntza eta direktorio pertsonalizatu guztiak berrezarri nahi '
                             'dituzu?',
        'mb_pref_reset_ok': 'Hasierako balioak berrezarri dira.',
        'mb_pref_reset_title': 'Berrezarri hobespenak',
        'mb_pref_saved_msg': 'Hobespenak ongi gorde dira.',
        'mb_prot_drag_msg': 'Zutabe hau babestuta dago eta ezin da arrastatuz bete.',
        'mb_prot_field_title': 'Babestutako eremua',
        'mb_prot_generic_msg': 'Eremu hau ezin da editatu.',
        'mb_prot_hdr_msg': 'Goiburuko zenbakia, data eta zenbatekoak ez dira zuzenean editatzen. Zenbakiak '
                           'aurrizkia edo atzizkia soilik onartzen du.',
        'mb_prot_line_msg': 'Eremu hau jatorrizko fakturatik dator eta ezin da FacturaeToAlma-n editatu.',
        'mb_reset_inv_ask': 'Faktura honetako aldaketa guztiak baztertu nahi dituzu?',
        'mb_reset_inv_title': 'Berrezarri faktura',
        'mb_save_changes_ask': 'Aldaketa guztiak gorde eta irteerako Excel-a orain eguneratu nahi duzu?\n'
                               '\n'
                               'Bihurtu fitxan adierazitako Excel-a gainidatziko da. Jatorrizko XMLa ez da '
                               'aldatuko.',
        'mb_save_changes_title': 'Gorde aldaketak',
        'mb_save_draft_empty': 'Ez dago fakturarik kargatuta zirriborro batean gordetzeko.',
        'mb_save_draft_title': 'Gorde zirriborroa',
        'mb_unsaved_ed_ask': 'Editatu fitxan egindako aldaketak daude (faktura-zenbakia, goiburua edo '
                             'lerroak) oraindik gorde gabe.\n'
                             '\n'
                             'Aldaketa horiek aplikatu eta gorde nahi dituzu orain bihurtzeko?',
        'mb_unsaved_ed_title': 'Gorde gabeko aldaketak Editatu fitxan',
        'mb_vend_clr_ask': 'Gordetako NIF → Vendor Account elkarketa guztiak ezabatu nahi dituzu?',
        'mb_vend_clr_title': 'Hustu gogoratutako hornitzaileak',
        'mb_vend_exp_err_title': 'Errorea hornitzaileak esportatzean',
        'mb_vend_exp_ok_msg': '{count} hornitzaile esportatu dira (NIF / CIF eta Vendor Account zutabeak) '
                              'fitxategi honetara:\n'
                              '{path}',
        'mb_vend_exp_ok_title': 'Hornitzaileen esportazioa osatuta',
        'mb_vend_full_ask': '«Karga osoa» moduak gaur egun gordetako {count} hornitzaileak hautatutako Excel '
                            'fitxategikoekin ordezkatuko ditu.\n'
                            '\n'
                            'Jarraitu nahi duzu?',
        'mb_vend_imp_err_title': 'Errorea hornitzaileak inportatzean',
        'mb_vend_imp_ok_msg': 'Aplikatutako modua: {mode}\n'
                              '\n'
                              '- Gehitutako berriak: {added}\n'
                              '- Eguneratutako lehendik zeudenak: {updated}\n'
                              '- Aldaketarik gabe / baztertuak: {skipped}\n'
                              '- Baztertutako errenkada osatugabeak: {invalid}\n'
                              '\n'
                              'Gordetako hornitzaileak guztira: {total}',
        'mb_vend_imp_ok_title': 'Hornitzaileen inportazioa osatuta',
        'mb_vend_missing': 'Adierazi bai NIF/CIFa bai Almako Vendor Account kodea.',
        'mb_vend_sel_del': 'Hautatu taulako errenkada bat ezabatzeko.',
        'mb_vend_title': 'Gogoratutako hornitzaileak',
        'mb_warn_title': 'Oharra',
        'numdlg_apply': 'Aplikatu saio honetan',
        'numdlg_desc': 'Litekeena da zenbakia aldian-aldian berrabiarazten den segida batekoa izatea.\n'
                       'Aldi baterako aurrizki edo atzizki bat gehi dezakezu anbiguotasuna kentzeko.\n'
                       'Jatorrizko XMLa ez da aldatuko.',
        'numdlg_orig': 'Jatorrizko zenbakia: {original}',
        'numdlg_title': 'Zuzendu faktura-zenbakia memorian',
        'poldlg_all_btn': 'Ikusi denak',
        'poldlg_assign_full': 'Esleitu PO Line + Analytics-eko kodeak eta funtsak',
        'poldlg_assign_only': 'Esleitu PO Line soilik',
        'poldlg_col_score': 'Kidetasuna',
        'poldlg_dblclick_hint': 'Klik bikoitzak PO Line bakarrik esleitzen du (Reporting Codes eta funtsak '
                                'aldatu gabe).',
        'poldlg_inv_title': 'Fakturako izenburua: {title}',
        'poldlg_line_frame': 'Faktura-lerroa #{line}',
        'poldlg_search_btn': 'Bilatu',
        'poldlg_search_lbl': 'Bilatu Analytics-en (Izenburua / ISBN / ISSN / PO Line):',
        'poldlg_summary': 'ISBN: {isbn}   |   ISSN: {issn}   |   Prezioa: {price}   |   Uneko PO Line: {pol}',
        'poldlg_title': 'Bilatu PO Line Almako katalogoan - {line}. lerroa',
        'poldlg_unassigned': '(esleitu gabe)',
        'st_alma_selected': 'Egoera: Alma Analytics-eko Excel fitxategia hautatuta.',
        'st_batch_count': 'Egoera: {count} faktura sortan.',
        'st_batch_empty': 'Egoera: sorta hutsik.',
        'st_batch_sorted': 'Egoera: sorta faktura-zenbakiaren arabera ordenatuta.',
        'st_check_inputs': 'Egoera: berrikusi sarrerako datuak.',
        'st_cleared': 'Egoera: datuak garbituta. Alma txantiloia mantentzen da.',
        'st_conv_err': 'Egoera: errorea bihurketan zehar.',
        'st_conv_ok': 'Egoera: bihurketa gorabeherarik gabe osatu da.',
        'st_conv_warn': 'Egoera: bihurketa oharrekin osatu da.',
        'st_converting': 'Egoera: bigarren planoan bihurtzen...',
        'st_copied': "Egoera: {col} zutabeko '{val}' balioa kopiatuta.",
        'st_draft_loaded': "Egoera: '{name}' zirriborroa ongi kargatu da.",
        'st_draft_ready_alma': 'Egoera: Almako Excel zirriborro editagarria prest (Alma Analytics-ekin '
                               'gurutzatuta).',
        'st_draft_ready_no_alma': 'Egoera: faktura Editatu fitxan kargatuta (hautatu Alma Analytics-eko '
                                  'Excel-a PO Line-ak gurutzatzeko).',
        'st_draft_saved': "Egoera: zirriborroa '{name}' fitxategian gordeta.",
        'st_fill_blanks': "Egoera: beheko {count} gelaxka huts bete dira '{val}' balioarekin.",
        'st_fill_drag': "Egoera: {count} gelaxka bete dira '{col}' zutabean '{val}' balioarekin.",
        'st_inv_selected': 'Egoera: {number} faktura hautatuta. Hornitzailea: {vendor}.',
        'st_lang_changed': 'Egoera: hizkuntza {lang} aukerara aldatuta.',
        'st_mode_batch': 'Egoera: sorta modua. Hautatu hornitzaile eta liburutegi bereko fakturak.',
        'st_mode_single': 'Egoera: faktura bakarreko modua.',
        'st_no_redo': 'Egoera: ez dago berregiteko ekintzarik.',
        'st_no_undo': 'Egoera: ez dago desegiteko ekintzarik.',
        'st_num_updated': "Egoera: faktura-zenbakia zirriborroan '{number}' baliora eguneratuta.",
        'st_open_in_edit': "Egoera: '{name}' faktura Editatu fitxan irekita.",
        'st_out_selected': 'Egoera: irteerako fitxategia zehaztuta.',
        'st_pending': 'Egoera: fitxategiak hautatzeko zain.',
        'st_redone': "Egoera: '{action}' ekintza berreginda.",
        'st_saving_edits': 'Egoera: {n_num} zenbaki-aldaketa, {n_hdr} goiburuko eta {n_line} lerroko '
                           'gordeta; irteerako Excel-a eguneratzen...',
        'st_tpl_selected': 'Egoera: Alma txantiloia hautatuta.',
        'st_undone': "Egoera: '{action}' ekintza deseginda.",
        'st_vendor_cleared': 'Egoera: gogoratutako hornitzaileen zerrenda hustuta.',
        'st_vendor_deleted': "Egoera: '{nif}' IFZren hornitzailea ezabatuta.",
        'st_vendor_exported': "Egoera: {count} hornitzaile '{name}' fitxategira esportatuta.",
        'st_vendor_imported': "Egoera: hornitzaileak '{name}' fitxategitik inportatuta (guztira: {total}).",
        'st_vendor_saved': 'Egoera: hornitzailea gordeta ({nif} → {vendor}).',
        'sum_body': '\n'
                    '\n'
                    'Lerroen laburpena:\n'
                    '- Prozesatutako lerroak: {total_lines}\n'
                    '- ISBN bidezko bat-etortzeak: {match_isbn}\n'
                    '- ISSN bidezko bat-etortzeak: {match_issn}\n'
                    '- Izenburu bidezko bat-etortzeak: {match_title}\n'
                    '{extra}- ISBN bikoiztuak: {duplicate_isbn}\n'
                    '- ISSN bikoiztuak: {duplicate_issn}\n'
                    '- Izenburu anbiguoak: {ambiguous_title}\n'
                    '- Bat-etortzerik gabe: {no_match}\n'
                    '- PO Line gabeko lerroak: {lines_without_po_line}',
        'sum_pol_man': '- Eskuzko PO Line esleipenak: {n}\n',
        'sum_pol_xml': '- XMLko PO Line bidezko bat-etortzeak: {n}\n',
        'tab_convert': 'Bihurtu',
        'tab_customize': 'Pertsonalizatu',
        'tab_edit': 'Editatu',
        'tab_help': 'Laguntza',
        'tab_visualize': 'Bistaratu',
        'val_bad_draft': 'Fitxategiak ez du FacturaeToAlma-ko baliozko zirriborro-formaturik.',
        'val_bad_ext': 'Faktura-luzapen baliogabea: {path}',
        'val_draft_no_files': 'Zirriborroan aipatutako faktura bat bera ere ez dago jatorrizko bidean.',
        'val_dup_numbers': 'Faktura-zenbaki bikoiztuak daude sortan: {nums}',
        'val_multi_vendor': 'Sortak hainbat hornitzaileren fakturak ditu. Hautatu hornitzaile bereko '
                            'fakturak soilik.\n'
                            '\n'
                            'Hautemandako hornitzaileak:\n'
                            '{details}',
        'val_need_alma': 'Alma Analytics-eko Excel fitxategia hautatu behar duzu.',
        'val_need_inv': 'Faktura bat edo faktura-sorta bat hautatu behar duzu.',
        'val_need_inv_edit': 'Lehenik faktura bat edo faktura-sorta bat hautatu behar duzu Bihurtu fitxan.',
        'val_need_out': 'Irteerako Excel-a zehaztu behar duzu.',
        'val_need_tpl': 'Alma txantiloia hautatu behar duzu.',
        'val_no_alma_file': 'Alma Analytics-eko Excel fitxategia ez da existitzen: {path}',
        'val_no_inv_file': 'Faktura ez da existitzen: {path}',
        'val_no_tpl_file': 'Alma txantiloia ez da existitzen: {path}',
        'val_out_collision': 'Irteerako Excel-a ezin da fakturarekin, Analytics-ekin edo txantiloiarekin bat '
                             'etorri. Aukeratu beste izen edo karpeta bat.',
        'val_single_inv_only': 'Fitxategiak faktura bakarra izan behar du eta {count} ditu.',
        'val_vend_empty_xlsx': 'Excel fitxategia hutsik dago.',
        'val_vend_po_cols': 'Fitxategiak eskaera-lerroko zutabeak ditu ({cols}). Hornitzaileen Excel-ak «NIF '
                            '/ CIF» eta «Vendor Account» 2 zutabeak soilik izan behar ditu.',
        'val_vend_two_cols': 'Hornitzaileen Excel-ak «NIF / CIF» eta «Vendor Account» 2 zutabeak zehazki '
                             'izan behar ditu (eskaera-lerroko eremurik gabe).',
        'vis_footer_note': 'Jatorrizko fakturaren kontsulta-ikuspegia. Almako Excel-aren aldaketak Editatu '
                           'fitxan egiten dira.',
        'vis_inv_lbl': 'Faktura:',
        'vw_addr': 'Helbidea: {addr}',
        'vw_card_buyer': ' HARTZAILEAREN DATUAK ',
        'vw_card_dir3': ' FACe ADMINISTRAZIO-ZENTROAK (DIR3) ',
        'vw_card_lines': ' XEHETASUN EKONOMIKOA (LERROAK) ',
        'vw_card_pay': ' ORDAINKETA DATUAK ',
        'vw_card_seller': ' JAULKITZAILEAREN DATUAK ',
        'vw_col_code': 'Artikulua / ISBN',
        'vw_col_desc': 'Deskribapena / Izenburua',
        'vw_col_gross': 'Zenbateko Gordina',
        'vw_col_price': 'Unit. Prezioa',
        'vw_col_qty': 'Kop.',
        'vw_col_tax': 'Jasanarazitako BEZa',
        'vw_contact': 'Harremana: {contact}',
        'vw_elec_inv_lbl': 'Faktura elektronikoa:',
        'vw_hdr_invoice': 'FAKTURA: {number}',
        'vw_hdr_meta': 'Seriea: {series}   |   Jaulkipen-data: {issue}   |   Eragiketa-data: {op}',
        'vw_load_btn': 'Kargatu faktura...',
        'vw_nif': 'IFK/IFZ: {nif}',
        'vw_no_number': '(zenbakirik gabe)',
        'vw_not_specified': 'Zehaztu gabe',
        'vw_pay_due': 'Epemuga-data: {date}',
        'vw_pay_iban': 'Kontua / IBAN: {iban}',
        'vw_pay_means': 'Ordainbidea: Transferentzia ({code} kodea)',
        'vw_placeholder': 'Hautatu faktura elektronikoko fitxategi bat (.xsig, .xml, .txt) bere datuak '
                          'ikusteko.',
        'vw_role_01': 'Kontabilitate Bulegoa',
        'vw_role_02': 'Organo Kudeatzailea',
        'vw_role_03': 'Izapidetze Unitatea',
        'vw_role_04': 'Organo Proposatzailea',
        'vw_role_other': '{code} rola',
        'vw_tab_pdf': '  Faktura ikuspegia (PDF estiloa)  ',
        'vw_tab_xml': '  XML kodea (FACe)  ',
        'vw_tot_base': 'Zerga Oinarria:',
        'vw_tot_grand': 'FAKTURA GUZTIRA:',
        'vw_tot_payable': 'Ordaintzeko Guztira:',
        'vw_tot_ret': 'Atxikitako Zergak:',
        'vw_tot_taxes': 'Jasanarazitako Zergak:',
        'vw_unknown_vendor': 'Identifikatu gabeko hornitzailea'},
 'gl': {'afdlg_apply_next': 'Aplicar e seguir filtrando outra columna',
        'afdlg_blank_chk': '(Baleiras / Sen valor)',
        'afdlg_check_all': 'Marcar todos',
        'afdlg_col_lbl': 'Columna a filtrar:',
        'afdlg_list_frame': 'Valores presentes na columna (marca os que desexas mostrar)',
        'afdlg_only_blank': 'Só baleiras',
        'afdlg_only_val': 'Só con valor',
        'afdlg_remove_col': 'Quitar filtro desta columna',
        'afdlg_search_lbl': 'Buscar texto / valor:',
        'afdlg_title': 'Filtro autofiltro por columnas (estilo Excel)',
        'afdlg_uncheck_all': 'Desmarcar todos',
        'bulkdlg_lbl': 'Valor para {col}:',
        'bulkdlg_title': 'Asignar {col}',
        'conv_actions_frame': '4. Accións',
        'conv_alma_btn': 'Seleccionar ficheiro...',
        'conv_alma_lbl': 'Ficheiro Excel de Alma Analytics con PO Lines *',
        'conv_batch_add': 'Engadir facturas',
        'conv_batch_clear': 'Baleirar lote',
        'conv_batch_count': '{count} facturas seleccionadas',
        'conv_batch_frame': 'Facturas do lote',
        'conv_batch_hint': '(Dobre clic nunha factura para abrila en Editar)',
        'conv_batch_remove': 'Quitar seleccionadas',
        'conv_batch_sort': 'Ordenar por número',
        'conv_clear_btn': 'Limpar datos',
        'conv_col_file': 'Ficheiro',
        'conv_col_invoice': 'Número de factura',
        'conv_col_vendor': 'Provedor',
        'conv_exit_btn': 'Saír',
        'conv_inputs_frame': '2. Arquivos de entrada',
        'conv_inv_batch_btn': 'Seleccionar facturas...',
        'conv_inv_batch_lbl': 'Facturas XSIG / XML / TXT do mesmo provedor e biblioteca *',
        'conv_inv_single_btn': 'Seleccionar factura...',
        'conv_inv_single_lbl': 'Factura electrónica XSIG / XML / TXT *',
        'conv_mode_batch': 'Lote de facturas do mesmo provedor e biblioteca',
        'conv_mode_frame': '1. Modo de conversión',
        'conv_mode_single': 'Factura individual',
        'conv_ok_heading': 'CONVERSIÓN COMPLETA SEN INCIDENCIAS',
        'conv_open_out_btn': 'Abrir cartafol de saída',
        'conv_output_btn': 'Escoller destino...',
        'conv_output_frame': '3. Arquivo de saída',
        'conv_result_body': '{heading}\n'
                            '\n'
                            'Facturas procesadas: {total_inv}\n'
                            'Avisos de validación: {total_iss}\n'
                            '\n'
                            'Excel Alma xerado:\n'
                            '{out}',
        'conv_run_btn': 'Converter a formato Alma',
        'conv_tpl_btn': 'Seleccionar modelo...',
        'conv_tpl_lbl': 'Modelo Alma .xlsx *',
        'conv_unmatched_warn': '\n'
                               '\n'
                               'Atención: hai liñas sen PO Line. Revisa o informe antes de cargar en Alma.',
        'conv_val_report': '\n\nInforme de validación:\n{path}',
        'conv_warn_heading': 'INCIDENCIAS NA CONVERSIÓN',
        'ctx_bulk_apply': 'Aplicar valor ás liñas seleccionadas...',
        'ctx_clear_cell': 'Baleirar cela (Supr)',
        'ctx_edit_cell': 'Editar cela',
        'ctx_fill_blanks_down': 'Encher brancos inferiores (■)',
        'ctx_fill_blanks_sel': 'Encher brancos seleccionados',
        'ctx_fill_down': 'Encher cara abaixo (Ctrl+D)',
        'ctx_flt_cell_val': 'Filtrar polo valor desta cela',
        'ctx_flt_clear_all': 'Quitar todos os filtros',
        'ctx_flt_col_excel': 'Filtrar esta columna (estilo Excel)...',
        'ctx_redo': 'Refacer (Ctrl+Y)',
        'ctx_restore_cell': 'Restablecer cela',
        'ctx_restore_orig': 'Restablecer cela orixinal',
        'ctx_search_pol': 'Buscar PO Line no catálogo de Alma...',
        'ctx_suggest_ctx': 'Suxerir valor polo contexto',
        'ctx_undo': 'Desfacer (Ctrl+Z)',
        'cust_choose_dir_btn': 'Seleccionar cartafol...',
        'cust_dir_alma': 'Directorio do ficheiro Excel de Alma',
        'cust_dir_inv': 'Directorio de facturas FACe',
        'cust_dir_out': 'Directorio de Excel finais',
        'cust_dir_tpl': 'Directorio de modelo Alma',
        'cust_inst_lbl': 'Institución',
        'cust_lang_lbl': 'Idioma da interface',
        'cust_logo_del_btn': 'Eliminar logo',
        'cust_logo_hint': 'PNG, JPG ou JPEG. Altura máxima: 100 px.',
        'cust_logo_invalid': 'O arquivo seleccionado non é unha imaxe válida.',
        'cust_logo_lbl': 'Logo institucional',
        'cust_logo_loaded': 'Logo cargado: {name}',
        'cust_logo_none': 'Sen logo. PNG, JPG ou JPEG; altura máxima: 100 px.',
        'cust_logo_sel_btn': 'Seleccionar logo...',
        'cust_logo_unavail': 'O logo configurado non está dispoñible.',
        'cust_pref_frame': 'Preferencias',
        'cust_pref_note': 'Todos os campos son opcionais. Se un cartafol personalizado queda baleiro, úsase '
                          'o último cartafol válido; na primeira execución úsase o cartafol da aplicación. A '
                          'saída gárdase xunto á factura, agás que se defina un directorio de saída.',
        'cust_reset_pref_btn': 'Restablecer valores',
        'cust_save_pref_btn': 'Gardar preferencias',
        'cust_vend_acc_lbl': 'Vendor Account en Alma:',
        'cust_vend_add_btn': 'Engadir / Actualizar',
        'cust_vend_clr_btn': 'Baleirar lista',
        'cust_vend_col_acc': 'Vendor Account en Alma',
        'cust_vend_col_nif': 'NIF / CIF normalizado',
        'cust_vend_del_btn': 'Eliminar seleccionado',
        'cust_vend_exp_btn': 'Exportar Excel...',
        'cust_vend_frame': 'Provedores lembrados automaticamente (NIF / CIF → Vendor Account en Alma)',
        'cust_vend_imp_btn': 'Importar Excel...',
        'cust_vend_imp_mode_lbl': 'Modo de importación Excel:',
        'cust_vend_mode_full': 'Carga completa (substituír lista actual)',
        'cust_vend_mode_inc': 'Incremental / acumulativo (engadir e actualizar)',
        'cust_vend_mode_new': 'Só engadir novos (non sobrescribir)',
        'cust_vend_nif_lbl': 'NIF / CIF:',
        'dlg_accept': 'Aceptar',
        'dlg_apply': 'Aplicar',
        'dlg_cancel': 'Cancelar',
        'ed_apply_num_btn': 'Aplicar número',
        'ed_bottom_hint': 'Escribe directo ou dobre clic para editar | Arrastra ■ para encher | Ctrl+F '
                          'buscar | Clic nas cabeceiras ▼ para filtrar.',
        'ed_col_invoice': 'Factura',
        'ed_col_line': 'Liña',
        'ed_diag_edited': '✎ Liña {line}: modificada manualmente nesta sesión.',
        'ed_diag_init': 'Selecciona unha liña para ver o seu diagnóstico de cruzamento e validación.',
        'ed_diag_manual': "✓ Liña {line}: PO Line '{pol}' asignada manualmente.",
        'ed_diag_matched': '✓ Liña {line}: {msg}.',
        'ed_diag_no_inv': 'Selecciona unha factura para revisar as súas liñas e avisos.',
        'ed_diag_ready': '✓ Liña {line}: lista para exportar.',
        'ed_diag_recon': '⚠ Cadre factura: {msg}',
        'ed_diag_unmatched': "⚠ Liña {line}: sen PO Line asignada (usa 'Buscar PO Line en Alma...' ou "
                             'escríbea na cela).',
        'ed_diag_warning': '⚠ Liña {line}: {notes}',
        'ed_flt_active_prefix': 'Filtros activos -> ',
        'ed_flt_add_btn': '+ Engadir filtro',
        'ed_flt_clear_btn': 'Limpar filtros',
        'ed_flt_col_lbl': 'Columna:',
        'ed_flt_contains_q': " contén '{q}'",
        'ed_flt_excel_btn': 'Filtro Excel...',
        'ed_flt_frame': 'Filtrar liñas de factura (por estado e por unha ou varias columnas estilo Excel)',
        'ed_flt_in_list': '{col} en [{items}]',
        'ed_flt_none': 'Sen filtros activos (clic en calquera cabeceira ▼ para filtrar por unha ou máis '
                       'columnas estilo Excel).',
        'ed_flt_op_blank': 'Baleiras',
        'ed_flt_op_contains': 'Contén / Igual',
        'ed_flt_op_exact': 'Igual exacto',
        'ed_flt_op_nblank': 'Non baleiras',
        'ed_flt_op_neq': 'Distinto de',
        'ed_flt_st_all': 'Todas as liñas',
        'ed_flt_st_edit': 'Só liñas modificadas manualmente',
        'ed_flt_st_ok': 'Só liñas correctas (con PO Line e sen avisos)',
        'ed_flt_st_prob': 'Só liñas con problemas (sen PO Line ou con avisos)',
        'ed_flt_st_unm': 'Só liñas sen PO Line (pendentes)',
        'ed_flt_state_lbl': 'Estado:',
        'ed_flt_tag': 'Filtro',
        'ed_flt_val_blank': '(Baleiras)',
        'ed_flt_val_lbl': 'Valor:',
        'ed_flt_val_nblank': '(Non baleiras)',
        'ed_flt_values_cnt': '{col} ({count} valores)',
        'ed_hdr_frame': 'Cabeceira de factura (INV)',
        'ed_load_draft_btn': 'Cargar borrador...',
        'ed_mode_orig': 'Orixinal',
        'ed_mode_prefix': 'Prefixo',
        'ed_mode_suffix': 'Sufixo',
        'ed_num_frame': 'Número de factura no Excel de Alma',
        'ed_orig_lbl': 'Orixinal:',
        'ed_refresh_btn': 'Actualizar borrador',
        'ed_reset_inv_btn': 'Restablecer factura',
        'ed_save_draft_btn': 'Gardar borrador...',
        'ed_save_excel_btn': 'Gardar cambios e actualizar Excel',
        'ed_search_pol_btn': 'Buscar PO Line en Alma...',
        'ed_stats_empty': 'Mostrando: 0/0 | Correctas: 0 | Con problemas: 0',
        'ed_stats_fmt': 'Mostrando: {shown}/{total} | Correctas: {ok} | Con problemas: {prob} (Sen POL: '
                        '{unm})',
        'ed_tb_bulk': 'Asignar valor ás seleccionadas...',
        'ed_tb_clear': 'Baleirar cela (Supr)',
        'ed_tb_copy': 'Copiar (Ctrl+C)',
        'ed_tb_fill_blanks': '⬇ Encher brancos (■)',
        'ed_tb_fill_down': '⬇ Encher abaixo (Ctrl+D)',
        'ed_tb_paste': 'Pegar (Ctrl+V)',
        'ed_tb_redo': '↪ Refacer (Ctrl+Y)',
        'ed_tb_restore': 'Restablecer cela',
        'ed_tb_search_lbl': 'Buscar (Ctrl+F):',
        'ed_tb_undo': '↩ Desfacer (Ctrl+Z)',
        'ed_tot_init': 'Importe visible: 0,00 €',
        'ed_tot_sel': '   |   Selección ({count}): {base}',
        'ed_tot_vis': 'Visible ({count}): Base {base} | IVE {vat}',
        'err_conv_title': 'Erro durante a conversión',
        'err_fnf_title': 'Ficheiro non atopado',
        'err_log_hint': "{err}\n\nConsulta o arquivo de rexistro '{log}' para máis detalles.",
        'err_perm_title': 'Ficheiro bloqueado ou sen permisos',
        'err_val_title': 'Datos ou formato non válidos',
        'err_xml_sec_title': 'XML bloqueado por seguridade',
        'fd_exp_vendors': 'Exportar provedores (NIF / CIF e Vendor Account) a Excel',
        'fd_ft_all': 'Todos os arquivos',
        'fd_ft_compat_inv': 'Facturas compatibles',
        'fd_ft_draft': 'Borrador FacturaeToAlma',
        'fd_ft_excel': 'Arquivos Excel',
        'fd_ft_img_compat': 'Imaxes compatibles',
        'fd_ft_jpg': 'Imaxe JPEG',
        'fd_ft_png': 'Imaxe PNG',
        'fd_ft_txt': 'Arquivos TXT',
        'fd_ft_xml': 'Arquivos XML',
        'fd_ft_xsig': 'Arquivos XSIG',
        'fd_imp_vendors': 'Importar provedores (NIF / CIF e Vendor Account) desde Excel',
        'fd_load_draft': 'Cargar borrador de sesión FacturaeToAlma',
        'fd_save_draft': 'Gardar borrador de sesión FacturaeToAlma',
        'fd_sel_alma': 'Seleccionar ficheiro Excel de Alma Analytics',
        'fd_sel_folder': 'Seleccionar cartafol',
        'fd_sel_inv_batch': 'Seleccionar facturas do mesmo provedor e biblioteca',
        'fd_sel_inv_face': 'Seleccionar factura electrónica FACe',
        'fd_sel_inv_single': 'Seleccionar factura XSIG / XML / TXT',
        'fd_sel_logo': 'Seleccionar logo institucional',
        'fd_sel_out': 'Definir arquivo Excel de saída',
        'fd_sel_tpl': 'Seleccionar modelo Alma .xlsx',
        'gctx_copy': 'Copiar',
        'gctx_copy_txt': 'Copiar texto',
        'gctx_cut': 'Cortar',
        'gctx_go_edit': 'Ir a Editar',
        'gctx_go_vis': 'Ir a Visualizar',
        'gctx_paste': 'Pegar',
        'gctx_sel_all': 'Seleccionar todo',
        'hlp_body_5_0': 'Factura individual: selecciona unha factura, o informe de Alma Analytics, o modelo '
                        'e o destino.\n'
                        '\n'
                        'Modo lote: todas as facturas deben pertencer ao mesmo provedor e biblioteca. A '
                        'lista permite engadir, quitar, ordenar e abrir cada factura con dobre clic en '
                        'Editar.\n'
                        '\n'
                        'Edición e filtrado: revisa a cabeceira (INV) e as liñas (IL), busca PO Lines no '
                        'catálogo de Alma Analytics, aplica autofiltros estilo Excel ou garda/carga '
                        'borradores (.ftalma.json).\n'
                        '\n'
                        'Provedores e idiomas: en Personalizar podes cambiar o idioma oficial/cooficial (es, '
                        'ca, va, gl, eu, oc) e xestionar, importar ou exportar en Excel a táboa de '
                        'provedores (NIF / CIF e Vendor Account) en modo incremental, só novos ou carga '
                        'completa.\n'
                        '\n'
                        'Informe de validación: xérase cando faltan PO Lines, hai coincidencias ambiguas, '
                        'duplicados, descadres contables ou diferenzas entre Quantity e Quantity for '
                        'Pricing.\n'
                        '\n'
                        'Limitacións confirmadas por Ex Libris: o modelo Excel non permite indicar Line '
                        'Exclusive nin declarar explicitamente facturación parcial ou completa.\n'
                        '\n'
                        'Licenza: GPL-3.0-only.',
        'hlp_body_5_1': 'Factura individual: selecciona unha factura, o informe de Alma Analytics, o modelo '
                        'e o destino.\n'
                        '\n'
                        'Modo lote: todas as facturas deben pertencer ao mesmo provedor e biblioteca. A '
                        'lista permite engadir, quitar, ordenar e abrir cada factura con dobre clic en '
                        'Editar.\n'
                        '\n'
                        'Importes e IVE compatibles con Liña exclusiva (lóxica exacta 4.0): a extracción e '
                        'exportación de importes e IVE reproducen exactamente o comportamento da versión 4.0 '
                        'para que o cambio manual a «Liña exclusiva» en Alma cadre sempre.\n'
                        '\n'
                        'Provedores e idiomas: en Personalizar podes seleccionar calquera das linguas '
                        'oficiais e cooficiais de España (es, ca, va, gl, eu, oc) e importar ou exportar en '
                        'Excel a táboa de provedores (só columnas NIF / CIF e Vendor Account) en modo '
                        'incremental/acumulativo, só engadir novos ou carga completa.\n'
                        '\n'
                        'Informe de validación: xérase cando faltan PO Lines, hai coincidencias ambiguas, '
                        'duplicados ou diferenzas entre Quantity e Quantity for Pricing.\n'
                        '\n'
                        'Limitacións confirmadas por Ex Libris: o modelo Excel non permite indicar Line '
                        'Exclusive nin declarar explicitamente facturación parcial ou completa.\n'
                        '\n'
                        'Licenza: GPL-3.0-only.',
        'hlp_guide_btn': 'Abrir Biblioguía / LibGuides',
        'hlp_links_frame': 'Documentación e soporte',
        'hlp_repo_btn': 'Abrir repositorio de GitHub',
        'match_msg_ambig': 'Varias coincidencias posibles por título',
        'match_msg_dup': 'Duplicado por {mtype}. Revisión manual necesaria.',
        'match_msg_manual': 'PO Line asignada manualmente',
        'match_msg_none': 'Sen coincidencia en Alma por ISBN, ISSN nin título',
        'match_msg_special': 'Liña de tipo {ltype} (non require PO Line bibliográfica)',
        'match_msg_unique': 'Coincidencia única por {mtype}',
        'mb_batch_clr_ask': 'Queres eliminar todas as facturas do lote?',
        'mb_batch_clr_title': 'Baleirar lote',
        'mb_batch_dup_num': 'Número de factura duplicado {num}: {name}',
        'mb_batch_none_added': 'Non se incorporaron facturas novas.',
        'mb_batch_read_err': 'Non se puido ler {name}: {err}',
        'mb_batch_rep_path': 'Ruta repetida: {name}',
        'mb_batch_sel_one': 'Selecciona polo menos unha factura da lista.',
        'mb_batch_title': 'Lote',
        'mb_batch_warn_title': 'Facturas non incorporadas',
        'mb_clr_alma_ask': 'Queres limpar tamén o ficheiro Excel de Alma Analytics?\n'
                           '\n'
                           'O modelo Alma conservarase.',
        'mb_clr_alma_title': 'Limpar ficheiro Excel de Alma',
        'mb_conv_close_ask': 'Hai unha conversión en curso. Queres pechar a aplicación de todos os xeitos?',
        'mb_conv_ok_title': 'Conversión completada',
        'mb_conv_running_msg': 'Xa hai unha conversión en curso.',
        'mb_conv_running_title': 'Conversión en curso',
        'mb_conv_warn_title': 'Incidencias na conversión',
        'mb_draft_loaded_msg': 'Restauráronse as facturas e os cambios do borrador.{warn}',
        'mb_draft_loaded_title': 'Borrador cargado',
        'mb_draft_missing_warn': '\n'
                                 '\n'
                                 'Aviso: non se atoparon algunhas facturas movidas ou eliminadas:\n'
                                 '- {files}',
        'mb_draft_outdated_msg': 'A factura, Analytics ou o modelo cambiaron despois de gardar. Volve a '
                                 'Editar, revisa e garda de novo.',
        'mb_draft_outdated_title': 'Borrador desactualizado',
        'mb_draft_rebuilt_alma': 'Reconstruíuse o borrador do Excel final de Alma coa factura e o informe de '
                                 'Analytics seleccionados actualmente.',
        'mb_draft_rebuilt_no_alma': 'Cargáronse as liñas da factura no editor.\n'
                                    '\n'
                                    'Se seleccionas tamén o ficheiro Excel de Alma Analytics, cruzaranse '
                                    'automaticamente as PO Lines, os Reporting Codes e os fondos.',
        'mb_draft_saved_msg': 'Gardouse o progreso actual en:\n'
                              '{path}\n'
                              '\n'
                              "Poderás retomar a edición máis tarde con 'Cargar borrador...'.",
        'mb_draft_saved_title': 'Borrador gardado',
        'mb_draft_unavail_msg': 'Non se puido preparar o borrador editable antes de converter.',
        'mb_draft_unavail_title': 'Borrador non dispoñible',
        'mb_empty_src_msg': 'A cela activa está baleira. Escribe ou selecciona unha cela con valor antes de '
                            'encher os brancos.',
        'mb_empty_src_title': 'Cela de orixe baleira',
        'mb_err_folder_msg': 'Non se puido abrir o cartafol:\n{path}\n\n{err}',
        'mb_err_folder_title': 'Erro ao abrir o cartafol',
        'mb_err_load_draft': 'Erro ao cargar o borrador',
        'mb_err_load_draft_msg': 'Non se puido cargar o borrador:\n{err}',
        'mb_err_prep_draft': 'Non se puido preparar o borrador',
        'mb_err_read_inv_msg': 'Non se puido cargar a factura seleccionada:\n{err}',
        'mb_err_read_inv_title': 'Erro ao ler a factura',
        'mb_err_reset_title': 'Erro ao restablecer',
        'mb_err_save_draft': 'Erro ao gardar o borrador',
        'mb_err_save_draft_msg': 'Non se puido gardar o borrador:\n{err}',
        'mb_err_save_title': 'Erro ao gardar',
        'mb_err_url_msg': 'Non se puido abrir a ligazón:\n{url}\n\n{err}',
        'mb_err_url_title': 'Erro ao abrir a ligazón',
        'mb_flt_cols_no_inv': 'Selecciona primeiro unha factura na lapela Editar.',
        'mb_flt_cols_title': 'Filtrar columnas',
        'mb_inv_invalid_title': 'Factura non válida',
        'mb_invalid_date_msg': 'Usa unha data válida, por exemplo 2026-10-19.',
        'mb_invalid_date_title': 'Data non válida',
        'mb_invalid_inputs_title': 'Datos de entrada non válidos',
        'mb_invalid_ltype': 'Selecciona un Line type permitido por Alma.',
        'mb_invalid_num': '{col} debe ser numérico.',
        'mb_invalid_val_title': 'Valor non válido',
        'mb_logo_invalid_msg': 'Non se puido cargar o logo:\n{path}\n\n{err}',
        'mb_logo_invalid_title': 'Logo non válido',
        'mb_logo_unavail_msg': 'Non se atopou o arquivo de logo:\n{path}',
        'mb_logo_unavail_title': 'Logo non dispoñible',
        'mb_no_active_inv': 'Non hai ningunha factura activa.',
        'mb_no_edit_inv': 'Non hai ningunha factura activa no editor.',
        'mb_no_out_dir_def': 'Non hai un cartafol de saída definido.',
        'mb_no_out_dir_found': 'Non se atopou o cartafol de saída.',
        'mb_no_sugg_msg': 'Non hai valores de contexto para suxerir nesta columna.',
        'mb_no_sugg_title': 'Sen suxestións',
        'mb_pol_ask_alma': 'Este buscador consulta as liñas de pedido (PO Lines) dentro do ficheiro Excel '
                           'exportado de Alma Analytics.\n'
                           '\n'
                           'Aínda non seleccionaches ningún ficheiro Excel de Alma Analytics.\n'
                           'Queres seleccionalo agora?',
        'mb_pol_no_inv': 'Selecciona primeiro unha factura (ou lote de facturas) na lapela Converter.',
        'mb_pol_no_lines': 'Non hai liñas visibles no editor.',
        'mb_pol_sel_row': 'Selecciona unha fila da lista.',
        'mb_pol_title': 'Buscar PO Line en Alma',
        'mb_pref_invalid_dir': '{label} non é un cartafol válido:\n{val}',
        'mb_pref_invalid_title': 'Preferencias non válidas',
        'mb_pref_reset_ask': 'Queres restablecer a institución, o idioma e todos os directorios '
                             'personalizados?',
        'mb_pref_reset_ok': 'Restablecéronse os valores iniciais.',
        'mb_pref_reset_title': 'Restablecer preferencias',
        'mb_pref_saved_msg': 'As preferencias gardáronse correctamente.',
        'mb_prot_drag_msg': 'Esta columna está protexida e non pode encherse arrastrando.',
        'mb_prot_field_title': 'Campo protexido',
        'mb_prot_generic_msg': 'Este campo non é editable.',
        'mb_prot_hdr_msg': 'O número, a data e os importes de cabeceira non se editan directamente. O número '
                           'só admite prefixo ou sufixo.',
        'mb_prot_line_msg': 'Este campo procede da factura orixinal e non é editable en FacturaeToAlma.',
        'mb_reset_inv_ask': 'Queres descartar todos os cambios desta factura?',
        'mb_reset_inv_title': 'Restablecer factura',
        'mb_save_changes_ask': 'Queres gardar todos os cambios e actualizar agora o Excel de saída?\n'
                               '\n'
                               'Sobrescribirase o Excel indicado en Converter. O XML orixinal non se '
                               'modificará.',
        'mb_save_changes_title': 'Gardar cambios',
        'mb_save_draft_empty': 'Non hai ningunha factura cargada para gardar nun borrador.',
        'mb_save_draft_title': 'Gardar borrador',
        'mb_unsaved_ed_ask': 'Hai cambios realizados en Editar (número de factura, cabeceira ou liñas) que '
                             'aínda non se gardaran.\n'
                             '\n'
                             'Queres aplicar e gardar estes cambios para converter agora?',
        'mb_unsaved_ed_title': 'Cambios pendentes en Editar',
        'mb_vend_clr_ask': 'Queres eliminar todas as asociacións NIF → Vendor Account gardadas?',
        'mb_vend_clr_title': 'Baleirar provedores lembrados',
        'mb_vend_exp_err_title': 'Erro ao exportar provedores',
        'mb_vend_exp_ok_msg': 'Exportáronse {count} provedores (columnas NIF / CIF e Vendor Account) ao '
                              'arquivo:\n'
                              '{path}',
        'mb_vend_exp_ok_title': 'Exportación de provedores completada',
        'mb_vend_full_ask': 'O modo «Carga completa» substituirá os {count} provedores gardados actualmente '
                            'polos do arquivo Excel seleccionado.\n'
                            '\n'
                            'Desexas continuar?',
        'mb_vend_imp_err_title': 'Erro ao importar provedores',
        'mb_vend_imp_ok_msg': 'Modo aplicado: {mode}\n'
                              '\n'
                              '- Novos engadidos: {added}\n'
                              '- Existentes actualizados: {updated}\n'
                              '- Sen cambios / omitidos: {skipped}\n'
                              '- Filas incompletas ignoradas: {invalid}\n'
                              '\n'
                              'Total de provedores gardados: {total}',
        'mb_vend_imp_ok_title': 'Importación de provedores completada',
        'mb_vend_missing': 'Indica tanto o NIF/CIF como o código Vendor Account de Alma.',
        'mb_vend_sel_del': 'Selecciona unha fila da táboa para eliminala.',
        'mb_vend_title': 'Provedores lembrados',
        'mb_warn_title': 'Aviso',
        'numdlg_apply': 'Aplicar nesta sesión',
        'numdlg_desc': 'É probable que o número pertenza a unha secuencia que se reinicia periodicamente.\n'
                       'Podes engadir temporalmente un prefixo ou sufixo para desambiguar.\n'
                       'O XML orixinal non se modificará.',
        'numdlg_orig': 'Número orixinal: {original}',
        'numdlg_title': 'Corrixir número de factura en memoria',
        'poldlg_all_btn': 'Ver todos',
        'poldlg_assign_full': 'Asignar PO Line + códigos e fondos de Analytics',
        'poldlg_assign_only': 'Asignar só PO Line',
        'poldlg_col_score': 'Afinidade',
        'poldlg_dblclick_hint': 'O dobre clic asigna unicamente a PO Line (sen alterar Reporting Codes nin '
                                'fondos).',
        'poldlg_inv_title': 'Título na factura: {title}',
        'poldlg_line_frame': 'Liña de factura #{line}',
        'poldlg_search_btn': 'Buscar',
        'poldlg_search_lbl': 'Buscar en Analytics (Título / ISBN / ISSN / PO Line):',
        'poldlg_summary': 'ISBN: {isbn}   |   ISSN: {issn}   |   Prezo: {price}   |   PO Line actual: {pol}',
        'poldlg_title': 'Buscar PO Line no catálogo de Alma - Liña {line}',
        'poldlg_unassigned': '(sen asignar)',
        'st_alma_selected': 'Estado: ficheiro Excel de Alma Analytics seleccionado.',
        'st_batch_count': 'Estado: {count} facturas no lote.',
        'st_batch_empty': 'Estado: lote baleiro.',
        'st_batch_sorted': 'Estado: lote ordenado por número de factura.',
        'st_check_inputs': 'Estado: revisa os datos de entrada.',
        'st_cleared': 'Estado: datos limpados. O modelo Alma consérvase.',
        'st_conv_err': 'Estado: erro durante a conversión.',
        'st_conv_ok': 'Estado: conversión completada sen incidencias.',
        'st_conv_warn': 'Estado: conversión completada con incidencias.',
        'st_converting': 'Estado: convertendo en segundo plano...',
        'st_copied': "Estado: copiado o valor '{val}' de {col}.",
        'st_draft_loaded': "Estado: borrador '{name}' cargado correctamente.",
        'st_draft_ready_alma': 'Estado: borrador editable do Excel de Alma preparado (cruzado con Alma '
                               'Analytics).',
        'st_draft_ready_no_alma': 'Estado: factura cargada en Editar (selecciona o Excel de Alma Analytics '
                                  'para cruzar PO Lines).',
        'st_draft_saved': "Estado: borrador gardado en '{name}'.",
        'st_fill_blanks': "Estado: enchidas {count} celas en branco inferiores con '{val}'.",
        'st_fill_drag': "Estado: enchidas {count} celas en '{col}' con '{val}'.",
        'st_inv_selected': 'Estado: factura {number} seleccionada. Provedor: {vendor}.',
        'st_lang_changed': 'Estado: idioma cambiado a {lang}.',
        'st_mode_batch': 'Estado: modo lote. Selecciona facturas do mesmo provedor e biblioteca.',
        'st_mode_single': 'Estado: modo factura individual.',
        'st_no_redo': 'Estado: non hai accións para refacer.',
        'st_no_undo': 'Estado: non hai accións para desfacer.',
        'st_num_updated': "Estado: número de factura actualizado no borrador a '{number}'.",
        'st_open_in_edit': "Estado: aberta a factura '{name}' en Editar.",
        'st_out_selected': 'Estado: arquivo de saída definido.',
        'st_pending': 'Estado: pendente de seleccionar arquivos.',
        'st_redone': "Estado: refeita a acción '{action}'.",
        'st_saving_edits': 'Estado: gardados {n_num} cambios de número, {n_hdr} de cabeceira e {n_line} de '
                           'liña; actualizando o Excel de saída...',
        'st_tpl_selected': 'Estado: modelo Alma seleccionado.',
        'st_undone': "Estado: desfeita a acción '{action}'.",
        'st_vendor_cleared': 'Estado: lista de provedores lembrados baleirada.',
        'st_vendor_deleted': "Estado: eliminado o provedor lembrado para o NIF '{nif}'.",
        'st_vendor_exported': "Estado: exportados {count} provedores a '{name}'.",
        'st_vendor_imported': "Estado: provedores importados desde '{name}' (total: {total}).",
        'st_vendor_saved': 'Estado: provedor gardado ({nif} → {vendor}).',
        'sum_body': '\n'
                    '\n'
                    'Resumo de liñas:\n'
                    '- Liñas procesadas: {total_lines}\n'
                    '- Coincidencias por ISBN: {match_isbn}\n'
                    '- Coincidencias por ISSN: {match_issn}\n'
                    '- Coincidencias por título: {match_title}\n'
                    '{extra}- Duplicados ISBN: {duplicate_isbn}\n'
                    '- Duplicados ISSN: {duplicate_issn}\n'
                    '- Títulos ambiguos: {ambiguous_title}\n'
                    '- Sen coincidencia: {no_match}\n'
                    '- Liñas sen PO Line: {lines_without_po_line}',
        'sum_pol_man': '- Asignacións manuais de PO Line: {n}\n',
        'sum_pol_xml': '- Coincidencias por PO Line no XML: {n}\n',
        'tab_convert': 'Converter',
        'tab_customize': 'Personalizar',
        'tab_edit': 'Editar',
        'tab_help': 'Axuda',
        'tab_visualize': 'Visualizar',
        'val_bad_draft': 'O ficheiro non ten un formato de borrador válido de FacturaeToAlma.',
        'val_bad_ext': 'Extensión de factura non válida: {path}',
        'val_draft_no_files': 'Ningunha das facturas referenciadas no borrador existe xa na súa ruta '
                              'orixinal.',
        'val_dup_numbers': 'Hai números de factura duplicados no lote: {nums}',
        'val_multi_vendor': 'O lote contén facturas de varios provedores. Selecciona unicamente facturas do '
                            'mesmo provedor.\n'
                            '\n'
                            'Provedores detectados:\n'
                            '{details}',
        'val_need_alma': 'Debes seleccionar o ficheiro Excel de Alma Analytics.',
        'val_need_inv': 'Debes seleccionar unha factura ou un lote de facturas.',
        'val_need_inv_edit': 'Debes seleccionar primeiro unha factura ou un lote de facturas na lapela '
                             'Converter.',
        'val_need_out': 'Debes definir o Excel de saída.',
        'val_need_tpl': 'Debes seleccionar o modelo Alma.',
        'val_no_alma_file': 'Non existe o ficheiro Excel de Alma Analytics: {path}',
        'val_no_inv_file': 'Non existe a factura: {path}',
        'val_no_tpl_file': 'Non existe o modelo Alma: {path}',
        'val_out_collision': 'O Excel de saída non pode coincidir coa factura, o Analytics nin o modelo. '
                             'Escolle outro nome ou cartafol.',
        'val_single_inv_only': 'O ficheiro debe conter unha soa factura e contén {count}.',
        'val_vend_empty_xlsx': 'O arquivo Excel está baleiro.',
        'val_vend_po_cols': 'O arquivo contén columnas de liña de pedido ({cols}). O Excel de provedores '
                            'debe ter unicamente as 2 columnas «NIF / CIF» e «Vendor Account».',
        'val_vend_two_cols': 'O Excel de provedores debe conter exactamente as 2 columnas «NIF / CIF» e '
                             '«Vendor Account» (sen campos de liña de pedido).',
        'vis_footer_note': 'Vista de consulta da factura orixinal. Os cambios do Excel de Alma realízanse na '
                           'lapela Editar.',
        'vis_inv_lbl': 'Factura:',
        'vw_addr': 'Enderezo: {addr}',
        'vw_card_buyer': ' DATOS DO RECEPTOR ',
        'vw_card_dir3': ' CENTROS ADMINISTRATIVOS FACe (DIR3) ',
        'vw_card_lines': ' DETALLE ECONÓMICO (LIÑAS) ',
        'vw_card_pay': ' DATOS DE PAGAMENTO ',
        'vw_card_seller': ' DATOS DO EMISOR ',
        'vw_col_code': 'Artigo / ISBN',
        'vw_col_desc': 'Descrición / Título',
        'vw_col_gross': 'Importe Bruto',
        'vw_col_price': 'Prezo Unit.',
        'vw_col_qty': 'Cant.',
        'vw_col_tax': 'IVE Repercutido',
        'vw_contact': 'Contacto: {contact}',
        'vw_elec_inv_lbl': 'Factura electrónica:',
        'vw_hdr_invoice': 'FACTURA: {number}',
        'vw_hdr_meta': 'Serie: {series}   |   Data de emisión: {issue}   |   Data de operación: {op}',
        'vw_load_btn': 'Cargar factura...',
        'vw_nif': 'CIF/NIF: {nif}',
        'vw_no_number': '(sen número)',
        'vw_not_specified': 'Non especificado',
        'vw_pay_due': 'Data de vencemento: {date}',
        'vw_pay_iban': 'Conta / IBAN: {iban}',
        'vw_pay_means': 'Medio de pagamento: Transferencia (Código {code})',
        'vw_placeholder': 'Seleccione un arquivo de factura electrónica (.xsig, .xml, .txt) para visualizar '
                          'os seus datos.',
        'vw_role_01': 'Oficina Contable',
        'vw_role_02': 'Órgano Xestor',
        'vw_role_03': 'Unidade Tramitadora',
        'vw_role_04': 'Órgano Propoñente',
        'vw_role_other': 'Rol {code}',
        'vw_tab_pdf': '  Vista Factura (Estilo PDF)  ',
        'vw_tab_xml': '  Código XML (FACe)  ',
        'vw_tot_base': 'Base Impoñible:',
        'vw_tot_grand': 'TOTAL FACTURA:',
        'vw_tot_payable': 'Total a Pagar:',
        'vw_tot_ret': 'Impostos Retidos:',
        'vw_tot_taxes': 'Impostos Repercutidos:',
        'vw_unknown_vendor': 'Provedor non identificado'},
 'oc': {'afdlg_apply_next': 'Aplicar e contunhar filtrant ua auta colomna',
        'afdlg_blank_chk': '(Vuedes / Sense valor)',
        'afdlg_check_all': 'Marcar toti',
        'afdlg_col_lbl': 'Colomna a filtrar:',
        'afdlg_list_frame': 'Valors presentes ena colomna (marca es que vòs mostrar)',
        'afdlg_only_blank': 'Sonque vuedes',
        'afdlg_only_val': 'Sonque damb valor',
        'afdlg_remove_col': "Trèir filtre d'aguesta colomna",
        'afdlg_search_lbl': 'Cercar tèxte / valor:',
        'afdlg_title': 'Filtre autofiltre per colomnes (estil Excel)',
        'afdlg_uncheck_all': 'Desmarcar toti',
        'bulkdlg_lbl': 'Valor entà {col}:',
        'bulkdlg_title': 'Assignar {col}',
        'conv_actions_frame': '4. Accions',
        'conv_alma_btn': 'Seleccionar archiu...',
        'conv_alma_lbl': "Archiu Excel d'Alma Analytics damb PO Lines *",
        'conv_batch_add': 'Ajustar factures',
        'conv_batch_clear': 'Vuedar lòt',
        'conv_batch_count': '{count} factures seleccionades',
        'conv_batch_frame': 'Factures deth lòt',
        'conv_batch_hint': '(Doble clic en ua factura entà daurir-la en Editar)',
        'conv_batch_remove': 'Trèir seleccionades',
        'conv_batch_sort': 'Ordenar per numèro',
        'conv_clear_btn': 'Netejar donades',
        'conv_col_file': 'Archiu',
        'conv_col_invoice': 'Numèro de factura',
        'conv_col_vendor': 'Provedidor',
        'conv_exit_btn': 'Gésser',
        'conv_inputs_frame': "2. Archius d'entrada",
        'conv_inv_batch_btn': 'Seleccionar factures...',
        'conv_inv_batch_lbl': 'Factures XSIG / XML / TXT deth madeish provedidor e bibliotèca *',
        'conv_inv_single_btn': 'Seleccionar factura...',
        'conv_inv_single_lbl': 'Factura electronica XSIG / XML / TXT *',
        'conv_mode_batch': 'Lòt de factures deth madeish provedidor e bibliotèca',
        'conv_mode_frame': '1. Mòde de conversion',
        'conv_mode_single': 'Factura individuau',
        'conv_ok_heading': 'CONVERSION COMPLÈTA SENSE INCIDÉNCIES',
        'conv_open_out_btn': 'Daurir dorsièr de gessuda',
        'conv_output_btn': 'Escuélher destinacion...',
        'conv_output_frame': '3. Archiu de gessuda',
        'conv_result_body': '{heading}\n'
                            '\n'
                            'Factures processades: {total_inv}\n'
                            'Auisi de validacion: {total_iss}\n'
                            '\n'
                            'Excel Alma generat:\n'
                            '{out}',
        'conv_run_btn': 'Convertir a format Alma',
        'conv_tpl_btn': 'Seleccionar modèl...',
        'conv_tpl_lbl': 'Modèl Alma .xlsx *',
        'conv_unmatched_warn': '\n'
                               '\n'
                               'Atencion: i a linhes sense PO Line. Revise er informe abans de cargar en '
                               'Alma.',
        'conv_val_report': '\n\nInforme de validacion:\n{path}',
        'conv_warn_heading': 'INCIDÉNCIES ENA CONVERSION',
        'ctx_bulk_apply': 'Aplicar valor as linhes seleccionades...',
        'ctx_clear_cell': 'Vuedar cèlla (Supr)',
        'ctx_edit_cell': 'Editar cèlla',
        'ctx_fill_blanks_down': 'Aumplir blancs inferiors (■)',
        'ctx_fill_blanks_sel': 'Aumplir blancs seleccionadi',
        'ctx_fill_down': 'Aumplir cap avath (Ctrl+D)',
        'ctx_flt_cell_val': "Filtrar pera valor d'aguesta cèlla",
        'ctx_flt_clear_all': 'Trèir toti es filtres',
        'ctx_flt_col_excel': 'Filtrar aguesta colomna (estil Excel)...',
        'ctx_redo': 'Tornar a hèr (Ctrl+Y)',
        'ctx_restore_cell': 'Restablir cèlla',
        'ctx_restore_orig': 'Restablir cèlla originau',
        'ctx_search_pol': "Cercar PO Line en catalòg d'Alma...",
        'ctx_suggest_ctx': 'Suggerir valor peth contèxte',
        'ctx_undo': 'Desfer (Ctrl+Z)',
        'cust_choose_dir_btn': 'Seleccionar dorsièr...',
        'cust_dir_alma': "Directòri der archiu Excel d'Alma",
        'cust_dir_inv': 'Directòri de factures FACe',
        'cust_dir_out': "Directòri d'Excel finaus",
        'cust_dir_tpl': 'Directòri de modèl Alma',
        'cust_inst_lbl': 'Institucion',
        'cust_lang_lbl': 'Lengua dera interfàcia',
        'cust_logo_del_btn': 'Eliminar logotipe',
        'cust_logo_hint': 'PNG, JPG o JPEG. Nautor maxima: 100 px.',
        'cust_logo_invalid': 'Er archiu seleccionat non ei ua imatge valida.',
        'cust_logo_lbl': 'Logotipe institucionau',
        'cust_logo_loaded': 'Logotipe cargat: {name}',
        'cust_logo_none': 'Sense logotipe. PNG, JPG o JPEG; nautor maxima: 100 px.',
        'cust_logo_sel_btn': 'Seleccionar logotipe...',
        'cust_logo_unavail': 'Eth logotipe configurat non ei disponible.',
        'cust_pref_frame': 'Preferéncies',
        'cust_pref_note': "Toti es camps son opcionaus. Se un dorsièr personalizat demore vued, s'emplegue "
                          "eth darrèr dorsièr valid; ena prumèra execucion s'emplegue eth dorsièr dera "
                          'aplicacion. Era gessuda se sauve ath costat dera factura, sauv que se definisque '
                          'un directòri de gessuda.',
        'cust_reset_pref_btn': 'Restablir valors',
        'cust_save_pref_btn': 'Sauvar preferéncies',
        'cust_vend_acc_lbl': 'Vendor Account en Alma:',
        'cust_vend_add_btn': 'Ajustar / Actualizar',
        'cust_vend_clr_btn': 'Vuedar lista',
        'cust_vend_col_acc': 'Vendor Account en Alma',
        'cust_vend_col_nif': 'NIF / CIF normalizat',
        'cust_vend_del_btn': 'Eliminar seleccionat',
        'cust_vend_exp_btn': 'Exportar Excel...',
        'cust_vend_frame': 'Provedidors rebrembadi automaticament (NIF / CIF → Vendor Account en Alma)',
        'cust_vend_imp_btn': 'Importar Excel...',
        'cust_vend_imp_mode_lbl': "Mòde d'importacion Excel:",
        'cust_vend_mode_full': 'Carga complèta (remplaçar lista actuau)',
        'cust_vend_mode_inc': 'Incrementau / acumulatiu (ajustar e actualizar)',
        'cust_vend_mode_new': 'Sonque ajustar naui (non susescríuer)',
        'cust_vend_nif_lbl': 'NIF / CIF:',
        'dlg_accept': 'Acceptar',
        'dlg_apply': 'Aplicar',
        'dlg_cancel': 'Anullar',
        'ed_apply_num_btn': 'Aplicar numèro',
        'ed_bottom_hint': 'Escriu dirècte o doble clic entà editar | Arrossegue ■ entà aumplir | Ctrl+F '
                          'cercar | Clic en capçalères ▼ entà filtrar.',
        'ed_col_invoice': 'Factura',
        'ed_col_line': 'Linha',
        'ed_diag_edited': '✎ Linha {line}: modificada manuaument en aguesta session.',
        'ed_diag_init': 'Selecciona ua linha entà véder eth sòn diagnostic de crotzament e validacion.',
        'ed_diag_manual': "✓ Linha {line}: PO Line '{pol}' assignada manuaument.",
        'ed_diag_matched': '✓ Linha {line}: {msg}.',
        'ed_diag_no_inv': 'Selecciona ua factura entà revisar es sues linhes e auisi.',
        'ed_diag_ready': '✓ Linha {line}: prèsta entà exportar.',
        'ed_diag_recon': '⚠ Quadrament factura: {msg}',
        'ed_diag_unmatched': "⚠ Linha {line}: sense PO Line assignada (emplega 'Cercar PO Line en Alma...' o "
                             'escriu-la ena cèlla).',
        'ed_diag_warning': '⚠ Linha {line}: {notes}',
        'ed_flt_active_prefix': 'Filtres actius -> ',
        'ed_flt_add_btn': '+ Ajustar filtre',
        'ed_flt_clear_btn': 'Netejar filtres',
        'ed_flt_col_lbl': 'Colomna:',
        'ed_flt_contains_q': " conten '{q}'",
        'ed_flt_excel_btn': 'Filtre Excel...',
        'ed_flt_frame': 'Filtrar linhes de factura (per estat e per ua o diuèrses colomnes estil Excel)',
        'ed_flt_in_list': '{col} en [{items}]',
        'ed_flt_none': 'Sense filtres actius (clic en quinsevolh capçalèra ▼ entà filtrar per ua o mès '
                       'colomnes estil Excel).',
        'ed_flt_op_blank': 'Vuedes',
        'ed_flt_op_contains': 'Conten / Parièr',
        'ed_flt_op_exact': 'Parièr exacte',
        'ed_flt_op_nblank': 'Non vuedes',
        'ed_flt_op_neq': 'Diferent de',
        'ed_flt_st_all': 'Totes es linhes',
        'ed_flt_st_edit': 'Sonque linhes modificades manuaument',
        'ed_flt_st_ok': 'Sonque linhes corrèctes (damb PO Line e sense auisi)',
        'ed_flt_st_prob': 'Sonque linhes damb problèmas (sense PO Line o damb auisi)',
        'ed_flt_st_unm': 'Sonque linhes sense PO Line (pendentes)',
        'ed_flt_state_lbl': 'Estat:',
        'ed_flt_tag': 'Filtre',
        'ed_flt_val_blank': '(Vuedes)',
        'ed_flt_val_lbl': 'Valor:',
        'ed_flt_val_nblank': '(Non vuedes)',
        'ed_flt_values_cnt': '{col} ({count} valors)',
        'ed_hdr_frame': 'Capçalèra de factura (INV)',
        'ed_load_draft_btn': 'Cargar borrador...',
        'ed_mode_orig': 'Originau',
        'ed_mode_prefix': 'Prefix',
        'ed_mode_suffix': 'Sufix',
        'ed_num_frame': "Numèro de factura en Excel d'Alma",
        'ed_orig_lbl': 'Originau:',
        'ed_refresh_btn': 'Actualizar borrador',
        'ed_reset_inv_btn': 'Restablir factura',
        'ed_save_draft_btn': 'Sauvar borrador...',
        'ed_save_excel_btn': 'Sauvar cambis e actualizar Excel',
        'ed_search_pol_btn': 'Cercar PO Line en Alma...',
        'ed_stats_empty': 'Mostrant: 0/0 | Corrèctes: 0 | Damb problèmas: 0',
        'ed_stats_fmt': 'Mostrant: {shown}/{total} | Corrèctes: {ok} | Damb problèmas: {prob} (Sense POL: '
                        '{unm})',
        'ed_tb_bulk': 'Assignar valor as seleccionades...',
        'ed_tb_clear': 'Vuedar cèlla (Supr)',
        'ed_tb_copy': 'Copiar (Ctrl+C)',
        'ed_tb_fill_blanks': '⬇ Aumplir blancs (■)',
        'ed_tb_fill_down': '⬇ Aumplir cap avath (Ctrl+D)',
        'ed_tb_paste': 'Pegar (Ctrl+V)',
        'ed_tb_redo': '↪ Tornar a hèr (Ctrl+Y)',
        'ed_tb_restore': 'Restablir cèlla',
        'ed_tb_search_lbl': 'Cercar (Ctrl+F):',
        'ed_tb_undo': '↩ Desfer (Ctrl+Z)',
        'ed_tot_init': 'Impòrt visible: 0,00 €',
        'ed_tot_sel': '   |   Seleccion ({count}): {base}',
        'ed_tot_vis': 'Visible ({count}): Basa {base} | IVA {vat}',
        'err_conv_title': 'Error pendent era conversion',
        'err_fnf_title': 'Archiu non trobat',
        'err_log_hint': "{err}\n\nConsulta er archiu de registre '{log}' entà mès detalhs.",
        'err_perm_title': 'Archiu blocat o sense permisi',
        'err_val_title': 'Donades o format non valids',
        'err_xml_sec_title': 'XML blocat per seguretat',
        'fd_exp_vendors': 'Exportar provedidors (NIF / CIF e Vendor Account) a Excel',
        'fd_ft_all': 'Toti es archius',
        'fd_ft_compat_inv': 'Factures compatibles',
        'fd_ft_draft': 'Borrador FacturaeToAlma',
        'fd_ft_excel': 'Archius Excel',
        'fd_ft_img_compat': 'Imatges compatibles',
        'fd_ft_jpg': 'Imatge JPEG',
        'fd_ft_png': 'Imatge PNG',
        'fd_ft_txt': 'Archius TXT',
        'fd_ft_xml': 'Archius XML',
        'fd_ft_xsig': 'Archius XSIG',
        'fd_imp_vendors': "Importar provedidors (NIF / CIF e Vendor Account) des d'Excel",
        'fd_load_draft': 'Cargar borrador de session FacturaeToAlma',
        'fd_save_draft': 'Sauvar borrador de session FacturaeToAlma',
        'fd_sel_alma': "Seleccionar archiu Excel d'Alma Analytics",
        'fd_sel_folder': 'Seleccionar dorsièr',
        'fd_sel_inv_batch': 'Seleccionar factures deth madeish provedidor e bibliotèca',
        'fd_sel_inv_face': 'Seleccionar factura electronica FACe',
        'fd_sel_inv_single': 'Seleccionar factura XSIG / XML / TXT',
        'fd_sel_logo': 'Seleccionar logotipe institucionau',
        'fd_sel_out': 'Definir archiu Excel de gessuda',
        'fd_sel_tpl': 'Seleccionar modèl Alma .xlsx',
        'gctx_copy': 'Copiar',
        'gctx_copy_txt': 'Copiar tèxte',
        'gctx_cut': 'Talhar',
        'gctx_go_edit': 'Anar a Editar',
        'gctx_go_vis': 'Anar a Visualizar',
        'gctx_paste': 'Pegar',
        'gctx_sel_all': 'Seleccionar tot',
        'hlp_body_5_0': "Factura individuau: selecciona ua factura, er informe d'Alma Analytics, eth modèl e "
                        'era destinacion.\n'
                        '\n'
                        'Mòde lòt: totes es factures an de pertànher ath madeish provedidor e bibliotèca. '
                        'Era lista permet ajustar, trèir, ordenar e daurir cada factura damb doble clic en '
                        'Editar.\n'
                        '\n'
                        'Edicion e filtratge: revise era capçalèra (INV) e es linhes (IL), cèrca PO Lines en '
                        "catalòg d'Alma Analytics, aplica autofiltres estil Excel o sauve/cargue borradors "
                        '(.ftalma.json).\n'
                        '\n'
                        'Provedidors e lengües: en Personalizar pòs cambiar era lengua oficiau/cooficiau '
                        '(es, ca, va, gl, eu, oc) e gestionar, importar o exportar en Excel era taula de '
                        'provedidors (NIF / CIF e Vendor Account) en mòde incrementau, sonque naui o carga '
                        'complèta.\n'
                        '\n'
                        'Informe de validacion: se genère quan manquen PO Lines, i a coincidéncies ambigües, '
                        'duplicats, desquadraments compdables o diferéncies entre Quantity e Quantity for '
                        'Pricing.\n'
                        '\n'
                        'Limitacions confirmades per Ex Libris: eth modèl Excel non permet indicar Line '
                        'Exclusive ne declarar explicitament facturacion parciau o complèta.\n'
                        '\n'
                        'Licéncia: GPL-3.0-only.',
        'hlp_body_5_1': "Factura individuau: selecciona ua factura, er informe d'Alma Analytics, eth modèl e "
                        'era destinacion.\n'
                        '\n'
                        'Mòde lòt: totes es factures an de pertànher ath madeish provedidor e bibliotèca. '
                        'Era lista permet ajustar, trèir, ordenar e daurir cada factura damb doble clic en '
                        'Editar.\n'
                        '\n'
                        "Impòrts e IVA compatibles damb Linha exclusiua (logica exacta 4.0): er'extraccion e "
                        "er'exportacion d'impòrts e IVA reprodusissen exactament eth compòrtament dera "
                        "version 4.0 entà qu'eth cambi manuau a «Linha exclusiua» en Alma quadre tostemps.\n"
                        '\n'
                        'Provedidors e lengües: en Personalizar pòs seleccionar quinsevolh des lengües '
                        "oficiaus e cooficiaus d'Espanha (es, ca, va, gl, eu, oc) e importar o exportar en "
                        'Excel era taula de provedidors (sonque colomnes NIF / CIF e Vendor Account) en mòde '
                        'incrementau/acumulatiu, sonque ajustar naui o carga complèta.\n'
                        '\n'
                        'Informe de validacion: se genère quan manquen PO Lines, i a coincidéncies ambigües, '
                        'duplicats o diferéncies entre Quantity e Quantity for Pricing.\n'
                        '\n'
                        'Limitacions confirmades per Ex Libris: eth modèl Excel non permet indicar Line '
                        'Exclusive ne declarar explicitament facturacion parciau o complèta.\n'
                        '\n'
                        'Licéncia: GPL-3.0-only.',
        'hlp_guide_btn': 'Daurir Biblioguia / LibGuides',
        'hlp_links_frame': 'Documentacion e supòrt',
        'hlp_repo_btn': 'Daurir repositòri de GitHub',
        'match_msg_ambig': 'Diuèrses coincidéncies possibles per títol',
        'match_msg_dup': 'Duplicat per {mtype}. Revision manuau necessària.',
        'match_msg_manual': 'PO Line assignada manuaument',
        'match_msg_none': 'Sense coincidéncia en Alma per ISBN, ISSN ne títol',
        'match_msg_special': 'Linha de tipe {ltype} (non requerís PO Line bibliografica)',
        'match_msg_unique': 'Coincidéncia unica per {mtype}',
        'mb_batch_clr_ask': 'Vòs eliminar totes es factures deth lòt?',
        'mb_batch_clr_title': 'Vuedar lòt',
        'mb_batch_dup_num': 'Numèro de factura duplicat {num}: {name}',
        'mb_batch_none_added': "Non s'an incorporat factures naues.",
        'mb_batch_read_err': "Non s'a pogut léger {name}: {err}",
        'mb_batch_rep_path': ' Rota repetida: {name}',
        'mb_batch_sel_one': 'Selecciona aumens ua factura dera lista.',
        'mb_batch_title': 'Lòt',
        'mb_batch_warn_title': 'Factures non incorporades',
        'mb_clr_alma_ask': "Vòs netejar tanben er archiu Excel d'Alma Analytics?\n"
                           '\n'
                           'Eth modèl Alma se conservarà.',
        'mb_clr_alma_title': "Netejar archiu Excel d'Alma",
        'mb_conv_close_ask': "I a ua conversion en cors. Vòs barrar er'aplicacion de totes manères?",
        'mb_conv_ok_title': 'Conversion completada',
        'mb_conv_running_msg': 'Ja i a ua conversion en cors.',
        'mb_conv_running_title': 'Conversion en cors',
        'mb_conv_warn_title': 'Incidéncies ena conversion',
        'mb_draft_loaded_msg': "S'an restaurat es factures e es cambis deth borrador.{warn}",
        'mb_draft_loaded_title': 'Borrador cargat',
        'mb_draft_missing_warn': '\n'
                                 '\n'
                                 "Auís: non s'an trobat quauques factures mogudes o eliminades:\n"
                                 '- {files}',
        'mb_draft_outdated_msg': 'Era factura, Analytics o eth modèl an cambiat dempús de sauvar. Torna a '
                                 'Editar, revise e sauve de nau.',
        'mb_draft_outdated_title': 'Borrador desactualizat',
        'mb_draft_rebuilt_alma': "S'a rebastit eth borrador der Excel finau d'Alma damb era factura e er "
                                 "informe d'Analytics seleccionadi actuaument.",
        'mb_draft_rebuilt_no_alma': "S'an cargat es linhes dera factura en editor.\n"
                                    '\n'
                                    "Se selecciones tanben er archiu Excel d'Alma Analytics, se crotzaran "
                                    'automaticament es PO Lines, es Reporting Codes e es honsi.',
        'mb_draft_saved_msg': "S'a sauvat eth progrès actuau en:\n"
                              '{path}\n'
                              '\n'
                              "Poigaràs repréner er'edicion mès tard damb 'Cargar borrador...'.",
        'mb_draft_saved_title': 'Borrador sauvat',
        'mb_draft_unavail_msg': "Non s'a pogut preparar eth borrador editable abans de convertir.",
        'mb_draft_unavail_title': 'Borrador non disponible',
        'mb_empty_src_msg': 'Era cèlla actiua ei vueda. Escriu o selecciona ua cèlla damb valor abans '
                            "d'aumplir es blancs.",
        'mb_empty_src_title': "Cèlla d'origina vueda",
        'mb_err_folder_msg': "Non s'a pogut daurir eth dorsièr:\n{path}\n\n{err}",
        'mb_err_folder_title': 'Error en daurir eth dorsièr',
        'mb_err_load_draft': 'Error en cargar eth borrador',
        'mb_err_load_draft_msg': "Non s'a pogut cargar eth borrador:\n{err}",
        'mb_err_prep_draft': "Non s'a pogut preparar eth borrador",
        'mb_err_read_inv_msg': "Non s'a pogut cargar era factura seleccionada:\n{err}",
        'mb_err_read_inv_title': 'Error en léger era factura',
        'mb_err_reset_title': 'Error en restablir',
        'mb_err_save_draft': 'Error en sauvar eth borrador',
        'mb_err_save_draft_msg': "Non s'a pogut sauvar eth borrador:\n{err}",
        'mb_err_save_title': 'Error en sauvar',
        'mb_err_url_msg': "Non s'a pogut daurir eth ligam:\n{url}\n\n{err}",
        'mb_err_url_title': 'Error en daurir eth ligam',
        'mb_flt_cols_no_inv': 'Selecciona prumèr ua factura ena pestanha Editar.',
        'mb_flt_cols_title': 'Filtrar colomnes',
        'mb_inv_invalid_title': 'Factura non valida',
        'mb_invalid_date_msg': 'Emplega ua data valida, per exemple 2026-10-19.',
        'mb_invalid_date_title': 'Data non valida',
        'mb_invalid_inputs_title': "Donades d'entrada non valides",
        'mb_invalid_ltype': 'Selecciona un Line type permetut per Alma.',
        'mb_invalid_num': "{col} a d'èster numeric.",
        'mb_invalid_val_title': 'Valor non valida',
        'mb_logo_invalid_msg': "Non s'a pogut cargar eth logotipe:\n{path}\n\n{err}",
        'mb_logo_invalid_title': 'Logotipe non valid',
        'mb_logo_unavail_msg': "Non s'a trobat er archiu de logotipe:\n{path}",
        'mb_logo_unavail_title': 'Logotipe non disponible',
        'mb_no_active_inv': 'Non i a cap factura actiua.',
        'mb_no_edit_inv': 'Non i a cap factura actiua en editor.',
        'mb_no_out_dir_def': 'Non i a cap dorsièr de gessuda definit.',
        'mb_no_out_dir_found': "Non s'a trobat eth dorsièr de gessuda.",
        'mb_no_sugg_msg': 'Non i a valors de contèxte entà suggerir en aguesta colomna.',
        'mb_no_sugg_title': 'Sense suggeriments',
        'mb_pol_ask_alma': 'Aguest cercador consulte es linhes de comanda (PO Lines) laguens der archiu '
                           "Excel exportat d'Alma Analytics.\n"
                           '\n'
                           "Encara non as seleccionat cap archiu Excel d'Alma Analytics.\n"
                           'Vòs seleccionar-lo ara?',
        'mb_pol_no_inv': 'Selecciona prumèr ua factura (o lòt de factures) ena pestanha Convertir.',
        'mb_pol_no_lines': 'Non i a linhes visibles en editor.',
        'mb_pol_sel_row': 'Selecciona ua linha dera lista.',
        'mb_pol_title': 'Cercar PO Line en Alma',
        'mb_pref_invalid_dir': '{label} non ei un dorsièr valid:\n{val}',
        'mb_pref_invalid_title': 'Preferéncies non valides',
        'mb_pref_reset_ask': 'Vòs restablir era institucion, era lengua e toti es directòris personalizadi?',
        'mb_pref_reset_ok': "S'an restablit es valors iniciaus.",
        'mb_pref_reset_title': 'Restablir preferéncies',
        'mb_pref_saved_msg': "Es preferéncies s'an sauvat corrèctament.",
        'mb_prot_drag_msg': 'Aguesta colomna ei protegida e non se pòt aumplir en tot arrossegar.',
        'mb_prot_field_title': 'Camp protegit',
        'mb_prot_generic_msg': 'Aguest camp non ei editable.',
        'mb_prot_hdr_msg': "Eth numèro, era data e es impòrts de capçalèra non s'editen dirèctament. Eth "
                           'numèro sonque admet prefix o sufix.',
        'mb_prot_line_msg': 'Aguest camp procedís dera factura originau e non ei editable en FacturaeToAlma.',
        'mb_reset_inv_ask': "Vòs descartar toti es cambis d'aguesta factura?",
        'mb_reset_inv_title': 'Restablir factura',
        'mb_save_changes_ask': 'Vòs sauvar toti es cambis e actualizar ara er Excel de gessuda?\n'
                               '\n'
                               'Se susescriurà er Excel indicat en Convertir. Er XML originau non se '
                               'modificarà.',
        'mb_save_changes_title': 'Sauvar cambis',
        'mb_save_draft_empty': 'Non i a cap factura cargada entà sauvar en un borrador.',
        'mb_save_draft_title': 'Sauvar borrador',
        'mb_unsaved_ed_ask': "I a cambis hèti en Editar (numèro de factura, capçalèra o linhes) qu'encara "
                             "non s'auien sauvat.\n"
                             '\n'
                             'Vòs aplicar e sauvar aguesti cambis entà convertir ara?',
        'mb_unsaved_ed_title': 'Cambis pendents en Editar',
        'mb_vend_clr_ask': 'Vòs eliminar totes es associacions NIF → Vendor Account sauvades?',
        'mb_vend_clr_title': 'Vuedar provedidors rebrembadi',
        'mb_vend_exp_err_title': 'Error en exportar provedidors',
        'mb_vend_exp_ok_msg': "S'an exportat {count} provedidors (colomnes NIF / CIF e Vendor Account) ar "
                              'archiu:\n'
                              '{path}',
        'mb_vend_exp_ok_title': 'Exportacion de provedidors completada',
        'mb_vend_full_ask': 'Eth mòde «Carga complèta» remplaçarà es {count} provedidors sauvati actuaument '
                            'pes der archiu Excel seleccionat.\n'
                            '\n'
                            'Vòs contunhar?',
        'mb_vend_imp_err_title': 'Error en importar provedidors',
        'mb_vend_imp_ok_msg': 'Mòde aplicat: {mode}\n'
                              '\n'
                              '- Naui ajustadi: {added}\n'
                              '- Existents actualizadi: {updated}\n'
                              '- Sense cambis / ometudi: {skipped}\n'
                              '- Linhes incomplètes ignorades: {invalid}\n'
                              '\n'
                              'Totau de provedidors sauvati: {total}',
        'mb_vend_imp_ok_title': 'Importacion de provedidors completada',
        'mb_vend_missing': "Indica tant eth NIF/CIF com eth còdi Vendor Account d'Alma.",
        'mb_vend_sel_del': 'Selecciona ua linha dera taula entà eliminar-la.',
        'mb_vend_title': 'Provedidors rebrembadi',
        'mb_warn_title': 'Auís',
        'numdlg_apply': 'Aplicar en aguesta session',
        'numdlg_desc': "Ei probable qu'eth numèro pertanhe a ua seqüéncia que se reinície periodicament.\n"
                       'Pòs ajustar temporaument un prefix o sufix entà desambiguar.\n'
                       'Er XML originau non se modificarà.',
        'numdlg_orig': 'Numèro originau: {original}',
        'numdlg_title': 'Corregir numèro de factura en memòria',
        'poldlg_all_btn': 'Véder toti',
        'poldlg_assign_full': "Assignar PO Line + còdis e honsi d'Analytics",
        'poldlg_assign_only': 'Assignar sonque PO Line',
        'poldlg_col_score': 'Afinitat',
        'poldlg_dblclick_hint': 'Eth doble clic assigne unicament era PO Line (sense alterar Reporting Codes '
                                'ne honsi).',
        'poldlg_inv_title': 'Títol ena factura: {title}',
        'poldlg_line_frame': 'Linha de factura #{line}',
        'poldlg_search_btn': 'Cercar',
        'poldlg_search_lbl': 'Cercar en Analytics (Títol / ISBN / ISSN / PO Line):',
        'poldlg_summary': 'ISBN: {isbn}   |   ISSN: {issn}   |   Prètz: {price}   |   PO Line actuau: {pol}',
        'poldlg_title': "Cercar PO Line en catalòg d'Alma - Linha {line}",
        'poldlg_unassigned': '(sense assignar)',
        'st_alma_selected': "Estat: archiu Excel d'Alma Analytics seleccionat.",
        'st_batch_count': 'Estat: {count} factures en lòt.',
        'st_batch_empty': 'Estat: lòt vued.',
        'st_batch_sorted': 'Estat: lòt ordenat per numèro de factura.',
        'st_check_inputs': "Estat: revise es donades d'entrada.",
        'st_cleared': 'Estat: donades netejades. Eth modèl Alma se consèrve.',
        'st_conv_err': 'Estat: error pendent era conversion.',
        'st_conv_ok': 'Estat: conversion completada sense incidéncies.',
        'st_conv_warn': 'Estat: conversion completada damb incidéncies.',
        'st_converting': 'Estat: en tot convertir en segon plan...',
        'st_copied': "Estat: copiada era valor '{val}' de {col}.",
        'st_draft_loaded': "Estat: borrador '{name}' cargat corrèctament.",
        'st_draft_ready_alma': "Estat: borrador editable der Excel d'Alma preparat ( crotzat damb Alma "
                               'Analytics).',
        'st_draft_ready_no_alma': "Estat: factura cargada en Editar (selecciona er Excel d'Alma Analytics "
                                  'entà crotzar PO Lines).',
        'st_draft_saved': "Estat: borrador sauvat en '{name}'.",
        'st_fill_blanks': "Estat: aumplides {count} cèlles vuedes inferiores damb '{val}'.",
        'st_fill_drag': "Estat: aumplides {count} cèlles en '{col}' damb '{val}'.",
        'st_inv_selected': 'Estat: factura {number} seleccionada. Provedidor: {vendor}.',
        'st_lang_changed': 'Estat: lengua cambiada a {lang}.',
        'st_mode_batch': 'Estat: mòde lòt. Selecciona factures deth madeish provedidor e bibliotèca.',
        'st_mode_single': 'Estat: mòde factura individuau.',
        'st_no_redo': 'Estat: non i a accions entà tornar a hèr.',
        'st_no_undo': 'Estat: non i a accions entà desfer.',
        'st_num_updated': "Estat: numèro de factura actualizat en borrador a '{number}'.",
        'st_open_in_edit': "Estat: daurida era factura '{name}' en Editar.",
        'st_out_selected': 'Estat: archiu de gessuda definit.',
        'st_pending': 'Estat: pendent de seleccionar archius.',
        'st_redone': "Estat: tornada a hèr er'accion '{action}'.",
        'st_saving_edits': 'Estat: sauvati {n_num} cambis de numèro, {n_hdr} de capçalèra e {n_line} de '
                           'linha; en tot actualizar er Excel de gessuda...',
        'st_tpl_selected': 'Estat: modèl Alma seleccionat.',
        'st_undone': "Estat: deshèta er'accion '{action}'.",
        'st_vendor_cleared': 'Estat: lista de provedidors rebrembadi vuedada.',
        'st_vendor_deleted': "Estat: eliminat eth provedidor rebrembat entath NIF '{nif}'.",
        'st_vendor_exported': "Estat: exportadi {count} provedidors a '{name}'.",
        'st_vendor_imported': "Estat: provedidors importadi des de '{name}' (totau: {total}).",
        'st_vendor_saved': 'Estat: provedidor sauvat ({nif} → {vendor}).',
        'sum_body': '\n'
                    '\n'
                    'Resum de linhes:\n'
                    '- Linhes processades: {total_lines}\n'
                    '- Coincidéncies per ISBN: {match_isbn}\n'
                    '- Coincidéncies per ISSN: {match_issn}\n'
                    '- Coincidéncies per títol: {match_title}\n'
                    '{extra}- Duplicats ISBN: {duplicate_isbn}\n'
                    '- Duplicats ISSN: {duplicate_issn}\n'
                    '- Títols ambigús: {ambiguous_title}\n'
                    '- Sense coincidéncia: {no_match}\n'
                    '- Linhes sense PO Line: {lines_without_po_line}',
        'sum_pol_man': '- Assignacions manuaus de PO Line: {n}\n',
        'sum_pol_xml': '- Coincidéncies per PO Line en XML: {n}\n',
        'tab_convert': 'Convertir',
        'tab_customize': 'Personalizar',
        'tab_edit': 'Editar',
        'tab_help': 'Ajuda',
        'tab_visualize': 'Visualizar',
        'val_bad_draft': 'Er archiu non a un format de borrador valid de FacturaeToAlma.',
        'val_bad_ext': 'Extension de factura non valida: {path}',
        'val_draft_no_files': 'Cap des factures referenciades en borrador existís ja ena sua rota originau.',
        'val_dup_numbers': 'I a numèros de factura duplicadi en lòt: {nums}',
        'val_multi_vendor': 'Eth lòt conten factures de diuèrsi provedidors. Selecciona unicament factures '
                            'deth madeish provedidor.\n'
                            '\n'
                            'Provedidors detectadi:\n'
                            '{details}',
        'val_need_alma': "As de seleccionar er archiu Excel d'Alma Analytics.",
        'val_need_inv': 'As de seleccionar ua factura o un lòt de factures.',
        'val_need_inv_edit': 'As de seleccionar prumèr ua factura o un lòt de factures ena pestanha '
                             'Convertir.',
        'val_need_out': 'As de definir er Excel de gessuda.',
        'val_need_tpl': 'As de seleccionar eth modèl Alma.',
        'val_no_alma_file': "Non existís er archiu Excel d'Alma Analytics: {path}",
        'val_no_inv_file': 'Non existís era factura: {path}',
        'val_no_tpl_file': 'Non existís eth modèl Alma: {path}',
        'val_out_collision': 'Er Excel de gessuda non pòt coincidir damb era factura, er Analytics ne eth '
                             'modèl. Escuelh un aute nòm o dorsièr.',
        'val_single_inv_only': 'Er archiu a de conténer ua soleta factura e en conten {count}.',
        'val_vend_empty_xlsx': 'Er archiu Excel ei vued.',
        'val_vend_po_cols': 'Er archiu conten colomnes de linha de comanda ({cols}). Er Excel de provedidors '
                            "a d'auer unicament es 2 colomnes «NIF / CIF» e «Vendor Account».",
        'val_vend_two_cols': 'Er Excel de provedidors a de conténer exactament es 2 colomnes «NIF / CIF» e '
                             '«Vendor Account» (sense camps de linha de comanda).',
        'vis_footer_note': "Vista de consulta dera factura originau. Es cambis der Excel d'Alma se hèn ena "
                           'pestanha Editar.',
        'vis_inv_lbl': 'Factura:',
        'vw_addr': 'Adreça: {addr}',
        'vw_card_buyer': ' DONADES DETH RECEPTOR ',
        'vw_card_dir3': ' CENTRES ADMINISTRATIUS FACe (DIR3) ',
        'vw_card_lines': ' DETALH ECONOMIC (LINHES) ',
        'vw_card_pay': ' DONADES DE PAGAMENT ',
        'vw_card_seller': ' DONADES DETH EMISSOR ',
        'vw_col_code': 'Article / ISBN',
        'vw_col_desc': 'Descripcion / Títol',
        'vw_col_gross': 'Impòrt Brut',
        'vw_col_price': 'Prètz Unit.',
        'vw_col_qty': 'Quant.',
        'vw_col_tax': 'IVA Repercutit',
        'vw_contact': 'Contacte: {contact}',
        'vw_elec_inv_lbl': 'Factura electronica:',
        'vw_hdr_invoice': 'FACTURA: {number}',
        'vw_hdr_meta': "Sèria: {series}   |   Data d'emission: {issue}   |   Data d'operacion: {op}",
        'vw_load_btn': 'Cargar factura...',
        'vw_nif': 'CIF/NIF: {nif}',
        'vw_no_number': '(sense numèro)',
        'vw_not_specified': 'Non especificat',
        'vw_pay_due': "Data d'escadença: {date}",
        'vw_pay_iban': 'Compte / IBAN: {iban}',
        'vw_pay_means': 'Mieja de pagament: Transferéncia (Còdi {code})',
        'vw_placeholder': 'Seleccionatz un archiu de factura electronica (.xsig, .xml, .txt) entà visualizar '
                          'es sues donades.',
        'vw_role_01': 'Oficina Compdabla',
        'vw_role_02': 'Òrgan Gestor',
        'vw_role_03': 'Unitat Tramitadora',
        'vw_role_04': 'Òrgan Proponent',
        'vw_role_other': 'Ròtle {code}',
        'vw_tab_pdf': '  Vista Factura (Estil PDF)  ',
        'vw_tab_xml': '  Còdi XML (FACe)  ',
        'vw_tot_base': 'Basa Imposabla:',
        'vw_tot_grand': 'TOTAU FACTURA:',
        'vw_tot_payable': 'Totau a Pagar:',
        'vw_tot_ret': 'Impòsti Retengudi:',
        'vw_tot_taxes': 'Impòsti Repercutidi:',
        'vw_unknown_vendor': 'Provedidor non identificat'},
 'va': {'afdlg_apply_next': 'Aplicar i continuar filtrant una altra columna',
        'afdlg_blank_chk': '(Buides / Sense valor)',
        'afdlg_check_all': 'Marcar tots',
        'afdlg_col_lbl': 'Columna a filtrar:',
        'afdlg_list_frame': 'Valors presents en la columna (marca els que desitges mostrar)',
        'afdlg_only_blank': 'Només buides',
        'afdlg_only_val': 'Només amb valor',
        'afdlg_remove_col': "Llevar filtre d'esta columna",
        'afdlg_search_lbl': 'Buscar text / valor:',
        'afdlg_title': 'Filtre autofiltre per columnes (estil Excel)',
        'afdlg_uncheck_all': 'Desmarcar tots',
        'bulkdlg_lbl': 'Valor per a {col}:',
        'bulkdlg_title': 'Assignar {col}',
        'conv_actions_frame': '4. Accions',
        'conv_alma_btn': 'Seleccionar arxiu...',
        'conv_alma_lbl': "Arxiu Excel d'Alma Analytics amb PO Lines *",
        'conv_batch_add': 'Afegir factures',
        'conv_batch_clear': 'Buidar lot',
        'conv_batch_count': '{count} factures seleccionades',
        'conv_batch_frame': 'Factures del lot',
        'conv_batch_hint': '(Doble clic en una factura per a obrir-la en Editar)',
        'conv_batch_remove': 'Llevar seleccionades',
        'conv_batch_sort': 'Ordenar per número',
        'conv_clear_btn': 'Netejar dades',
        'conv_col_file': 'Arxiu',
        'conv_col_invoice': 'Número de factura',
        'conv_col_vendor': 'Proveïdor',
        'conv_exit_btn': 'Eixir',
        'conv_inputs_frame': "2. Arxius d'entrada",
        'conv_inv_batch_btn': 'Seleccionar factures...',
        'conv_inv_batch_lbl': 'Factures XSIG / XML / TXT del mateix proveïdor i biblioteca *',
        'conv_inv_single_btn': 'Seleccionar factura...',
        'conv_inv_single_lbl': 'Factura electrònica XSIG / XML / TXT *',
        'conv_mode_batch': 'Lot de factures del mateix proveïdor i biblioteca',
        'conv_mode_frame': '1. Mode de conversió',
        'conv_mode_single': 'Factura individual',
        'conv_ok_heading': 'CONVERSIÓ COMPLETA SENSE INCIDÈNCIES',
        'conv_open_out_btn': "Obrir carpeta d'eixida",
        'conv_output_btn': 'Triar destinació...',
        'conv_output_frame': "3. Arxiu d'eixida",
        'conv_result_body': '{heading}\n'
                            '\n'
                            'Factures processades: {total_inv}\n'
                            'Avisos de validació: {total_iss}\n'
                            '\n'
                            'Excel Alma generat:\n'
                            '{out}',
        'conv_run_btn': 'Convertir a format Alma',
        'conv_tpl_btn': 'Seleccionar plantilla...',
        'conv_tpl_lbl': 'Plantilla Alma .xlsx *',
        'conv_unmatched_warn': '\n'
                               '\n'
                               "Atenció: hi ha línies sense PO Line. Revisa l'informe abans de carregar en "
                               'Alma.',
        'conv_val_report': '\n\nInforme de validació:\n{path}',
        'conv_warn_heading': 'INCIDÈNCIES EN LA CONVERSIÓ',
        'ctx_bulk_apply': 'Aplicar valor a les línies seleccionades...',
        'ctx_clear_cell': 'Buidar cel·la (Supr)',
        'ctx_edit_cell': 'Editar cel·la',
        'ctx_fill_blanks_down': 'Omplir blancs inferiors (■)',
        'ctx_fill_blanks_sel': 'Omplir blancs seleccionats',
        'ctx_fill_down': 'Omplir cap avall (Ctrl+D)',
        'ctx_flt_cell_val': "Filtrar pel valor d'esta cel·la",
        'ctx_flt_clear_all': 'Llevar tots els filtres',
        'ctx_flt_col_excel': 'Filtrar esta columna (estil Excel)...',
        'ctx_redo': 'Refer (Ctrl+Y)',
        'ctx_restore_cell': 'Restablir cel·la',
        'ctx_restore_orig': 'Restablir cel·la original',
        'ctx_search_pol': "Buscar PO Line en catàleg d'Alma...",
        'ctx_suggest_ctx': 'Suggerir valor pel context',
        'ctx_undo': 'Desfer (Ctrl+Z)',
        'cust_choose_dir_btn': 'Seleccionar carpeta...',
        'cust_dir_alma': "Directori de l'arxiu Excel d'Alma",
        'cust_dir_inv': 'Directori de factures FACe',
        'cust_dir_out': "Directori d'Excel finals",
        'cust_dir_tpl': 'Directori de plantilla Alma',
        'cust_inst_lbl': 'Institució',
        'cust_lang_lbl': 'Idioma de la interfície',
        'cust_logo_del_btn': 'Eliminar logo',
        'cust_logo_hint': 'PNG, JPG o JPEG. Altura màxima: 100 px.',
        'cust_logo_invalid': "L'arxiu seleccionat no és una imatge vàlida.",
        'cust_logo_lbl': 'Logo institucional',
        'cust_logo_loaded': 'Logo carregat: {name}',
        'cust_logo_none': 'Sense logo. PNG, JPG o JPEG; altura màxima: 100 px.',
        'cust_logo_sel_btn': 'Seleccionar logo...',
        'cust_logo_unavail': 'El logo configurat no està disponible.',
        'cust_pref_frame': 'Preferències',
        'cust_pref_note': "Tots els camps són opcionals. Si una carpeta personalitzada queda buida, s'usa "
                          "l'última carpeta vàlida; en la primera execució s'usa la carpeta de l'aplicació. "
                          "L'eixida es guarda junt amb la factura, llevat que es definisca un directori "
                          "d'eixida.",
        'cust_reset_pref_btn': 'Restablir valors',
        'cust_save_pref_btn': 'Guardar preferències',
        'cust_vend_acc_lbl': 'Vendor Account en Alma:',
        'cust_vend_add_btn': 'Afegir / Actualitzar',
        'cust_vend_clr_btn': 'Buidar llista',
        'cust_vend_col_acc': 'Vendor Account en Alma',
        'cust_vend_col_nif': 'NIF / CIF normalitzat',
        'cust_vend_del_btn': 'Eliminar seleccionat',
        'cust_vend_exp_btn': 'Exportar Excel...',
        'cust_vend_frame': 'Proveïdors recordats automàticament (NIF / CIF → Vendor Account en Alma)',
        'cust_vend_imp_btn': 'Importar Excel...',
        'cust_vend_imp_mode_lbl': "Mode d'importació Excel:",
        'cust_vend_mode_full': 'Càrrega completa (reemplaçar llista actual)',
        'cust_vend_mode_inc': 'Incremental / acumulatiu (afegir i actualitzar)',
        'cust_vend_mode_new': 'Només afegir nous (no sobreescriure)',
        'cust_vend_nif_lbl': 'NIF / CIF:',
        'dlg_accept': 'Acceptar',
        'dlg_apply': 'Aplicar',
        'dlg_cancel': 'Cancel·lar',
        'ed_apply_num_btn': 'Aplicar número',
        'ed_bottom_hint': 'Escriu directe o doble clic per a editar | Arrossega ■ per a omplir | Ctrl+F '
                          'buscar | Clic en capçaleres ▼ per a filtrar.',
        'ed_col_invoice': 'Factura',
        'ed_col_line': 'Línia',
        'ed_diag_edited': '✎ Línia {line}: modificada manualment en esta sessió.',
        'ed_diag_init': "Selecciona una línia per a vore el seu diagnòstic d'encreuament i validació.",
        'ed_diag_manual': "✓ Línia {line}: PO Line '{pol}' assignada manualment.",
        'ed_diag_matched': '✓ Línia {line}: {msg}.',
        'ed_diag_no_inv': 'Selecciona una factura per a revisar les seues línies i avisos.',
        'ed_diag_ready': '✓ Línia {line}: llesta per a exportar.',
        'ed_diag_recon': '⚠ Quadrament factura: {msg}',
        'ed_diag_unmatched': "⚠ Línia {line}: sense PO Line assignada (usa 'Buscar PO Line en Alma...' o "
                             'escriu-la en la cel·la).',
        'ed_diag_warning': '⚠ Línia {line}: {notes}',
        'ed_flt_active_prefix': 'Filtres actius -> ',
        'ed_flt_add_btn': '+ Afegir filtre',
        'ed_flt_clear_btn': 'Netejar filtres',
        'ed_flt_col_lbl': 'Columna:',
        'ed_flt_contains_q': " conté '{q}'",
        'ed_flt_excel_btn': 'Filtre Excel...',
        'ed_flt_frame': 'Filtrar línies de factura (per estat i per una o diverses columnes estil Excel)',
        'ed_flt_in_list': '{col} en [{items}]',
        'ed_flt_none': 'Sense filtres actius (clic en qualsevol capçalera ▼ per a filtrar per una o més '
                       'columnes estil Excel).',
        'ed_flt_op_blank': 'Buides',
        'ed_flt_op_contains': 'Conté / Igual',
        'ed_flt_op_exact': 'Igual exacte',
        'ed_flt_op_nblank': 'No buides',
        'ed_flt_op_neq': 'Diferent de',
        'ed_flt_st_all': 'Totes les línies',
        'ed_flt_st_edit': 'Només línies modificades manualment',
        'ed_flt_st_ok': 'Només línies correctes (amb PO Line i sense avisos)',
        'ed_flt_st_prob': 'Només línies amb problemes (sense PO Line o amb avisos)',
        'ed_flt_st_unm': 'Només línies sense PO Line (pendents)',
        'ed_flt_state_lbl': 'Estat:',
        'ed_flt_tag': 'Filtre',
        'ed_flt_val_blank': '(Buides)',
        'ed_flt_val_lbl': 'Valor:',
        'ed_flt_val_nblank': '(No buides)',
        'ed_flt_values_cnt': '{col} ({count} valors)',
        'ed_hdr_frame': 'Capçalera de factura (INV)',
        'ed_load_draft_btn': 'Carregar esborrany...',
        'ed_mode_orig': 'Original',
        'ed_mode_prefix': 'Prefix',
        'ed_mode_suffix': 'Sufix',
        'ed_num_frame': "Número de factura en l'Excel d'Alma",
        'ed_orig_lbl': 'Original:',
        'ed_refresh_btn': 'Actualitzar esborrany',
        'ed_reset_inv_btn': 'Restablir factura',
        'ed_save_draft_btn': 'Guardar esborrany...',
        'ed_save_excel_btn': 'Guardar canvis i actualitzar Excel',
        'ed_search_pol_btn': 'Buscar PO Line en Alma...',
        'ed_stats_empty': 'Mostrant: 0/0 | Correctes: 0 | Amb problemes: 0',
        'ed_stats_fmt': 'Mostrant: {shown}/{total} | Correctes: {ok} | Amb problemes: {prob} (Sense POL: '
                        '{unm})',
        'ed_tb_bulk': 'Assignar valor a seleccionades...',
        'ed_tb_clear': 'Buidar cel·la (Supr)',
        'ed_tb_copy': 'Copiar (Ctrl+C)',
        'ed_tb_fill_blanks': '⬇ Omplir blancs (■)',
        'ed_tb_fill_down': '⬇ Omplir avall (Ctrl+D)',
        'ed_tb_paste': 'Pegar (Ctrl+V)',
        'ed_tb_redo': '↪ Refer (Ctrl+Y)',
        'ed_tb_restore': 'Restablir cel·la',
        'ed_tb_search_lbl': 'Buscar (Ctrl+F):',
        'ed_tb_undo': '↩ Desfer (Ctrl+Z)',
        'ed_tot_init': 'Import visible: 0,00 €',
        'ed_tot_sel': '   |   Selecció ({count}): {base}',
        'ed_tot_vis': 'Visible ({count}): Base {base} | IVA {vat}',
        'err_conv_title': 'Error durant la conversió',
        'err_fnf_title': 'Arxiu no trobat',
        'err_log_hint': "{err}\n\nConsulta l'arxiu de log '{log}' per a més detalls.",
        'err_perm_title': 'Arxiu bloquejat o sense permisos',
        'err_val_title': 'Dades o format no vàlids',
        'err_xml_sec_title': 'XML bloquejat per seguretat',
        'fd_exp_vendors': 'Exportar proveïdors (NIF / CIF i Vendor Account) a Excel',
        'fd_ft_all': 'Tots els arxius',
        'fd_ft_compat_inv': 'Factures compatibles',
        'fd_ft_draft': 'Esborrany FacturaeToAlma',
        'fd_ft_excel': 'Arxius Excel',
        'fd_ft_img_compat': 'Imatges compatibles',
        'fd_ft_jpg': 'Imatge JPEG',
        'fd_ft_png': 'Imatge PNG',
        'fd_ft_txt': 'Arxius TXT',
        'fd_ft_xml': 'Arxius XML',
        'fd_ft_xsig': 'Arxius XSIG',
        'fd_imp_vendors': "Importar proveïdors (NIF / CIF i Vendor Account) des d'Excel",
        'fd_load_draft': 'Carregar esborrany de sessió FacturaeToAlma',
        'fd_save_draft': 'Guardar esborrany de sessió FacturaeToAlma',
        'fd_sel_alma': "Seleccionar arxiu Excel d'Alma Analytics",
        'fd_sel_folder': 'Seleccionar carpeta',
        'fd_sel_inv_batch': 'Seleccionar factures del mateix proveïdor i biblioteca',
        'fd_sel_inv_face': 'Seleccionar factura electrònica FACe',
        'fd_sel_inv_single': 'Seleccionar factura XSIG / XML / TXT',
        'fd_sel_logo': 'Seleccionar logo institucional',
        'fd_sel_out': "Definir arxiu Excel d'eixida",
        'fd_sel_tpl': 'Seleccionar plantilla Alma .xlsx',
        'gctx_copy': 'Copiar',
        'gctx_copy_txt': 'Copiar text',
        'gctx_cut': 'Retallar',
        'gctx_go_edit': 'Anar a Editar',
        'gctx_go_vis': 'Anar a Visualitzar',
        'gctx_paste': 'Pegar',
        'gctx_sel_all': 'Seleccionar tot',
        'hlp_body_5_0': "Factura individual: selecciona una factura, l'informe d'Alma Analytics, la "
                        'plantilla i la destinació.\n'
                        '\n'
                        'Mode lot: totes les factures han de pertànyer al mateix proveïdor i biblioteca. La '
                        'llista permet afegir, llevar, ordenar i obrir cada factura amb doble clic en '
                        'Editar.\n'
                        '\n'
                        'Edició i filtratge: revisa la capçalera (INV) i les línies (IL), busca PO Lines en '
                        "el catàleg d'Alma Analytics, aplica autofiltres estil Excel o guarda/carrega "
                        'esborranys (.ftalma.json).\n'
                        '\n'
                        "Proveïdors i idiomes: en Personalitzar pots canviar l'idioma oficial/cooficial (es, "
                        'ca, va, gl, eu, oc) i gestionar, importar o exportar en Excel la taula de '
                        'proveïdors (NIF / CIF i Vendor Account) en mode incremental, només nous o càrrega '
                        'completa.\n'
                        '\n'
                        'Informe de validació: es genera quan falten PO Lines, hi ha coincidències ambigües, '
                        'duplicats, desquadraments comptables o diferències entre Quantity i Quantity for '
                        'Pricing.\n'
                        '\n'
                        'Limitacions confirmades per Ex Libris: la plantilla Excel no permet indicar Line '
                        'Exclusive ni declarar explícitament facturació parcial o completa.\n'
                        '\n'
                        'Llicència: GPL-3.0-only.',
        'hlp_body_5_1': "Factura individual: selecciona una factura, l'informe d'Alma Analytics, la "
                        'plantilla i la destinació.\n'
                        '\n'
                        'Mode lot: totes les factures han de pertànyer al mateix proveïdor i biblioteca. La '
                        'llista permet afegir, llevar, ordenar i obrir cada factura amb doble clic en '
                        'Editar.\n'
                        '\n'
                        "Imports i IVA compatibles amb Línia exclusiva (lògica exacta 4.0): l'extracció i "
                        "l'exportació d'imports i IVA reproduïxen exactament el comportament de la versió "
                        '4.0 perquè el canvi manual a «Línia exclusiva» en Alma quadre sempre.\n'
                        '\n'
                        'Proveïdors i idiomes: en Personalitzar pots seleccionar qualsevol de les llengües '
                        "oficials i cooficials d'Espanya (es, ca, va, gl, eu, oc) i importar o exportar en "
                        'Excel la taula de proveïdors (només columnes NIF / CIF i Vendor Account) en mode '
                        'incremental/acumulatiu, només afegir nous o càrrega completa.\n'
                        '\n'
                        'Informe de validació: es genera quan falten PO Lines, hi ha coincidències ambigües, '
                        'duplicats o diferències entre Quantity i Quantity for Pricing.\n'
                        '\n'
                        'Limitacions confirmades per Ex Libris: la plantilla Excel no permet indicar Line '
                        'Exclusive ni declarar explícitament facturació parcial o completa.\n'
                        '\n'
                        'Llicència: GPL-3.0-only.',
        'hlp_guide_btn': 'Obrir Biblioguia / LibGuides',
        'hlp_links_frame': 'Documentació i suport',
        'hlp_repo_btn': 'Obrir repositori de GitHub',
        'match_msg_ambig': 'Diverses coincidències possibles per títol',
        'match_msg_dup': 'Duplicat per {mtype}. Revisió manual necessària.',
        'match_msg_manual': 'PO Line assignada manualment',
        'match_msg_none': 'Sense coincidència en Alma per ISBN, ISSN ni títol',
        'match_msg_special': 'Línia de tipus {ltype} (no requerix PO Line bibliogràfica)',
        'match_msg_unique': 'Coincidència única per {mtype}',
        'mb_batch_clr_ask': 'Vols eliminar totes les factures del lot?',
        'mb_batch_clr_title': 'Buidar lot',
        'mb_batch_dup_num': 'Número de factura duplicat {num}: {name}',
        'mb_batch_none_added': "No s'han incorporat factures noves.",
        'mb_batch_read_err': "No s'ha pogut llegir {name}: {err}",
        'mb_batch_rep_path': 'Ruta repetida: {name}',
        'mb_batch_sel_one': 'Selecciona almenys una factura de la llista.',
        'mb_batch_title': 'Lot',
        'mb_batch_warn_title': 'Factures no incorporades',
        'mb_clr_alma_ask': "Vols netejar també l'arxiu Excel d'Alma Analytics?\n"
                           '\n'
                           'La plantilla Alma es conservarà.',
        'mb_clr_alma_title': "Netejar arxiu Excel d'Alma",
        'mb_conv_close_ask': "Hi ha una conversió en curs. Vols tancar l'aplicació de totes les maneres?",
        'mb_conv_ok_title': 'Conversió completada',
        'mb_conv_running_msg': 'Ja hi ha una conversió en curs.',
        'mb_conv_running_title': 'Conversió en curs',
        'mb_conv_warn_title': 'Incidències en la conversió',
        'mb_draft_loaded_msg': "S'han restaurat les factures i els canvis de l'esborrany.{warn}",
        'mb_draft_loaded_title': 'Esborrany carregat',
        'mb_draft_missing_warn': '\n'
                                 '\n'
                                 "Avís: no s'han trobat algunes factures mogudes o eliminades:\n"
                                 '- {files}',
        'mb_draft_outdated_msg': 'La factura, Analytics o la plantilla han canviat després de guardar. Torna '
                                 'a Editar, revisa-ho i guarda de nou.',
        'mb_draft_outdated_title': 'Esborrany desactualitzat',
        'mb_draft_rebuilt_alma': "S'ha reconstruït l'esborrany de l'Excel final d'Alma amb la factura i "
                                 "l'informe d'Analytics actualment seleccionats.",
        'mb_draft_rebuilt_no_alma': "S'han carregat les línies de la factura en l'editor.\n"
                                    '\n'
                                    "Si selecciones també l'arxiu Excel d'Alma Analytics, s'encreuaran "
                                    'automàticament les PO Lines, els Reporting Codes i els fons.',
        'mb_draft_saved_msg': "S'ha guardat el progrés actual en:\n"
                              '{path}\n'
                              '\n'
                              "Podràs reprendre l'edició més tard amb 'Carregar esborrany...'.",
        'mb_draft_saved_title': 'Esborrany guardat',
        'mb_draft_unavail_msg': "No s'ha pogut preparar l'esborrany editable abans de convertir.",
        'mb_draft_unavail_title': 'Esborrany no disponible',
        'mb_empty_src_msg': 'La cel·la activa està buida. Escriu o selecciona una cel·la amb valor abans '
                            "d'omplir els blancs.",
        'mb_empty_src_title': "Cel·la d'origen buida",
        'mb_err_folder_msg': "No s'ha pogut obrir la carpeta:\n{path}\n\n{err}",
        'mb_err_folder_title': 'Error en obrir carpeta',
        'mb_err_load_draft': 'Error en carregar esborrany',
        'mb_err_load_draft_msg': "No s'ha pogut carregar l'esborrany:\n{err}",
        'mb_err_prep_draft': "No s'ha pogut preparar l'esborrany",
        'mb_err_read_inv_msg': "No s'ha pogut carregar la factura seleccionada:\n{err}",
        'mb_err_read_inv_title': 'Error en llegir factura',
        'mb_err_reset_title': 'Error en restablir',
        'mb_err_save_draft': 'Error en guardar esborrany',
        'mb_err_save_draft_msg': "No s'ha pogut guardar l'esborrany:\n{err}",
        'mb_err_save_title': 'Error en guardar',
        'mb_err_url_msg': "No s'ha pogut obrir l'enllaç:\n{url}\n\n{err}",
        'mb_err_url_title': "Error en obrir l'enllaç",
        'mb_flt_cols_no_inv': 'Selecciona primer una factura en la pestanya Editar.',
        'mb_flt_cols_title': 'Filtrar columnes',
        'mb_inv_invalid_title': 'Factura no vàlida',
        'mb_invalid_date_msg': 'Usa una data vàlida, per exemple 2026-10-19.',
        'mb_invalid_date_title': 'Data no vàlida',
        'mb_invalid_inputs_title': "Dades d'entrada no vàlides",
        'mb_invalid_ltype': 'Selecciona un Line type permés per Alma.',
        'mb_invalid_num': '{col} ha de ser numèric.',
        'mb_invalid_val_title': 'Valor no vàlid',
        'mb_logo_invalid_msg': "No s'ha pogut carregar el logo:\n{path}\n\n{err}",
        'mb_logo_invalid_title': 'Logo no vàlid',
        'mb_logo_unavail_msg': "No s'ha trobat l'arxiu de logo:\n{path}",
        'mb_logo_unavail_title': 'Logo no disponible',
        'mb_no_active_inv': 'No hi ha cap factura activa.',
        'mb_no_edit_inv': "No hi ha cap factura activa en l'editor.",
        'mb_no_out_dir_def': "No hi ha una carpeta d'eixida definida.",
        'mb_no_out_dir_found': "No s'ha trobat la carpeta d'eixida.",
        'mb_no_sugg_msg': 'No hi ha valors de context per a suggerir en esta columna.',
        'mb_no_sugg_title': 'Sense suggeriments',
        'mb_pol_ask_alma': "Este buscador consulta les línies de comanda (PO Lines) dins de l'arxiu Excel "
                           "exportat d'Alma Analytics.\n"
                           '\n'
                           "Encara no has seleccionat cap arxiu Excel d'Alma Analytics.\n"
                           'Vols seleccionar-lo ara?',
        'mb_pol_no_inv': 'Selecciona primer una factura (o lot de factures) en la pestanya Convertir.',
        'mb_pol_no_lines': "No hi ha línies visibles en l'editor.",
        'mb_pol_sel_row': 'Selecciona una fila de la llista.',
        'mb_pol_title': 'Buscar PO Line en Alma',
        'mb_pref_invalid_dir': '{label} no és una carpeta vàlida:\n{val}',
        'mb_pref_invalid_title': 'Preferències no vàlides',
        'mb_pref_reset_ask': "Vols restablir la institució, l'idioma i tots els directoris personalitzats?",
        'mb_pref_reset_ok': "S'han restablit els valors inicials.",
        'mb_pref_reset_title': 'Restablir preferències',
        'mb_pref_saved_msg': "Les preferències s'han guardat correctament.",
        'mb_prot_drag_msg': 'Esta columna està protegida i no pot omplir-se arrossegant.',
        'mb_prot_field_title': 'Camp protegit',
        'mb_prot_generic_msg': 'Este camp no és editable.',
        'mb_prot_hdr_msg': "El número, la data i els imports de capçalera no s'editen directament. El número "
                           'només admet prefix o sufix.',
        'mb_prot_line_msg': 'Este camp procedix de la factura original i no és editable en FacturaeToAlma.',
        'mb_reset_inv_ask': "Vols descartar tots els canvis d'esta factura?",
        'mb_reset_inv_title': 'Restablir factura',
        'mb_save_changes_ask': "Vols guardar tots els canvis i actualitzar ara l'Excel d'eixida?\n"
                               '\n'
                               "Es sobreescriurà l'Excel indicat en Convertir. L'XML original no es "
                               'modificarà.',
        'mb_save_changes_title': 'Guardar canvis',
        'mb_save_draft_empty': 'No hi ha cap factura carregada per a guardar en un esborrany.',
        'mb_save_draft_title': 'Guardar esborrany',
        'mb_unsaved_ed_ask': 'Hi ha canvis realitzats en Editar (número de factura, capçalera o línies) que '
                             "encara no s'havien guardat.\n"
                             '\n'
                             'Vols aplicar i guardar estos canvis per a convertir ara?',
        'mb_unsaved_ed_title': 'Canvis pendents en Editar',
        'mb_vend_clr_ask': 'Vols eliminar totes les associacions NIF → Vendor Account guardades?',
        'mb_vend_clr_title': 'Buidar proveïdors recordats',
        'mb_vend_exp_err_title': 'Error en exportar proveïdors',
        'mb_vend_exp_ok_msg': "S'han exportat {count} proveïdors (columnes NIF / CIF i Vendor Account) a "
                              "l'arxiu:\n"
                              '{path}',
        'mb_vend_exp_ok_title': 'Exportació de proveïdors completada',
        'mb_vend_full_ask': 'El mode «Càrrega completa» reemplaçarà els {count} proveïdors guardats '
                            "actualment pels de l'arxiu Excel seleccionat.\n"
                            '\n'
                            'Desitges continuar?',
        'mb_vend_imp_err_title': 'Error en importar proveïdors',
        'mb_vend_imp_ok_msg': 'Mode aplicat: {mode}\n'
                              '\n'
                              '- Nous afegits: {added}\n'
                              '- Existents actualitzats: {updated}\n'
                              '- Sense canvis / omesos: {skipped}\n'
                              '- Files incompletes ignorades: {invalid}\n'
                              '\n'
                              'Total de proveïdors guardats: {total}',
        'mb_vend_imp_ok_title': 'Importació de proveïdors completada',
        'mb_vend_missing': "Indica tant el NIF/CIF com el codi Vendor Account d'Alma.",
        'mb_vend_sel_del': 'Selecciona una fila de la taula per a eliminar-la.',
        'mb_vend_title': 'Proveïdors recordats',
        'mb_warn_title': 'Avís',
        'numdlg_apply': 'Aplicar en esta sessió',
        'numdlg_desc': 'És probable que el número pertanga a una seqüència que es reinicia periòdicament.\n'
                       'Pots afegir temporalment un prefix o sufix per a desambiguar-lo.\n'
                       "L'XML original no es modificarà.",
        'numdlg_orig': 'Número original: {original}',
        'numdlg_title': 'Corregir número de factura en memòria',
        'poldlg_all_btn': 'Vore tots',
        'poldlg_assign_full': "Assignar PO Line + codis i fons d'Analytics",
        'poldlg_assign_only': 'Assignar només PO Line',
        'poldlg_col_score': 'Afinitat',
        'poldlg_dblclick_hint': 'Doble clic assigna únicament la PO Line (sense alterar Reporting Codes ni '
                                'fons).',
        'poldlg_inv_title': 'Títol en factura: {title}',
        'poldlg_line_frame': 'Línia de factura #{line}',
        'poldlg_search_btn': 'Buscar',
        'poldlg_search_lbl': 'Buscar en Analytics (Títol / ISBN / ISSN / PO Line):',
        'poldlg_summary': 'ISBN: {isbn}   |   ISSN: {issn}   |   Preu: {price}   |   PO Line actual: {pol}',
        'poldlg_title': "Buscar PO Line en catàleg d'Alma - Línia {line}",
        'poldlg_unassigned': '(sense assignar)',
        'st_alma_selected': "Estat: arxiu Excel d'Alma Analytics seleccionat.",
        'st_batch_count': 'Estat: {count} factures en el lot.',
        'st_batch_empty': 'Estat: lot buit.',
        'st_batch_sorted': 'Estat: lot ordenat per número de factura.',
        'st_check_inputs': "Estat: revisa les dades d'entrada.",
        'st_cleared': 'Estat: dades netejades. La plantilla Alma es conserva.',
        'st_conv_err': 'Estat: error durant la conversió.',
        'st_conv_ok': 'Estat: conversió completada sense incidències.',
        'st_conv_warn': 'Estat: conversió completada amb incidències.',
        'st_converting': 'Estat: convertint en segon pla...',
        'st_copied': "Estat: copiat el valor '{val}' de {col}.",
        'st_draft_loaded': "Estat: esborrany '{name}' carregat correctament.",
        'st_draft_ready_alma': "Estat: esborrany editable de l'Excel d'Alma preparat (encreuat amb Alma "
                               'Analytics).',
        'st_draft_ready_no_alma': "Estat: factura carregada en Editar (selecciona l'Excel d'Alma Analytics "
                                  'per a encreuar PO Lines).',
        'st_draft_saved': "Estat: esborrany guardat en '{name}'.",
        'st_fill_blanks': "Estat: omplides {count} cel·les en blanc inferiors amb '{val}'.",
        'st_fill_drag': "Estat: omplides {count} cel·les en '{col}' amb '{val}'.",
        'st_inv_selected': 'Estat: factura {number} seleccionada. Proveïdor: {vendor}.',
        'st_lang_changed': 'Estat: idioma canviat a {lang}.',
        'st_mode_batch': 'Estat: mode lot. Selecciona factures del mateix proveïdor i biblioteca.',
        'st_mode_single': 'Estat: mode factura individual.',
        'st_no_redo': 'Estat: no hi ha accions per a refer.',
        'st_no_undo': 'Estat: no hi ha accions per a desfer.',
        'st_num_updated': "Estat: número de factura actualitzat en l'esborrany a '{number}'.",
        'st_open_in_edit': "Estat: oberta la factura '{name}' en Editar.",
        'st_out_selected': "Estat: arxiu d'eixida definit.",
        'st_pending': 'Estat: pendent de seleccionar arxius.',
        'st_redone': "Estat: refeta l'acció '{action}'.",
        'st_saving_edits': 'Estat: guardats {n_num} canvis de número, {n_hdr} de capçalera i {n_line} de '
                           "línia; actualitzant l'Excel d'eixida...",
        'st_tpl_selected': 'Estat: plantilla Alma seleccionada.',
        'st_undone': "Estat: desfeta l'acció '{action}'.",
        'st_vendor_cleared': 'Estat: llista de proveïdors recordats buidada.',
        'st_vendor_deleted': "Estat: eliminat el proveïdor recordat per al NIF '{nif}'.",
        'st_vendor_exported': "Estat: exportats {count} proveïdors a '{name}'.",
        'st_vendor_imported': "Estat: proveïdors importats des de '{name}' (total: {total}).",
        'st_vendor_saved': 'Estat: proveïdor guardat ({nif} → {vendor}).',
        'sum_body': '\n'
                    '\n'
                    'Resum de línies:\n'
                    '- Línies processades: {total_lines}\n'
                    '- Coincidències per ISBN: {match_isbn}\n'
                    '- Coincidències per ISSN: {match_issn}\n'
                    '- Coincidències per títol: {match_title}\n'
                    '{extra}- Duplicats ISBN: {duplicate_isbn}\n'
                    '- Duplicats ISSN: {duplicate_issn}\n'
                    '- Títols ambigus: {ambiguous_title}\n'
                    '- Sense coincidència: {no_match}\n'
                    '- Línies sense PO Line: {lines_without_po_line}',
        'sum_pol_man': '- Assignacions manuals de PO Line: {n}\n',
        'sum_pol_xml': '- Coincidències per PO Line en XML: {n}\n',
        'tab_convert': 'Convertir',
        'tab_customize': 'Personalitzar',
        'tab_edit': 'Editar',
        'tab_help': 'Ajuda',
        'tab_visualize': 'Visualitzar',
        'val_bad_draft': "L'arxiu no té un format d'esborrany vàlid de FacturaeToAlma.",
        'val_bad_ext': 'Extensió de factura no vàlida: {path}',
        'val_draft_no_files': "Cap de les factures referenciades en l'esborrany existix ja en la seua ruta "
                              'original.',
        'val_dup_numbers': 'Hi ha números de factura duplicats en el lot: {nums}',
        'val_multi_vendor': 'El lot conté factures de diversos proveïdors. Selecciona únicament factures del '
                            'mateix proveïdor.\n'
                            '\n'
                            'Proveïdors detectats:\n'
                            '{details}',
        'val_need_alma': "Has de seleccionar l'arxiu Excel d'Alma Analytics.",
        'val_need_inv': 'Has de seleccionar una factura o un lot de factures.',
        'val_need_inv_edit': 'Has de seleccionar primer una factura o un lot de factures en la pestanya '
                             'Convertir.',
        'val_need_out': "Has de definir l'Excel d'eixida.",
        'val_need_tpl': 'Has de seleccionar la plantilla Alma.',
        'val_no_alma_file': "No existix l'arxiu Excel d'Alma Analytics: {path}",
        'val_no_inv_file': 'No existix la factura: {path}',
        'val_no_tpl_file': 'No existix la plantilla Alma: {path}',
        'val_out_collision': "L'Excel d'eixida no pot coincidir amb la factura, l'Analytics ni la plantilla. "
                             'Tria un altre nom o carpeta.',
        'val_single_inv_only': "L'arxiu ha de contindre una sola factura i en conté {count}.",
        'val_vend_empty_xlsx': "L'arxiu Excel està buit.",
        'val_vend_po_cols': "L'arxiu conté columnes de línia de comanda ({cols}). L'Excel de proveïdors ha "
                            'de tindre únicament les 2 columnes «NIF / CIF» i «Vendor Account».',
        'val_vend_two_cols': "L'Excel de proveïdors ha de contindre exactament les 2 columnes «NIF / CIF» i "
                             '«Vendor Account» (sense camps de línia de comanda).',
        'vis_footer_note': "Vista de consulta de la factura original. Els canvis de l'Excel d'Alma es "
                           'realitzen en la pestanya Editar.',
        'vis_inv_lbl': 'Factura:',
        'vw_addr': 'Adreça: {addr}',
        'vw_card_buyer': ' DADES DEL RECEPTOR ',
        'vw_card_dir3': ' CENTRES ADMINISTRATIUS FACe (DIR3) ',
        'vw_card_lines': ' DETALL ECONÒMIC (LÍNIES) ',
        'vw_card_pay': ' DADES DE PAGAMENT ',
        'vw_card_seller': " DADES DE L'EMISSOR ",
        'vw_col_code': 'Article / ISBN',
        'vw_col_desc': 'Descripció / Títol',
        'vw_col_gross': 'Import Brut',
        'vw_col_price': 'Preu Unit.',
        'vw_col_qty': 'Quant.',
        'vw_col_tax': 'IVA Repercutit',
        'vw_contact': 'Contacte: {contact}',
        'vw_elec_inv_lbl': 'Factura electrònica:',
        'vw_hdr_invoice': 'FACTURA: {number}',
        'vw_hdr_meta': "Sèrie: {series}   |   Data d'emissió: {issue}   |   Data d'operació: {op}",
        'vw_load_btn': 'Carregar factura...',
        'vw_nif': 'CIF/NIF: {nif}',
        'vw_no_number': '(sense número)',
        'vw_not_specified': 'No especificat',
        'vw_pay_due': 'Data de venciment: {date}',
        'vw_pay_iban': 'Compte / IBAN: {iban}',
        'vw_pay_means': 'Mitjà de pagament: Transferència (Codi {code})',
        'vw_placeholder': 'Seleccione un arxiu de factura electrònica (.xsig, .xml, .txt) per a visualitzar '
                          'les seues dades.',
        'vw_role_01': 'Oficina Comptable',
        'vw_role_02': 'Òrgan Gestor',
        'vw_role_03': 'Unitat Tramitadora',
        'vw_role_04': 'Òrgan Proponent',
        'vw_role_other': 'Rol {code}',
        'vw_tab_pdf': '  Vista Factura (Estil PDF)  ',
        'vw_tab_xml': '  Codi XML (FACe)  ',
        'vw_tot_base': 'Base Imposable:',
        'vw_tot_grand': 'TOTAL FACTURA:',
        'vw_tot_payable': 'Total a Pagar:',
        'vw_tot_ret': 'Impostos Retinguts:',
        'vw_tot_taxes': 'Impostos Repercutits:',
        'vw_unknown_vendor': 'Proveïdor no identificat'}}

_CURRENT_LANG = "es"


def set_current_language(lang_code):
    global _CURRENT_LANG
    code = str(lang_code or "es").strip().lower()
    _CURRENT_LANG = code if code in I18N else "es"
    return _CURRENT_LANG


def get_current_language():
    return _CURRENT_LANG


def tr(key, **kwargs):
    lang_dict = I18N.get(_CURRENT_LANG) or I18N["es"]
    template = lang_dict.get(key)
    if template is None:
        template = I18N["es"].get(key, key)
    if kwargs:
        try:
            return template.format(**kwargs)
        except (KeyError, ValueError, IndexError):
            return template
    return template


def localize_role_label(role_code):
    key = f"vw_role_{role_code}"
    if key in I18N["es"]:
        return tr(key)
    return tr("vw_role_other", code=role_code)


def localize_match_status(match_status, line_type="REGULAR"):
    status = str(match_status or "")
    if status == "MATCH_MANUAL":
        return tr("match_msg_manual")
    if status.startswith("MATCH_"):
        return tr("match_msg_unique", mtype=status.split("_", 1)[1])
    if status == "AMBIGUOUS_TITLE":
        return tr("match_msg_ambig")
    if status.startswith("DUPLICATE_"):
        return tr("match_msg_dup", mtype=status.split("_", 1)[1])
    if status.startswith("SPECIAL_"):
        return tr("match_msg_special", ltype=line_type or status.split("_", 1)[1])
    if status == "NO_MATCH":
        return tr("match_msg_none")
    return ""


_FORBIDDEN_VENDOR_PO_HEADERS = {
    "po line", "poline", "po_line", "reporting code", "secondary reporting code",
    "tertiary reporting code", "fourth reporting code", "fifth reporting code",
    "fund and percent", "quantity for pricing", "start subs date", "end subs date",
    "line type", "line number", "vendor ref number", "isbn", "issn", "title",
    "price", "quantity", "vat amount", "vat percentage",
}

_NIF_HEADER_ALIASES = {
    "nif / cif", "nif/cif", "cif / nif", "cif/nif", "nif", "cif",
    "nif / cif normalizado", "nif / cif normalitzat", "nif / cif normalizatua",
    "nif / cif normalizat", "ifz / ifk", "ifk / ifz", "ifz", "ifk",
    "tax id", "seller tax id", "identificacion fiscal",
}

_VENDOR_HEADER_ALIASES = {
    "vendor account", "vendor account en alma", "vendor account a alma",
    "vendor account alman", "vendor_account", "vendor", "codigo proveedor",
    "codi proveidor", "cuenta proveedor", "proveedor alma",
}


def _norm_vendor_col_header(value):
    cleaned = clean_cell(value)
    if not cleaned:
        return ""
    text = remove_accents(cleaned.lower())
    return re.sub(r"\s+", " ", text).strip()


def export_vendor_map_to_excel(vendor_map, output_path):
    """Exporta proveedores a un Excel (.xlsx) exclusivamente con las 2 columnas NIF / CIF y Vendor Account."""
    vendor_map = vendor_map if isinstance(vendor_map, dict) else {}
    output_dir = os.path.dirname(os.path.abspath(output_path)) or os.getcwd()
    os.makedirs(output_dir, exist_ok=True)
    temp_fd, temp_path = tempfile.mkstemp(prefix="Proveedores_Alma_", suffix=".xlsx", dir=output_dir)
    os.close(temp_fd)
    try:
        os.remove(temp_path)
    except FileNotFoundError:
        pass
    wb = Workbook()
    ws = wb.active
    ws.title = "Proveedores"
    ws.append(list(VENDOR_EXCEL_HEADERS))
    for col_idx in (1, 2):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)
    count = 0
    for nif in sorted(vendor_map):
        norm_nif = normalize_tax_id(nif)
        vendor_acc = clean_cell(vendor_map[nif])
        if not norm_nif or not vendor_acc:
            continue
        ws.append([sanitize_excel_value(norm_nif), sanitize_excel_value(vendor_acc)])
        count += 1
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 42
    try:
        wb.save(temp_path)
        os.replace(temp_path, output_path)
    except PermissionError as e:
        raise ValueError(f"No se pudo guardar el archivo '{output_path}'. Puede que esté abierto en Excel.") from e
    finally:
        wb.close()
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError as rm_exc:
                logging.debug("No se pudo eliminar temporal de proveedores: %s", rm_exc)
    return count


def import_vendor_map_from_excel(excel_path, existing_map=None, mode="incremental"):
    """
    Importa proveedores desde un Excel (.xlsx) que contenga únicamente las 2 columnas
    NIF / CIF y Vendor Account (sin campos de línea de pedido).
    Modos soportados: 'incremental', 'only_new', 'full'.
    """
    if mode not in VENDOR_IMPORT_MODES:
        raise ValueError(f"Modo de importación no válido: {mode}")
    existing = {
        normalize_tax_id(k): str(v).strip()
        for k, v in (existing_map or {}).items()
        if normalize_tax_id(k) and clean_cell(v)
    }
    wb = None
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        header_row = None
        for candidate_row in rows_iter:
            if candidate_row and any(clean_cell(c) is not None for c in candidate_row):
                header_row = candidate_row
                break
        if header_row is None:
            raise ValueError(tr("val_vend_empty_xlsx"))

        non_empty_headers = [
            (idx, clean_cell(val), _norm_vendor_col_header(val))
            for idx, val in enumerate(header_row)
            if clean_cell(val) is not None
        ]
        forbidden_found = [
            raw for _idx, raw, norm in non_empty_headers
            if norm in _FORBIDDEN_VENDOR_PO_HEADERS
        ]
        if forbidden_found:
            raise ValueError(tr("val_vend_po_cols", cols=", ".join(forbidden_found)))

        if len(non_empty_headers) != 2:
            raise ValueError(tr("val_vend_two_cols"))

        nif_col = None
        vendor_col = None
        for idx, _raw, norm in non_empty_headers:
            if norm in _NIF_HEADER_ALIASES and nif_col is None:
                nif_col = idx
            elif norm in _VENDOR_HEADER_ALIASES and vendor_col is None:
                vendor_col = idx

        if nif_col is None or vendor_col is None or nif_col == vendor_col:
            raise ValueError(tr("val_vend_two_cols"))

        file_entries = {}
        invalid = 0
        for row in rows_iter:
            if not row:
                continue
            raw_nif = row[nif_col] if nif_col < len(row) else None
            raw_vendor = row[vendor_col] if vendor_col < len(row) else None
            if clean_cell(raw_nif) is None and clean_cell(raw_vendor) is None:
                continue
            norm_nif = normalize_tax_id(raw_nif)
            clean_vendor = clean_cell(raw_vendor)
            if not norm_nif or not clean_vendor:
                invalid += 1
                continue
            if clean_vendor.startswith("'"):
                clean_vendor = clean_vendor[1:].strip()
            if mode == "only_new" and norm_nif in file_entries:
                continue
            file_entries[norm_nif] = clean_vendor
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception as close_exc:
                logging.debug("No se pudo cerrar el workbook de proveedores: %s", close_exc)

    added = 0
    updated = 0
    skipped = 0
    if mode == "full":
        result_map = dict(file_entries)
        for nif, vendor in file_entries.items():
            if nif not in existing:
                added += 1
            elif existing[nif] != vendor:
                updated += 1
            else:
                skipped += 1
    elif mode == "only_new":
        result_map = dict(existing)
        for nif, vendor in file_entries.items():
            if nif not in result_map:
                result_map[nif] = vendor
                added += 1
            else:
                skipped += 1
    else:  # incremental
        result_map = dict(existing)
        for nif, vendor in file_entries.items():
            if nif not in result_map:
                result_map[nif] = vendor
                added += 1
            elif result_map[nif] != vendor:
                result_map[nif] = vendor
                updated += 1
            else:
                skipped += 1

    stats = {
        "added": added,
        "updated": updated,
        "skipped": skipped,
        "invalid": invalid,
        "total": len(result_map),
    }
    return result_map, stats


def get_vendor_account(nif, seller_name=None, custom_vendor_map=None):
    nif_tax = normalize_tax_id(nif) if nif else ""
    nif_clean = normalize_identifier(nif) if nif else ""
    if custom_vendor_map and isinstance(custom_vendor_map, dict):
        for candidate in (nif_tax, nif.strip() if nif else "", nif_clean):
            if candidate and candidate in custom_vendor_map and custom_vendor_map[candidate]:
                return str(custom_vendor_map[candidate]).strip()
    for candidate in (nif.strip() if nif else "", nif_tax, nif_clean):
        if candidate and candidate in VENDOR_MAP:
            return VENDOR_MAP[candidate]
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

BIBLIOGRAPHIC_ABBREVIATIONS = {
    "adm": "administrativo", "admin": "administracion", "anatom": "anatomia",
    "antropol": "antropologia", "aprox": "aproximacion", "arq": "arquitectura",
    "biolog": "biologia", "cient": "cientifica", "civ": "civil",
    "coment": "comentarios", "comp": "compendio", "const": "constitucional",
    "contemp": "contemporanea", "contab": "contabilidad", "crim": "criminologia",
    "cuad": "cuadernos", "der": "derecho", "dicc": "diccionario",
    "dir": "direccion", "doc": "documentacion", "econ": "economia",
    "ed": "edicion", "educ": "educacion", "elem": "elementos",
    "encicl": "enciclopedia", "esp": "espanol", "estad": "estadistica",
    "est": "estudios", "filos": "filosofia", "finan": "financiero",
    "fis": "fisica", "fund": "fundamentos", "gen": "general",
    "geogr": "geografia", "geol": "geologia", "gest": "gestion",
    "gram": "gramatica", "hist": "historia", "ig": "general",
    "ing": "ingenieria", "inst": "instituciones", "int": "internacional",
    "intr": "introduccion", "intro": "introduccion", "introd": "introduccion",
    "inv": "investigacion", "jur": "juridico",
    "leg": "legislacion", "leng": "lengua", "lit": "literatura",
    "man": "manual", "mat": "matematicas", "mec": "mecanica",
    "med": "medicina", "mercan": "mercantil", "metod": "metodologia",
    "mod": "moderna", "org": "organizacion", "pedag": "pedagogia",
    "pen": "penal", "pol": "politica", "pract": "practico",
    "prin": "principios", "proc": "procesal", "prog": "programacion",
    "psicol": "psicologia", "pub": "publico", "quim": "quimica",
    "rev": "revista", "sist": "sistemas", "soc": "sociologia",
    "sup": "superior", "tec": "tecnica", "tecnol": "tecnologia",
    "teor": "teoria", "trab": "trabajo", "trat": "tratado",
    "trib": "tributario", "univ": "universitaria", "vol": "volumen",
}

def expand_bibliographic_words(words):
    expanded = []
    for w in words:
        mapped = BIBLIOGRAPHIC_ABBREVIATIONS.get(w, w)
        if mapped not in STOPWORDS:
            expanded.append(mapped)
    return expanded

def extract_main_title(title):
    if not title:
        return ""
    text = str(title)
    main_part = re.split(r"\s*[:/;]\s*|\s+-\s+|\s+–\s+", text, maxsplit=1)[0].strip()
    return main_part if len(main_part) >= 4 else text

def normalize_title_for_matching(title, expand_abbrevs=True):
    if not title:
        return []
    text = remove_accents(html.unescape(str(title)).lower()).replace("’", "'").replace("´", "'")
    for phrase in STOPWORD_PHRASES:
        normalized_phrase = re.sub(r"[^a-z0-9]+", " ", remove_accents(phrase.lower())).strip()
        if normalized_phrase:
            text = re.sub(rf"\b{re.escape(normalized_phrase)}\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    words = [w for w in text.split() if len(w) > 1 and w not in STOPWORDS]
    return expand_bibliographic_words(words) if expand_abbrevs else words

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

    coverage = len(shorter) / len(longer)
    return 0.94 + min(0.05, coverage * 0.05)


def extract_significant_numbers(words):
    """Extrae numeros que suelen distinguir curso, volumen o nivel."""
    values = set()
    roman = {"i": "1", "ii": "2", "iii": "3", "iv": "4", "v": "5", "vi": "6", "vii": "7", "viii": "8", "ix": "9", "x": "10"}
    for word in words:
        if word.isdigit():
            values.add(word)
        elif word in roman:
            values.add(roman[word])
        else:
            match = re.fullmatch(r"(\d+)(?:ep|o|a|ed)?", word)
            if match:
                values.add(match.group(1))
    return values


def diagnostic_title_score(invoice_title, record, word_idf=None):
    """Puntuacion explicable mejorada con IDF, comparacion de titulo principal MARC 245 y autor."""
    invoice_words = normalize_title_for_matching(invoice_title)
    alma_words = record.normalized_title_words
    if not invoice_words or not alma_words:
        return 0.0, []
    invoice_set, alma_set = set(invoice_words), set(alma_words)
    common = invoice_set & alma_set
    union = invoice_set | alma_set
    if word_idf:
        common_w = sum(word_idf.get(w, 1.0) for w in common)
        union_w = sum(word_idf.get(w, 1.0) for w in union)
        min_w = min(sum(word_idf.get(w, 1.0) for w in invoice_set), sum(word_idf.get(w, 1.0) for w in alma_set))
        jaccard = common_w / union_w if union_w else 0.0
        containment = common_w / min_w if min_w else 0.0
    else:
        jaccard = len(common) / len(union) if union else 0.0
        containment = len(common) / min(len(invoice_set), len(alma_set))
    sequence = SequenceMatcher(None, " ".join(invoice_words), " ".join(alma_words)).ratio()
    fuzzy_hits = 0
    for word in invoice_set:
        if len(word) < 4:
            continue
        if any(SequenceMatcher(None, word, other).ratio() >= 0.84 for other in alma_set if len(other) >= 4):
            fuzzy_hits += 1
    fuzzy = fuzzy_hits / max(1, min(len(invoice_set), len(alma_set)))
    score = 0.38 * containment + 0.27 * sequence + 0.20 * jaccard + 0.15 * fuzzy
    reasons = [f"lexico={score:.3f}"]

    main_alma_words = getattr(record, "main_title_words", None) or normalize_title_for_matching(extract_main_title(record.title))
    main_inv_words = normalize_title_for_matching(extract_main_title(invoice_title))
    if main_alma_words and main_inv_words and (len(main_alma_words) < len(alma_words) or len(main_inv_words) < len(invoice_words)):
        m_inv_set, m_alma_set = set(main_inv_words), set(main_alma_words)
        m_common = m_inv_set & m_alma_set
        m_union = m_inv_set | m_alma_set
        m_jacc = len(m_common) / len(m_union) if m_union else 0.0
        m_cont = len(m_common) / max(1, min(len(m_inv_set), len(m_alma_set)))
        m_seq = SequenceMatcher(None, " ".join(main_inv_words), " ".join(main_alma_words)).ratio()
        main_score = 0.42 * m_cont + 0.33 * m_seq + 0.25 * m_jacc
        if main_score > score:
            score = 0.75 * main_score + 0.25 * score
            reasons.append(f"titulo_principal={main_score:.3f}")

    if "/" in str(record.title or "") and "-" in str(invoice_title or ""):
        alma_resp = normalize_title_for_matching(str(record.title).split("/", 1)[1])
        inv_prefix = normalize_title_for_matching(str(invoice_title).split("-", 1)[0])
        if set(alma_resp) & set(inv_prefix):
            score += 0.08
            reasons.append("autor coincidente")

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
        return normalize_tax_id(invoice.seller_tax_id)
    if invoice.vendor_account:
        return normalize_provider_for_comparison(invoice.vendor_account)
    return ""

def build_conversion_summary(invoices, validation_issues):
    summary = {
        "total_lines": 0, "match_isbn": 0, "match_issn": 0, "match_title": 0,
        "match_pol": 0, "match_manual": 0,
        "duplicate_isbn": 0, "duplicate_issn": 0, "duplicate_title": 0,
        "ambiguous_title": 0, "no_match": 0,
        "validation_issues": len(validation_issues), "lines_without_po_line": 0,
    }
    for invoice in invoices:
        for line in invoice.lines:
            summary["total_lines"] += 1
            status = line.match_status or ""
            if status == "MATCH_ISBN": summary["match_isbn"] += 1
            elif status == "MATCH_ISSN": summary["match_issn"] += 1
            elif status == "MATCH_TITLE": summary["match_title"] += 1
            elif status == "MATCH_POL": summary["match_pol"] += 1
            elif status == "MATCH_MANUAL": summary["match_manual"] += 1
            elif status == "DUPLICATE_ISBN": summary["duplicate_isbn"] += 1
            elif status == "DUPLICATE_ISSN": summary["duplicate_issn"] += 1
            elif status == "DUPLICATE_TITLE": summary["duplicate_title"] += 1
            elif status == "AMBIGUOUS_TITLE": summary["ambiguous_title"] += 1
            elif status == "NO_MATCH": summary["no_match"] += 1
            if not line.po_line and (line.line_type or "REGULAR") == "REGULAR":
                summary["lines_without_po_line"] += 1
    return summary

_ALMA_LOOKUP_CACHE = {}

def get_cached_alma_lookup(alma_excel_path):
    """Devuelve un AlmaPOLineLookup reutilizable en caché según ruta, tamaño y fecha de modificación."""
    norm_path = os.path.normcase(os.path.abspath(alma_excel_path))
    try:
        st = os.stat(alma_excel_path)
        cache_key = (norm_path, st.st_size, st.st_mtime_ns)
    except OSError:
        cache_key = None
    if cache_key is not None and cache_key in _ALMA_LOOKUP_CACHE:
        return _ALMA_LOOKUP_CACHE[cache_key]
    lookup = AlmaPOLineLookup(alma_excel_path)
    lookup.load()
    if cache_key is not None:
        # Limitar el tamaño de caché en memoria para no acumular ficheros antiguos
        if len(_ALMA_LOOKUP_CACHE) >= 6:
            _ALMA_LOOKUP_CACHE.clear()
        _ALMA_LOOKUP_CACHE[cache_key] = lookup
    return lookup

class AlmaPOLineLookup:
    REQUIRED_COLUMNS = ["PO Line", "Reporting Code", "Secondary Reporting Code", "Tertiary Reporting Code", "Start subs date", "Title", "ISSN", "ISBN", "Quantity for Pricing"]
    def __init__(self, alma_excel_path):
        self.alma_excel_path = alma_excel_path
        self.records = []
        self.isbn_index = {}
        self.issn_index = {}
        self.po_line_index = {}
        self.word_doc_freq = {}
        self.word_idf = {}
    def load(self):
        wb = None
        try:
            wb = load_workbook(self.alma_excel_path, read_only=True, data_only=True)
            ws = wb.active
            rows_iter = ws.iter_rows(values_only=True)
            header_row = next(rows_iter, None)
            if not header_row:
                raise ValueError("El Excel de Alma está vacío y no contiene encabezados.")
            headers = [clean_cell(val) or "" for val in header_row]
            self._validate_headers(headers)
            col_map = {}
            col_map_multi = {}
            for idx, header in enumerate(headers):
                normalized = normalize_header(header)
                if normalized not in col_map:
                    col_map[normalized] = idx
                col_map_multi.setdefault(normalized, []).append(idx)
            fund_columns = col_map_multi.get("Fund and percent", [])
            price_header = next((header for header in (
                "PO Line Price", "Net Price", "Price", "Amount", "Line Amount", "Encumbrance"
            ) if header in col_map), None)

            idx_po = col_map["PO Line"]
            idx_title = col_map["Title"]
            idx_rc1 = col_map["Reporting Code"]
            idx_rc2 = col_map["Secondary Reporting Code"]
            idx_rc3 = col_map["Tertiary Reporting Code"]
            idx_rc4 = col_map.get("Fourth Reporting Code")
            idx_rc5 = col_map.get("Fifth Reporting Code")
            idx_start = col_map["Start subs date"]
            idx_issn = col_map["ISSN"]
            idx_qty = col_map["Quantity for Pricing"]
            idx_price = col_map.get(price_header) if price_header else None
            idx_isbn = col_map["ISBN"]

            def _get(row, col_i):
                if col_i is None or col_i >= len(row):
                    return None
                return row[col_i]

            for row in rows_iter:
                if not row:
                    continue
                po_line = clean_cell(_get(row, idx_po))
                if not po_line:
                    continue
                title = clean_cell(_get(row, idx_title))
                title_words = normalize_title_for_matching(title)
                main_words = normalize_title_for_matching(extract_main_title(title))
                funds = []
                for fund_col in fund_columns:
                    fund_value = alma_empty_to_none(_get(row, fund_col))
                    if fund_value:
                        funds.append(fund_value)
                rec = AlmaPOLRecord(
                    po_line=po_line,
                    reporting_code=alma_empty_to_none(_get(row, idx_rc1)),
                    secondary_reporting_code=alma_empty_to_none(_get(row, idx_rc2)),
                    tertiary_reporting_code=alma_empty_to_none(_get(row, idx_rc3)),
                    fourth_reporting_code=alma_empty_to_none(_get(row, idx_rc4)),
                    fifth_reporting_code=alma_empty_to_none(_get(row, idx_rc5)),
                    start_subs_date=parse_date_for_excel(_get(row, idx_start)),
                    title=title,
                    issn_raw=clean_cell(_get(row, idx_issn)),
                    quantity_for_pricing=decimal_or_none(_get(row, idx_qty)),
                    po_price=decimal_or_none(_get(row, idx_price)),
                    isbn_raw=clean_cell(_get(row, idx_isbn)),
                    funds=funds,
                    normalized_title_words=title_words,
                    normalized_title_text=" ".join(title_words),
                    main_title_words=main_words,
                )
                self.records.append(rec)
                self.po_line_index[rec.po_line.upper()] = rec
                self.po_line_index[rec.po_line] = rec
                for word in set(title_words):
                    self.word_doc_freq[word] = self.word_doc_freq.get(word, 0) + 1
                for isbn in split_identifiers(rec.isbn_raw):
                    for variant in expand_isbn_variants(isbn):
                        if rec not in self.isbn_index.setdefault(variant, []):
                            self.isbn_index[variant].append(rec)
                for issn in split_identifiers(rec.issn_raw):
                    if rec not in self.issn_index.setdefault(issn, []):
                        self.issn_index[issn].append(rec)
            total_docs = max(1, len(self.records))
            self.word_idf = {
                w: math.log(1.0 + total_docs / (1.0 + df))
                for w, df in self.word_doc_freq.items()
            }
            logging.info(f"Excel Alma cargado: {len(self.records)} registros, {len(self.isbn_index)} ISBN indexados, {len(self.issn_index)} ISSN indexados")
        except ValueError:
            raise
        except Exception as e:
            logging.exception("Error cargando Excel de Alma")
            raise ValueError(f"No se pudo leer el Excel de Alma con PO Lines: {e}") from e
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception as close_exc:
                    logging.debug("No se pudo cerrar el workbook de Alma: %s", close_exc)
    def _validate_headers(self, headers):
        normalized = {normalize_header(h) for h in headers}
        missing = [c for c in self.REQUIRED_COLUMNS if c not in normalized]
        if missing:
            raise ValueError("El Excel de Alma no contiene las columnas esperadas. Faltan: " + ", ".join(missing))
    def find_by_po_line(self, po_line):
        cleaned = clean_cell(po_line)
        if not cleaned:
            return None
        return self.po_line_index.get(cleaned) or self.po_line_index.get(cleaned.upper())

    def find_by_isbn(self, isbn):
        results = []
        for variant in expand_isbn_variants(isbn):
            for rec in self.isbn_index.get(variant, []):
                if rec not in results:
                    results.append(rec)
        return results

    def find_by_issn(self, issn):
        return self.issn_index.get(normalize_identifier(issn), []) if normalize_identifier(issn) else []

    def rank_title_candidates(self, title, limit=8):
        ranked = []
        for record in self.records:
            score, reasons = diagnostic_title_score(title, record, self.word_idf)
            if score >= 0.20:
                ranked.append((record, score, reasons))
        ranked.sort(key=lambda item: item[1], reverse=True)
        return ranked[:limit]

    def search_catalog(self, query="", ref_title="", limit=30):
        """Busqueda interactiva multicampo por PO Line, ISBN, ISSN, titulo o reporting code."""
        query = (query or "").strip()
        if not query:
            if ref_title:
                ranked = self.rank_title_candidates(ref_title, limit=limit)
                if ranked:
                    return ranked
            return [(rec, 0.0, ["catalogo"]) for rec in self.records[:limit]]

        q_norm = remove_accents(query.lower()).strip()
        q_id = normalize_identifier(query)
        q_words = normalize_title_for_matching(query)
        results = []
        for rec in self.records:
            score = 0.0
            reasons = []
            if q_norm in rec.po_line.lower():
                score = 1.0 if q_norm == rec.po_line.lower() else 0.95
                reasons.append("PO Line")
            if q_id and len(q_id) >= 4:
                if rec.isbn_raw and q_id in normalize_identifier(rec.isbn_raw):
                    score = max(score, 0.98)
                    reasons.append("ISBN")
                if rec.issn_raw and q_id in normalize_identifier(rec.issn_raw):
                    score = max(score, 0.98)
                    reasons.append("ISSN")
            if rec.title:
                t_norm = remove_accents(rec.title.lower())
                if q_norm in t_norm:
                    score = max(score, 0.88)
                    reasons.append("substring titulo")
                elif q_words:
                    t_score, t_reasons = diagnostic_title_score(query, rec, self.word_idf)
                    if t_score >= 0.22:
                        score = max(score, t_score)
                        reasons.extend(t_reasons)
            for rc in (rec.reporting_code, rec.secondary_reporting_code, rec.tertiary_reporting_code):
                if rc and q_norm in rc.lower():
                    score = max(score, 0.75)
                    reasons.append(f"codigo={rc}")
            if score > 0:
                results.append((rec, score, reasons))
        results.sort(key=lambda item: (item[1], item[0].po_line), reverse=True)
        return results[:limit]

    def find_by_title(self, title):
        if not title:
            return {"status": "NO_MATCH", "candidates": [], "scores": [], "method": None}

        invoice_words = normalize_title_for_matching(title)
        if not invoice_words:
            return {"status": "NO_MATCH", "candidates": [], "scores": [], "method": None}

        main_inv_words = normalize_title_for_matching(extract_main_title(title))

        # 1. Coincidencia exacta tras normalizar mayusculas, tildes, abreviaturas y titulo principal.
        exact = self._unique_by_po_line([
            record for record in self.records
            if record.normalized_title_words == invoice_words
            or (main_inv_words and len(main_inv_words) >= 2 and record.main_title_words == main_inv_words)
        ])
        if exact:
            return self._title_result(exact, [1.0] * len(exact), "EXACT")

        # 2. Inclusion segura por secuencias completas de palabras.
        contained = []
        for record in self.records:
            alma_words = record.normalized_title_words
            if not alma_words:
                continue
            score = title_containment_score(invoice_words, alma_words)
            if score == 0.0 and record.main_title_words and len(record.main_title_words) >= 2:
                score = title_containment_score(invoice_words, record.main_title_words)
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

        # 3. Algoritmo de similitud lexico.
        scored = []
        for record in self.records:
            if record.normalized_title_words:
                score = title_similarity_score(title, record)
                if score >= TITLE_MATCH_MIN_SCORE:
                    scored.append((record, score))
        scored.sort(key=lambda item: item[1], reverse=True)
        if scored:
            unique_records, unique_scores = self._unique_scored_by_po_line(scored)
            if len(unique_records) == 1:
                return {"status": "MATCH", "candidates": unique_records, "scores": unique_scores, "method": "SIMILARITY"}
            best_score = unique_scores[0]
            if best_score - unique_scores[1] >= TITLE_MATCH_MIN_MARGIN:
                return {"status": "MATCH", "candidates": [unique_records[0]], "scores": [best_score], "method": "SIMILARITY"}
            ambiguous = [
                (record, score)
                for record, score in zip(unique_records, unique_scores, strict=True)
                if best_score - score < TITLE_MATCH_MIN_MARGIN
            ]
            return {
                "status": "AMBIGUOUS",
                "candidates": [record for record, _ in ambiguous],
                "scores": [score for _, score in ambiguous],
                "method": "SIMILARITY",
            }

        # 4. Puntuacion diagnostica ponderada por IDF y titulo principal MARC 245 / ISBD.
        diag_scored = []
        for record in self.records:
            if record.normalized_title_words:
                d_score, _ = diagnostic_title_score(title, record, self.word_idf)
                if d_score >= 0.72:
                    diag_scored.append((record, d_score))
        diag_scored.sort(key=lambda item: item[1], reverse=True)
        if not diag_scored:
            return {"status": "NO_MATCH", "candidates": [], "scores": [], "method": None}

        u_records, u_scores = self._unique_scored_by_po_line(diag_scored)
        if len(u_records) == 1:
            return {"status": "MATCH", "candidates": u_records, "scores": u_scores, "method": "DIAGNOSTIC_IDF"}
        best_d = u_scores[0]
        if best_d - u_scores[1] >= TITLE_MATCH_MIN_MARGIN:
            return {"status": "MATCH", "candidates": [u_records[0]], "scores": [best_d], "method": "DIAGNOSTIC_IDF"}
        ambiguous_d = [
            (rec, sc) for rec, sc in zip(u_records, u_scores, strict=True)
            if best_d - sc < TITLE_MATCH_MIN_MARGIN
        ]
        return {
            "status": "AMBIGUOUS",
            "candidates": [rec for rec, _ in ambiguous_d],
            "scores": [sc for _, sc in ambiguous_d],
            "method": "DIAGNOSTIC_IDF",
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
    def __init__(self, invoice_path, custom_vendor_map=None):
        self.invoice_path = invoice_path
        self.custom_vendor_map = custom_vendor_map
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
            raise ValueError("El archivo XML/XSIG contiene estructuras XML no permitidas por seguridad y no se procesara. " + str(e)) from e
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
                raise ValueError("El archivo XML/XSIG contiene estructuras XML no permitidas por seguridad y no se procesara. " + str(e)) from e
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
        for candidate in [
            self.find_direct_text(legal, "CorporateName"),
            self.find_direct_text(individual, "CorporateName"),
            self.find_direct_text(legal, "TradeName"),
            self.find_direct_text(individual, "TradeName"),
        ]:
            if candidate:
                return candidate
        parts = [
            self.find_direct_text(individual, "Name"),
            self.find_direct_text(individual, "FirstSurname"),
            self.find_direct_text(individual, "SecondSurname"),
        ]
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
        vendor_account = get_vendor_account(seller_tax_id, seller_name, self.custom_vendor_map)
        # Gestión exacta de importes e IVA de FacturaeToAlma 4.0
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
        article_code = self.find_text(line, "ArticleCode") or None
        xml_po_line = extract_po_line_from_text(
            self.find_text(line, "ReceiverTransactionReference"),
            self.find_text(line, "IssuerTransactionReference"),
            self.find_text(line, "FileReference"),
            self.find_text(line, "ReceiverContractReference"),
            article_code,
            raw_title,
        )
        desc_isbn, desc_issn, cleaned_title = extract_identifiers_from_text(raw_title)
        title = cleaned_title or raw_title
        # Extracción exacta de importes e IVA de línea de FacturaeToAlma 4.0
        quantity = safe_decimal(self.find_text(line, "Quantity", "1"), Decimal("1.00"))
        gross_amount_text = self.find_path_text(line, "GrossAmount")
        unit_price_text = self.find_text(line, "UnitPriceWithoutTax")
        price = safe_decimal(gross_amount_text, Decimal("0.00")) if gross_amount_text else safe_decimal(unit_price_text, Decimal("0.00"))
        issn, isbn, vendor_ref = classify_article_code(article_code)
        if not isbn and desc_isbn:
            isbn = desc_isbn
        if not issn and desc_issn:
            issn = desc_issn
        line_type = detect_special_line_type(title, price=price, isbn=isbn, issn=issn)
        return InvoiceLineData(
            line_number=seq_num,
            po_line=None,
            line_type=line_type,
            vendor_ref=vendor_ref,
            issn=issn,
            isbn=isbn,
            title=title,
            price=price,
            quantity=quantity,
            vat_amount=decimal_or_none(self.find_path_text(line, "TaxesOutputs/Tax/TaxAmount/TotalAmount")),
            vat_percentage=decimal_or_none(self.find_path_text(line, "TaxesOutputs/Tax/TaxRate")),
            xml_po_line=xml_po_line,
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

    def _check_invoice_accounting_balance(self, invoice):
        if not invoice.lines or invoice.total_amount is None:
            return
        sum_prices = sum((ln.price or Decimal("0.00")) for ln in invoice.lines)
        sum_vats = sum((ln.vat_amount or Decimal("0.00")) for ln in invoice.lines)
        header_vat = invoice.vat_amount if invoice.vat_amount is not None else sum_vats
        use_vat = sum_vats if any(ln.vat_amount is not None for ln in invoice.lines) else (header_vat or Decimal("0.00"))
        discounts = getattr(invoice, "general_discounts", Decimal("0.00")) or Decimal("0.00")
        surcharges = getattr(invoice, "general_surcharges", Decimal("0.00")) or Decimal("0.00")
        computed_total = (sum_prices - discounts + surcharges + use_vat).quantize(Decimal("0.01"))
        diff = abs(computed_total - invoice.total_amount)
        if diff > Decimal("0.05"):
            extra_notes = []
            if discounts > Decimal("0.00"):
                extra_notes.append(f"Descuentos generales cabecera={discounts}")
            if surcharges > Decimal("0.00"):
                extra_notes.append(f"Cargos generales cabecera={surcharges}")
            extra_str = (" (" + "; ".join(extra_notes) + ")") if extra_notes else ""
            self.validation_issues.append(
                ValidationIssue(
                    issue_type="REVISAR CUADRE IMPORTE",
                    invoice_number=invoice.invoice_number,
                    invoice_line_number=0,
                    isbn=None,
                    issn=None,
                    title="Cabecera vs. Suma de lineas",
                    candidates=[],
                    message=(
                        f"El Total amount de la factura es {invoice.total_amount}, mientras que la suma "
                        f"de lineas (base={sum_prices} + IVA={use_vat}){extra_str} es {computed_total} "
                        f"(diferencia={diff}). Revisar posibles descuentos, portes o retenciones de cabecera."
                    ),
                )
            )

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
        manual_entry = self.manual_assignments.get((source_key, line.line_number))
        if manual_entry:
            if isinstance(manual_entry, tuple):
                manual_po_line, copy_analytics = manual_entry[0], bool(manual_entry[1])
            else:
                manual_po_line, copy_analytics = str(manual_entry).strip(), False
            if manual_po_line:
                manual_record = self.lookup.find_by_po_line(manual_po_line)
                if manual_record and copy_analytics:
                    self._resolve_candidates(invoice, line, [manual_record], "MANUAL")
                    alma_qty = manual_record.quantity_for_pricing
                else:
                    # Regla del usuario: una PO Line introducida a mano NO sobreescribe
                    # reporting codes ni fondos con valores supuestos.
                    line.po_line = manual_po_line
                    line.match_status = "MATCH_MANUAL"
                    line.match_message = "PO Line asignada manualmente"
                    line.match_candidates = [manual_po_line]
                    alma_qty = manual_record.quantity_for_pricing if manual_record else None
                self.validation_issues.append(ValidationIssue(
                    "ASIGNACION MANUAL", invoice.invoice_number, line.line_number,
                    line.isbn, line.issn, line.title, [manual_po_line],
                    "PO Line asignada manualmente durante esta sesion; el XML original no ha sido modificado.",
                    po_line=manual_po_line,
                    invoice_quantity=line.quantity,
                    alma_quantity_for_pricing=alma_qty,
                ))
                return

        if getattr(line, "xml_po_line", None):
            xml_rec = self.lookup.find_by_po_line(line.xml_po_line)
            if xml_rec:
                self._resolve_candidates(invoice, line, [xml_rec], "POL")
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
                score_text = "; ".join(f"{r.po_line}: {s:.3f}" for r, s in zip(result["candidates"], result["scores"], strict=True))
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
        if (line.line_type or "REGULAR") != "REGULAR":
            line.match_status = f"SPECIAL_{line.line_type}"
            line.match_message = f"Linea de tipo {line.line_type} (no requiere PO Line bibliografica)"
            return
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
        except PermissionError as e:
            logging.exception("No se puede guardar el Excel porque esta abierto o bloqueado")
            raise ValueError(f"No se pudo guardar el archivo '{self.output_path}'. Puede que este abierto en Excel.") from e
        except Exception as e:
            logging.exception("Error guardando Excel")
            raise ValueError(f"Error al guardar el Excel: {e}") from e

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

def parse_invoice_paths(invoice_paths, invoice_number_overrides=None, custom_vendor_map=None):
    if isinstance(invoice_paths, str):
        invoice_paths = [invoice_paths]
    invoices = []
    invoice_number_overrides = invoice_number_overrides or {}
    for p in invoice_paths:
        parsed = FacturaeParser(p, custom_vendor_map=custom_vendor_map).parse()
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
        label = inv.vendor_account or inv.seller_name or tr("vw_unknown_vendor")
        keys.setdefault(key, set()).add(label)
    meaningful_keys = {k for k in keys if k}
    if len(meaningful_keys) <= 1:
        return
    details = []
    for labels in keys.values():
        for label in sorted(labels):
            details.append(f"- {label}")
    raise ValueError(tr("val_multi_vendor", details="\n".join(details)))

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
        for field_name, value in header_changes.items():
            normalized = normalize_header(header_map.get(field_name, ""))
            columns = hinv_columns.get(normalized, [])
            if not columns:
                raise ValueError(f"No existe la columna de cabecera necesaria para guardar '{field_name}'.")
            cell = worksheet.cell(row=inv_row, column=columns[0])
            cell.value = sanitize_excel_value(value)
            expected.append((worksheet.title, inv_row, columns[0], cell.value, f"cabecera {field_name}"))

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
        by_order = {line.line_number: row for line, row in zip(invoice.lines, il_rows, strict=False)}

        for (override_source, line_number), changes in line_overrides.items():
            if override_source != source_key:
                continue
            row = by_number.get(line_number) or by_order.get(line_number)
            if row is None:
                raise ValueError(f"No se encontro la linea {line_number} de la factura {invoice.invoice_number} en el Excel final.")
            for field_name, value in changes.items():
                if field_name in {"fund_1", "fund_2"}:
                    columns = hil_columns.get(normalize_header("Fund and percent"), [])
                    occurrence = int(field_name.split("_", 1)[1]) - 1
                    if occurrence >= len(columns):
                        raise ValueError(f"No existe la columna necesaria para guardar '{field_name}'.")
                    column = columns[occurrence]
                else:
                    normalized = normalize_header(line_map.get(field_name, ""))
                    columns = hil_columns.get(normalized, [])
                    if not columns:
                        raise ValueError(f"No existe la columna necesaria para guardar '{field_name}'.")
                    column = columns[0]
                if field_name in {"price", "quantity", "vat_amount", "vat_percentage"} and value not in (None, ""):
                    value = float(safe_decimal(value))
                elif field_name in {"start_date", "end_date"} and value:
                    value = parse_date_for_excel(value)
                cell = worksheet.cell(row=row, column=column)
                cell.value = sanitize_excel_value(value)
                expected.append((worksheet.title, row, column, cell.value, f"linea {line_number} {field_name}"))

    audit_name = "Cambios FacturaeToAlma"
    if audit_name in workbook.sheetnames:
        del workbook[audit_name]
    audit = workbook.create_sheet(audit_name)
    audit.append(["Tipo", "Factura/Fichero", "Linea", "Campo", "Valor guardado"])
    for invoice in invoices:
        source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
        if source_key in invoice_number_overrides:
            audit.append(["Cabecera", os.path.basename(invoice.source_file or ""), "", "Invoice Number", invoice_number_overrides[source_key]])
        for field_name, value in invoice_header_overrides.get(source_key, {}).items():
            audit.append(["Cabecera", os.path.basename(invoice.source_file or ""), "", field_name, value])
        for (override_source, line_number), changes in line_overrides.items():
            if override_source != source_key:
                continue
            for field_name, value in changes.items():
                audit.append(["Linea", os.path.basename(invoice.source_file or ""), line_number, field_name, value])
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


def convert_facturae_to_alma_excel(invoice_paths, alma_pol_excel_path, excel_output_path, template_path=None, invoice_number_overrides=None, manual_assignments=None, line_overrides=None, invoice_header_overrides=None, custom_vendor_map=None):
    invoices = parse_invoice_paths(invoice_paths, invoice_number_overrides, custom_vendor_map=custom_vendor_map)
    if not invoices:
        raise ValueError("No se encontro ninguna factura para procesar.")
    validate_single_vendor_batch(invoices)
    lookup = get_cached_alma_lookup(alma_pol_excel_path)
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
        lang_code = set_current_language(self.preferences.get("language", "es"))
        self.preferences["language"] = lang_code
        self.language_label_var = tk.StringVar(value=LANGUAGE_CODE_TO_LABEL.get(lang_code, LANGUAGE_CODE_TO_LABEL["es"]))
        self.conversion_mode = tk.StringVar(value="single")
        self.invoice_file = tk.StringVar()
        self.alma_pol_file = tk.StringVar()
        self.template_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self._last_status = ("st_pending", {})
        self.status_text = tk.StringVar(value=tr("st_pending"))
        self.institution_var = tk.StringVar(value=self.preferences.get("institution", self.DEFAULT_INSTITUTION))
        self.custom_invoice_dir = tk.StringVar(value=self.preferences.get("custom_invoice_dir", ""))
        self.custom_alma_dir = tk.StringVar(value=self.preferences.get("custom_alma_dir", ""))
        self.custom_template_dir = tk.StringVar(value=self.preferences.get("custom_template_dir", ""))
        self.custom_output_dir = tk.StringVar(value=self.preferences.get("custom_output_dir", ""))
        self.logo_path_var = tk.StringVar(value=self.preferences.get("logo_path", ""))
        self._logo_status_state = ("cust_logo_hint", {})
        self.logo_photo = None
        self.invoice_files = []
        self.invoice_metadata = {}
        self.invoice_number_overrides = {}
        self.invoice_number_mode_by_source = {}
        self._updating_number_controls = False
        self.manual_po_assignments = {}
        self.line_overrides = {}
        self.invoice_header_overrides = {}
        self.edit_original_values = {}
        self.edit_line_diagnostics = {}
        self.edit_invoice_reconciliation = {}
        self.edit_header_original_values = {}
        self.edit_invoice_by_label = {}
        self.edit_lookup = None
        self.edit_filter_code = "all"
        self.edit_filter_var = tk.StringVar(value=tr("ed_flt_st_all"))
        self.edit_stats_var = tk.StringVar(value=tr("ed_stats_empty"))
        self.edit_line_diag_var = tk.StringVar(value=tr("ed_diag_init"))
        self.edit_totals_var = tk.StringVar(value=tr("ed_tot_init"))
        self.edit_quick_search_var = tk.StringVar(value="")
        self.filter_col_key = "po_line"
        self.filter_col_var = tk.StringVar(value="PO Line")
        self.filter_op_code = "contains"
        self.filter_op_var = tk.StringVar(value=tr("ed_flt_op_contains"))
        self._syncing_lang_combos = False
        self.edit_filter_var.trace_add("write", self._on_edit_filter_var_trace)
        self.filter_val_var = tk.StringVar(value="")
        self.active_filters_label_var = tk.StringVar(value=tr("ed_flt_none"))
        self.active_column_filters = {}
        self.vendor_import_mode_code = "incremental"
        self.vendor_import_mode_var = tk.StringVar(value=tr("cust_vend_mode_inc"))
        self.last_edited_column = "po_line"
        self.active_cell = None
        self.selection_anchor = None
        self._clipboard_cell_value = None
        self.undo_stack = []
        self.redo_stack = []
        self.edit_active_editor = None
        self.edit_active_context = None
        self.header_active_editor = None
        self.edit_combo_editor = None
        self.edit_combo_context = None
        self.header_cell_editor = None
        self.header_editor_context = None
        self.edit_draft_signature = None
        self._last_draft_error = ""
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

    def _set_status(self, key, **kwargs):
        self._last_status = (key, kwargs)
        self.status_text.set(tr(key, **kwargs))

    def _set_logo_status(self, key, **kwargs):
        self._logo_status_state = (key, kwargs)
        if hasattr(self, "logo_status_label"):
            self.logo_status_label.config(text=tr(key, **kwargs))

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
            "language": "es",
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
            "vendor_map": {},
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
        if not isinstance(data.get("vendor_map"), dict):
            data["vendor_map"] = {}
        if data.get("language") not in LANGUAGE_CODE_TO_LABEL:
            data["language"] = "es"
        return data

    def _save_preferences_file(self):
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.preferences, f, ensure_ascii=False, indent=2)
        except OSError as e:
            logging.exception("No se pudo guardar la configuracion")
            raise ValueError(f"No se pudo guardar la configuracion: {e}") from e

    def _remember_vendor_for_nif(self, seller_tax_id, vendor_account):
        norm_nif = normalize_tax_id(seller_tax_id)
        clean_vendor = (vendor_account or "").strip()
        if not norm_nif or not clean_vendor:
            return
        vendor_map = self.preferences.setdefault("vendor_map", {})
        if vendor_map.get(norm_nif) != clean_vendor:
            vendor_map[norm_nif] = clean_vendor
            try:
                self._save_preferences_file()
                logging.info("Recordado Vendor Account '%s' para NIF '%s'", clean_vendor, norm_nif)
            except ValueError:
                logging.warning("No se pudo guardar el mapeo NIF -> Vendor Account en preferencias")
            self._refresh_vendor_map_tree()

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
        self.customize_tab = ttk.Frame(self.notebook, padding=14, style="App.TFrame")
        self.help_tab = ttk.Frame(self.notebook, padding=18, style="App.TFrame")
        self.notebook.add(self.convert_tab, text=tr("tab_convert"))
        self.notebook.add(self.visualize_tab, text=tr("tab_visualize"))
        self.notebook.add(self.edit_tab, text=tr("tab_edit"))
        self.notebook.add(self.customize_tab, text=tr("tab_customize"))
        self.notebook.add(self.help_tab, text=tr("tab_help"))
        self._build_convert_tab()
        self._build_visualize_tab()
        self._build_edit_tab()
        self._build_customize_tab()
        self._build_help_tab()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _on_tab_changed(self, _event=None):
        if hasattr(self, "_close_all_cell_editors"):
            self._close_all_cell_editors(commit=True)
        try:
            current_tab = self.notebook.nametowidget(self.notebook.select())
        except Exception:
            return
        if current_tab == self.edit_tab:
            if (
                self.invoice_files
                and self.edit_draft_signature != self._current_edit_signature()
            ):
                self.prepare_edit_draft(silent=True)

    def _build_convert_tab(self):
        tab = self.convert_tab
        tab.columnconfigure(0, weight=1)
        self.conv_mode_lf = ttk.LabelFrame(tab, text=tr("conv_mode_frame"), padding=10, style="App.TLabelframe")
        self.conv_mode_lf.grid(row=0, column=0, sticky="ew")
        self.conv_single_rb = ttk.Radiobutton(self.conv_mode_lf, text=tr("conv_mode_single"), variable=self.conversion_mode, value="single", command=self.on_mode_change, style="App.TRadiobutton")
        self.conv_single_rb.grid(row=0, column=0, padx=(0, 24), sticky="w")
        self.conv_batch_rb = ttk.Radiobutton(self.conv_mode_lf, text=tr("conv_mode_batch"), variable=self.conversion_mode, value="batch", command=self.on_mode_change, style="App.TRadiobutton")
        self.conv_batch_rb.grid(row=0, column=1, sticky="w")

        self.conv_input_lf = ttk.LabelFrame(tab, text=tr("conv_inputs_frame"), padding=10, style="App.TLabelframe")
        self.conv_input_lf.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.conv_input_lf.columnconfigure(0, weight=1)
        self.invoice_label = ttk.Label(self.conv_input_lf, text=tr("conv_inv_single_lbl"), style="App.TLabel")
        self.invoice_label.grid(row=0, column=0, sticky="w")
        ttk.Entry(self.conv_input_lf, textvariable=self.invoice_file, state="readonly").grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(2, 8))
        self.invoice_button = ttk.Button(self.conv_input_lf, text=tr("conv_inv_single_btn"), command=self.select_invoice_or_batch)
        self.invoice_button.grid(row=1, column=1, sticky="ew", pady=(2, 8))
        self.conv_alma_lbl = ttk.Label(self.conv_input_lf, text=tr("conv_alma_lbl"), style="App.TLabel")
        self.conv_alma_lbl.grid(row=2, column=0, sticky="w")
        ttk.Entry(self.conv_input_lf, textvariable=self.alma_pol_file).grid(row=3, column=0, sticky="ew", padx=(0, 10), pady=(2, 8))
        self.conv_alma_btn = ttk.Button(self.conv_input_lf, text=tr("conv_alma_btn"), command=self.select_alma_pol)
        self.conv_alma_btn.grid(row=3, column=1, sticky="ew", pady=(2, 8))
        self.conv_tpl_lbl = ttk.Label(self.conv_input_lf, text=tr("conv_tpl_lbl"), style="App.TLabel")
        self.conv_tpl_lbl.grid(row=4, column=0, sticky="w")
        ttk.Entry(self.conv_input_lf, textvariable=self.template_file).grid(row=5, column=0, sticky="ew", padx=(0, 10), pady=(2, 0))
        self.conv_tpl_btn = ttk.Button(self.conv_input_lf, text=tr("conv_tpl_btn"), command=self.select_template)
        self.conv_tpl_btn.grid(row=5, column=1, sticky="ew", pady=(2, 0))

        self.batch_frame = ttk.LabelFrame(tab, text=tr("conv_batch_frame"), padding=8, style="App.TLabelframe")
        self.batch_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        self.batch_frame.columnconfigure(0, weight=1)
        self.batch_frame.rowconfigure(0, weight=1)
        cols = ("invoice", "vendor", "file")
        self.batch_tree = ttk.Treeview(self.batch_frame, columns=cols, show="headings", height=5, selectmode="extended")
        self.batch_tree.heading("invoice", text=tr("conv_col_invoice"))
        self.batch_tree.heading("vendor", text=tr("conv_col_vendor"))
        self.batch_tree.heading("file", text=tr("conv_col_file"))
        self.batch_tree.column("invoice", width=150, anchor="w")
        self.batch_tree.column("vendor", width=310, anchor="w")
        self.batch_tree.column("file", width=520, anchor="w")
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        self.batch_tree.bind("<Double-1>", self._on_batch_tree_double_click)
        scroll = ttk.Scrollbar(self.batch_frame, orient="vertical", command=self.batch_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.batch_tree.configure(yscrollcommand=scroll.set)
        batch_buttons = ttk.Frame(self.batch_frame, style="App.TFrame")
        batch_buttons.grid(row=1, column=0, columnspan=2, sticky="w", pady=(7, 0))
        self.batch_add_btn = ttk.Button(batch_buttons, text=tr("conv_batch_add"), command=self.add_batch_files)
        self.batch_add_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.batch_rm_btn = ttk.Button(batch_buttons, text=tr("conv_batch_remove"), command=self.remove_selected_batch_files)
        self.batch_rm_btn.pack(side=tk.LEFT, padx=6)
        self.batch_sort_btn = ttk.Button(batch_buttons, text=tr("conv_batch_sort"), command=self.sort_batch_files)
        self.batch_sort_btn.pack(side=tk.LEFT, padx=6)
        self.batch_clr_btn = ttk.Button(batch_buttons, text=tr("conv_batch_clear"), command=self.clear_batch_files)
        self.batch_clr_btn.pack(side=tk.LEFT, padx=6)
        self.batch_hint_lbl = ttk.Label(batch_buttons, text=tr("conv_batch_hint"), style="App.TLabel", font=("Segoe UI", 8, "italic"))
        self.batch_hint_lbl.pack(side=tk.LEFT, padx=(10, 0))

        self.conv_output_lf = ttk.LabelFrame(tab, text=tr("conv_output_frame"), padding=10, style="App.TLabelframe")
        self.conv_output_lf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.conv_output_lf.columnconfigure(0, weight=1)
        ttk.Entry(self.conv_output_lf, textvariable=self.output_file).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.conv_output_btn = ttk.Button(self.conv_output_lf, text=tr("conv_output_btn"), command=self.select_output)
        self.conv_output_btn.grid(row=0, column=1)

        self.conv_actions_lf = ttk.LabelFrame(tab, text=tr("conv_actions_frame"), padding=10, style="App.TLabelframe")
        self.conv_actions_lf.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        self.conv_actions_lf.columnconfigure(0, weight=2)
        for c in (1, 2, 3):
            self.conv_actions_lf.columnconfigure(c, weight=1)
        self.convert_button = ttk.Button(self.conv_actions_lf, text=tr("conv_run_btn"), command=self.run_conversion, style="Primary.TButton")
        self.convert_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.conv_clear_btn = ttk.Button(self.conv_actions_lf, text=tr("conv_clear_btn"), command=self.clear_fields)
        self.conv_clear_btn.grid(row=0, column=1, sticky="ew", padx=4)
        self.conv_open_out_btn = ttk.Button(self.conv_actions_lf, text=tr("conv_open_out_btn"), command=self.open_output_folder)
        self.conv_open_out_btn.grid(row=0, column=2, sticky="ew", padx=4)
        self.conv_exit_btn = ttk.Button(self.conv_actions_lf, text=tr("conv_exit_btn"), command=self._on_close)
        self.conv_exit_btn.grid(row=0, column=3, sticky="ew", padx=(4, 0))
        ttk.Label(tab, textvariable=self.status_text, style="App.TLabel", font=("Segoe UI", 9, "italic")).grid(row=5, column=0, sticky="w", pady=(10, 0))
        tab.rowconfigure(2, weight=1)
        tab.rowconfigure(3, weight=0)
        self.on_mode_change(initial=True)

    def _on_batch_tree_double_click(self, _event=None):
        sel = self.batch_tree.selection()
        if not sel:
            return
        path = sel[0]
        if path not in self.invoice_files:
            return
        self._refresh_edit_selector(select_path=path)
        self.notebook.select(self.edit_tab)
        self._set_status("st_open_in_edit", name=os.path.basename(path))

    def _build_visualize_tab(self):
        tab = self.visualize_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        controls = ttk.Frame(tab, style="App.TFrame")
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        controls.columnconfigure(1, weight=1)
        self.vis_inv_lbl = ttk.Label(controls, text=tr("vis_inv_lbl"), style="App.TLabel", font=("Segoe UI", 9, "bold"))
        self.vis_inv_lbl.grid(row=0, column=0, padx=(0, 8))
        self.viewer_invoice_var = tk.StringVar()
        self.viewer_selector = ttk.Combobox(controls, textvariable=self.viewer_invoice_var, state="readonly")
        self.viewer_selector.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.viewer_selector.bind("<<ComboboxSelected>>", self._on_viewer_invoice_selected)
        self.invoice_viewer = FacturaeViewerTab(tab, show_loader=False)
        self.invoice_viewer.grid(row=1, column=0, sticky="nsew")
        self.vis_footer_lbl = ttk.Label(tab, text=tr("vis_footer_note"), style="App.TLabel")
        self.vis_footer_lbl.grid(row=2, column=0, sticky="w", pady=(6, 0))

    def _refresh_viewer_selector(self, select_path=None):
        labels = []
        self.viewer_path_by_label = {}
        for path in self.invoice_files:
            meta = self.invoice_metadata.get(path, {})
            number = meta.get("number") or tr("vw_no_number")
            vendor = meta.get("vendor", "")
            label = f"{number} | {vendor} | {os.path.basename(path)}"
            labels.append(label)
            self.viewer_path_by_label[label] = path
        self.viewer_selector["values"] = labels
        if not labels:
            self.viewer_invoice_var.set("")
            if hasattr(self, "invoice_viewer"):
                self.invoice_viewer.clear_view()
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
            messagebox.showinfo(tr("tab_visualize"), tr("mb_no_active_inv"))
            return
        original = self.invoice_metadata.get(path, {}).get("original_number") or self.invoice_metadata.get(path, {}).get("number", "")
        dialog = tk.Toplevel(self.root)
        dialog.title(tr("numdlg_title"))
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(1, weight=1)
        ttk.Label(dialog, text=tr("numdlg_desc"), wraplength=520, justify="left").grid(row=0, column=0, columnspan=2, padx=16, pady=12, sticky="w")
        ttk.Label(dialog, text=tr("numdlg_orig", original=original)).grid(row=1, column=0, columnspan=2, padx=16, sticky="w")
        current_year = str(datetime.now().year)
        prefix = tk.StringVar(value=f"[{current_year}]")
        suffix = tk.StringVar()
        preview = tk.StringVar(value=f"[{current_year}]{original}")

        def update(*_):
            preview.set(f"{prefix.get()}{original}{suffix.get()}")

        prefix.trace_add("write", update)
        suffix.trace_add("write", update)
        ttk.Label(dialog, text=tr("ed_mode_prefix")).grid(row=2, column=0, padx=16, pady=6, sticky="w")
        ttk.Entry(dialog, textvariable=prefix).grid(row=2, column=1, padx=16, pady=6, sticky="ew")
        ttk.Label(dialog, text=tr("ed_mode_suffix")).grid(row=3, column=0, padx=16, pady=6, sticky="w")
        ttk.Entry(dialog, textvariable=suffix).grid(row=3, column=1, padx=16, pady=6, sticky="ew")
        ttk.Label(dialog, textvariable=preview, font=("Segoe UI", 10, "bold")).grid(row=4, column=0, columnspan=2, padx=16, pady=10)

        def accept():
            value = preview.get().strip()
            if not value:
                return
            key = os.path.normcase(os.path.abspath(path))
            self.invoice_number_overrides[key] = value
            if prefix.get() and not suffix.get():
                self.invoice_number_mode_by_source[key] = ("prefix", prefix.get())
            elif suffix.get() and not prefix.get():
                self.invoice_number_mode_by_source[key] = ("suffix", suffix.get())
            self.invoice_metadata[path]["number"] = value
            self.edit_changes_saved = False
            self._refresh_batch_tree()
            self._refresh_viewer_selector(select_path=path)
            self._refresh_edit_selector(select_path=path)
            dialog.destroy()

        buttons = ttk.Frame(dialog)
        buttons.grid(row=5, column=0, columnspan=2, pady=(0, 14))
        ttk.Button(buttons, text=tr("numdlg_apply"), command=accept).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons, text=tr("dlg_cancel"), command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def reset_active_invoice_number(self):
        path = self._active_viewer_path() or self._active_edit_path()
        if not path:
            return False
        key = os.path.normcase(os.path.abspath(path))
        original = self.invoice_metadata.get(path, {}).get("original_number")
        if original:
            self.invoice_metadata[path]["number"] = original
        self.invoice_number_overrides.pop(key, None)
        self.invoice_number_mode_by_source.pop(key, None)
        self._refresh_batch_tree()
        self._refresh_viewer_selector(select_path=path)

    def _rebuild_edit_menus(self):
        if hasattr(self, "header_context_menu"):
            self.header_context_menu.delete(0, tk.END)
            self.header_context_menu.add_command(label=tr("ctx_edit_cell"), command=self._edit_selected_header_cell)
            self.header_context_menu.add_command(label=tr("ctx_restore_cell"), command=self._restore_selected_header_cell)
        if hasattr(self, "edit_context_menu"):
            self.edit_context_menu.delete(0, tk.END)
            self.edit_context_menu.add_command(label=tr("ctx_edit_cell"), command=self._edit_selected_cell)
            self.edit_context_menu.add_command(label=tr("ctx_search_pol"), command=self.open_po_line_assistant)
            self.edit_context_menu.add_command(label=tr("ctx_clear_cell"), command=self._clear_selected_edit_cell)
            self.edit_context_menu.add_command(label=tr("ctx_restore_orig"), command=self._restore_selected_edit_cell)
            self.edit_context_menu.add_separator()
            self.edit_context_menu.add_command(label=tr("ctx_fill_down"), command=self._excel_fill_down)
            self.edit_context_menu.add_command(label=tr("ctx_fill_blanks_down"), command=self._fill_blanks_downward_action)
            self.edit_context_menu.add_command(label=tr("ctx_fill_blanks_sel"), command=self._fill_selected_blanks)
            self.edit_context_menu.add_command(label=tr("ctx_suggest_ctx"), command=self._suggest_selected_context_value)
            self.edit_context_menu.add_command(label=tr("ctx_bulk_apply"), command=self._bulk_apply_selected_value)
            self.edit_context_menu.add_separator()
            self.edit_context_menu.add_command(label=tr("ctx_flt_cell_val"), command=self._filter_by_selected_cell_value)
            self.edit_context_menu.add_command(label=tr("ctx_flt_col_excel"), command=self._open_filter_for_context_column)
            self.edit_context_menu.add_command(label=tr("ctx_flt_clear_all"), command=self.clear_all_edit_filters)
            self.edit_context_menu.add_separator()
            self.edit_context_menu.add_command(label=tr("ctx_undo"), command=self.undo_edit)
            self.edit_context_menu.add_command(label=tr("ctx_redo"), command=self.redo_edit)

    def _build_edit_tab(self):
        tab = self.edit_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)

        toolbar = ttk.Frame(tab, style="App.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        toolbar.columnconfigure(1, weight=1)
        self.ed_inv_lbl = ttk.Label(toolbar, text=tr("vis_inv_lbl"), style="App.TLabel", font=("Segoe UI", 9, "bold"))
        self.ed_inv_lbl.grid(row=0, column=0, padx=(0, 8))
        self.edit_invoice_var = tk.StringVar()
        self.edit_invoice_selector = ttk.Combobox(toolbar, textvariable=self.edit_invoice_var, state="readonly")
        self.edit_invoice_selector.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.edit_invoice_selector.bind("<<ComboboxSelected>>", self._on_edit_invoice_selected)

        ttk.Label(toolbar, textvariable=self.edit_stats_var, style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=(0, 8))
        self.ed_search_pol_btn = ttk.Button(toolbar, text=tr("ed_search_pol_btn"), command=self.open_po_line_assistant)
        self.ed_search_pol_btn.grid(row=0, column=3, padx=(0, 6))
        self.ed_refresh_btn = ttk.Button(toolbar, text=tr("ed_refresh_btn"), command=lambda: self.prepare_edit_draft(silent=False))
        self.ed_refresh_btn.grid(row=0, column=4)

        self.ed_num_lf = ttk.LabelFrame(tab, text=tr("ed_num_frame"), padding=6, style="App.TLabelframe")
        self.ed_num_lf.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        self.ed_num_lf.columnconfigure(6, weight=1)
        self.ed_orig_lbl = ttk.Label(self.ed_num_lf, text=tr("ed_orig_lbl"), style="App.TLabel")
        self.ed_orig_lbl.grid(row=0, column=0, padx=(0, 4))
        self.edit_original_number_var = tk.StringVar()
        ttk.Label(self.ed_num_lf, textvariable=self.edit_original_number_var, style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, padx=(0, 12))
        self.number_mode_var = tk.StringVar(value="original")
        self.ed_rb_orig = ttk.Radiobutton(self.ed_num_lf, text=tr("ed_mode_orig"), variable=self.number_mode_var, value="original", command=self._update_number_preview, style="App.TRadiobutton")
        self.ed_rb_orig.grid(row=0, column=2)
        self.ed_rb_suffix = ttk.Radiobutton(self.ed_num_lf, text=tr("ed_mode_suffix"), variable=self.number_mode_var, value="suffix", command=self._update_number_preview, style="App.TRadiobutton")
        self.ed_rb_suffix.grid(row=0, column=3)
        self.ed_rb_prefix = ttk.Radiobutton(self.ed_num_lf, text=tr("ed_mode_prefix"), variable=self.number_mode_var, value="prefix", command=self._update_number_preview, style="App.TRadiobutton")
        self.ed_rb_prefix.grid(row=0, column=4)
        self.number_affix_var = tk.StringVar(value=f"[{datetime.now().year}]")
        self.number_affix_var.trace_add("write", self._update_number_preview)
        self.number_affix_entry = ttk.Entry(self.ed_num_lf, textvariable=self.number_affix_var, width=14)
        self.number_affix_entry.grid(row=0, column=5, padx=8)
        self.number_preview_var = tk.StringVar()
        ttk.Label(self.ed_num_lf, textvariable=self.number_preview_var, style="App.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky="w")
        self.ed_apply_num_btn = ttk.Button(self.ed_num_lf, text=tr("ed_apply_num_btn"), command=self.apply_edit_invoice_number)
        self.ed_apply_num_btn.grid(row=0, column=7, padx=(8, 0))

        self.ed_hdr_lf = ttk.LabelFrame(tab, text=tr("ed_hdr_frame"), padding=6, style="App.TLabelframe")
        self.ed_hdr_lf.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        self.ed_hdr_lf.columnconfigure(0, weight=1)
        self.header_columns = ("invoice_number", "date", "total_amount", "currency", "vendor_account", "note", "vat_amount", "inclusive")
        header_titles = {
            "invoice_number": "Invoice Number", "date": "Date", "total_amount": "Total amount",
            "currency": "Currency", "vendor_account": "Vendor Account", "note": "Note",
            "vat_amount": "VAT Amount", "inclusive": "Inclusive",
        }
        self.header_tree = ttk.Treeview(self.ed_hdr_lf, columns=self.header_columns, show="headings", height=1, selectmode="browse")
        for col in self.header_columns:
            self.header_tree.heading(col, text=header_titles[col])
            self.header_tree.column(col, width=140 if col != "note" else 260, minwidth=75, anchor="w")
        self.header_tree.grid(row=0, column=0, sticky="ew")
        self.header_tree.bind("<Button-1>", lambda _e: self._close_all_cell_editors(commit=True), add="+")
        self.header_tree.bind("<Double-1>", self._begin_edit_header_cell)
        self.header_tree.bind("<Button-3>", self._show_header_context_menu)
        self.header_context_menu = tk.Menu(self.header_tree, tearoff=0)

        self.edit_columns = (
            "invoice", "line", "line_type", "po_line", "vendor_ref", "issn", "isbn", "title",
            "price", "quantity", "reporting_code", "secondary_reporting_code", "tertiary_reporting_code",
            "fourth_reporting_code", "fifth_reporting_code", "start_date", "end_date", "note",
            "fund_1", "fund_2", "vat_amount", "vat_percentage",
        )
        self.edit_headings = {
            "invoice": tr("ed_col_invoice"), "line": tr("ed_col_line"), "line_type": "Line type", "po_line": "PO Line",
            "vendor_ref": "Vendor Ref Number", "issn": "ISSN", "isbn": "ISBN", "title": "Title",
            "price": "Price", "quantity": "Quantity", "reporting_code": "Reporting Code",
            "secondary_reporting_code": "Secondary Reporting Code", "tertiary_reporting_code": "Tertiary Reporting Code",
            "fourth_reporting_code": "Fourth Reporting Code", "fifth_reporting_code": "Fifth Reporting Code",
            "start_date": "Start subs date", "end_date": "End subs date", "note": "Note",
            "fund_1": "Fund and percent 1", "fund_2": "Fund and percent 2",
            "vat_amount": "VAT Amount", "vat_percentage": "VAT Percentage",
        }
        self.heading_to_col = {label: col for col, label in self.edit_headings.items()}

        self.ed_flt_lf = ttk.LabelFrame(
            tab,
            text=tr("ed_flt_frame"),
            padding=6,
            style="App.TLabelframe",
        )
        self.ed_flt_lf.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        self.ed_flt_lf.columnconfigure(6, weight=1)

        self.ed_flt_state_lbl = ttk.Label(self.ed_flt_lf, text=tr("ed_flt_state_lbl"), style="App.TLabel", font=("Segoe UI", 9, "bold"))
        self.ed_flt_state_lbl.grid(row=0, column=0, padx=(0, 4), sticky="w")
        self.edit_filter_combo = ttk.Combobox(
            self.ed_flt_lf,
            textvariable=self.edit_filter_var,
            values=[tr(STATUS_FILTER_KEYS[c]) for c in STATUS_FILTER_CODES],
            state="readonly",
            width=42,
        )
        self.edit_filter_combo.grid(row=0, column=1, padx=(0, 10), sticky="w")
        self.edit_filter_combo.bind("<<ComboboxSelected>>", self._on_status_filter_selected)

        self.ed_flt_col_lbl = ttk.Label(self.ed_flt_lf, text=tr("ed_flt_col_lbl"), style="App.TLabel", font=("Segoe UI", 9, "bold"))
        self.ed_flt_col_lbl.grid(row=0, column=2, padx=(4, 4), sticky="w")
        self.filter_col_combo = ttk.Combobox(
            self.ed_flt_lf,
            textvariable=self.filter_col_var,
            values=[self.edit_headings[c] for c in self.edit_columns],
            state="readonly",
            width=20,
        )
        self.filter_col_combo.grid(row=0, column=3, padx=(0, 6), sticky="w")
        self.filter_col_combo.bind("<<ComboboxSelected>>", self._on_filter_col_changed)

        self.filter_op_combo = ttk.Combobox(
            self.ed_flt_lf,
            textvariable=self.filter_op_var,
            values=[tr(FILTER_OP_KEYS[c]) for c in FILTER_OP_CODES],
            state="readonly",
            width=14,
        )
        self.filter_op_combo.grid(row=0, column=4, padx=(0, 6), sticky="w")
        self.filter_op_combo.bind("<<ComboboxSelected>>", self._on_filter_op_changed)

        self.ed_flt_val_lbl = ttk.Label(self.ed_flt_lf, text=tr("ed_flt_val_lbl"), style="App.TLabel")
        self.ed_flt_val_lbl.grid(row=0, column=5, padx=(0, 4), sticky="w")
        self.filter_val_combo = ttk.Combobox(
            self.ed_flt_lf,
            textvariable=self.filter_val_var,
            state="normal",
            width=22,
        )
        self.filter_val_combo.grid(row=0, column=6, padx=(0, 6), sticky="ew")
        self.filter_val_combo.bind("<Return>", lambda _e: self._apply_quick_column_filter())

        self.ed_flt_add_btn = ttk.Button(self.ed_flt_lf, text=tr("ed_flt_add_btn"), command=self._apply_quick_column_filter)
        self.ed_flt_add_btn.grid(row=0, column=7, padx=(0, 4))
        self.ed_flt_excel_btn = ttk.Button(self.ed_flt_lf, text=tr("ed_flt_excel_btn"), command=lambda: self.open_excel_column_filter_dialog())
        self.ed_flt_excel_btn.grid(row=0, column=8, padx=(0, 4))
        self.ed_flt_clear_btn = ttk.Button(self.ed_flt_lf, text=tr("ed_flt_clear_btn"), command=self.clear_all_edit_filters)
        self.ed_flt_clear_btn.grid(row=0, column=9)

        ttk.Label(
            self.ed_flt_lf,
            textvariable=self.active_filters_label_var,
            style="App.TLabel",
            font=("Segoe UI", 8, "italic"),
        ).grid(row=1, column=0, columnspan=10, sticky="w", pady=(4, 0))

        grid_container = ttk.Frame(tab, style="App.TFrame")
        grid_container.grid(row=4, column=0, sticky="nsew")
        grid_container.columnconfigure(0, weight=1)
        grid_container.rowconfigure(1, weight=1)

        excel_tb = ttk.Frame(grid_container, style="App.TFrame")
        excel_tb.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        self.ed_tb_undo_btn = ttk.Button(excel_tb, text=tr("ed_tb_undo"), command=self.undo_edit)
        self.ed_tb_undo_btn.pack(side=tk.LEFT, padx=(0, 3))
        self.ed_tb_redo_btn = ttk.Button(excel_tb, text=tr("ed_tb_redo"), command=self.redo_edit)
        self.ed_tb_redo_btn.pack(side=tk.LEFT, padx=3)
        ttk.Separator(excel_tb, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        self.ed_tb_fill_down_btn = ttk.Button(excel_tb, text=tr("ed_tb_fill_down"), command=self._excel_fill_down)
        self.ed_tb_fill_down_btn.pack(side=tk.LEFT, padx=3)
        self.ed_tb_fill_blanks_btn = ttk.Button(excel_tb, text=tr("ed_tb_fill_blanks"), command=self._fill_blanks_downward_action)
        self.ed_tb_fill_blanks_btn.pack(side=tk.LEFT, padx=3)
        ttk.Separator(excel_tb, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        self.ed_tb_copy_btn = ttk.Button(excel_tb, text=tr("ed_tb_copy"), command=self._copy_cell_value)
        self.ed_tb_copy_btn.pack(side=tk.LEFT, padx=3)
        self.ed_tb_paste_btn = ttk.Button(excel_tb, text=tr("ed_tb_paste"), command=self._paste_cell_value)
        self.ed_tb_paste_btn.pack(side=tk.LEFT, padx=3)
        self.ed_tb_clear_btn = ttk.Button(excel_tb, text=tr("ed_tb_clear"), command=self._clear_selected_edit_cell)
        self.ed_tb_clear_btn.pack(side=tk.LEFT, padx=3)
        self.ed_tb_restore_btn = ttk.Button(excel_tb, text=tr("ed_tb_restore"), command=self._restore_selected_edit_cell)
        self.ed_tb_restore_btn.pack(side=tk.LEFT, padx=3)
        ttk.Separator(excel_tb, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        self.ed_tb_bulk_btn = ttk.Button(excel_tb, text=tr("ed_tb_bulk"), command=self._bulk_apply_selected_value)
        self.ed_tb_bulk_btn.pack(side=tk.LEFT, padx=3)

        quick_clear_btn = ttk.Button(excel_tb, text="✕", width=2, command=self._clear_quick_search)
        quick_clear_btn.pack(side=tk.RIGHT, padx=(2, 0))
        self.edit_quick_search_entry = ttk.Entry(excel_tb, textvariable=self.edit_quick_search_var, width=22)
        self.edit_quick_search_entry.pack(side=tk.RIGHT, padx=(4, 0))
        self.edit_quick_search_entry.bind("<Escape>", lambda _e: self._clear_quick_search())
        self.edit_quick_search_var.trace_add("write", lambda *_: self._refresh_edit_tree())
        self.ed_tb_search_lbl = ttk.Label(excel_tb, text=tr("ed_tb_search_lbl"), style="App.TLabel", font=("Segoe UI", 8, "bold"))
        self.ed_tb_search_lbl.pack(side=tk.RIGHT, padx=(8, 0))

        grid_frame = ttk.Frame(grid_container, style="App.TFrame")
        grid_frame.grid(row=1, column=0, sticky="nsew")
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)

        self.edit_tree = ttk.Treeview(grid_frame, columns=self.edit_columns, show="headings", selectmode="extended")
        for col in self.edit_columns:
            self.edit_tree.heading(
                col,
                text=f"{self.edit_headings[col]} ▼",
                command=lambda c=col: self.open_excel_column_filter_dialog(c),
            )
            width = 80 if col in {"invoice", "line", "price", "quantity"} else 130
            if col == "title":
                width = 320
            self.edit_tree.column(col, width=width, minwidth=55, anchor="w")

        self.edit_tree.tag_configure("unmatched", background="#fff8dc")
        self.edit_tree.tag_configure("warning", background="#fde8e8")
        self.edit_tree.tag_configure("edited", background="#eff6ff")
        self.edit_tree.grid(row=0, column=0, sticky="nsew")

        self.edit_yscroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self._on_tree_yscroll)
        self.edit_xscroll = ttk.Scrollbar(grid_frame, orient="horizontal", command=self._on_tree_xscroll)
        self.edit_yscroll.grid(row=0, column=1, sticky="ns")
        self.edit_xscroll.grid(row=1, column=0, sticky="ew")
        self.edit_tree.configure(yscrollcommand=self.edit_yscroll.set, xscrollcommand=self.edit_xscroll.set)

        self._setup_excel_selection_overlay()

        self.edit_tree.bind("<Button-1>", self._on_tree_click)
        self.edit_tree.bind("<B1-Motion>", self._on_tree_drag)
        self.edit_tree.bind("<Double-1>", self._begin_edit_cell)
        self.edit_tree.bind("<<TreeviewSelect>>", self._on_edit_tree_selection_changed)
        self.edit_tree.bind("<Configure>", lambda _e: self.root.after_idle(self._update_excel_selection_overlay))
        self.edit_tree.bind("<MouseWheel>", self._on_edit_mousewheel)
        self.edit_tree.bind("<Button-4>", self._on_edit_mousewheel)
        self.edit_tree.bind("<Button-5>", self._on_edit_mousewheel)
        self.edit_tree.bind("<Button-3>", self._show_edit_context_menu)
        self.edit_tree.bind("<Control-f>", self._focus_quick_search)
        self.edit_tree.bind("<Control-F>", self._focus_quick_search)
        self.edit_tab.bind("<Control-f>", self._focus_quick_search)
        self.edit_tab.bind("<Control-F>", self._focus_quick_search)
        self.edit_tree.bind("<Control-z>", self.undo_edit)
        self.edit_tree.bind("<Control-Z>", self.undo_edit)
        self.edit_tree.bind("<Control-y>", self.redo_edit)
        self.edit_tree.bind("<Control-Y>", self.redo_edit)
        self.edit_tree.bind("<Control-d>", self._fill_down_shortcut)
        self.edit_tree.bind("<Control-D>", self._fill_down_shortcut)
        self.edit_tree.bind("<Control-b>", lambda _e: self._fill_blanks_downward_action())
        self.edit_tree.bind("<Control-B>", lambda _e: self._fill_blanks_downward_action())
        self.edit_tree.bind("<Control-c>", lambda _e: self._copy_cell_value())
        self.edit_tree.bind("<Control-C>", lambda _e: self._copy_cell_value())
        self.edit_tree.bind("<Control-v>", lambda _e: self._paste_cell_value())
        self.edit_tree.bind("<Control-V>", lambda _e: self._paste_cell_value())
        self.edit_tree.bind("<Delete>", lambda _e: self._clear_selected_edit_cell())
        self.edit_tree.bind("<BackSpace>", lambda _e: self._clear_selected_edit_cell())
        self.edit_tree.bind("<Return>", self._on_tree_return_pressed)
        self.edit_tree.bind("<KP_Enter>", self._on_tree_return_pressed)
        self.edit_tree.bind("<Tab>", self._on_tree_tab_pressed)
        self.edit_tree.bind("<Shift-Tab>", self._on_tree_shift_tab_pressed)
        self.edit_tree.bind("<Up>", self._on_tree_arrow_up)
        self.edit_tree.bind("<Down>", self._on_tree_arrow_down)
        self.edit_tree.bind("<Key>", self._on_tree_key_press)

        self.edit_context_menu = tk.Menu(self.edit_tree, tearoff=0)
        self._rebuild_edit_menus()

        diag_bar = ttk.Frame(tab, style="App.TFrame")
        diag_bar.grid(row=5, column=0, sticky="ew", pady=(5, 0))
        diag_bar.columnconfigure(0, weight=1)
        self.edit_line_diag_label = ttk.Label(
            diag_bar,
            textvariable=self.edit_line_diag_var,
            style="App.TLabel",
            font=("Segoe UI", 8),
        )
        self.edit_line_diag_label.grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Label(
            diag_bar,
            textvariable=self.edit_totals_var,
            style="App.TLabel",
            font=("Segoe UI", 8, "bold"),
        ).grid(row=0, column=1, sticky="e")

        edit_bottom = ttk.Frame(tab, style="App.TFrame")
        edit_bottom.grid(row=6, column=0, sticky="ew", pady=(6, 0))
        edit_bottom.columnconfigure(0, weight=1)
        self.ed_bottom_hint_lbl = ttk.Label(
            edit_bottom,
            text=tr("ed_bottom_hint"),
            style="App.TLabel",
            font=("Segoe UI", 8),
        )
        self.ed_bottom_hint_lbl.grid(row=0, column=0, sticky="w")
        self.ed_save_draft_btn = ttk.Button(edit_bottom, text=tr("ed_save_draft_btn"), command=self.save_edit_session_draft)
        self.ed_save_draft_btn.grid(row=0, column=1, padx=(8, 3))
        self.ed_load_draft_btn = ttk.Button(edit_bottom, text=tr("ed_load_draft_btn"), command=self.load_edit_session_draft)
        self.ed_load_draft_btn.grid(row=0, column=2, padx=(3, 8))
        self.ed_save_excel_btn = ttk.Button(edit_bottom, text=tr("ed_save_excel_btn"), command=self.save_edit_changes, style="Primary.TButton")
        self.ed_save_excel_btn.grid(row=0, column=3, padx=(4, 6))
        self.ed_reset_inv_btn = ttk.Button(edit_bottom, text=tr("ed_reset_inv_btn"), command=self.reset_edit_invoice)
        self.ed_reset_inv_btn.grid(row=0, column=4)

    def _directory_row(self, parent, row, label_key, variable, command):
        lbl = ttk.Label(parent, text=tr(label_key), style="App.TLabel")
        lbl.grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=10, pady=4)
        btn = ttk.Button(parent, text=tr("cust_choose_dir_btn"), command=command)
        btn.grid(row=row, column=2, pady=4, sticky="ew")
        self._dir_row_widgets.append((label_key, lbl, btn))

    def _build_customize_tab(self):
        tab = self.customize_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        self._dir_row_widgets = []
        self.cust_pref_lf = ttk.LabelFrame(tab, text=tr("cust_pref_frame"), padding=12, style="App.TLabelframe")
        self.cust_pref_lf.grid(row=0, column=0, sticky="ew")
        self.cust_pref_lf.columnconfigure(1, weight=1)

        self.cust_lang_lbl = ttk.Label(self.cust_pref_lf, text=tr("cust_lang_lbl"), style="App.TLabel")
        self.cust_lang_lbl.grid(row=0, column=0, sticky="w", pady=4)
        self.language_combo = ttk.Combobox(
            self.cust_pref_lf,
            textvariable=self.language_label_var,
            values=[label for _code, label in SUPPORTED_LANGUAGES],
            state="readonly",
            width=34,
        )
        self.language_combo.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=4)
        self.language_combo.bind("<<ComboboxSelected>>", self._on_language_selected)

        self.cust_inst_lbl = ttk.Label(self.cust_pref_lf, text=tr("cust_inst_lbl"), style="App.TLabel")
        self.cust_inst_lbl.grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(self.cust_pref_lf, textvariable=self.institution_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=4)
        self._directory_row(self.cust_pref_lf, 2, "cust_dir_inv", self.custom_invoice_dir, lambda: self._choose_preference_dir(self.custom_invoice_dir))
        self._directory_row(self.cust_pref_lf, 3, "cust_dir_alma", self.custom_alma_dir, lambda: self._choose_preference_dir(self.custom_alma_dir))
        self._directory_row(self.cust_pref_lf, 4, "cust_dir_tpl", self.custom_template_dir, lambda: self._choose_preference_dir(self.custom_template_dir))
        self._directory_row(self.cust_pref_lf, 5, "cust_dir_out", self.custom_output_dir, lambda: self._choose_preference_dir(self.custom_output_dir))
        self.cust_logo_lbl = ttk.Label(self.cust_pref_lf, text=tr("cust_logo_lbl"), style="App.TLabel")
        self.cust_logo_lbl.grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(self.cust_pref_lf, textvariable=self.logo_path_var).grid(row=6, column=1, sticky="ew", padx=10, pady=4)
        logo_buttons = ttk.Frame(self.cust_pref_lf, style="App.TFrame")
        logo_buttons.grid(row=6, column=2, sticky="e", pady=4)
        self.cust_logo_sel_btn = ttk.Button(logo_buttons, text=tr("cust_logo_sel_btn"), command=self.select_logo)
        self.cust_logo_sel_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.cust_logo_del_btn = ttk.Button(logo_buttons, text=tr("cust_logo_del_btn"), command=self.remove_logo)
        self.cust_logo_del_btn.pack(side=tk.LEFT)
        self.logo_status_label = ttk.Label(self.cust_pref_lf, text=tr("cust_logo_hint"), style="App.TLabel")
        self.logo_status_label.grid(row=7, column=1, columnspan=2, sticky="w", padx=10)
        self.cust_pref_note_lbl = ttk.Label(self.cust_pref_lf, text=tr("cust_pref_note"), wraplength=880, justify="left", style="App.TLabel")
        self.cust_pref_note_lbl.grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 4))
        buttons = ttk.Frame(self.cust_pref_lf, style="App.TFrame")
        buttons.grid(row=9, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.cust_save_pref_btn = ttk.Button(buttons, text=tr("cust_save_pref_btn"), command=self.save_preferences)
        self.cust_save_pref_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.cust_reset_pref_btn = ttk.Button(buttons, text=tr("cust_reset_pref_btn"), command=self.reset_preferences)
        self.cust_reset_pref_btn.pack(side=tk.LEFT)

        self.cust_vend_lf = ttk.LabelFrame(
            tab,
            text=tr("cust_vend_frame"),
            padding=10,
            style="App.TLabelframe",
        )
        self.cust_vend_lf.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.cust_vend_lf.columnconfigure(0, weight=1)
        self.cust_vend_lf.rowconfigure(2, weight=1)

        form_row = ttk.Frame(self.cust_vend_lf, style="App.TFrame")
        form_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        form_row.columnconfigure(3, weight=1)
        self.cust_vend_nif_lbl = ttk.Label(form_row, text=tr("cust_vend_nif_lbl"), style="App.TLabel")
        self.cust_vend_nif_lbl.grid(row=0, column=0, padx=(0, 6), sticky="w")
        self.vendor_nif_var = tk.StringVar()
        ttk.Entry(form_row, textvariable=self.vendor_nif_var, width=20).grid(row=0, column=1, padx=(0, 12), sticky="w")
        self.cust_vend_acc_lbl = ttk.Label(form_row, text=tr("cust_vend_acc_lbl"), style="App.TLabel")
        self.cust_vend_acc_lbl.grid(row=0, column=2, padx=(0, 6), sticky="w")
        self.vendor_acc_var = tk.StringVar()
        ttk.Entry(form_row, textvariable=self.vendor_acc_var).grid(row=0, column=3, padx=(0, 8), sticky="ew")
        self.cust_vend_add_btn = ttk.Button(form_row, text=tr("cust_vend_add_btn"), command=self._add_or_update_vendor_map_entry)
        self.cust_vend_add_btn.grid(row=0, column=4, padx=(0, 4))
        self.cust_vend_del_btn = ttk.Button(form_row, text=tr("cust_vend_del_btn"), command=self._delete_selected_vendor_map_entry)
        self.cust_vend_del_btn.grid(row=0, column=5, padx=(0, 4))
        self.cust_vend_clr_btn = ttk.Button(form_row, text=tr("cust_vend_clr_btn"), command=self._clear_vendor_map)
        self.cust_vend_clr_btn.grid(row=0, column=6)

        excel_row = ttk.Frame(self.cust_vend_lf, style="App.TFrame")
        excel_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        self.cust_vend_imp_mode_lbl = ttk.Label(excel_row, text=tr("cust_vend_imp_mode_lbl"), style="App.TLabel")
        self.cust_vend_imp_mode_lbl.grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.vendor_import_mode_combo = ttk.Combobox(
            excel_row,
            textvariable=self.vendor_import_mode_var,
            values=[tr(VENDOR_IMPORT_MODE_KEYS[m]) for m in VENDOR_IMPORT_MODES],
            state="readonly",
            width=44,
        )
        self.vendor_import_mode_combo.grid(row=0, column=1, padx=(0, 8), sticky="w")
        self.vendor_import_mode_combo.bind("<<ComboboxSelected>>", self._on_vendor_import_mode_selected)
        self.cust_vend_imp_btn = ttk.Button(excel_row, text=tr("cust_vend_imp_btn"), command=self.import_vendors_excel_dialog)
        self.cust_vend_imp_btn.grid(row=0, column=2, padx=(0, 6))
        self.cust_vend_exp_btn = ttk.Button(excel_row, text=tr("cust_vend_exp_btn"), command=self.export_vendors_excel_dialog)
        self.cust_vend_exp_btn.grid(row=0, column=3)

        self.vendor_map_tree = ttk.Treeview(self.cust_vend_lf, columns=("nif", "vendor"), show="headings", height=5, selectmode="browse")
        self.vendor_map_tree.heading("nif", text=tr("cust_vend_col_nif"))
        self.vendor_map_tree.heading("vendor", text=tr("cust_vend_col_acc"))
        self.vendor_map_tree.column("nif", width=220, anchor="w")
        self.vendor_map_tree.column("vendor", width=480, anchor="w")
        self.vendor_map_tree.grid(row=2, column=0, sticky="nsew")
        self.vendor_map_tree.bind("<<TreeviewSelect>>", self._on_vendor_map_selected)
        v_scroll = ttk.Scrollbar(self.cust_vend_lf, orient="vertical", command=self.vendor_map_tree.yview)
        v_scroll.grid(row=2, column=1, sticky="ns")
        self.vendor_map_tree.configure(yscrollcommand=v_scroll.set)
        self._refresh_vendor_map_tree()

    def _build_help_tab(self):
        tab = self.help_tab
        tab.columnconfigure(0, weight=1)
        self.hlp_title_lbl = ttk.Label(tab, text=f"{APP_NAME} {APP_VERSION}", style="HelpTitle.TLabel")
        self.hlp_title_lbl.grid(row=0, column=0, sticky="w", pady=(0, 12))
        self.hlp_body_lbl = ttk.Label(tab, text=tr("hlp_body_5_1"), wraplength=940, justify="left", style="App.TLabel")
        self.hlp_body_lbl.grid(row=1, column=0, sticky="nw")
        self.hlp_links_lf = ttk.LabelFrame(tab, text=tr("hlp_links_frame"), padding=14, style="App.TLabelframe")
        self.hlp_links_lf.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        self.hlp_repo_btn = ttk.Button(self.hlp_links_lf, text=tr("hlp_repo_btn"), command=lambda: self._open_url(PROJECT_REPOSITORY_URL))
        self.hlp_repo_btn.pack(anchor="w", pady=4)
        self.hlp_guide_btn = ttk.Button(self.hlp_links_lf, text=tr("hlp_guide_btn"), command=lambda: self._open_url(PROJECT_GUIDE_URL))
        self.hlp_guide_btn.pack(anchor="w", pady=4)
        ttk.Label(self.hlp_links_lf, text=PROJECT_REPOSITORY_URL, style="App.TLabel").pack(anchor="w", pady=(10, 2))
        ttk.Label(self.hlp_links_lf, text=PROJECT_GUIDE_URL, style="App.TLabel").pack(anchor="w")

    def _on_language_selected(self, _event=None):
        selected_label = self.language_label_var.get()
        lang_code = LANGUAGE_LABEL_TO_CODE.get(selected_label, "es")
        self.set_ui_language(lang_code, persist=True, announce=True)

    def set_ui_language(self, lang_code, persist=True, announce=False):
        code = set_current_language(lang_code)
        self.preferences["language"] = code
        self.language_label_var.set(LANGUAGE_CODE_TO_LABEL.get(code, LANGUAGE_CODE_TO_LABEL["es"]))
        if persist:
            try:
                self._save_preferences_file()
            except ValueError:
                logging.warning("No se pudo guardar el idioma en preferencias")
        self.apply_language()
        if announce:
            self._set_status("st_lang_changed", lang=LANGUAGE_CODE_TO_LABEL.get(code, code))

    def apply_language(self):
        # Pestañas principales
        self.notebook.tab(self.convert_tab, text=tr("tab_convert"))
        self.notebook.tab(self.visualize_tab, text=tr("tab_visualize"))
        self.notebook.tab(self.edit_tab, text=tr("tab_edit"))
        self.notebook.tab(self.customize_tab, text=tr("tab_customize"))
        self.notebook.tab(self.help_tab, text=tr("tab_help"))

        # Pestaña Convertir
        self.conv_mode_lf.config(text=tr("conv_mode_frame"))
        self.conv_single_rb.config(text=tr("conv_mode_single"))
        self.conv_batch_rb.config(text=tr("conv_mode_batch"))
        self.conv_input_lf.config(text=tr("conv_inputs_frame"))
        is_batch = self.conversion_mode.get() == "batch"
        self.invoice_label.config(text=tr("conv_inv_batch_lbl" if is_batch else "conv_inv_single_lbl"))
        self.invoice_button.config(text=tr("conv_inv_batch_btn" if is_batch else "conv_inv_single_btn"))
        if is_batch and self.invoice_files:
            self.invoice_file.set(tr("conv_batch_count", count=len(self.invoice_files)))
        self.conv_alma_lbl.config(text=tr("conv_alma_lbl"))
        self.conv_alma_btn.config(text=tr("conv_alma_btn"))
        self.conv_tpl_lbl.config(text=tr("conv_tpl_lbl"))
        self.conv_tpl_btn.config(text=tr("conv_tpl_btn"))
        self.batch_frame.config(text=tr("conv_batch_frame"))
        self.batch_tree.heading("invoice", text=tr("conv_col_invoice"))
        self.batch_tree.heading("vendor", text=tr("conv_col_vendor"))
        self.batch_tree.heading("file", text=tr("conv_col_file"))
        self.batch_add_btn.config(text=tr("conv_batch_add"))
        self.batch_rm_btn.config(text=tr("conv_batch_remove"))
        self.batch_sort_btn.config(text=tr("conv_batch_sort"))
        self.batch_clr_btn.config(text=tr("conv_batch_clear"))
        self.batch_hint_lbl.config(text=tr("conv_batch_hint"))
        self.conv_output_lf.config(text=tr("conv_output_frame"))
        self.conv_output_btn.config(text=tr("conv_output_btn"))
        self.conv_actions_lf.config(text=tr("conv_actions_frame"))
        self.convert_button.config(text=tr("conv_run_btn"))
        self.conv_clear_btn.config(text=tr("conv_clear_btn"))
        self.conv_open_out_btn.config(text=tr("conv_open_out_btn"))
        self.conv_exit_btn.config(text=tr("conv_exit_btn"))

        # Pestaña Visualizar y Visor integrado
        self.vis_inv_lbl.config(text=tr("vis_inv_lbl"))
        self.vis_footer_lbl.config(text=tr("vis_footer_note"))
        self.invoice_viewer.refresh_language()

        # Pestaña Editar y desplegables de filtros
        self.ed_inv_lbl.config(text=tr("vis_inv_lbl"))
        self.ed_search_pol_btn.config(text=tr("ed_search_pol_btn"))
        self.ed_refresh_btn.config(text=tr("ed_refresh_btn"))
        self.ed_num_lf.config(text=tr("ed_num_frame"))
        self.ed_orig_lbl.config(text=tr("ed_orig_lbl"))
        self.ed_rb_orig.config(text=tr("ed_mode_orig"))
        self.ed_rb_suffix.config(text=tr("ed_mode_suffix"))
        self.ed_rb_prefix.config(text=tr("ed_mode_prefix"))
        self.ed_apply_num_btn.config(text=tr("ed_apply_num_btn"))
        self.ed_hdr_lf.config(text=tr("ed_hdr_frame"))

        self.edit_headings["invoice"] = tr("ed_col_invoice")
        self.edit_headings["line"] = tr("ed_col_line")
        self.heading_to_col = {label: col for col, label in self.edit_headings.items()}

        self._syncing_lang_combos = True
        try:
            self.ed_flt_lf.config(text=tr("ed_flt_frame"))
            self.ed_flt_state_lbl.config(text=tr("ed_flt_state_lbl"))
            self.edit_filter_combo["values"] = [tr(STATUS_FILTER_KEYS[c]) for c in STATUS_FILTER_CODES]
            self.edit_filter_var.set(tr(STATUS_FILTER_KEYS.get(self.edit_filter_code, "ed_flt_st_all")))

            self.ed_flt_col_lbl.config(text=tr("ed_flt_col_lbl"))
            self.filter_col_combo["values"] = [self.edit_headings[c] for c in self.edit_columns]
            self.filter_col_var.set(self.edit_headings.get(self.filter_col_key, "PO Line"))

            self.filter_op_combo["values"] = [tr(FILTER_OP_KEYS[c]) for c in FILTER_OP_CODES]
            self.filter_op_var.set(tr(FILTER_OP_KEYS.get(self.filter_op_code, "ed_flt_op_contains")))
            if self.filter_op_code == "blank":
                self.filter_val_var.set(tr("ed_flt_val_blank"))
            elif self.filter_op_code == "not_blank":
                self.filter_val_var.set(tr("ed_flt_val_nblank"))
        finally:
            self._syncing_lang_combos = False

        self.ed_flt_val_lbl.config(text=tr("ed_flt_val_lbl"))
        self.ed_flt_add_btn.config(text=tr("ed_flt_add_btn"))
        self.ed_flt_excel_btn.config(text=tr("ed_flt_excel_btn"))
        self.ed_flt_clear_btn.config(text=tr("ed_flt_clear_btn"))

        self.ed_tb_undo_btn.config(text=tr("ed_tb_undo"))
        self.ed_tb_redo_btn.config(text=tr("ed_tb_redo"))
        self.ed_tb_fill_down_btn.config(text=tr("ed_tb_fill_down"))
        self.ed_tb_fill_blanks_btn.config(text=tr("ed_tb_fill_blanks"))
        self.ed_tb_copy_btn.config(text=tr("ed_tb_copy"))
        self.ed_tb_paste_btn.config(text=tr("ed_tb_paste"))
        self.ed_tb_clear_btn.config(text=tr("ed_tb_clear"))
        self.ed_tb_restore_btn.config(text=tr("ed_tb_restore"))
        self.ed_tb_bulk_btn.config(text=tr("ed_tb_bulk"))
        self.ed_tb_search_lbl.config(text=tr("ed_tb_search_lbl"))
        self._rebuild_edit_menus()

        self.ed_bottom_hint_lbl.config(text=tr("ed_bottom_hint"))
        self.ed_save_draft_btn.config(text=tr("ed_save_draft_btn"))
        self.ed_load_draft_btn.config(text=tr("ed_load_draft_btn"))
        self.ed_save_excel_btn.config(text=tr("ed_save_excel_btn"))
        self.ed_reset_inv_btn.config(text=tr("ed_reset_inv_btn"))

        # Pestaña Personalizar (incluyendo etiquetas de directorios, botones, notas y proveedores)
        self.cust_pref_lf.config(text=tr("cust_pref_frame"))
        self.cust_lang_lbl.config(text=tr("cust_lang_lbl"))
        self.cust_inst_lbl.config(text=tr("cust_inst_lbl"))
        for lbl_key, lbl_w, btn_w in self._dir_row_widgets:
            lbl_w.config(text=tr(lbl_key))
            btn_w.config(text=tr("cust_choose_dir_btn"))
        self.cust_logo_lbl.config(text=tr("cust_logo_lbl"))
        self.cust_logo_sel_btn.config(text=tr("cust_logo_sel_btn"))
        self.cust_logo_del_btn.config(text=tr("cust_logo_del_btn"))
        logo_key, logo_kw = self._logo_status_state
        self.logo_status_label.config(text=tr(logo_key, **logo_kw))
        self.cust_pref_note_lbl.config(text=tr("cust_pref_note"))
        self.cust_save_pref_btn.config(text=tr("cust_save_pref_btn"))
        self.cust_reset_pref_btn.config(text=tr("cust_reset_pref_btn"))

        self.cust_vend_lf.config(text=tr("cust_vend_frame"))
        self.cust_vend_nif_lbl.config(text=tr("cust_vend_nif_lbl"))
        self.cust_vend_acc_lbl.config(text=tr("cust_vend_acc_lbl"))
        self.cust_vend_add_btn.config(text=tr("cust_vend_add_btn"))
        self.cust_vend_del_btn.config(text=tr("cust_vend_del_btn"))
        self.cust_vend_clr_btn.config(text=tr("cust_vend_clr_btn"))
        self.cust_vend_imp_mode_lbl.config(text=tr("cust_vend_imp_mode_lbl"))
        self.vendor_import_mode_combo["values"] = [tr(VENDOR_IMPORT_MODE_KEYS[m]) for m in VENDOR_IMPORT_MODES]
        self.vendor_import_mode_var.set(tr(VENDOR_IMPORT_MODE_KEYS.get(self.vendor_import_mode_code, "cust_vend_mode_inc")))
        self.cust_vend_imp_btn.config(text=tr("cust_vend_imp_btn"))
        self.cust_vend_exp_btn.config(text=tr("cust_vend_exp_btn"))
        self.vendor_map_tree.heading("nif", text=tr("cust_vend_col_nif"))
        self.vendor_map_tree.heading("vendor", text=tr("cust_vend_col_acc"))

        # Pestaña Ayuda
        self.hlp_body_lbl.config(text=tr("hlp_body_5_1"))
        self.hlp_links_lf.config(text=tr("hlp_links_frame"))
        self.hlp_repo_btn.config(text=tr("hlp_repo_btn"))
        self.hlp_guide_btn.config(text=tr("hlp_guide_btn"))

        # Refrescar barra de estado, diagnóstico, filtros y totales
        self._refresh_edit_tree()
        if not self._active_edit_path():
            self.edit_totals_var.set(tr("ed_tot_init"))
            self.edit_line_diag_var.set(tr("ed_diag_init"))
        st_key, st_kw = self._last_status
        self.status_text.set(tr(st_key, **st_kw))

    # --------------------------------------------------------------------------
    # OVERLAY DE SELECCIÓN Y ASA DE ARRASTRE TIPO EXCEL (■)
    # --------------------------------------------------------------------------
    def _setup_excel_selection_overlay(self):
        sel_color = "#1e40af"
        self.border_top = tk.Frame(self.edit_tree, bg=sel_color, height=2)
        self.border_bottom = tk.Frame(self.edit_tree, bg=sel_color, height=2)
        self.border_left = tk.Frame(self.edit_tree, bg=sel_color, width=2)
        self.border_right = tk.Frame(self.edit_tree, bg=sel_color, width=2)

        self.fill_handle = tk.Frame(self.edit_tree, bg=sel_color, width=7, height=7, cursor="crosshair")
        self.fill_handle.bind("<Button-1>", self._start_fill_handle_drag)
        self.fill_handle.bind("<B1-Motion>", self._on_fill_handle_drag)
        self.fill_handle.bind("<ButtonRelease-1>", self._end_fill_handle_drag)
        self.fill_handle.bind("<Double-1>", self._on_fill_handle_double_click)

        self._fill_drag_start_cell = None
        self._fill_drag_target_row = None

    def _hide_excel_selection_overlay(self):
        if not hasattr(self, "border_top"):
            return
        for w in (self.border_top, self.border_bottom, self.border_left, self.border_right, self.fill_handle):
            if w.winfo_ismapped():
                w.place_forget()

    def _update_excel_selection_overlay(self):
        try:
            if not hasattr(self, "border_top") or not self.edit_tree.winfo_exists():
                return
        except tk.TclError:
            return
        if self.edit_active_editor is not None or not self.active_cell:
            self._hide_excel_selection_overlay()
            return
        row_id, column = self.active_cell
        if not self.edit_tree.exists(row_id) or column not in self.edit_columns:
            self._hide_excel_selection_overlay()
            return
        col_id = f"#{self.edit_columns.index(column) + 1}"
        bbox = self.edit_tree.bbox(row_id, col_id)
        if not bbox:
            self._hide_excel_selection_overlay()
            return
        x, y, w, h = bbox
        selected_rows = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if selected_rows and row_id in selected_rows and len(selected_rows) > 1:
            all_bboxes = [self.edit_tree.bbox(r, col_id) for r in selected_rows]
            valid_bboxes = [b for b in all_bboxes if b]
            if valid_bboxes:
                min_y = min(b[1] for b in valid_bboxes)
                max_y = max(b[1] + b[3] for b in valid_bboxes)
                y = min_y
                h = max_y - min_y
        if y < 18:
            self._hide_excel_selection_overlay()
            return
        self.border_top.place(x=x, y=y, width=w, height=2)
        self.border_bottom.place(x=x, y=y + h - 2, width=w, height=2)
        self.border_left.place(x=x, y=y, width=2, height=h)
        self.border_right.place(x=x + w - 2, y=y, width=2, height=h)
        self.fill_handle.place(x=x + w - 5, y=y + h - 5, width=7, height=7)
        self.fill_handle.lift()

    def _start_fill_handle_drag(self, _event):
        self._close_all_cell_editors(commit=True)
        if not self.active_cell:
            return "break"
        self._fill_drag_start_cell = self.active_cell
        self._fill_drag_target_row = self.active_cell[0]
        return "break"

    def _on_fill_handle_drag(self, event):
        if not self._fill_drag_start_cell:
            return "break"
        target_y = event.y_root - self.edit_tree.winfo_rooty()
        row_id = self.edit_tree.identify_row(target_y)
        if row_id and row_id != self._fill_drag_target_row:
            self._fill_drag_target_row = row_id
            start_row = self._fill_drag_start_cell[0]
            children = list(self.edit_tree.get_children())
            if start_row in children and row_id in children:
                idx1 = children.index(start_row)
                idx2 = children.index(row_id)
                min_i, max_i = min(idx1, idx2), max(idx1, idx2)
                self.edit_tree.selection_set(children[min_i:max_i + 1])
                self._update_excel_selection_overlay()
        return "break"

    def _end_fill_handle_drag(self, _event):
        if not self._fill_drag_start_cell:
            return "break"
        start_row, column = self._fill_drag_start_cell
        target_row = self._fill_drag_target_row
        self._fill_drag_start_cell = None
        self._fill_drag_target_row = None
        if not target_row or target_row == start_row:
            return "break"
        if column in READ_ONLY_LINE_FIELDS:
            messagebox.showinfo(tr("mb_prot_field_title"), tr("mb_prot_drag_msg"))
            return "break"
        children = list(self.edit_tree.get_children())
        if start_row not in children or target_row not in children:
            return "break"
        idx1 = children.index(start_row)
        idx2 = children.index(target_row)
        min_i, max_i = min(idx1, idx2), max(idx1, idx2)
        source_val = self.edit_tree.set(start_row, column)
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        count = 0
        for r_id in children[min_i:max_i + 1]:
            if r_id != start_row:
                self._set_edit_value(r_id, column, source_val, update_widget=True, record_undo=False)
                count += 1
        if count > 0:
            self._set_status("st_fill_drag", count=count, col=self.edit_headings.get(column, column), val=source_val)
        self._update_excel_selection_overlay()
        return "break"

    def _on_fill_handle_double_click(self, _event):
        self._fill_blanks_downward_action()
        return "break"

    # --------------------------------------------------------------------------
    # NAVEGACIÓN Y SELECCIÓN TIPO EXCEL EN TREEVIEW
    # --------------------------------------------------------------------------
    def _on_tree_click(self, event):
        self._close_all_cell_editors(commit=True)
        region = self.edit_tree.identify_region(event.x, event.y)
        if region == "heading":
            return
        row_id = self.edit_tree.identify_row(event.y)
        col_id = self.edit_tree.identify_column(event.x)
        if not row_id or not col_id:
            self._hide_excel_selection_overlay()
            return
        try:
            col_index = int(col_id[1:]) - 1
            column = self.edit_columns[col_index]
        except (ValueError, IndexError):
            return
        self.last_edited_column = column
        self.edit_context_column = column
        self.active_cell = (row_id, column)

        state = getattr(event, "state", 0)
        shift_pressed = bool(state & 0x0001)
        ctrl_pressed = bool(state & 0x0004)
        if shift_pressed and self.selection_anchor:
            anchor_row = self.selection_anchor[0]
            children = list(self.edit_tree.get_children())
            if anchor_row in children and row_id in children:
                i1, i2 = children.index(anchor_row), children.index(row_id)
                lo, hi = min(i1, i2), max(i1, i2)
                self.edit_tree.selection_set(children[lo:hi + 1])
                self.root.after_idle(self._update_excel_selection_overlay)
                return "break"
        elif not ctrl_pressed:
            self.selection_anchor = (row_id, column)
            self.edit_tree.selection_set(row_id)
            self.edit_tree.focus(row_id)
            self.root.after_idle(self._update_excel_selection_overlay)
        else:
            self.selection_anchor = (row_id, column)
            self.root.after_idle(self._update_excel_selection_overlay)

    def _on_tree_drag(self, event):
        if self._fill_drag_start_cell is not None or self.edit_active_editor is not None:
            return
        row_id = self.edit_tree.identify_row(event.y)
        if row_id and self.selection_anchor:
            anchor_row, anchor_col = self.selection_anchor
            children = list(self.edit_tree.get_children())
            if anchor_row in children and row_id in children:
                i1, i2 = children.index(anchor_row), children.index(row_id)
                lo, hi = min(i1, i2), max(i1, i2)
                self.edit_tree.selection_set(children[lo:hi + 1])
                self.active_cell = (anchor_row, anchor_col)
                self._update_excel_selection_overlay()

    def _on_tree_yscroll(self, *args):
        self._close_all_cell_editors(commit=True)
        self.edit_tree.yview(*args)
        self.root.after_idle(self._update_excel_selection_overlay)

    def _on_tree_xscroll(self, *args):
        self._close_all_cell_editors(commit=True)
        self.edit_tree.xview(*args)
        self.root.after_idle(self._update_excel_selection_overlay)

    def _on_edit_mousewheel(self, event):
        self._close_all_cell_editors(commit=True)
        if getattr(event, "num", None) == 4:
            units = -3
        elif getattr(event, "num", None) == 5:
            units = 3
        else:
            delta = getattr(event, "delta", 0)
            units = -int(delta / 120) if delta else 0
            if units == 0 and delta:
                units = -1 if delta > 0 else 1
        self.edit_tree.yview_scroll(units, "units")
        self.root.after_idle(self._update_excel_selection_overlay)
        return "break"

    def _scroll_tree_to_column(self, column):
        if column not in self.edit_columns:
            return
        cols = list(self.edit_columns)
        idx = cols.index(column)
        fraction = idx / max(1, len(cols))
        self.edit_tree.xview_moveto(max(0.0, fraction - 0.1))

    def _on_tree_return_pressed(self, _event=None):
        if self.active_cell:
            row_id, column = self.active_cell
            self._open_tree_cell_editor(row_id, column)
            return "break"

    def _on_tree_tab_pressed(self, _event=None):
        if not self.active_cell:
            return "break"
        row_id, column = self.active_cell
        editable = [c for c in self.edit_columns if c in EDITABLE_LINE_FIELDS]
        if column in editable:
            idx = editable.index(column)
            if idx + 1 < len(editable):
                next_col = editable[idx + 1]
                self.active_cell = (row_id, next_col)
                self.last_edited_column = next_col
                self.edit_context_column = next_col
                self._scroll_tree_to_column(next_col)
                self._update_excel_selection_overlay()
        return "break"

    def _on_tree_shift_tab_pressed(self, _event=None):
        if not self.active_cell:
            return "break"
        row_id, column = self.active_cell
        editable = [c for c in self.edit_columns if c in EDITABLE_LINE_FIELDS]
        if column in editable:
            idx = editable.index(column)
            if idx > 0:
                prev_col = editable[idx - 1]
                self.active_cell = (row_id, prev_col)
                self.last_edited_column = prev_col
                self.edit_context_column = prev_col
                self._scroll_tree_to_column(prev_col)
                self._update_excel_selection_overlay()
        return "break"

    def _on_tree_arrow_up(self, _event=None):
        if not self.active_cell:
            return
        row_id, column = self.active_cell
        children = list(self.edit_tree.get_children())
        if row_id in children:
            idx = children.index(row_id)
            if idx > 0:
                prev_row = children[idx - 1]
                self.active_cell = (prev_row, column)
                self.selection_anchor = (prev_row, column)
                self.edit_tree.selection_set(prev_row)
                self.edit_tree.focus(prev_row)
                self.edit_tree.see(prev_row)
                self._update_excel_selection_overlay()
        return "break"

    def _on_tree_arrow_down(self, _event=None):
        if not self.active_cell:
            return
        row_id, column = self.active_cell
        children = list(self.edit_tree.get_children())
        if row_id in children:
            idx = children.index(row_id)
            if idx + 1 < len(children):
                next_row = children[idx + 1]
                self.active_cell = (next_row, column)
                self.selection_anchor = (next_row, column)
                self.edit_tree.selection_set(next_row)
                self.edit_tree.focus(next_row)
                self.edit_tree.see(next_row)
                self._update_excel_selection_overlay()
        return "break"

    def _focus_quick_search(self, _event=None):
        if hasattr(self, "edit_quick_search_entry"):
            self.edit_quick_search_entry.focus_set()
            self.edit_quick_search_entry.select_range(0, tk.END)
        return "break"

    def _clear_quick_search(self):
        if self.edit_quick_search_var.get():
            self.edit_quick_search_var.set("")
        if hasattr(self, "edit_tree"):
            self.edit_tree.focus_set()

    def _on_tree_key_press(self, event):
        if self.edit_active_editor is not None:
            return
        if getattr(event, "state", 0) & (0x0004 | 0x0008 | 0x0080):
            return
        ch = getattr(event, "char", "")
        if not ch or not ch.isprintable() or ch in ("\r", "\n", "\t"):
            return
        if not self.active_cell:
            sel = self.edit_tree.selection()
            col = self._active_editable_column() or "po_line"
            if sel:
                self.active_cell = (sel[0], col)
            else:
                return
        row_id, column = self.active_cell
        if not self.edit_tree.exists(row_id) or column in READ_ONLY_LINE_FIELDS:
            return
        self._open_tree_cell_editor(row_id, column, initial_text=ch)
        return "break"

    def _on_edit_tree_selection_changed(self, _event=None):
        self._update_selected_line_diagnostic()
        self._update_visible_and_selected_totals()

    def _update_selected_line_diagnostic(self):
        path = self._active_edit_path()
        if not path:
            self.edit_line_diag_var.set(tr("ed_diag_no_inv"))
            return
        source_key = os.path.normcase(os.path.abspath(path))
        recon_msg = self.edit_invoice_reconciliation.get(source_key, "")
        sel = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        row_id = sel[0] if sel else (self.active_cell[0] if self.active_cell and self.edit_tree.exists(self.active_cell[0]) else None)
        if not row_id:
            if recon_msg:
                self.edit_line_diag_var.set(tr("ed_diag_recon", msg=recon_msg))
            else:
                self.edit_line_diag_var.set(tr("ed_diag_init"))
            return
        try:
            line_num = int(row_id)
        except ValueError:
            return
        key = (source_key, line_num)
        original = self.edit_original_values.get(key, {})
        values = dict(original)
        values.update(self.line_overrides.get(key, {}))
        is_unmatched, has_warning, _is_problem, is_edited = self._evaluate_line_state(key, values)
        diag = self.edit_line_diagnostics.get(key, {})
        notes = diag.get("validation_notes", [])
        match_status = diag.get("match_status", "")
        localized_match = localize_match_status(match_status, values.get("line_type", "REGULAR")) or diag.get("match_message", "")
        if is_unmatched:
            base = tr("ed_diag_unmatched", line=line_num)
        elif has_warning and notes:
            base = tr("ed_diag_warning", line=line_num, notes=" | ".join(notes))
        elif key in self.manual_po_assignments and values.get("po_line"):
            base = tr("ed_diag_manual", line=line_num, pol=values.get("po_line"))
        elif is_edited:
            base = tr("ed_diag_edited", line=line_num)
        elif localized_match:
            base = tr("ed_diag_matched", line=line_num, msg=localized_match)
        else:
            base = tr("ed_diag_ready", line=line_num)
        if recon_msg:
            base += f"   ||   ⚠ {recon_msg}"
        if len(base) > 190:
            base = base[:187] + "..."
        self.edit_line_diag_var.set(base)

    def _update_visible_and_selected_totals(self):
        if not hasattr(self, "edit_tree"):
            return
        path = self._active_edit_path()
        currency = "€"
        if path:
            source_key = os.path.normcase(os.path.abspath(path))
            hdr = self.edit_header_original_values.get(source_key, {})
            cur_code = (self.invoice_header_overrides.get(source_key, {}).get("currency") or hdr.get("currency") or "EUR").upper()
            currency = "€" if cur_code == "EUR" else cur_code

        visible_rows = list(self.edit_tree.get_children())
        vis_price = Decimal("0.00")
        vis_vat = Decimal("0.00")
        for r in visible_rows:
            vis_price += safe_decimal(self.edit_tree.set(r, "price"), Decimal("0.00"))
            vis_vat += safe_decimal(self.edit_tree.set(r, "vat_amount"), Decimal("0.00"))

        summary = tr(
            "ed_tot_vis",
            count=len(visible_rows),
            base=format_currency(vis_price, currency),
            vat=format_currency(vis_vat, currency),
        )
        sel_rows = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if len(sel_rows) > 1:
            sel_price = sum((safe_decimal(self.edit_tree.set(r, "price"), Decimal("0.00")) for r in sel_rows), Decimal("0.00"))
            summary += tr("ed_tot_sel", count=len(sel_rows), base=format_currency(sel_price, currency))
        self.edit_totals_var.set(summary)

    def save_edit_session_draft(self):
        self._close_all_cell_editors(commit=True)
        if not self.invoice_files:
            messagebox.showinfo(tr("mb_save_draft_title"), tr("mb_save_draft_empty"))
            return
        initial_dir = self._initial_dir("custom_output_dir", "last_output_dir", os.path.dirname(self.invoice_files[0]))
        base_name = os.path.splitext(os.path.basename(self.invoice_files[0]))[0]
        default_name = f"{base_name}_borrador.ftalma.json"
        path = filedialog.asksaveasfilename(
            title=tr("fd_save_draft"),
            initialdir=initial_dir,
            initialfile=default_name,
            defaultextension=".ftalma.json",
            filetypes=[(tr("fd_ft_draft"), "*.ftalma.json *.json"), (tr("fd_ft_all"), "*.*")],
        )
        if not path:
            return
        payload = {
            "app": APP_NAME,
            "version": APP_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "conversion_mode": self.conversion_mode.get(),
            "invoice_files": list(self.invoice_files),
            "alma_pol_file": self.alma_pol_file.get().strip(),
            "template_file": self.template_file.get().strip(),
            "output_file": self.output_file.get().strip(),
            "invoice_number_overrides": dict(self.invoice_number_overrides),
            "invoice_number_mode_by_source": {k: list(v) for k, v in self.invoice_number_mode_by_source.items()},
            "invoice_header_overrides": copy.deepcopy(self.invoice_header_overrides),
            "line_overrides": [
                {"source": k[0], "line": k[1], "changes": v}
                for k, v in self.line_overrides.items()
            ],
            "manual_po_assignments": [
                {"source": k[0], "line": k[1], "value": list(v) if isinstance(v, tuple) else v}
                for k, v in self.manual_po_assignments.items()
            ],
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self._set_status("st_draft_saved", name=os.path.basename(path))
            messagebox.showinfo(tr("mb_draft_saved_title"), tr("mb_draft_saved_msg", path=path))
        except OSError as e:
            messagebox.showerror(tr("mb_err_save_draft"), tr("mb_err_save_draft_msg", err=e))

    def load_edit_session_draft(self):
        self._close_all_cell_editors(commit=False)
        initial_dir = self._initial_dir("custom_output_dir", "last_output_dir")
        path = filedialog.askopenfilename(
            title=tr("fd_load_draft"),
            initialdir=initial_dir,
            filetypes=[(tr("fd_ft_draft"), "*.ftalma.json *.json"), (tr("fd_ft_all"), "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            if not isinstance(payload, dict) or not isinstance(payload.get("invoice_files"), list):
                raise ValueError(tr("val_bad_draft"))
            missing_invoices = [p for p in payload["invoice_files"] if not os.path.isfile(p)]
            existing_invoices = [p for p in payload["invoice_files"] if os.path.isfile(p)]
            if not existing_invoices:
                raise FileNotFoundError(tr("val_draft_no_files"))

            mode = payload.get("conversion_mode", "batch" if len(existing_invoices) > 1 else "single")
            self.conversion_mode.set(mode)
            self.on_mode_change(initial=True)

            self.invoice_files = []
            self.invoice_metadata = {}
            for inv_path in existing_invoices:
                meta = self._read_invoice_metadata(inv_path)
                self.invoice_files.append(inv_path)
                self.invoice_metadata[inv_path] = meta

            if mode == "batch":
                self.invoice_file.set(tr("conv_batch_count", count=len(self.invoice_files)))
            else:
                self.invoice_file.set(self.invoice_files[0])

            alma_path = payload.get("alma_pol_file", "")
            if alma_path and os.path.isfile(alma_path):
                self.alma_pol_file.set(alma_path)
            tpl_path = payload.get("template_file", "")
            if tpl_path and os.path.isfile(tpl_path):
                self.template_file.set(tpl_path)
            out_path = payload.get("output_file", "")
            if out_path:
                self.output_file.set(out_path)

            self.invoice_number_overrides = dict(payload.get("invoice_number_overrides", {}))
            self.invoice_number_mode_by_source = {
                k: tuple(v) for k, v in payload.get("invoice_number_mode_by_source", {}).items() if isinstance(v, (list, tuple)) and len(v) == 2
            }
            self.invoice_header_overrides = copy.deepcopy(payload.get("invoice_header_overrides", {}))
            self.line_overrides = {
                (item["source"], int(item["line"])): dict(item["changes"])
                for item in payload.get("line_overrides", [])
                if isinstance(item, dict) and "source" in item and "line" in item
            }
            self.manual_po_assignments = {}
            for item in payload.get("manual_po_assignments", []):
                if isinstance(item, dict) and "source" in item and "line" in item:
                    val = item.get("value")
                    self.manual_po_assignments[(item["source"], int(item["line"]))] = tuple(val) if isinstance(val, list) else val

            for inv_path in self.invoice_files:
                sk = os.path.normcase(os.path.abspath(inv_path))
                if sk in self.invoice_number_overrides:
                    self.invoice_metadata[inv_path]["number"] = self.invoice_number_overrides[sk]

            self.undo_stack.clear()
            self.redo_stack.clear()
            self.edit_changes_saved = False
            self._refresh_batch_tree()
            self._refresh_viewer_selector(select_path=self.invoice_files[0])
            self.prepare_edit_draft(silent=True)
            warn_msg = ""
            if missing_invoices:
                warn_msg = tr("mb_draft_missing_warn", files="\n- ".join(missing_invoices))
            self._set_status("st_draft_loaded", name=os.path.basename(path))
            messagebox.showinfo(tr("mb_draft_loaded_title"), tr("mb_draft_loaded_msg", warn=warn_msg))
        except Exception as e:
            messagebox.showerror(tr("mb_err_load_draft"), tr("mb_err_load_draft_msg", err=e))

    # --------------------------------------------------------------------------
    # FILTRADO MULTICRITERIO ESTILO EXCEL Y POR ESTADO DE LÍNEA
    # --------------------------------------------------------------------------
    def _on_edit_filter_var_trace(self, *_):
        if getattr(self, "_syncing_lang_combos", False):
            return
        raw = self.edit_filter_var.get()
        if raw in STATUS_FILTER_CODES:
            self.edit_filter_code = raw
            return
        for code, key in STATUS_FILTER_KEYS.items():
            for lang_dict in I18N.values():
                if raw == lang_dict.get(key):
                    self.edit_filter_code = code
                    return

    def _resolve_status_filter_code(self):
        return self.edit_filter_code or "all"

    def _resolve_filter_op_code(self, raw=None):
        text = raw if raw is not None else self.filter_op_var.get()
        if text in FILTER_OP_CODES:
            return text
        for code, key in FILTER_OP_KEYS.items():
            for lang_dict in I18N.values():
                if text == lang_dict.get(key):
                    return code
        return self.filter_op_code or "contains"

    def _is_blank_sentinel(self, text):
        if text == "__BLANK__":
            return True
        return any(text == d.get("ed_flt_val_blank") for d in I18N.values())

    def _is_not_blank_sentinel(self, text):
        return any(text == d.get("ed_flt_val_nblank") for d in I18N.values())

    def _on_status_filter_selected(self, _event=None):
        idx = self.edit_filter_combo.current()
        if 0 <= idx < len(STATUS_FILTER_CODES):
            self.edit_filter_code = STATUS_FILTER_CODES[idx]
        else:
            self.edit_filter_code = self._resolve_status_filter_code()
        self._refresh_edit_tree()

    def _on_filter_col_changed(self, _event=None):
        col_label = self.filter_col_var.get()
        col_key = self.heading_to_col.get(col_label)
        if col_key:
            self.filter_col_key = col_key
        else:
            col_key = self.filter_col_key or "po_line"
        distinct_vals = self._distinct_values_for_column(col_key)
        self.filter_val_combo["values"] = [tr("ed_flt_val_blank"), tr("ed_flt_val_nblank")] + distinct_vals

    def _on_filter_op_changed(self, _event=None):
        idx = self.filter_op_combo.current()
        if 0 <= idx < len(FILTER_OP_CODES):
            self.filter_op_code = FILTER_OP_CODES[idx]
        else:
            self.filter_op_code = self._resolve_filter_op_code()
        if self.filter_op_code == "blank":
            self.filter_val_var.set(tr("ed_flt_val_blank"))
        elif self.filter_op_code == "not_blank":
            self.filter_val_var.set(tr("ed_flt_val_nblank"))

    def _distinct_values_for_column(self, column):
        path = self._active_edit_path()
        if not path:
            return []
        source_key = os.path.normcase(os.path.abspath(path))
        effective_number = (
            self.invoice_number_overrides.get(source_key)
            or self.invoice_metadata.get(path, {}).get("number")
            or ""
        )
        seen = set()
        for key, original in self.edit_original_values.items():
            if key[0] != source_key:
                continue
            values = dict(original)
            values.update(self.line_overrides.get(key, {}))
            if effective_number:
                values["invoice"] = effective_number
            val = str(values.get(column, "") or "").strip()
            if val:
                seen.add(val)
        return sorted(seen, key=lambda s: s.casefold())

    def _apply_quick_column_filter(self):
        self._close_all_cell_editors(commit=True)
        col_label = self.filter_col_var.get()
        col_key = self.heading_to_col.get(col_label) or self.filter_col_key
        if not col_key:
            return
        op_code = self._resolve_filter_op_code()
        self.filter_op_code = op_code
        val = self.filter_val_var.get().strip()
        if op_code not in {"blank", "not_blank"} and not val:
            self.active_column_filters.pop(col_key, None)
        else:
            self.active_column_filters[col_key] = {
                "mode": "condition",
                "op_code": op_code,
                "value": val,
            }
        self._refresh_edit_tree()

    def clear_all_edit_filters(self):
        self._close_all_cell_editors(commit=True)
        self.edit_filter_code = "all"
        self.edit_filter_var.set(tr("ed_flt_st_all"))
        self.active_column_filters.clear()
        self.filter_val_var.set("")
        if self.edit_quick_search_var.get():
            self.edit_quick_search_var.set("")
        self._refresh_edit_tree()

    def _filter_by_selected_cell_value(self):
        self._close_all_cell_editors(commit=True)
        column = getattr(self, "edit_context_column", None) or (self.active_cell[1] if self.active_cell else None)
        selected = self.edit_tree.selection()
        row_id = selected[0] if selected else (self.active_cell[0] if self.active_cell else None)
        if not column or not row_id or not self.edit_tree.exists(row_id):
            return
        cell_val = str(self.edit_tree.set(row_id, column) or "").strip()
        if cell_val == "":
            self.active_column_filters[column] = {"mode": "condition", "op_code": "blank", "value": "__BLANK__"}
            self.filter_op_code = "blank"
            self.filter_op_var.set(tr("ed_flt_op_blank"))
            self.filter_val_var.set(tr("ed_flt_val_blank"))
        else:
            self.active_column_filters[column] = {"mode": "condition", "op_code": "exact", "value": cell_val}
            self.filter_op_code = "exact"
            self.filter_op_var.set(tr("ed_flt_op_exact"))
            self.filter_val_var.set(cell_val)
        self.filter_col_key = column
        self.filter_col_var.set(self.edit_headings.get(column, column))
        self._on_filter_col_changed()
        self._refresh_edit_tree()

    def _open_filter_for_context_column(self):
        column = getattr(self, "edit_context_column", None) or (self.active_cell[1] if self.active_cell else "po_line")
        self.open_excel_column_filter_dialog(column)

    def open_excel_column_filter_dialog(self, initial_column=None):
        self._close_all_cell_editors(commit=True)
        path = self._active_edit_path()
        if not path:
            messagebox.showinfo(tr("mb_flt_cols_title"), tr("mb_flt_cols_no_inv"))
            return
        if initial_column is None:
            col_label = self.filter_col_var.get()
            initial_column = self.heading_to_col.get(col_label, self.filter_col_key or "po_line")

        dialog = tk.Toplevel(self.root)
        dialog.title(tr("afdlg_title"))
        dialog.geometry("580x520")
        dialog.minsize(480, 420)
        dialog.configure(bg=UI_BG)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(2, weight=1)

        top = ttk.Frame(dialog, padding=10, style="App.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text=tr("afdlg_col_lbl"), style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=(0, 8), sticky="w")
        col_var = tk.StringVar(value=self.edit_headings.get(initial_column, "PO Line"))
        col_combo = ttk.Combobox(
            top,
            textvariable=col_var,
            values=[self.edit_headings[c] for c in self.edit_columns],
            state="readonly",
        )
        col_combo.grid(row=0, column=1, sticky="ew")

        search_box = ttk.Frame(dialog, padding=(10, 0, 10, 6), style="App.TFrame")
        search_box.grid(row=1, column=0, sticky="ew")
        search_box.columnconfigure(1, weight=1)
        ttk.Label(search_box, text=tr("afdlg_search_lbl"), style="App.TLabel").grid(row=0, column=0, padx=(0, 8), sticky="w")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_box, textvariable=search_var)
        search_entry.grid(row=0, column=1, sticky="ew")

        list_frame = ttk.LabelFrame(dialog, text=tr("afdlg_list_frame"), padding=8, style="App.TLabelframe")
        list_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)

        sel_btns = ttk.Frame(list_frame, style="App.TFrame")
        sel_btns.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        canvas = tk.Canvas(list_frame, bg="#ffffff", highlightthickness=1, highlightbackground="#cccccc")
        vscroll = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="App.TFrame")
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        vscroll.grid(row=1, column=1, sticky="ns")

        check_vars = {}
        blank_label = tr("afdlg_blank_chk")

        def populate_checkboxes(*_):
            for child in inner.winfo_children():
                child.destroy()
            check_vars.clear()
            curr_col = self.heading_to_col.get(col_var.get(), "po_line")
            distinct = self._distinct_values_for_column(curr_col)
            existing_rule = self.active_column_filters.get(curr_col)
            existing_allowed = existing_rule.get("allowed") if (existing_rule and existing_rule.get("mode") == "values") else None
            q = search_var.get().strip().casefold()

            if not q or q in blank_label.casefold():
                is_checked = True if existing_allowed is None else ("__BLANK__" in existing_allowed)
                var_blank = tk.BooleanVar(value=is_checked)
                check_vars["__BLANK__"] = var_blank
                tk.Checkbutton(
                    inner,
                    text=blank_label,
                    variable=var_blank,
                    bg="#ffffff",
                    anchor="w",
                ).pack(fill=tk.X, anchor="w", padx=4, pady=1)

            for val in distinct:
                if q and q not in val.casefold():
                    continue
                is_checked = True if existing_allowed is None else (val in existing_allowed)
                var_item = tk.BooleanVar(value=is_checked)
                check_vars[val] = var_item
                tk.Checkbutton(
                    inner,
                    text=val,
                    variable=var_item,
                    bg="#ffffff",
                    anchor="w",
                ).pack(fill=tk.X, anchor="w", padx=4, pady=1)

        def set_all_checks(state_bool):
            for var in check_vars.values():
                var.set(state_bool)

        ttk.Button(sel_btns, text=tr("afdlg_check_all"), command=lambda: set_all_checks(True)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(sel_btns, text=tr("afdlg_uncheck_all"), command=lambda: set_all_checks(False)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(sel_btns, text=tr("afdlg_only_blank"), command=lambda: (set_all_checks(False), check_vars.get("__BLANK__", tk.BooleanVar()).set(True))).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(sel_btns, text=tr("afdlg_only_val"), command=lambda: (set_all_checks(True), check_vars.get("__BLANK__", tk.BooleanVar()).set(False))).pack(side=tk.LEFT)

        col_combo.bind("<<ComboboxSelected>>", populate_checkboxes)
        search_var.trace_add("write", populate_checkboxes)

        def apply_column_filter(close_dialog=True):
            curr_col = self.heading_to_col.get(col_var.get(), "po_line")
            selected_vals = {k for k, v in check_vars.items() if v.get()}
            all_vals = set(check_vars.keys())
            q_text = search_var.get().strip()
            if selected_vals == all_vals and not q_text:
                self.active_column_filters.pop(curr_col, None)
            else:
                self.active_column_filters[curr_col] = {
                    "mode": "values",
                    "allowed": selected_vals,
                    "text_contains": q_text,
                }
            self._refresh_edit_tree()
            if close_dialog:
                dialog.destroy()

        def remove_this_column_filter():
            curr_col = self.heading_to_col.get(col_var.get(), "po_line")
            self.active_column_filters.pop(curr_col, None)
            self._refresh_edit_tree()
            populate_checkboxes()

        btn_bar = ttk.Frame(dialog, padding=10, style="App.TFrame")
        btn_bar.grid(row=3, column=0, sticky="ew")
        ttk.Button(btn_bar, text=tr("afdlg_remove_col"), command=remove_this_column_filter).pack(side=tk.LEFT)
        ttk.Button(btn_bar, text=tr("dlg_cancel"), command=dialog.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_bar, text=tr("afdlg_apply_next"), command=lambda: apply_column_filter(close_dialog=False)).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btn_bar, text=tr("dlg_accept"), command=lambda: apply_column_filter(close_dialog=True)).pack(side=tk.RIGHT, padx=6)

        populate_checkboxes()
        search_entry.focus_set()

    def _describe_active_filters(self):
        parts = []
        status_code = self._resolve_status_filter_code()
        if status_code != "all":
            parts.append(f"{tr('ed_flt_state_lbl')} {tr(STATUS_FILTER_KEYS.get(status_code, 'ed_flt_st_all'))}")
        for col in self.edit_columns:
            rule = self.active_column_filters.get(col)
            if not rule:
                continue
            label = self.edit_headings.get(col, col)
            if rule.get("mode") == "condition":
                op_code = rule.get("op_code") or self._resolve_filter_op_code(rule.get("op"))
                op_lbl = tr(FILTER_OP_KEYS.get(op_code, "ed_flt_op_contains"))
                val = rule.get("value", "")
                if op_code in {"blank", "not_blank"}:
                    parts.append(f"{label}: {op_lbl}")
                else:
                    parts.append(f"{label} ({op_lbl}: '{val}')")
            else:
                allowed = rule.get("allowed", set())
                q = rule.get("text_contains", "")
                if allowed == {"__BLANK__"}:
                    parts.append(f"{label}: {tr('ed_flt_val_blank')}")
                elif len(allowed) <= 2:
                    shown = ", ".join(tr("ed_flt_val_blank") if v == "__BLANK__" else f"'{v}'" for v in sorted(allowed))
                    parts.append(tr("ed_flt_in_list", col=label, items=shown))
                else:
                    desc = tr("ed_flt_values_cnt", col=label, count=len(allowed))
                    if q:
                        desc += tr("ed_flt_contains_q", q=q)
                    parts.append(desc)
        if not parts:
            return tr("ed_flt_none")
        return tr("ed_flt_active_prefix") + "  |  ".join(parts)

    def _update_column_filter_headings(self):
        if not hasattr(self, "edit_tree"):
            return
        tag_lbl = tr("ed_flt_tag")
        for col in self.edit_columns:
            base = self.edit_headings[col]
            if col in self.active_column_filters:
                self.edit_tree.heading(col, text=f"▼ [{tag_lbl}] {base}")
            else:
                self.edit_tree.heading(col, text=f"{base} ▼")

    def _row_matches_column_filters(self, values):
        for col, rule in self.active_column_filters.items():
            cell_str = str(values.get(col, "") or "").strip()
            mode = rule.get("mode")
            if mode == "condition":
                op_code = rule.get("op_code") or self._resolve_filter_op_code(rule.get("op"))
                target = str(rule.get("value", "") or "").strip()
                if op_code == "blank" or self._is_blank_sentinel(target):
                    if op_code == "not_equal":
                        if cell_str == "":
                            return False
                    elif cell_str != "":
                        return False
                elif op_code == "not_blank" or self._is_not_blank_sentinel(target):
                    if op_code == "not_equal":
                        if cell_str != "":
                            return False
                    elif cell_str == "":
                        return False
                elif op_code == "exact":
                    if cell_str.casefold() != target.casefold():
                        return False
                elif op_code == "not_equal":
                    if cell_str.casefold() == target.casefold():
                        return False
                else:
                    if target.casefold() not in cell_str.casefold():
                        return False
            elif mode == "values":
                allowed = rule.get("allowed", set())
                if cell_str == "":
                    if "__BLANK__" not in allowed:
                        return False
                else:
                    if cell_str not in allowed:
                        return False
        return True

    def _evaluate_line_state(self, key, values):
        po_val = str(values.get("po_line") or "").strip()
        ltype = str(values.get("line_type") or "REGULAR").strip().upper()
        is_unmatched = (not po_val) and (ltype == "REGULAR")
        diag = self.edit_line_diagnostics.get(key, {})
        notes = diag.get("validation_notes", [])
        if key in self.manual_po_assignments and po_val:
            has_warning = False
        elif ltype != "REGULAR":
            has_warning = False
        else:
            has_warning = bool(notes)
        is_problem = is_unmatched or has_warning
        is_edited = bool(self.line_overrides.get(key))
        return is_unmatched, has_warning, is_problem, is_edited

    # --------------------------------------------------------------------------
    # GESTIÓN DE FACTURA Y BORRADOR EN EDITAR
    # --------------------------------------------------------------------------
    def _on_edit_invoice_selected(self, _event=None):
        self._close_all_cell_editors(commit=True)
        path = self._active_edit_path()
        if path:
            self._sync_number_controls_for_path(path)
        self._refresh_edit_tree()

    def _sync_number_controls_for_path(self, path):
        if not path:
            return
        source_key = os.path.normcase(os.path.abspath(path))
        original = self.invoice_metadata.get(path, {}).get("original_number", "")
        self.edit_original_number_var.set(original)
        saved_mode_affix = self.invoice_number_mode_by_source.get(source_key)
        if saved_mode_affix:
            mode, affix = saved_mode_affix
        else:
            override = self.invoice_number_overrides.get(source_key)
            if override and original and override != original:
                if override.endswith(original):
                    mode = "prefix"
                    affix = override[: len(override) - len(original)]
                elif override.startswith(original):
                    mode = "suffix"
                    affix = override[len(original) :]
                else:
                    mode = "prefix"
                    affix = self.number_affix_var.get() or f"[{datetime.now().year}]"
            else:
                mode = "original"
                affix = self.number_affix_var.get() or f"[{datetime.now().year}]"
        self._updating_number_controls = True
        try:
            self.number_mode_var.set(mode)
            self.number_affix_var.set(affix)
            if mode == "suffix":
                preview_val = f"{original}{affix}"
            elif mode == "prefix":
                preview_val = f"{affix}{original}"
            else:
                preview_val = original
            self.number_preview_var.set(preview_val)
            if hasattr(self, "number_affix_entry"):
                self.number_affix_entry.configure(state="normal" if mode in {"suffix", "prefix"} else "disabled")
        finally:
            self._updating_number_controls = False

    def _refresh_edit_selector(self, select_path=None):
        labels = []
        self.edit_invoice_by_label = {}
        for path in self.invoice_files:
            meta = self.invoice_metadata.get(path, {})
            number = meta.get("number") or tr("vw_no_number")
            label = f"{number} | {meta.get('vendor', '')} | {os.path.basename(path)}"
            labels.append(label)
            self.edit_invoice_by_label[label] = path
        self.edit_invoice_selector["values"] = labels
        if not labels:
            self.edit_invoice_var.set("")
            self.edit_original_number_var.set("")
            self.number_preview_var.set("")
            self._refresh_edit_tree()
            return
        target = next((label for label, path in self.edit_invoice_by_label.items() if path == select_path), labels[0])
        self.edit_invoice_var.set(target)
        path = self.edit_invoice_by_label[target]
        self._sync_number_controls_for_path(path)
        self._refresh_edit_tree()

    def _active_edit_path(self):
        return self.edit_invoice_by_label.get(self.edit_invoice_var.get())

    def _apply_effective_number_to_active_invoice(self, effective_number, refresh_selectors=False):
        path = self._active_edit_path()
        if not path:
            return
        source_key = os.path.normcase(os.path.abspath(path))
        original = self.invoice_metadata.get(path, {}).get("original_number", effective_number)
        mode = self.number_mode_var.get()
        affix = self.number_affix_var.get()
        if mode == "original" or not effective_number or effective_number == original:
            self.invoice_number_overrides.pop(source_key, None)
            self.invoice_number_mode_by_source[source_key] = ("original", affix)
            effective_number = original
        else:
            self.invoice_number_overrides[source_key] = effective_number
            self.invoice_number_mode_by_source[source_key] = (mode, affix)
        if path in self.invoice_metadata:
            self.invoice_metadata[path]["number"] = effective_number
        if hasattr(self, "header_tree") and self.header_tree.exists("header"):
            self.header_tree.set("header", "invoice_number", effective_number)
        if hasattr(self, "edit_tree"):
            for row_id in self.edit_tree.get_children():
                self.edit_tree.set(row_id, "invoice", effective_number)
        self._refresh_batch_tree()
        if refresh_selectors:
            self._refresh_viewer_selector(select_path=path)
            self._refresh_edit_selector(select_path=path)

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
        if hasattr(self, "number_affix_entry"):
            self.number_affix_entry.configure(state="normal" if mode in {"suffix", "prefix"} else "disabled")
        if self._updating_number_controls:
            return
        self.edit_changes_saved = False
        self._apply_effective_number_to_active_invoice(value.strip() or original, refresh_selectors=False)

    def apply_edit_invoice_number(self):
        path = self._active_edit_path()
        if not path:
            return
        value = self.number_preview_var.get().strip()
        if not value:
            return
        self.edit_changes_saved = False
        self._apply_effective_number_to_active_invoice(value, refresh_selectors=True)
        self._set_status("st_num_updated", number=self.invoice_metadata[path]["number"])

    def _current_edit_signature(self):
        invoice_signature = tuple(os.path.normcase(os.path.abspath(path)) for path in self.invoice_files)
        return (
            invoice_signature,
            os.path.normcase(os.path.abspath(self.alma_pol_file.get().strip())) if self.alma_pol_file.get().strip() else "",
            os.path.normcase(os.path.abspath(self.template_file.get().strip())) if self.template_file.get().strip() else "",
        )

    def _invalidate_edit_draft(self, clear_overrides=True):
        self._close_all_cell_editors(commit=False)
        self._hide_excel_selection_overlay()
        self.edit_draft_signature = None
        self.edit_original_values.clear()
        self.edit_line_diagnostics.clear()
        self.edit_header_original_values.clear()
        if clear_overrides:
            self.line_overrides.clear()
            self.manual_po_assignments.clear()
            self.invoice_number_overrides.clear()
            self.invoice_number_mode_by_source.clear()
            self.invoice_header_overrides.clear()
            self.saved_line_overrides.clear()
            self.saved_manual_assignments.clear()
            self.saved_number_overrides.clear()
            self.saved_header_overrides.clear()
            self.undo_stack.clear()
            self.redo_stack.clear()
            self.active_column_filters.clear()
            self.active_cell = None
            self.selection_anchor = None
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
        self.invoice_number_mode_by_source = {key: value for key, value in self.invoice_number_mode_by_source.items() if key in valid_sources}
        self.invoice_header_overrides = {key: value for key, value in self.invoice_header_overrides.items() if key in valid_sources}

    def prepare_edit_draft(self, silent=False):
        self._close_all_cell_editors(commit=True)
        self._last_draft_error = ""
        try:
            if not self.invoice_files:
                raise ValueError(tr("val_need_inv_edit"))
            allowed = {".xsig", ".xml", ".txt"}
            for path in self.invoice_files:
                if os.path.splitext(path)[1].lower() not in allowed:
                    raise ValueError(tr("val_bad_ext", path=path))
                if not os.path.isfile(path):
                    raise FileNotFoundError(tr("val_no_inv_file", path=path))

            self._prune_edit_state_to_current_invoices()
            vendor_map = self.preferences.get("vendor_map", {})
            invoices = parse_invoice_paths(self.invoice_files, self.invoice_number_overrides, custom_vendor_map=vendor_map)
            validate_single_vendor_batch(invoices)

            alma_path = self.alma_pol_file.get().strip()
            issues_by_inv_line = {}
            reconciliation_by_inv = {}
            if alma_path and os.path.isfile(alma_path):
                lookup = get_cached_alma_lookup(alma_path)
                self.edit_lookup = lookup
                enricher = AlmaLineEnricher(lookup, self.manual_po_assignments)
                enricher.enrich(invoices)
                for issue in enricher.validation_issues:
                    if issue.issue_type == "REVISAR CUADRE IMPORTE":
                        reconciliation_by_inv[issue.invoice_number] = issue.message
                    elif issue.issue_type != "ASIGNACION MANUAL" and issue.invoice_line_number:
                        issues_by_inv_line.setdefault(
                            (issue.invoice_number, issue.invoice_line_number), []
                        ).append(f"{issue.issue_type}: {issue.message}")
            else:
                self.edit_lookup = None

            refreshed_originals = {}
            refreshed_diagnostics = {}
            refreshed_reconciliation = {}
            for invoice in invoices:
                source_key = os.path.normcase(os.path.abspath(invoice.source_file or ""))
                if invoice.invoice_number in reconciliation_by_inv:
                    refreshed_reconciliation[source_key] = reconciliation_by_inv[invoice.invoice_number]
                for line in invoice.lines:
                    key = (source_key, line.line_number)
                    refreshed_originals[key] = self._editable_line_values(invoice, line)
                    refreshed_diagnostics[key] = {
                        "match_status": line.match_status,
                        "match_message": line.match_message or "",
                        "validation_notes": issues_by_inv_line.get(
                            (invoice.invoice_number, line.line_number), []
                        ),
                    }

            valid_keys = set(refreshed_originals)
            self.line_overrides = {key: value for key, value in self.line_overrides.items() if key in valid_keys}
            self.manual_po_assignments = {key: value for key, value in self.manual_po_assignments.items() if key in valid_keys}
            self.edit_original_values = refreshed_originals
            self.edit_line_diagnostics = refreshed_diagnostics
            self.edit_invoice_reconciliation = refreshed_reconciliation
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
            self._on_filter_col_changed()
            if self.edit_lookup is not None:
                self._set_status("st_draft_ready_alma")
            else:
                self._set_status("st_draft_ready_no_alma")
            if not silent:
                if self.edit_lookup is not None:
                    messagebox.showinfo(tr("tab_edit"), tr("mb_draft_rebuilt_alma"))
                else:
                    messagebox.showinfo(tr("tab_edit"), tr("mb_draft_rebuilt_no_alma"))
            return True
        except (ValueError, FileNotFoundError) as exc:
            self._last_draft_error = str(exc)
            self.edit_draft_signature = None
            self.edit_original_values.clear()
            self.edit_line_diagnostics.clear()
            self.edit_header_original_values.clear()
            self._refresh_edit_tree()
            if not silent:
                messagebox.showerror(tr("mb_err_prep_draft"), str(exc))
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
        if not hasattr(self, "edit_tree"):
            return
        self._close_all_cell_editors(commit=True)
        for item in self.edit_tree.get_children():
            self.edit_tree.delete(item)
        if hasattr(self, "header_tree"):
            for item in self.header_tree.get_children():
                self.header_tree.delete(item)
        self._update_column_filter_headings()
        self.active_filters_label_var.set(self._describe_active_filters())
        path = self._active_edit_path()
        if not path:
            self.edit_stats_var.set(tr("ed_stats_empty"))
            self._hide_excel_selection_overlay()
            return
        source_key = os.path.normcase(os.path.abspath(path))
        header_values = dict(self.edit_header_original_values.get(source_key, {}))
        header_values.update(self.invoice_header_overrides.get(source_key, {}))
        effective_number = (
            self.invoice_number_overrides.get(source_key)
            or self.invoice_metadata.get(path, {}).get("number")
            or header_values.get("invoice_number", "")
        )
        if header_values:
            header_values["invoice_number"] = effective_number
            self.header_tree.insert("", "end", iid="header", values=[header_values.get(col, "") for col in self.header_columns])

        status_code = self._resolve_status_filter_code()
        self.edit_filter_code = status_code
        quick_query = self.edit_quick_search_var.get().strip().casefold()
        total_lines = 0
        ok_lines = 0
        problem_lines = 0
        unmatched_lines = 0
        shown_lines = 0

        for key, original in sorted(self.edit_original_values.items(), key=lambda item: item[0][1]):
            if key[0] != source_key:
                continue
            values = dict(original)
            values.update(self.line_overrides.get(key, {}))
            if effective_number:
                values["invoice"] = effective_number
            total_lines += 1
            is_unmatched, has_warning, is_problem, is_edited = self._evaluate_line_state(key, values)
            if is_unmatched:
                unmatched_lines += 1
            if is_problem:
                problem_lines += 1
            else:
                ok_lines += 1

            if status_code == "problems" and not is_problem:
                continue
            if status_code == "unmatched" and not is_unmatched:
                continue
            if status_code == "ok" and is_problem:
                continue
            if status_code == "edited" and not is_edited:
                continue

            if not self._row_matches_column_filters(values):
                continue

            if quick_query:
                if not any(quick_query in str(values.get(c, "") or "").casefold() for c in self.edit_columns):
                    continue

            iid = f"{key[1]}"
            if is_unmatched:
                row_tags = ("unmatched",)
            elif has_warning:
                row_tags = ("warning",)
            elif is_edited:
                row_tags = ("edited",)
            else:
                row_tags = ()
            self.edit_tree.insert("", "end", iid=iid, values=[values.get(col, "") for col in self.edit_columns], tags=row_tags)
            shown_lines += 1

        self.edit_stats_var.set(
            tr("ed_stats_fmt", shown=shown_lines, total=total_lines, ok=ok_lines, prob=problem_lines, unm=unmatched_lines)
        )
        self._on_filter_col_changed()
        self._update_visible_and_selected_totals()
        self._update_selected_line_diagnostic()
        self.root.after_idle(self._update_excel_selection_overlay)

    def _update_row_visual_tags_and_stats(self):
        path = self._active_edit_path()
        if not path:
            return
        source_key = os.path.normcase(os.path.abspath(path))
        total_lines = 0
        ok_lines = 0
        problem_lines = 0
        unmatched_lines = 0
        for key, original in self.edit_original_values.items():
            if key[0] != source_key:
                continue
            values = dict(original)
            values.update(self.line_overrides.get(key, {}))
            total_lines += 1
            is_unmatched, has_warning, is_problem, is_edited = self._evaluate_line_state(key, values)
            if is_unmatched:
                unmatched_lines += 1
            if is_problem:
                problem_lines += 1
            else:
                ok_lines += 1
            row_id = f"{key[1]}"
            if self.edit_tree.exists(row_id):
                if is_unmatched:
                    tags = ("unmatched",)
                elif has_warning:
                    tags = ("warning",)
                elif is_edited:
                    tags = ("edited",)
                else:
                    tags = ()
                self.edit_tree.item(row_id, tags=tags)
        shown_lines = len(self.edit_tree.get_children())
        self.edit_stats_var.set(
            tr("ed_stats_fmt", shown=shown_lines, total=total_lines, ok=ok_lines, prob=problem_lines, unm=unmatched_lines)
        )

    # --------------------------------------------------------------------------
    # CICLO DE VIDA LIMPIO DEL EDITOR EN LÍNEA
    # --------------------------------------------------------------------------
    def _close_all_cell_editors(self, commit=True):
        if self.edit_active_editor is not None:
            editor_widget, row_id, column, committed_ref = self.edit_active_editor
            self.edit_active_editor = None
            self.edit_active_context = None
            self.edit_combo_editor = None
            self.edit_combo_context = None
            if not committed_ref[0] and commit and editor_widget.winfo_exists():
                committed_ref[0] = True
                val = editor_widget.get()
                self._set_edit_value(row_id, column, val, update_widget=True)
            if editor_widget.winfo_exists():
                editor_widget.destroy()

        if self.header_active_editor is not None:
            h_widget, h_col, h_committed_ref = self.header_active_editor
            self.header_active_editor = None
            self.header_cell_editor = None
            self.header_editor_context = None
            if not h_committed_ref[0] and commit and h_widget.winfo_exists():
                h_committed_ref[0] = True
                val = h_widget.get()
                self._set_header_value(h_col, val)
            if h_widget.winfo_exists():
                h_widget.destroy()

    def _begin_edit_cell(self, event):
        region = self.edit_tree.identify_region(event.x, event.y)
        if region == "heading":
            return
        row_id = self.edit_tree.identify_row(event.y)
        col_id = self.edit_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_index = int(col_id[1:]) - 1
        column = self.edit_columns[col_index]
        self.last_edited_column = column
        self.edit_context_column = column
        self.active_cell = (row_id, column)
        if column in READ_ONLY_LINE_FIELDS:
            messagebox.showinfo(tr("mb_prot_field_title"), tr("mb_prot_line_msg"))
            return
        self._open_tree_cell_editor(row_id, column, col_id)

    def _open_tree_cell_editor(self, row_id, column, col_id=None, initial_text=None):
        self._close_all_cell_editors(commit=True)
        if column in READ_ONLY_LINE_FIELDS:
            messagebox.showinfo(tr("mb_prot_field_title"), tr("mb_prot_line_msg"))
            return
        if col_id is None:
            col_id = f"#{self.edit_columns.index(column) + 1}"
        bbox = self.edit_tree.bbox(row_id, col_id)
        if not bbox:
            self._scroll_tree_to_column(column)
            self.root.update_idletasks()
            bbox = self.edit_tree.bbox(row_id, col_id)
            if not bbox:
                return
        self._hide_excel_selection_overlay()
        x, y, width, height = bbox
        current = self.edit_tree.set(row_id, column)
        start_val = initial_text if initial_text is not None else current
        values = None
        if column == "line_type":
            values = LINE_TYPE_OPTIONS
        elif column in {"reporting_code", "secondary_reporting_code", "tertiary_reporting_code", "fourth_reporting_code", "fifth_reporting_code", "fund_1", "fund_2"}:
            values = self._context_suggestions(column)
        if values:
            editor = ttk.Combobox(self.edit_tree, values=values, state="normal", exportselection=False)
            editor.set(start_val)
            editor.bind("<MouseWheel>", lambda _e: "break")
            editor.bind("<Button-4>", lambda _e: "break")
            editor.bind("<Button-5>", lambda _e: "break")
        else:
            editor = ttk.Entry(self.edit_tree, exportselection=False)
            editor.insert(0, start_val)
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        if initial_text is not None:
            if hasattr(editor, "icursor"):
                editor.icursor(tk.END)
        elif hasattr(editor, "select_range"):
            editor.select_range(0, tk.END)

        committed_flag = [False]
        self.edit_active_editor = (editor, row_id, column, committed_flag)
        self.edit_active_context = (row_id, column)
        self.edit_combo_editor = editor
        self.edit_combo_context = (row_id, column)
        self.active_cell = (row_id, column)

        def commit_and_close(_event=None):
            if committed_flag[0]:
                return "break"
            committed_flag[0] = True
            val = editor.get() if editor.winfo_exists() else current
            self._set_edit_value(row_id, column, val, update_widget=True)
            self._close_all_cell_editors(commit=False)
            self.root.after_idle(self._update_excel_selection_overlay)
            return "break"

        def cancel_and_close(_event=None):
            committed_flag[0] = True
            self._close_all_cell_editors(commit=False)
            self.root.after_idle(self._update_excel_selection_overlay)
            return "break"

        def on_return(_event=None):
            commit_and_close()
            children = list(self.edit_tree.get_children())
            if row_id in children:
                idx = children.index(row_id)
                if idx + 1 < len(children):
                    next_row = children[idx + 1]
                    self.active_cell = (next_row, column)
                    self.selection_anchor = (next_row, column)
                    self.edit_tree.selection_set(next_row)
                    self.edit_tree.focus(next_row)
                    self.edit_tree.see(next_row)
            self.root.after_idle(self._update_excel_selection_overlay)
            return "break"

        def on_tab(_event=None):
            commit_and_close()
            self.root.after_idle(lambda: self._focus_next_editable_cell(row_id, column))
            return "break"

        def on_shift_tab(_event=None):
            commit_and_close()
            self._on_tree_shift_tab_pressed()
            return "break"

        def on_focus_out(_event=None):
            def _deferred_focus_check():
                if committed_flag[0] or not editor.winfo_exists():
                    return
                try:
                    focused = self.root.focus_get()
                except Exception:
                    focused = None
                if focused is not None and (focused == editor or str(focused).startswith(str(editor))):
                    return
                commit_and_close()
            self.root.after(90, _deferred_focus_check)

        editor.bind("<Return>", on_return)
        editor.bind("<KP_Enter>", on_return)
        editor.bind("<Tab>", on_tab)
        editor.bind("<Shift-Tab>", on_shift_tab)
        editor.bind("<Escape>", cancel_and_close)
        editor.bind("<FocusOut>", on_focus_out)
        if isinstance(editor, ttk.Combobox):
            editor.bind("<<ComboboxSelected>>", commit_and_close)

    def _commit_active_line_editor(self):
        self._close_all_cell_editors(commit=True)
        return True

    def _focus_next_editable_cell(self, row_id, column):
        editable = [name for name in self.edit_columns if name in EDITABLE_LINE_FIELDS]
        if column not in editable:
            return
        rows = list(self.edit_tree.get_children())
        if row_id not in rows:
            return
        column_index = editable.index(column)
        row_index = rows.index(row_id)
        if column_index + 1 < len(editable):
            next_row, next_column = row_id, editable[column_index + 1]
        elif row_index + 1 < len(rows):
            next_row, next_column = rows[row_index + 1], editable[0]
        else:
            return
        self.active_cell = (next_row, next_column)
        self.selection_anchor = (next_row, next_column)
        self.last_edited_column = next_column
        self.edit_context_column = next_column
        self.edit_tree.selection_set(next_row)
        self.edit_tree.focus(next_row)
        self.edit_tree.see(next_row)
        self._scroll_tree_to_column(next_column)
        self._update_excel_selection_overlay()

    def _focus_next_header_cell(self, column):
        editable = [name for name in self.header_columns if name in EDITABLE_HEADER_FIELDS]
        if column not in editable:
            return
        index = editable.index(column)
        if index + 1 < len(editable):
            self._open_header_cell_editor(editable[index + 1])

    def _get_current_line_cell_value(self, key, column):
        if key in self.line_overrides and column in self.line_overrides[key]:
            return self.line_overrides[key][column]
        return self.edit_original_values.get(key, {}).get(column, "")

    def _push_undo_state(self, action_name="Edición"):
        snapshot = (
            "snapshot",
            action_name,
            copy.deepcopy(self.line_overrides),
            copy.deepcopy(self.manual_po_assignments),
        )
        self.undo_stack.append(snapshot)
        self.redo_stack.clear()
        if len(self.undo_stack) > 40:
            self.undo_stack.pop(0)

    def _set_edit_value(self, row_id, column, value, update_widget=True, record_undo=True):
        path = self._active_edit_path()
        if not path:
            return False
        source_key = os.path.normcase(os.path.abspath(path))
        key = (source_key, int(row_id))
        if column == "line_type" and value not in LINE_TYPE_OPTIONS:
            messagebox.showerror(tr("mb_invalid_val_title"), tr("mb_invalid_ltype"))
            return False
        if column in {"price", "quantity", "vat_amount", "vat_percentage"} and str(value).strip() != "":
            try:
                Decimal(str(value).replace(",", "."))
            except InvalidOperation:
                messagebox.showerror(tr("mb_invalid_val_title"), tr("mb_invalid_num", col=column))
                return False
        clean_value = sanitize_excel_value(value.strip()) if isinstance(value, str) else value
        previous_value = self._get_current_line_cell_value(key, column)
        if record_undo and str(previous_value) != str(clean_value):
            self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        self.last_edited_column = column
        original_val = self.edit_original_values.get(key, {}).get(column, "")
        if str(clean_value) == str(original_val):
            self.line_overrides.get(key, {}).pop(column, None)
            if not self.line_overrides.get(key):
                self.line_overrides.pop(key, None)
        else:
            self.line_overrides.setdefault(key, {})[column] = clean_value
        if update_widget and self.edit_tree.exists(row_id):
            self.edit_tree.set(row_id, column, clean_value)
        if column == "po_line":
            if str(clean_value) != str(original_val):
                self.manual_po_assignments[key] = clean_value if clean_value else None
            else:
                self.manual_po_assignments.pop(key, None)
        self._update_row_visual_tags_and_stats()
        self.edit_changes_saved = False
        return True

    def undo_edit(self, _event=None):
        self._close_all_cell_editors(commit=False)
        if not self.undo_stack:
            self._set_status("st_no_undo")
            return "break"
        entry = self.undo_stack.pop()
        if entry[0] == "snapshot":
            _, action_name, prev_overrides, prev_manual = entry
            self.redo_stack.append((
                "snapshot",
                action_name,
                copy.deepcopy(self.line_overrides),
                copy.deepcopy(self.manual_po_assignments),
            ))
            self.line_overrides = prev_overrides
            self.manual_po_assignments = prev_manual
            self.edit_changes_saved = False
            self._refresh_edit_tree()
            self._set_status("st_undone", action=action_name)
        return "break"

    def redo_edit(self, _event=None):
        self._close_all_cell_editors(commit=False)
        if not self.redo_stack:
            self._set_status("st_no_redo")
            return "break"
        entry = self.redo_stack.pop()
        if entry[0] == "snapshot":
            _, action_name, next_overrides, next_manual = entry
            self.undo_stack.append((
                "snapshot",
                action_name,
                copy.deepcopy(self.line_overrides),
                copy.deepcopy(self.manual_po_assignments),
            ))
            self.line_overrides = next_overrides
            self.manual_po_assignments = next_manual
            self.edit_changes_saved = False
            self._refresh_edit_tree()
            self._set_status("st_redone", action=action_name)
        return "break"

    # --------------------------------------------------------------------------
    # OPERACIONES TIPO EXCEL: RELLENAR ABAJO, BLANCOS, COPIAR, PEGAR, VACIAR
    # --------------------------------------------------------------------------
    def _active_editable_column(self):
        if self.active_cell and self.active_cell[1]:
            return self.active_cell[1]
        return getattr(self, "edit_context_column", None) or getattr(self, "last_edited_column", None)

    def _excel_fill_down(self, _event=None):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        if not column or column in READ_ONLY_LINE_FIELDS:
            return "break"
        selected = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if not selected and self.active_cell and self.edit_tree.exists(self.active_cell[0]):
            selected = [self.active_cell[0]]
        if not selected:
            return "break"
        all_rows = list(self.edit_tree.get_children())
        if len(selected) == 1:
            row_id = selected[0]
            if row_id in all_rows:
                idx = all_rows.index(row_id)
                if idx > 0:
                    prev_val = self.edit_tree.set(all_rows[idx - 1], column)
                    self._set_edit_value(row_id, column, prev_val, update_widget=True, record_undo=True)
        else:
            source_val = self.edit_tree.set(selected[0], column)
            self._push_undo_state(f"{self.edit_headings.get(column, column)} (Ctrl+D)")
            for row_id in selected[1:]:
                self._set_edit_value(row_id, column, source_val, update_widget=True, record_undo=False)
        self._update_excel_selection_overlay()
        return "break"

    def _fill_down_shortcut(self, _event=None):
        return self._excel_fill_down()

    def _fill_blanks_downward_action(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        if not column or column in READ_ONLY_LINE_FIELDS:
            return
        selected = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        start_row = (self.active_cell[0] if self.active_cell else None) or (selected[0] if selected else None)
        if not start_row or not self.edit_tree.exists(start_row):
            return
        source_val = self.edit_tree.set(start_row, column)
        if not source_val:
            messagebox.showinfo(tr("mb_empty_src_title"), tr("mb_empty_src_msg"))
            return
        children = list(self.edit_tree.get_children())
        if start_row not in children:
            return
        start_idx = children.index(start_row)
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        changed = 0
        for r_id in children[start_idx + 1:]:
            if not self.edit_tree.set(r_id, column):
                self._set_edit_value(r_id, column, source_val, update_widget=True, record_undo=False)
                changed += 1
            else:
                break
        if changed > 0:
            self._set_status("st_fill_blanks", count=changed, val=source_val)
        self._update_excel_selection_overlay()

    def _copy_cell_value(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        selected = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        row_id = (self.active_cell[0] if self.active_cell else None) or (selected[0] if selected else None)
        if not row_id or not column or not self.edit_tree.exists(row_id):
            return
        val = self.edit_tree.set(row_id, column)
        self._clipboard_cell_value = val
        self.root.clipboard_clear()
        self.root.clipboard_append(val or "")
        self._set_status("st_copied", val=val, col=self.edit_headings.get(column, column))

    def _paste_cell_value(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        if not column or column in READ_ONLY_LINE_FIELDS:
            return
        val = self._clipboard_cell_value
        if val is None:
            try:
                val = self.root.clipboard_get()
            except Exception:
                val = ""
        targets = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if not targets and self.active_cell and self.edit_tree.exists(self.active_cell[0]):
            targets = [self.active_cell[0]]
        if not targets:
            return
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        for row_id in targets:
            self._set_edit_value(row_id, column, val, update_widget=True, record_undo=False)
        self._update_excel_selection_overlay()

    def open_po_line_assistant(self):
        self._close_all_cell_editors(commit=True)
        if not self.invoice_files:
            messagebox.showinfo(tr("mb_pol_title"), tr("mb_pol_no_inv"))
            return
        if not self.edit_original_values:
            self.prepare_edit_draft(silent=True)
        if not self.edit_lookup or not getattr(self.edit_lookup, "records", None):
            alma_path = self.alma_pol_file.get().strip()
            if not alma_path or not os.path.isfile(alma_path):
                if messagebox.askyesno(tr("mb_pol_title"), tr("mb_pol_ask_alma")):
                    self.select_alma_pol()
                alma_path = self.alma_pol_file.get().strip()
            if alma_path and os.path.isfile(alma_path):
                self.prepare_edit_draft(silent=True)
            if not self.edit_lookup or not getattr(self.edit_lookup, "records", None):
                return
        selected = self.edit_tree.selection()
        if not selected:
            children = self.edit_tree.get_children()
            if not children:
                messagebox.showinfo(tr("mb_pol_title"), tr("mb_pol_no_lines"))
                return
            row_id = children[0]
            self.edit_tree.selection_set(row_id)
        else:
            row_id = selected[0]

        line_num = self.edit_tree.set(row_id, "line")
        line_title = self.edit_tree.set(row_id, "title")
        line_isbn = self.edit_tree.set(row_id, "isbn")
        line_issn = self.edit_tree.set(row_id, "issn")
        line_price = self.edit_tree.set(row_id, "price")
        current_pol = self.edit_tree.set(row_id, "po_line")

        dialog = tk.Toplevel(self.root)
        dialog.title(tr("poldlg_title", line=line_num))
        dialog.geometry("980x540")
        dialog.minsize(820, 420)
        dialog.configure(bg=UI_BG)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(2, weight=1)

        info_frame = ttk.LabelFrame(dialog, text=tr("poldlg_line_frame", line=line_num), padding=10, style="App.TLabelframe")
        info_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        info_frame.columnconfigure(1, weight=1)
        id_summary = tr(
            "poldlg_summary",
            isbn=line_isbn or "-",
            issn=line_issn or "-",
            price=line_price,
            pol=current_pol or tr("poldlg_unassigned"),
        )
        ttk.Label(info_frame, text=tr("poldlg_inv_title", title=line_title), style="App.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(info_frame, text=id_summary, style="App.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        search_frame = ttk.Frame(dialog, style="App.TFrame")
        search_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 6))
        search_frame.columnconfigure(1, weight=1)
        ttk.Label(search_frame, text=tr("poldlg_search_lbl"), style="App.TLabel").grid(row=0, column=0, padx=(0, 8))
        initial_query = extract_main_title(line_title) or line_isbn or line_issn
        query_var = tk.StringVar(value=initial_query)
        search_entry = ttk.Entry(search_frame, textvariable=query_var)
        search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))

        results_frame = ttk.Frame(dialog, style="App.TFrame")
        results_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 8))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        cols = ("po_line", "score", "isbn", "issn", "title", "reporting_code", "fund")
        tree = ttk.Treeview(results_frame, columns=cols, show="headings", selectmode="browse")
        tree.heading("po_line", text="PO Line")
        tree.heading("score", text=tr("poldlg_col_score"))
        tree.heading("isbn", text="ISBN")
        tree.heading("issn", text="ISSN")
        tree.heading("title", text="Title (Alma Analytics)")
        tree.heading("reporting_code", text="Reporting Code")
        tree.heading("fund", text="Fund")
        tree.column("po_line", width=110, anchor="w")
        tree.column("score", width=75, anchor="e")
        tree.column("isbn", width=120, anchor="w")
        tree.column("issn", width=90, anchor="w")
        tree.column("title", width=360, anchor="w")
        tree.column("reporting_code", width=110, anchor="w")
        tree.column("fund", width=120, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(results_frame, orient="vertical", command=tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=yscroll.set)

        records_by_iid = {}

        def refresh_results(*_):
            for item in tree.get_children():
                tree.delete(item)
            records_by_iid.clear()
            matches = self.edit_lookup.search_catalog(query_var.get().strip(), ref_title=line_title, limit=80)
            for idx, (rec, score, _reasons) in enumerate(matches):
                iid = f"rec_{idx}"
                records_by_iid[iid] = rec
                score_txt = f"{score * 100:.0f}%" if score > 0 else "-"
                fund_txt = rec.funds[0] if rec.funds else ""
                tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(
                        rec.po_line,
                        score_txt,
                        rec.isbn_raw or "",
                        rec.issn_raw or "",
                        rec.title or "",
                        rec.reporting_code or "",
                        fund_txt,
                    ),
                )
            children = tree.get_children()
            if children:
                tree.selection_set(children[0])

        ttk.Button(search_frame, text=tr("poldlg_search_btn"), command=refresh_results).grid(row=0, column=2, padx=(0, 4))
        ttk.Button(search_frame, text=tr("poldlg_all_btn"), command=lambda: (query_var.set(""), refresh_results())).grid(row=0, column=3)
        query_var.trace_add("write", refresh_results)

        def assign_selected(include_metadata=False):
            sel = tree.selection()
            if not sel:
                messagebox.showinfo(tr("mb_pol_title"), tr("mb_pol_sel_row"), parent=dialog)
                return
            rec = records_by_iid.get(sel[0])
            if not rec:
                return
            self._push_undo_state(f"PO Line {rec.po_line}")
            self._set_edit_value(row_id, "po_line", rec.po_line, record_undo=False)
            if include_metadata:
                for field_name in ("reporting_code", "secondary_reporting_code", "tertiary_reporting_code", "fourth_reporting_code", "fifth_reporting_code"):
                    val = getattr(rec, field_name, None)
                    if val:
                        self._set_edit_value(row_id, field_name, val, record_undo=False)
                if rec.funds and rec.funds[0]:
                    self._set_edit_value(row_id, "fund_1", rec.funds[0], record_undo=False)
                if len(rec.funds) > 1 and rec.funds[1]:
                    self._set_edit_value(row_id, "fund_2", rec.funds[1], record_undo=False)
            dialog.destroy()
            self._update_excel_selection_overlay()

        tree.bind("<Double-1>", lambda _e: assign_selected(include_metadata=False))

        btn_frame = ttk.Frame(dialog, style="App.TFrame")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 12))
        ttk.Label(
            btn_frame,
            text=tr("poldlg_dblclick_hint"),
            style="App.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text=tr("dlg_cancel"), command=dialog.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(
            btn_frame,
            text=tr("poldlg_assign_full"),
            command=lambda: assign_selected(include_metadata=True),
        ).pack(side=tk.RIGHT, padx=6)
        ttk.Button(
            btn_frame,
            text=tr("poldlg_assign_only"),
            command=lambda: assign_selected(include_metadata=False),
        ).pack(side=tk.RIGHT, padx=6)

        refresh_results()
        search_entry.focus_set()
        search_entry.select_range(0, tk.END)

    def _context_suggestions(self, column):
        counts = {}
        for original in self.edit_original_values.values():
            value = original.get(column)
            if value:
                counts[value] = counts.get(value, 0) + 1
        for changes in self.line_overrides.values():
            value = changes.get(column)
            if value:
                counts[value] = counts.get(value, 0) + 1
        return [value for value, _count in sorted(counts.items(), key=lambda item: (-item[1], str(item[0])))]

    def _reporting_code_suggestions(self, column):
        return self._context_suggestions(column)

    def _show_edit_context_menu(self, event):
        self._close_all_cell_editors(commit=True)
        row_id = self.edit_tree.identify_row(event.y)
        col_id = self.edit_tree.identify_column(event.x)
        if row_id and row_id not in self.edit_tree.selection():
            self.edit_tree.selection_set(row_id)
        self.edit_context_column = self.edit_columns[int(col_id[1:]) - 1] if col_id else None
        if self.edit_context_column:
            self.last_edited_column = self.edit_context_column
            if row_id:
                self.active_cell = (row_id, self.edit_context_column)
                self._update_excel_selection_overlay()
        try:
            self.edit_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.edit_context_menu.grab_release()

    def _edit_selected_cell(self):
        selected = self.edit_tree.selection()
        col = self._active_editable_column()
        row_id = selected[0] if selected else (self.active_cell[0] if self.active_cell else None)
        if row_id and col:
            self._open_tree_cell_editor(row_id, col)

    def _clear_selected_edit_cell(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        if column in READ_ONLY_LINE_FIELDS or column is None:
            return
        targets = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if not targets and self.active_cell and self.edit_tree.exists(self.active_cell[0]):
            targets = [self.active_cell[0]]
        if not targets:
            return
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        for row_id in targets:
            self._set_edit_value(row_id, column, "", update_widget=True, record_undo=False)
        self._update_excel_selection_overlay()

    def _restore_selected_edit_cell(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        path = self._active_edit_path()
        if not column or not path or column in READ_ONLY_LINE_FIELDS:
            return
        source_key = os.path.normcase(os.path.abspath(path))
        targets = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if not targets and self.active_cell and self.edit_tree.exists(self.active_cell[0]):
            targets = [self.active_cell[0]]
        if not targets:
            return
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        for row_id in targets:
            key = (source_key, int(row_id))
            original = self.edit_original_values.get(key, {}).get(column, "")
            self._set_edit_value(row_id, column, original, update_widget=True, record_undo=False)
        self._update_excel_selection_overlay()

    def _suggest_selected_context_value(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        if not column or column in READ_ONLY_LINE_FIELDS:
            return
        suggestions = self._context_suggestions(column)
        if suggestions:
            targets = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
            if not targets and self.active_cell and self.edit_tree.exists(self.active_cell[0]):
                targets = [self.active_cell[0]]
            self._push_undo_state(f"{self.edit_headings.get(column, column)}")
            for row_id in targets:
                self._set_edit_value(row_id, column, suggestions[0], update_widget=True, record_undo=False)
            self._update_excel_selection_overlay()
        else:
            messagebox.showinfo(tr("mb_no_sugg_title"), tr("mb_no_sugg_msg"))

    def _prompt_bulk_value(self, column, initial=""):
        col_title = self.edit_headings.get(column, column)
        dialog = tk.Toplevel(self.root)
        dialog.title(tr("bulkdlg_title", col=col_title))
        dialog.transient(self.root)
        dialog.grab_set()
        value_var = tk.StringVar(value=initial)
        suggestions = LINE_TYPE_OPTIONS if column == "line_type" else self._context_suggestions(column)
        ttk.Label(dialog, text=tr("bulkdlg_lbl", col=col_title)).pack(anchor="w", padx=14, pady=(12, 4))
        editor = ttk.Combobox(dialog, textvariable=value_var, values=suggestions, state="normal", width=55)
        editor.pack(fill=tk.X, padx=14, pady=4)
        editor.focus_set()
        result = {"value": None}

        def accept():
            result["value"] = value_var.get()
            dialog.destroy()

        buttons = ttk.Frame(dialog)
        buttons.pack(pady=(6, 12))
        ttk.Button(buttons, text=tr("dlg_apply"), command=accept).pack(side=tk.LEFT, padx=4)
        ttk.Button(buttons, text=tr("dlg_cancel"), command=dialog.destroy).pack(side=tk.LEFT, padx=4)
        dialog.bind("<Return>", lambda _e: accept())
        self.root.wait_window(dialog)
        return result["value"]

    def _bulk_apply_selected_value(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        selected = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if not selected or column in READ_ONLY_LINE_FIELDS or column is None:
            return
        value = self._prompt_bulk_value(column, self.edit_tree.set(selected[0], column))
        if value is None:
            return
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        for row_id in selected:
            self._set_edit_value(row_id, column, value, update_widget=True, record_undo=False)
        self._update_excel_selection_overlay()

    def _fill_selected_blanks(self):
        self._close_all_cell_editors(commit=True)
        column = self._active_editable_column()
        selected = [r for r in self.edit_tree.selection() if self.edit_tree.exists(r)]
        if not selected or column in READ_ONLY_LINE_FIELDS or column is None:
            return
        source = next((self.edit_tree.set(row, column) for row in selected if self.edit_tree.set(row, column)), "")
        if not source:
            source = self._prompt_bulk_value(column, "")
        if source is None or source == "":
            return
        self._push_undo_state(f"{self.edit_headings.get(column, column)}")
        for row_id in selected:
            if not self.edit_tree.set(row_id, column):
                self._set_edit_value(row_id, column, source, update_widget=True, record_undo=False)
        self._update_excel_selection_overlay()

    def _fill_blanks_below(self):
        self._fill_blanks_downward_action()

    def _begin_edit_header_cell(self, event):
        self._close_all_cell_editors(commit=True)
        row_id = self.header_tree.identify_row(event.y)
        col_id = self.header_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        column = self.header_columns[int(col_id[1:]) - 1]
        if column in READ_ONLY_HEADER_FIELDS:
            messagebox.showinfo(tr("mb_prot_field_title"), tr("mb_prot_hdr_msg"))
            return
        self._open_header_cell_editor(column, col_id)

    def _open_header_cell_editor(self, column, col_id=None):
        self._close_all_cell_editors(commit=True)
        if col_id is None:
            col_id = f"#{self.header_columns.index(column) + 1}"
        bbox = self.header_tree.bbox("header", col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        current = self.header_tree.set("header", column)
        if column == "inclusive":
            editor = ttk.Combobox(self.header_tree, values=["False", "True"], state="readonly", exportselection=False)
            editor.set(current)
        else:
            editor = ttk.Entry(self.header_tree, exportselection=False)
            editor.insert(0, current)
        editor.place(x=x, y=y, width=w, height=h)
        editor.focus_set()
        if hasattr(editor, "select_range"):
            editor.select_range(0, tk.END)

        committed_flag = [False]
        self.header_active_editor = (editor, column, committed_flag)
        self.header_cell_editor = editor
        self.header_editor_context = column

        def commit_and_close(_event=None):
            if committed_flag[0]:
                return "break"
            committed_flag[0] = True
            val = editor.get() if editor.winfo_exists() else current
            self._set_header_value(column, val)
            self._close_all_cell_editors(commit=False)
            return "break"

        def cancel_and_close(_event=None):
            committed_flag[0] = True
            self._close_all_cell_editors(commit=False)
            return "break"

        def commit_tab(_event=None):
            commit_and_close()
            self.root.after_idle(lambda: self._focus_next_header_cell(column))
            return "break"

        def on_focus_out(_event=None):
            def _deferred_focus_check():
                if committed_flag[0] or not editor.winfo_exists():
                    return
                try:
                    focused = self.root.focus_get()
                except Exception:
                    focused = None
                if focused is not None and (focused == editor or str(focused).startswith(str(editor))):
                    return
                commit_and_close()
            self.root.after(90, _deferred_focus_check)

        editor.bind("<Return>", commit_and_close)
        editor.bind("<KP_Enter>", commit_and_close)
        editor.bind("<Tab>", commit_tab)
        editor.bind("<Escape>", cancel_and_close)
        editor.bind("<FocusOut>", on_focus_out)
        if isinstance(editor, ttk.Combobox):
            editor.bind("<<ComboboxSelected>>", commit_and_close)

    def _commit_active_header_editor(self):
        self._close_all_cell_editors(commit=True)
        return True

    def _flush_editors(self):
        self._close_all_cell_editors(commit=True)
        self.root.update_idletasks()
        return True

    def _set_header_value(self, column, value):
        path = self._active_edit_path()
        if not path:
            return False
        source_key = os.path.normcase(os.path.abspath(path))
        if column in {"total_amount", "vat_amount"} and value:
            try:
                Decimal(str(value).replace(",", "."))
            except InvalidOperation:
                messagebox.showerror(tr("mb_invalid_val_title"), tr("mb_invalid_num", col=column))
                return False
        if column == "date" and value and parse_date_for_excel(value) is None:
            messagebox.showerror(tr("mb_invalid_date_title"), tr("mb_invalid_date_msg"))
            return False
        clean_value = sanitize_excel_value(value.strip()) if isinstance(value, str) else value
        if column == "invoice_number":
            self.invoice_metadata[path]["number"] = clean_value
            self.invoice_number_overrides[source_key] = clean_value
        else:
            original = self.edit_header_original_values.get(source_key, {}).get(column, "")
            if str(clean_value) == str(original):
                self.invoice_header_overrides.get(source_key, {}).pop(column, None)
            else:
                self.invoice_header_overrides.setdefault(source_key, {})[column] = clean_value
            if column == "vendor_account" and clean_value:
                nif = self.invoice_metadata.get(path, {}).get("seller_tax_id", "")
                self._remember_vendor_for_nif(nif, clean_value)
                if path in self.invoice_metadata:
                    self.invoice_metadata[path]["vendor"] = clean_value
                self._refresh_batch_tree()
        self.header_tree.set("header", column, clean_value)
        self.edit_changes_saved = False
        return True

    def _show_header_context_menu(self, event):
        self._close_all_cell_editors(commit=True)
        row = self.header_tree.identify_row(event.y)
        col = self.header_tree.identify_column(event.x)
        if row:
            self.header_tree.selection_set(row)
        self.header_context_column = self.header_columns[int(col[1:]) - 1] if col else None
        try:
            self.header_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.header_context_menu.grab_release()

    def _edit_selected_header_cell(self):
        column = getattr(self, "header_context_column", None)
        if column in READ_ONLY_HEADER_FIELDS:
            messagebox.showinfo(tr("mb_prot_field_title"), tr("mb_prot_generic_msg"))
            return
        if column:
            self._open_header_cell_editor(column)

    def _restore_selected_header_cell(self):
        self._close_all_cell_editors(commit=True)
        path = self._active_edit_path()
        column = getattr(self, "header_context_column", None)
        if not path or not column:
            return
        source_key = os.path.normcase(os.path.abspath(path))
        original = self.edit_header_original_values.get(source_key, {}).get(column, "")
        self.invoice_header_overrides.get(source_key, {}).pop(column, None)
        self.header_tree.set("header", column, original)
        self.edit_changes_saved = False

    def _harvest_visible_editor(self):
        if not self._flush_editors():
            return False
        path = self._active_edit_path()
        if not path:
            return False
        source_key = os.path.normcase(os.path.abspath(path))

        if self.header_tree.exists("header"):
            for column in EDITABLE_HEADER_FIELDS:
                current = self.header_tree.set("header", column)
                original = self.edit_header_original_values.get(source_key, {}).get(column, "")
                if str(current) != str(original):
                    clean_current = sanitize_excel_value(current)
                    self.invoice_header_overrides.setdefault(source_key, {})[column] = clean_current
                    if column == "vendor_account" and clean_current:
                        nif = self.invoice_metadata.get(path, {}).get("seller_tax_id", "")
                        self._remember_vendor_for_nif(nif, clean_current)
                else:
                    self.invoice_header_overrides.get(source_key, {}).pop(column, None)
            if not self.invoice_header_overrides.get(source_key):
                self.invoice_header_overrides.pop(source_key, None)

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

        original_number = self.invoice_metadata.get(path, {}).get("original_number", "")
        mode = self.number_mode_var.get()
        affix = self.number_affix_var.get()
        if mode == "prefix":
            effective_number = f"{affix}{original_number}"
        elif mode == "suffix":
            effective_number = f"{original_number}{affix}"
        else:
            effective_number = original_number
        self._apply_effective_number_to_active_invoice(effective_number, refresh_selectors=False)
        return True

    def _snapshot_editor_changes(self):
        self.saved_line_overrides = copy.deepcopy(self.line_overrides)
        self.saved_header_overrides = copy.deepcopy(self.invoice_header_overrides)
        self.saved_number_overrides = copy.deepcopy(self.invoice_number_overrides)
        self.saved_manual_assignments = copy.deepcopy(self.manual_po_assignments)
        self.edit_draft_signature = self._current_edit_signature()
        self.edit_changes_saved = True

    def _has_unsaved_editor_changes(self):
        if self.edit_active_editor is not None or self.header_active_editor is not None:
            return True
        return (
            self.line_overrides != self.saved_line_overrides
            or self.invoice_header_overrides != self.saved_header_overrides
            or self.invoice_number_overrides != self.saved_number_overrides
            or self.manual_po_assignments != self.saved_manual_assignments
        )

    def save_edit_changes(self):
        path = self._active_edit_path()
        if not path:
            if not self.prepare_edit_draft(silent=False):
                return
            path = self._active_edit_path()
            if not path:
                messagebox.showinfo(tr("tab_edit"), tr("mb_no_edit_inv"))
                return
        if not self._harvest_visible_editor():
            return
        if not messagebox.askyesno(tr("mb_save_changes_title"), tr("mb_save_changes_ask")):
            return
        self._snapshot_editor_changes()
        saved_line_fields = sum(len(values) for values in self.saved_line_overrides.values())
        saved_header_fields = sum(len(values) for values in self.saved_header_overrides.values())
        saved_number_fields = len(self.saved_number_overrides)
        logging.info(
            "Editor guardado: %d campos de linea, %d campos de cabecera, %d numeros, %d POL manuales",
            saved_line_fields, saved_header_fields, saved_number_fields, len(self.saved_manual_assignments),
        )
        self._set_status("st_saving_edits", n_num=saved_number_fields, n_hdr=saved_header_fields, n_line=saved_line_fields)
        self.root.after_idle(self.run_conversion)

    def reset_edit_invoice(self):
        self._close_all_cell_editors(commit=False)
        path = self._active_edit_path()
        if not path:
            return
        if not messagebox.askyesno(tr("mb_reset_inv_title"), tr("mb_reset_inv_ask")):
            return
        source_key = os.path.normcase(os.path.abspath(path))
        self.line_overrides = {key: value for key, value in self.line_overrides.items() if key[0] != source_key}
        self.manual_po_assignments = {key: value for key, value in self.manual_po_assignments.items() if key[0] != source_key}
        self.invoice_header_overrides.pop(source_key, None)
        self.saved_line_overrides = {key: value for key, value in self.saved_line_overrides.items() if key[0] != source_key}
        self.saved_manual_assignments = {key: value for key, value in self.saved_manual_assignments.items() if key[0] != source_key}
        self.saved_header_overrides.pop(source_key, None)
        self.saved_number_overrides.pop(source_key, None)
        self.invoice_number_mode_by_source.pop(source_key, None)
        self.reset_active_invoice_number()
        self.edit_changes_saved = False
        self.number_mode_var.set("original")
        self._update_number_preview()
        self._refresh_edit_selector(select_path=path)

    # --------------------------------------------------------------------------
    # GESTIÓN DE PROVEEDORES E IMPORTACIÓN / EXPORTACIÓN EXCEL (SOLO 2 COLUMNAS)
    # --------------------------------------------------------------------------
    def _refresh_vendor_map_tree(self):
        if not hasattr(self, "vendor_map_tree"):
            return
        for item in self.vendor_map_tree.get_children():
            self.vendor_map_tree.delete(item)
        vendor_map = self.preferences.get("vendor_map", {})
        for nif in sorted(vendor_map):
            self.vendor_map_tree.insert("", "end", iid=nif, values=(nif, vendor_map[nif]))

    def _on_vendor_map_selected(self, _event=None):
        sel = self.vendor_map_tree.selection()
        if not sel:
            return
        nif = sel[0]
        vendor = self.vendor_map_tree.set(nif, "vendor")
        self.vendor_nif_var.set(nif)
        self.vendor_acc_var.set(vendor)

    def _sync_loaded_invoices_with_vendor_map(self):
        vendor_map = self.preferences.get("vendor_map", {})
        if not self.invoice_files:
            return
        for path in self.invoice_files:
            meta = self.invoice_metadata.get(path)
            if not meta:
                continue
            nif = normalize_tax_id(meta.get("seller_tax_id", ""))
            if nif and nif in vendor_map:
                meta["vendor"] = vendor_map[nif]
        self._refresh_batch_tree()
        self.edit_draft_signature = None

    def _add_or_update_vendor_map_entry(self):
        norm_nif = normalize_tax_id(self.vendor_nif_var.get())
        vendor = self.vendor_acc_var.get().strip()
        if not norm_nif or not vendor:
            messagebox.showwarning(tr("mb_vend_title"), tr("mb_vend_missing"))
            return
        self.preferences.setdefault("vendor_map", {})[norm_nif] = vendor
        try:
            self._save_preferences_file()
        except ValueError as e:
            messagebox.showerror(tr("mb_err_save_title"), str(e))
            return
        self._refresh_vendor_map_tree()
        self._sync_loaded_invoices_with_vendor_map()
        self.vendor_nif_var.set("")
        self.vendor_acc_var.set("")
        self._set_status("st_vendor_saved", nif=norm_nif, vendor=vendor)

    def _delete_selected_vendor_map_entry(self):
        sel = self.vendor_map_tree.selection()
        if not sel:
            messagebox.showinfo(tr("mb_vend_title"), tr("mb_vend_sel_del"))
            return
        nif = sel[0]
        self.preferences.setdefault("vendor_map", {}).pop(nif, None)
        try:
            self._save_preferences_file()
        except ValueError as e:
            messagebox.showerror(tr("mb_err_save_title"), str(e))
            return
        self._refresh_vendor_map_tree()
        self._set_status("st_vendor_deleted", nif=nif)

    def _clear_vendor_map(self):
        vendor_map = self.preferences.get("vendor_map", {})
        if not vendor_map:
            return
        if not messagebox.askyesno(tr("mb_vend_clr_title"), tr("mb_vend_clr_ask")):
            return
        self.preferences["vendor_map"] = {}
        try:
            self._save_preferences_file()
        except ValueError as e:
            messagebox.showerror(tr("mb_err_save_title"), str(e))
            return
        self._refresh_vendor_map_tree()
        self._set_status("st_vendor_cleared")

    def _resolve_vendor_import_mode_code(self):
        raw = self.vendor_import_mode_var.get()
        if raw in VENDOR_IMPORT_MODES:
            return raw
        for code, key in VENDOR_IMPORT_MODE_KEYS.items():
            for lang_dict in I18N.values():
                if raw == lang_dict.get(key):
                    return code
        return self.vendor_import_mode_code or "incremental"

    def _on_vendor_import_mode_selected(self, _event=None):
        idx = self.vendor_import_mode_combo.current()
        if 0 <= idx < len(VENDOR_IMPORT_MODES):
            self.vendor_import_mode_code = VENDOR_IMPORT_MODES[idx]
        else:
            self.vendor_import_mode_code = self._resolve_vendor_import_mode_code()

    def import_vendors_excel_dialog(self):
        mode = self._resolve_vendor_import_mode_code()
        self.vendor_import_mode_code = mode
        current_map = self.preferences.get("vendor_map", {})
        if mode == "full" and current_map:
            if not messagebox.askyesno(tr("mb_vend_title"), tr("mb_vend_full_ask", count=len(current_map))):
                return
        initial = self._initial_dir("custom_output_dir", "last_output_dir")
        path = filedialog.askopenfilename(
            title=tr("fd_imp_vendors"),
            initialdir=initial,
            filetypes=[(tr("fd_ft_excel"), "*.xlsx"), (tr("fd_ft_all"), "*.*")],
        )
        if not path:
            return
        self.import_vendors_from_excel_path(path, mode=mode, show_dialog=True)

    def import_vendors_from_excel_path(self, path, mode="incremental", show_dialog=True):
        try:
            current_map = self.preferences.get("vendor_map", {})
            updated_map, stats = import_vendor_map_from_excel(path, existing_map=current_map, mode=mode)
            self.preferences["vendor_map"] = updated_map
            self._save_preferences_file()
            self._remember_dir("last_output_dir", path)
            self._refresh_vendor_map_tree()
            self._sync_loaded_invoices_with_vendor_map()
            self._set_status("st_vendor_imported", name=os.path.basename(path), total=stats["total"])
            if show_dialog:
                mode_label = tr(VENDOR_IMPORT_MODE_KEYS.get(mode, "cust_vend_mode_inc"))
                messagebox.showinfo(
                    tr("mb_vend_imp_ok_title"),
                    tr(
                        "mb_vend_imp_ok_msg",
                        mode=mode_label,
                        added=stats["added"],
                        updated=stats["updated"],
                        skipped=stats["skipped"],
                        invalid=stats["invalid"],
                        total=stats["total"],
                    ),
                )
            return stats
        except Exception as e:
            if show_dialog:
                messagebox.showerror(tr("mb_vend_imp_err_title"), str(e))
                return None
            raise

    def export_vendors_excel_dialog(self):
        initial = self._initial_dir("custom_output_dir", "last_output_dir")
        path = filedialog.asksaveasfilename(
            title=tr("fd_exp_vendors"),
            initialdir=initial,
            initialfile="proveedores_alma.xlsx",
            defaultextension=".xlsx",
            filetypes=[(tr("fd_ft_excel"), "*.xlsx"), (tr("fd_ft_all"), "*.*")],
        )
        if not path:
            return
        self.export_vendors_to_excel_path(path, show_dialog=True)

    def export_vendors_to_excel_path(self, path, show_dialog=True):
        try:
            vendor_map = self.preferences.get("vendor_map", {})
            count = export_vendor_map_to_excel(vendor_map, path)
            self._remember_dir("last_output_dir", path)
            self._set_status("st_vendor_exported", count=count, name=os.path.basename(path))
            if show_dialog:
                messagebox.showinfo(tr("mb_vend_exp_ok_title"), tr("mb_vend_exp_ok_msg", count=count, path=path))
            return count
        except Exception as e:
            if show_dialog:
                messagebox.showerror(tr("mb_vend_exp_err_title"), str(e))
                return None
            raise

    # --------------------------------------------------------------------------
    # LOGO, PREFERENCIAS, MENÚ GLOBAL Y ARCHIVOS
    # --------------------------------------------------------------------------
    def select_logo(self):
        current = self.logo_path_var.get().strip()
        initial = os.path.dirname(current) if current and os.path.isfile(current) else self._application_dir()
        path = filedialog.askopenfilename(
            title=tr("fd_sel_logo"),
            initialdir=initial,
            filetypes=[
                (tr("fd_ft_img_compat"), "*.png *.jpg *.jpeg"),
                (tr("fd_ft_png"), "*.png"),
                (tr("fd_ft_jpg"), "*.jpg *.jpeg"),
                (tr("fd_ft_all"), "*.*"),
            ],
        )
        if not path:
            return
        self.logo_path_var.set(path)
        if self._apply_logo(path, show_warning=True):
            self._set_logo_status("cust_logo_loaded", name=os.path.basename(path))

    def remove_logo(self):
        self.logo_path_var.set("")
        self.logo_photo = None
        self.logo_label.config(image="")
        self._set_logo_status("cust_logo_none")

    def _load_logo_from_preferences(self, show_warning=False):
        path = self.logo_path_var.get().strip()
        if not path:
            self.logo_photo = None
            if hasattr(self, "logo_label"):
                self.logo_label.config(image="")
            return False
        ok = self._apply_logo(path, show_warning=show_warning)
        if ok:
            self._set_logo_status("cust_logo_loaded", name=os.path.basename(path))
        return ok

    def _apply_logo(self, path, show_warning=False):
        if not hasattr(self, "logo_label"):
            return False
        if not os.path.isfile(path):
            self.logo_photo = None
            self.logo_label.config(image="")
            self._set_logo_status("cust_logo_unavail")
            if show_warning:
                messagebox.showwarning(tr("mb_logo_unavail_title"), tr("mb_logo_unavail_msg", path=path))
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
            self._set_logo_status("cust_logo_invalid")
            if show_warning:
                messagebox.showwarning(tr("mb_logo_invalid_title"), tr("mb_logo_invalid_msg", path=path, err=e))
            return False

    def _choose_preference_dir(self, variable):
        initial = self._valid_dir(variable.get()) or self._application_dir()
        folder = filedialog.askdirectory(initialdir=initial, title=tr("fd_sel_folder"))
        if folder:
            variable.set(folder)

    def save_preferences(self):
        institution = self.institution_var.get().strip() or self.DEFAULT_INSTITUTION
        self.institution_var.set(institution)
        lang_code = LANGUAGE_LABEL_TO_CODE.get(self.language_label_var.get(), get_current_language())
        set_current_language(lang_code)
        for value, label_key in [
            (self.custom_invoice_dir.get().strip(), "cust_dir_inv"),
            (self.custom_alma_dir.get().strip(), "cust_dir_alma"),
            (self.custom_template_dir.get().strip(), "cust_dir_tpl"),
            (self.custom_output_dir.get().strip(), "cust_dir_out"),
        ]:
            if value and not os.path.isdir(value):
                messagebox.showerror(tr("mb_pref_invalid_title"), tr("mb_pref_invalid_dir", label=tr(label_key), val=value))
                return
        logo_path = self.logo_path_var.get().strip()
        if logo_path and not self._apply_logo(logo_path, show_warning=True):
            return
        self.preferences.update({
            "language": get_current_language(),
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
            self.apply_language()
            messagebox.showinfo(tr("cust_pref_frame"), tr("mb_pref_saved_msg"))
        except ValueError as e:
            messagebox.showerror(tr("mb_err_save_title"), str(e))

    def reset_preferences(self):
        if not messagebox.askyesno(tr("mb_pref_reset_title"), tr("mb_pref_reset_ask")):
            return
        defaults = self._default_preferences()
        self.preferences = defaults
        set_current_language(defaults["language"])
        self.language_label_var.set(LANGUAGE_CODE_TO_LABEL[defaults["language"]])
        self.institution_var.set(defaults["institution"])
        self.custom_invoice_dir.set("")
        self.custom_alma_dir.set("")
        self.custom_template_dir.set("")
        self.custom_output_dir.set("")
        self.logo_path_var.set("")
        self.logo_photo = None
        self.logo_label.config(image="")
        self._set_logo_status("cust_logo_none")
        self.template_file.set("")
        try:
            self._save_preferences_file()
        except ValueError as e:
            messagebox.showerror(tr("mb_err_reset_title"), str(e))
            return
        self._update_title()
        self._refresh_vendor_map_tree()
        self.apply_language()
        messagebox.showinfo(tr("cust_pref_frame"), tr("mb_pref_reset_ok"))

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
            menu.add_command(label=tr("gctx_cut"), command=lambda: widget.event_generate("<<Cut>>"))
            menu.add_command(label=tr("gctx_copy"), command=lambda: widget.event_generate("<<Copy>>"))
            menu.add_command(label=tr("gctx_paste"), command=lambda: widget.event_generate("<<Paste>>"))
            menu.add_separator()
            menu.add_command(label=tr("gctx_sel_all"), command=lambda: self._select_all_widget(widget))
        else:
            menu.add_command(label=tr("gctx_copy_txt"), command=lambda: self._copy_widget_text(widget))
        if hasattr(self, "visualize_tab"):
            menu.add_separator()
            menu.add_command(label=tr("gctx_go_vis"), command=lambda: self.notebook.select(self.visualize_tab))
            menu.add_command(label=tr("gctx_go_edit"), command=lambda: self.notebook.select(self.edit_tab))
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
            self.root.clipboard_clear()
            self.root.clipboard_append(str(text))

    def _open_url(self, url):
        try:
            webbrowser.open_new(url)
        except Exception as e:
            messagebox.showerror(tr("mb_err_url_title"), tr("mb_err_url_msg", url=url, err=e))

    def _open_folder(self, folder_path):
        folder_path = os.path.realpath(os.path.abspath(folder_path))
        if not os.path.isdir(folder_path):
            messagebox.showwarning(tr("mb_warn_title"), tr("mb_no_out_dir_found"))
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
            messagebox.showerror(tr("mb_err_folder_title"), tr("mb_err_folder_msg", path=folder_path, err=e))

    def open_output_folder(self):
        output_path = self.output_file.get().strip()
        folder = os.path.dirname(output_path) if output_path else ""
        if not folder:
            messagebox.showwarning(tr("mb_warn_title"), tr("mb_no_out_dir_def"))
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
            self.invoice_label.config(text=tr("conv_inv_batch_lbl"))
            self.invoice_button.config(text=tr("conv_inv_batch_btn"))
            self.convert_tab.rowconfigure(2, weight=1)
            self.convert_tab.rowconfigure(3, weight=0)
            self.batch_frame.grid()
            self._set_status("st_mode_batch")
        else:
            self.invoice_label.config(text=tr("conv_inv_single_lbl"))
            self.invoice_button.config(text=tr("conv_inv_single_btn"))
            self.convert_tab.rowconfigure(2, weight=0)
            self.convert_tab.rowconfigure(3, weight=0)
            self.batch_frame.grid_remove()
            if initial:
                self._set_status("st_pending")
            else:
                self._set_status("st_mode_single")

    def select_invoice_or_batch(self):
        if self.conversion_mode.get() == "batch":
            self.add_batch_files(replace=True)
        else:
            self.select_invoice_single()

    def _invoice_filetypes(self):
        return [
            (tr("fd_ft_compat_inv"), "*.xsig *.xml *.txt"),
            (tr("fd_ft_xsig"), "*.xsig"),
            (tr("fd_ft_xml"), "*.xml"),
            (tr("fd_ft_txt"), "*.txt"),
            (tr("fd_ft_all"), "*.*"),
        ]

    def select_invoice_single(self):
        initial = self._initial_dir("custom_invoice_dir", "last_invoice_dir")
        path = filedialog.askopenfilename(title=tr("fd_sel_inv_single"), initialdir=initial, filetypes=self._invoice_filetypes())
        if not path:
            return
        try:
            metadata = self._read_invoice_metadata(path)
        except Exception as e:
            messagebox.showerror(tr("mb_inv_invalid_title"), str(e))
            return
        self._invalidate_edit_draft(clear_overrides=True)
        self.invoice_files = [path]
        self.invoice_metadata = {path: metadata}
        self.invoice_file.set(path)
        self._remember_dir("last_invoice_dir", path)
        self._set_default_output(path, batch=False)
        self._set_status("st_inv_selected", number=metadata["number"], vendor=metadata["vendor"])
        self._refresh_viewer_selector(select_path=path)
        self._refresh_edit_selector(select_path=path)

    def add_batch_files(self, replace=False):
        initial = self._initial_dir("custom_invoice_dir", "last_invoice_dir")
        paths = filedialog.askopenfilenames(title=tr("fd_sel_inv_batch"), initialdir=initial, filetypes=self._invoice_filetypes())
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
                warnings.append(tr("mb_batch_rep_path", name=os.path.basename(path)))
                continue
            try:
                metadata = self._read_invoice_metadata(path)
            except Exception as e:
                warnings.append(tr("mb_batch_read_err", name=os.path.basename(path), err=e))
                continue
            number = metadata.get("number")
            if number and number in existing_numbers:
                warnings.append(tr("mb_batch_dup_num", num=number, name=os.path.basename(path)))
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
            self.invoice_file.set(tr("conv_batch_count", count=len(self.invoice_files)))
            self._remember_dir("last_invoice_dir", self.invoice_files[0])
            self._set_default_output(self.invoice_files[0], batch=True)
        self._set_status("st_batch_count", count=len(self.invoice_files))
        if warnings:
            messagebox.showwarning(tr("mb_batch_warn_title"), "\n".join(warnings))
        elif added == 0:
            messagebox.showinfo(tr("mb_batch_title"), tr("mb_batch_none_added"))

    def select_invoice_batch(self):
        self.add_batch_files(replace=True)

    def _read_invoice_metadata(self, path):
        vendor_map = self.preferences.get("vendor_map", {})
        parsed = FacturaeParser(path, custom_vendor_map=vendor_map).parse()
        if len(parsed) != 1:
            raise ValueError(tr("val_single_inv_only", count=len(parsed)))
        inv = parsed[0]
        number = inv.invoice_number or tr("vw_no_number")
        return {
            "number": number,
            "original_number": number,
            "vendor": inv.vendor_account or inv.seller_name or tr("vw_unknown_vendor"),
            "seller_tax_id": inv.seller_tax_id or "",
        }

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
            messagebox.showinfo(tr("mb_batch_title"), tr("mb_batch_sel_one"))
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
        self.invoice_file.set(tr("conv_batch_count", count=len(self.invoice_files)) if self.invoice_files else "")
        self._set_status("st_batch_count", count=len(self.invoice_files))

    def clear_batch_files(self):
        if self.invoice_files and not messagebox.askyesno(tr("mb_batch_clr_title"), tr("mb_batch_clr_ask")):
            return
        self.invoice_files = []
        self.invoice_metadata = {}
        self.invoice_file.set("")
        self.output_file.set("")
        self._refresh_batch_tree()
        self._invalidate_edit_draft(clear_overrides=True)
        self._refresh_viewer_selector()
        self._refresh_edit_selector()
        self._set_status("st_batch_empty")

    def sort_batch_files(self):
        self.invoice_files.sort(key=lambda p: (self.invoice_metadata.get(p, {}).get("number", "").casefold(), os.path.basename(p).casefold()))
        self._refresh_batch_tree()
        self._set_status("st_batch_sorted")

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
        path = filedialog.askopenfilename(title=tr("fd_sel_alma"), initialdir=initial, filetypes=[(tr("fd_ft_excel"), "*.xlsx"), (tr("fd_ft_all"), "*.*")])
        if path:
            self.alma_pol_file.set(path)
            self._remember_dir("last_alma_dir", path)
            self._invalidate_edit_draft(clear_overrides=False)
            self._set_status("st_alma_selected")

    def select_template(self):
        fallback = os.path.dirname(self.alma_pol_file.get()) if self.alma_pol_file.get() else None
        initial = self._initial_dir("custom_template_dir", "last_template_dir", fallback)
        path = filedialog.askopenfilename(title=tr("fd_sel_tpl"), initialdir=initial, filetypes=[(tr("fd_ft_excel"), "*.xlsx"), (tr("fd_ft_all"), "*.*")])
        if path:
            self.template_file.set(path)
            self.preferences["last_template_path"] = path
            self._remember_dir("last_template_dir", path)
            self._invalidate_edit_draft(clear_overrides=False)
            self._set_status("st_tpl_selected")

    def select_output(self):
        fallback = os.path.dirname(self.invoice_files[0]) if self.invoice_files else None
        initial = self._initial_dir("custom_output_dir", "last_output_dir", fallback)
        initialfile = os.path.basename(self.output_file.get()) if self.output_file.get() else "FacturaeToAlma.xlsx"
        path = filedialog.asksaveasfilename(title=tr("fd_sel_out"), initialdir=initial, initialfile=initialfile, defaultextension=".xlsx", filetypes=[(tr("fd_ft_excel"), "*.xlsx"), (tr("fd_ft_all"), "*.*")])
        if path:
            self.output_file.set(path)
            self._remember_dir("last_output_dir", path)
            self._set_status("st_out_selected")

    def clear_fields(self):
        self.invoice_files = []
        self.invoice_metadata = {}
        self.invoice_file.set("")
        self.output_file.set("")
        self._refresh_batch_tree()
        self._refresh_viewer_selector()
        self._refresh_edit_selector()
        self.invoice_number_overrides.clear()
        self.invoice_number_mode_by_source.clear()
        self.manual_po_assignments.clear()
        self.line_overrides.clear()
        self.edit_original_values.clear()
        self.edit_header_original_values.clear()
        self.invoice_header_overrides.clear()
        self.saved_line_overrides.clear()
        self.saved_header_overrides.clear()
        self.saved_number_overrides.clear()
        self.saved_manual_assignments.clear()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.edit_draft_signature = None
        self.edit_changes_saved = False
        self.number_mode_var.set("original")
        if self.alma_pol_file.get().strip() and messagebox.askyesno(tr("mb_clr_alma_title"), tr("mb_clr_alma_ask")):
            self.alma_pol_file.set("")
        self._set_status("st_cleared")

    def _validate_inputs(self):
        if not self.invoice_files:
            raise ValueError(tr("val_need_inv"))
        allowed = {".xsig", ".xml", ".txt"}
        for path in self.invoice_files:
            if os.path.splitext(path)[1].lower() not in allowed:
                raise ValueError(tr("val_bad_ext", path=path))
            if not os.path.isfile(path):
                raise FileNotFoundError(tr("val_no_inv_file", path=path))
        alma = self.alma_pol_file.get().strip()
        template = self.template_file.get().strip()
        output = self.output_file.get().strip()
        if not alma:
            raise ValueError(tr("val_need_alma"))
        if not template:
            raise ValueError(tr("val_need_tpl"))
        if not output:
            raise ValueError(tr("val_need_out"))
        if not os.path.isfile(alma):
            raise FileNotFoundError(tr("val_no_alma_file", path=alma))
        if not os.path.isfile(template):
            raise FileNotFoundError(tr("val_no_tpl_file", path=template))
        output_norm = os.path.normcase(os.path.abspath(output))
        protected_paths = {
            os.path.normcase(os.path.abspath(alma)),
            os.path.normcase(os.path.abspath(template)),
            *(os.path.normcase(os.path.abspath(path)) for path in self.invoice_files),
        }
        if output_norm in protected_paths:
            raise ValueError(tr("val_out_collision"))
        duplicate_numbers = {}
        for path in self.invoice_files:
            number = self.invoice_metadata.get(path, {}).get("number")
            if number:
                duplicate_numbers.setdefault(number, []).append(path)
        repeated = {n: paths for n, paths in duplicate_numbers.items() if len(paths) > 1}
        if repeated:
            raise ValueError(tr("val_dup_numbers", nums=", ".join(sorted(repeated))))
        return list(self.invoice_files), alma, output, template

    def run_conversion(self):
        if not self._flush_editors():
            return
        if self.edit_draft_signature is not None:
            self._harvest_visible_editor()
        if self._has_unsaved_editor_changes() and not self.edit_changes_saved:
            if messagebox.askyesno(tr("mb_unsaved_ed_title"), tr("mb_unsaved_ed_ask")):
                if self.edit_draft_signature != self._current_edit_signature():
                    if not self.prepare_edit_draft(silent=True):
                        detail = self._last_draft_error or tr("mb_draft_unavail_msg")
                        messagebox.showerror(tr("mb_err_prep_draft"), detail)
                        self._set_status("st_check_inputs")
                        return
                    self._harvest_visible_editor()
                self._snapshot_editor_changes()
            else:
                self.notebook.select(self.edit_tab)
                return
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(tr("mb_conv_running_title"), tr("mb_conv_running_msg"))
            return
        try:
            inputs = self._validate_inputs()
        except (ValueError, FileNotFoundError) as e:
            messagebox.showerror(tr("mb_invalid_inputs_title"), str(e))
            self._set_status("st_check_inputs")
            return
        if self.edit_draft_signature != self._current_edit_signature():
            if self.edit_changes_saved:
                messagebox.showerror(tr("mb_draft_outdated_title"), tr("mb_draft_outdated_msg"))
                self.notebook.select(self.edit_tab)
                return
            if not self.prepare_edit_draft(silent=True):
                detail = self._last_draft_error or tr("mb_draft_unavail_msg")
                messagebox.showerror(tr("mb_err_prep_draft"), detail)
                self._set_status("st_check_inputs")
                return
        self.convert_button.config(state="disabled")
        self._set_status("st_converting")
        self.worker_thread = threading.Thread(target=self._conversion_worker, args=inputs + (
            copy.deepcopy(self.saved_number_overrides), copy.deepcopy(self.saved_manual_assignments),
            copy.deepcopy(self.saved_line_overrides), copy.deepcopy(self.saved_header_overrides),
            copy.deepcopy(self.preferences.get("vendor_map", {})),
        ), daemon=True)
        self.worker_thread.start()
        self.root.after(100, self._poll_worker)

    def _conversion_worker(self, invoice_paths, alma_path, output_path, template_path, invoice_number_overrides, manual_assignments, line_overrides, invoice_header_overrides, custom_vendor_map):
        try:
            result = convert_facturae_to_alma_excel(
                invoice_paths, alma_path, output_path, template_path,
                invoice_number_overrides, manual_assignments, line_overrides, invoice_header_overrides,
                custom_vendor_map=custom_vendor_map,
            )
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
        extra_lines = ""
        if summary.get("match_pol", 0) > 0:
            extra_lines += tr("sum_pol_xml", n=summary["match_pol"])
        if summary.get("match_manual", 0) > 0:
            extra_lines += tr("sum_pol_man", n=summary["match_manual"])
        return tr(
            "sum_body",
            total_lines=summary["total_lines"],
            match_isbn=summary["match_isbn"],
            match_issn=summary["match_issn"],
            match_title=summary["match_title"],
            extra=extra_lines,
            duplicate_isbn=summary["duplicate_isbn"],
            duplicate_issn=summary["duplicate_issn"],
            ambiguous_title=summary["ambiguous_title"],
            no_match=summary["no_match"],
            lines_without_po_line=summary["lines_without_po_line"],
        )

    def _handle_conversion_success(self, result, output_path):
        total_invoices, total_issues, validation_path, summary = result
        heading = tr("conv_ok_heading") if not validation_path else tr("conv_warn_heading")
        message = tr(
            "conv_result_body",
            heading=heading,
            total_inv=total_invoices,
            total_iss=total_issues,
            out=output_path,
        ) + self._build_summary_message(summary)
        if validation_path:
            message += tr("conv_val_report", path=validation_path)
        if summary.get("lines_without_po_line", 0) > 0:
            message += tr("conv_unmatched_warn")
        self._remember_dir("last_output_dir", output_path)
        self._set_status("st_conv_warn" if validation_path else "st_conv_ok")
        if validation_path:
            messagebox.showwarning(tr("mb_conv_warn_title"), message)
        else:
            messagebox.showinfo(tr("mb_conv_ok_title"), message)

    def _handle_conversion_error(self, error, trace):
        if isinstance(error, DefusedXmlException):
            title = tr("err_xml_sec_title")
        elif isinstance(error, PermissionError):
            title = tr("err_perm_title")
        elif isinstance(error, FileNotFoundError):
            title = tr("err_fnf_title")
        elif isinstance(error, ValueError):
            title = tr("err_val_title")
        else:
            title = tr("err_conv_title")
        logging.error(trace)
        self._set_status("st_conv_err")
        messagebox.showerror(title, tr("err_log_hint", err=error, log=LOG_FILE))

    def _on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno(tr("mb_conv_running_title"), tr("mb_conv_close_ask")):
                return
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = FacturaeApp(root)
    root.mainloop()
