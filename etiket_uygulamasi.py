# -*- coding: utf-8 -*-
"""Koli Etiketi Olusturucu
Excel'deki her satirdan JPEG'deki duzende etiket iceren bir PDF uretir.
"""
import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
from reportlab.lib.pagesizes import landscape, A5
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import eanbc
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF

EXPECTED_HEADERS = ["NS PO", "Purchase Order", "Style Code", "Style Name",
                    "Colour", "Size", "Carton Qty", "EAN/UPC Barcode", "Carton Number"]


def read_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel dosyasi bos.")
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    idx = {}
    for name in EXPECTED_HEADERS:
        for i, h in enumerate(header):
            if h.lower() == name.lower():
                idx[name] = i
                break
        else:
            raise ValueError(f"Excel'de '{name}' sutunu bulunamadi.\nBulunan basliklar: {header}")
    data = []
    for r in rows[1:]:
        if r is None or all(c is None or str(c).strip() == "" for c in r):
            continue
        data.append({name: r[idx[name]] for name in EXPECTED_HEADERS})
    if not data:
        raise ValueError("Excel'de veri satiri bulunamadi.")
    return data


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def draw_label(c, row, W, H):
    m = 8 * mm                      # dis bosluk
    x0, y1 = m, H - m               # sol ust
    x1, y0 = W - m, m               # sag alt
    tw = x1 - x0
    th = y1 - y0

    # kolon sinirlari (jpeg orani): sol basliklar %27, orta %42, sag %31
    cA = x0 + tw * 0.27
    cB = x0 + tw * 0.69

    # satir yukseklikleri: ust NS PO ve alt Carton Number biraz kucuk
    heights = [0.115, 0.11, 0.115, 0.115, 0.115, 0.115, 0.225, 0.09]
    ys = [y1]
    for hh in heights:
        ys.append(ys[-1] - th * hh)
    ys[-1] = y0

    def cell_text(text, xa, xb, yt, yb, size=13, bold=True):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawCentredString((xa + xb) / 2, (yt + yb) / 2 - size * 0.35, text)

    c.setLineWidth(1.2)
    c.rect(x0, y0, tw, th)

    # yatay cizgiler
    for y in ys[1:-1]:
        c.line(x0, y, x1, y)
    # NS PO satiri tam genislik oldugu icin ust satirda dikey cizgi yok
    c.line(cA, y0, cA, ys[1])
    c.line(cB, y0, cB, ys[1])
    # Carton Qty hucresi (sag kolon) 2. satirdan alta kadar birlesik:
    # sag kolonda ara yatay cizgileri sil -> beyaz kutu cizerek yapmak yerine
    # cizgileri bastan sadece sol+orta bolgeye cizmek gerekirdi; kolayi:
    c.setFillColorRGB(1, 1, 1)
    c.rect(cB + 0.7, y0 + 0.7, x1 - cB - 1.4, ys[2] - y0 - 1.4, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)

    # icerik
    cell_text(f"NS PO : {fmt(row['NS PO'])}", x0, x1, ys[0], ys[1], 15)

    labels = ["Purchase Order", "Style Code", "Style Name", "Colour", "Size"]
    for i, name in enumerate(labels):
        yt, yb = ys[i + 1], ys[i + 2]
        cell_text(name, x0, cA, yt, yb, 13)
        cell_text(fmt(row[name]) or "N/A", cA, cB, yt, yb, 13)

    # Carton Qty
    cell_text("Carton Qty", cB, x1, ys[1], ys[2], 13)
    c.setFont("Helvetica-Bold", 52)
    c.drawCentredString((cB + x1) / 2, (ys[2] + y0) / 2 - 18, fmt(row["Carton Qty"]))

    # Barkod satiri
    yt, yb = ys[6], ys[7]
    cell_text("EAN/UPC Barcode", x0, cA, yt, yb, 13)
    code = fmt(row["EAN/UPC Barcode"]).split(".")[0].zfill(12)
    try:
        bc = eanbc.Ean13BarcodeWidget(code)
        bc.barHeight = (yt - yb) * 0.55
        bc.barWidth = 0.33 * mm
        bc.fontSize = 8
        bnds = bc.getBounds()
        bw, bh = bnds[2] - bnds[0], bnds[3] - bnds[1]
        d = Drawing(bw, bh)
        d.add(bc)
        renderPDF.draw(d, c, (cA + cB) / 2 - bw / 2, (yt + yb) / 2 - bh / 2)
    except Exception:
        cell_text(code, cA, cB, yt, yb, 12)

    # Carton Number
    yt, yb = ys[7], ys[8]
    cell_text("Carton Number", x0, cA, yt, yb, 12)
    cell_text(fmt(row["Carton Number"]), cA, cB, yt, yb, 12)


def make_pdf(xlsx_path, pdf_path):
    rows = read_rows(xlsx_path)
    W, H = landscape(A5)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))
    for row in rows:
        draw_label(c, row, W, H)
        c.showPage()
    c.save()
    return len(rows)


def main():
    root = tk.Tk()
    root.title("Koli Etiketi Olusturucu")
    root.geometry("460x220")
    root.resizable(False, False)

    tk.Label(root, text="Koli Etiketi Olusturucu", font=("Arial", 16, "bold")).pack(pady=(18, 4))
    tk.Label(root, text="Excel dosyasini secin; her satir icin bir etiket\niceren PDF ayni klasore kaydedilir.",
             font=("Arial", 11)).pack(pady=(0, 12))
    status = tk.Label(root, text="", font=("Arial", 10), fg="green")

    def run():
        path = filedialog.askopenfilename(title="Excel dosyasi secin",
                                          filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if not path:
            return
        pdf = os.path.splitext(path)[0] + "_etiketler.pdf"
        try:
            n = make_pdf(path, pdf)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Hata", str(e))
            return
        status.config(text=f"{n} etiket olusturuldu:\n{pdf}")
        try:
            if sys.platform == "win32":
                os.startfile(pdf)
            elif sys.platform == "darwin":
                os.system(f'open "{pdf}"')
        except Exception:
            pass

    tk.Button(root, text="Excel Sec ve Etiket Olustur", font=("Arial", 12, "bold"),
              command=run, width=28, height=2).pack()
    status.pack(pady=8)
    root.mainloop()


if __name__ == "__main__":
    main()
