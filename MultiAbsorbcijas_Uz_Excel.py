#Lai palaistu kodu no sākuma terminālī jāievada šis ja vel nav un jauzinstalē python:
# pip install openpyxl tkinterdnd2
#
import re
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook
from tkinterdnd2 import TkinterDnD, DND_FILES


class TxtUzExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TXT faili no absorcijas mērijumiem uz Excel")
        self.root.geometry("700x500")

        self.faili = []

        tk.Label(
            root,
            text="Veidoja: RalfsLu",
            font=("Arial", 10, "bold")
        ).pack(pady=(15, 8))

        self.drop_area = tk.Label(
            root,
            text="IEVELC .TXT FAILUS ŠEIT",
            bg="#d9edf7",
            relief="groove",
            borderwidth=3,
            height=5,
            font=("Arial", 14, "bold")
        )
        self.drop_area.pack(fill="x", padx=20, pady=5)

        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind(
            "<<Drop>>",
            self.pievienot_failus_ar_dragdrop
        )

        pogas = tk.Frame(root)
        pogas.pack(pady=10)

        tk.Button(
            pogas,
            text="Pievienot failus",
            command=self.izveleties_failus,
            width=20,
            height=2,
            font=("Arial", 11)
        ).grid(row=0, column=0, padx=5)

        tk.Button(
            pogas,
            text="Notīrīt sarakstu",
            command=self.notirit_sarakstu,
            width=20,
            height=2,
            font=("Arial", 11)
        ).grid(row=0, column=1, padx=5)

        self.saraksts = tk.Listbox(
            root,
            height=6,
            font=("Arial", 10)
        )
        self.saraksts.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=5
        )

        # Lielāka Excel izveidošanas poga
        tk.Button(
            root,
            text="IZVEIDOT EXCEL FAILU",
            command=self.izveidot_excel,
            bg="#4CAF50",
            fg="white",
            activebackground="#388E3C",
            activeforeground="white",
            font=("Arial", 17, "bold"),
            height=2,
            padx=30,
            pady=15
        ).pack(
            fill="x",
            padx=20,
            pady=(12, 20)
        )

    def pievienot_failus_ar_dragdrop(self, event):
        # Pareizi apstrādā failu ceļus, arī ja tajos ir atstarpes
        failu_celi = self.root.tk.splitlist(event.data)
        self.pievienot_failus(failu_celi)

    def izveleties_failus(self):
        faili = filedialog.askopenfilenames(
            title="Izvēlies TXT failus",
            filetypes=[
                ("TXT faili", "*.txt"),
                ("Visi faili", "*.*")
            ]
        )
        self.pievienot_failus(faili)

    def pievienot_failus(self, faili):
        for fails in faili:
            if (
                fails.lower().endswith(".txt")
                and fails not in self.faili
            ):
                self.faili.append(fails)
                self.saraksts.insert(tk.END, fails)

    def notirit_sarakstu(self):
        self.faili = []
        self.saraksts.delete(0, tk.END)

    def parveidot_vertibu(self, vertiba):

        vertiba = vertiba.strip()

        if not vertiba:
            return ""

        skaitla_paraugs = r"[+-]?\d+(?:[.,]\d+)?"

        if re.fullmatch(skaitla_paraugs, vertiba):

            skaitlis = vertiba.replace(",", ".")

            if "." in skaitlis:
                return float(skaitlis)
            else:
                return int(skaitlis)

        return vertiba.replace(".", ",")

    def nolasit_txt_failu(self, faila_cels):

        rindas = None

        for kodējums in (
            "utf-8-sig",
            "utf-8",
            "cp1257",
            "latin-1"
        ):
            try:
                with open(
                    faila_cels,
                    "r",
                    encoding=kodējums
                ) as fails:
                    rindas = fails.readlines()
                break
            except UnicodeDecodeError:
                continue

        if rindas is None:
            raise ValueError(
                f"Neizdevās nolasīt failu:\n{faila_cels}"
            )

        dati = []

        for rinda in rindas:
            rinda = rinda.strip()

            if not rinda:
                dati.append([])
                continue

            vertibas = re.split(r"\s+", rinda)

            apstradatas_vertibas = [
                self.parveidot_vertibu(vertiba)
                for vertiba in vertibas
            ]

            dati.append(apstradatas_vertibas)

        return dati

    def izveidot_excel(self):
        if not self.faili:
            messagebox.showwarning(
                "Nav failu",
                "Vispirms pievieno vismaz vienu .txt failu."
            )
            return

        saglabat_ka = filedialog.asksaveasfilename(
            title="Saglabāt Excel failu",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel fails", "*.xlsx")
            ],
            initialfile="absorcijas_mērijumi.xlsx"
        )

        if not saglabat_ka:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Dati"

            sakuma_kolonna = 1

            for faila_cels in self.faili:
                dati = self.nolasit_txt_failu(faila_cels)

                maksimalais_kolonnu_skaits = 1

                for rindas_nr, rinda in enumerate(
                    dati,
                    start=1
                ):
                    for kolonnas_nr, vertiba in enumerate(
                        rinda,
                        start=0
                    ):
                        cell = ws.cell(
                            row=rindas_nr,
                            column=sakuma_kolonna + kolonnas_nr,
                            value=vertiba
                        )

                        if isinstance(vertiba, (int, float)):
                            cell.number_format = "0.############"

                    maksimalais_kolonnu_skaits = max(
                        maksimalais_kolonnu_skaits,
                        len(rinda)
                    )

                sakuma_kolonna += (
                    maksimalais_kolonnu_skaits + 3
                )

            wb.save(saglabat_ka)

            messagebox.showinfo(
                "Gatavs",
                "Excel fails veiksmīgi izveidots:\n\n"
                + saglabat_ka
            )

        except Exception as kļūda:
            messagebox.showerror(
                "Kļūda",
                f"Radās kļūda:\n{str(kļūda)}"
            )


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    programma = TxtUzExcelApp(root)
    root.mainloop()