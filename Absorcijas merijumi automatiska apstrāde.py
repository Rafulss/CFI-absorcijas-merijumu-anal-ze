# -*- coding: utf-8 -*-
"""Izvēlies Excel failu — skripts uz tās pašas lapas izveidos 3 Scatter grafikus.
Nepieciešams Windows, Excel un pywin32: py -m pip install pywin32
"""
from pathlib import Path
from datetime import datetime
import builtins as _builtins
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
try:
    import win32com.client as win32
except ImportError:
    win32 = None

XL_UP = -4162
XL_SCATTER = -4169
XL_SCATTER_LINES = 74
XL_XLSX = 51
XL_LEGEND_RIGHT = -4152
FIRST_ROW = 15
COLORS = [(31,119,180), (255,127,14), (44,160,44), (214,39,40),
          (148,103,189), (140,86,75), (227,119,194), (127,127,127),
          (188,189,34), (23,190,207), (57,73,171)]


def rgb(color):
    return color[0] + 256 * color[1] + 65536 * color[2]


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def col_letter(col):
    """Pārvērš Excel kolonnas numuru par burtu, neizmantojot COM Range.Address."""
    result = ""
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_ref(row, col):
    return f"{col_letter(col)}{row}"


def read_col(sheet, start, end, col):
    values = sheet.Range(sheet.Cells(start, col), sheet.Cells(end, col)).Value
    if isinstance(values, tuple):
        return [v[0] if isinstance(v, tuple) else v for v in values]
    return [values]


def to_seconds(text):
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            t = datetime.strptime(_builtins.str(text).strip(), fmt)
            return t.hour * 3600 + t.minute * 60 + t.second
        except ValueError:
            pass
    return None


def ask_wavelength_ranges(root):
    """Parāda dialogu divu viļņa garuma intervālu ievadīšanai (nm)."""
    dialog = tk.Toplevel(root)
    dialog.title("Izvēlies viļņa garuma intervālus")
    dialog.resizable(False, False)
    # Neiestatām transient(root): root ir apzināti paslēpts, un Windows vidē
    # šāds dialogs dažkārt var atvērties aiz citām lietotnēm.

    tk.Label(
        dialog,
        text=("Ievadi divus viļņa garuma intervālus (nm).\n"
              "Programma aprēķinās vidējo intensitāti katrā intervālā\n"
              "un to starpību salīdzinājumam laikā."),
        justify="left",
    ).grid(row=0, column=0, columnspan=3, padx=12, pady=(12, 8), sticky="w")

    tk.Label(dialog, text="Intervāls").grid(row=1, column=0, padx=8, pady=4, sticky="w")
    tk.Label(dialog, text="No (nm)").grid(row=1, column=1, padx=8, pady=4)
    tk.Label(dialog, text="Līdz (nm)").grid(row=1, column=2, padx=8, pady=4)

    defaults = (("1. intervāls", "400", "600"), ("2. intervāls", "700", "800"))
    entries = []
    for row, (label, start, end) in enumerate(defaults, start=2):
        tk.Label(dialog, text=label).grid(row=row, column=0, padx=8, pady=4, sticky="w")
        start_entry = tk.Entry(dialog, width=12)
        end_entry = tk.Entry(dialog, width=12)
        start_entry.insert(0, start)
        end_entry.insert(0, end)
        start_entry.grid(row=row, column=1, padx=8, pady=4)
        end_entry.grid(row=row, column=2, padx=8, pady=4)
        entries.append((start_entry, end_entry))

    result = {"ranges": None}

    def submit():
        try:
            ranges = []
            for start_entry, end_entry in entries:
                start = float(start_entry.get().strip().replace(",", "."))
                end = float(end_entry.get().strip().replace(",", "."))
                if start < 0 or end < 0:
                    raise ValueError("Viļņa garumam jābūt pozitīvam skaitlim.")
                if start > end:
                    raise ValueError("Katrā intervālā vērtībai 'No' jābūt mazākai vai vienādai ar 'Līdz'.")
                ranges.append((start, end))
        except ValueError as exc:
            messagebox.showerror("Nederīga ievade", str(exc), parent=dialog)
            return
        if ranges[0] == ranges[1]:
            messagebox.showerror("Nederīga ievade", "Abiem intervāliem jābūt atšķirīgiem.", parent=dialog)
            return
        result["ranges"] = tuple(ranges)
        dialog.destroy()

    def cancel():
        dialog.destroy()

    buttons = tk.Frame(dialog)
    buttons.grid(row=4, column=0, columnspan=3, pady=(8, 12))
    tk.Button(buttons, text="Turpināt", width=12, command=submit).pack(side="left", padx=5)
    tk.Button(buttons, text="Atcelt", width=12, command=cancel).pack(side="left", padx=5)
    dialog.protocol("WM_DELETE_WINDOW", cancel)

    # Novieto dialogu ekrāna centrā un izvirza priekšplānā.
    dialog.update_idletasks()
    width = dialog.winfo_reqwidth()
    height = dialog.winfo_reqheight()
    x = max(0, (dialog.winfo_screenwidth() - width) // 2)
    y = max(0, (dialog.winfo_screenheight() - height) // 2)
    dialog.geometry(f"+{x}+{y}")
    dialog.attributes("-topmost", True)
    dialog.lift()
    dialog.wait_visibility()
    dialog.grab_set()
    entries[0][0].focus_set()
    root.wait_window(dialog)
    return result["ranges"]


def new_chart(sheet, first_cell, last_cell, chart_type, title):
    a, b = sheet.Range(first_cell), sheet.Range(last_cell)
    obj = sheet.ChartObjects().Add(a.Left, a.Top, b.Left + b.Width - a.Left,
                                   b.Top + b.Height - a.Top)
    chart = obj.Chart
    chart.ChartType = chart_type
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.ChartTitle.Font.Size = 11
    chart.HasLegend = False
    return chart


def format_axes(chart, x_title, y_title, x_min=None, x_max=None, y_min=None, y_max=None):
    for axis_num, title in ((1, x_title), (2, y_title)):
        try:
            axis = chart.Axes(axis_num)
            axis.HasTitle = True
            axis.AxisTitle.Text = title
            axis.AxisTitle.Font.Size = 9
            axis.TickLabels.Font.Size = 8
            axis.MajorGridlines.Format.Line.ForeColor.RGB = rgb((225, 230, 236))
            if axis_num == 1:
                if x_min is not None: axis.MinimumScale = x_min
                if x_max is not None: axis.MaximumScale = x_max
            else:
                if y_min is not None: axis.MinimumScale = y_min
                if y_max is not None: axis.MaximumScale = y_max
        except Exception:
            pass


def available_output_path(source, numbered=False):
    """Atrod brīvu izvades faila nosaukumu, nepārrakstot esošus Excel failus."""
    source = Path(source).resolve()
    base_stem = source.stem
    # Ja par ievadi nejauši izvēlēts jau apstrādāts fails *_apstrādāts.xlsx,
    # neveido nosaukumu *_apstrādāts_apstrādāts.xlsx.
    if base_stem.lower().endswith("_apstrādāts"):
        base_stem = base_stem[:-len("_apstrādāts")]

    suffix = 2 if numbered else 1
    while True:
        if suffix == 1:
            candidate = source.with_name(base_stem + "_apstrādāts.xlsx")
        else:
            candidate = source.with_name(base_stem + f"_apstrādāts_{suffix}.xlsx")
        if candidate.resolve() != source and not candidate.exists():
            return candidate
        suffix = max(2, suffix + 1)


def run(path, wavelength_ranges):
    if win32 is None:
        raise RuntimeError("CMD logā izpildi: py -m pip install pywin32")
    source = Path(path).resolve()
    # Nekad nemēģina pārrakstīt jau esošu failu: tas arī novērš kļūdu,
    # ja iepriekšējais *_apstrādāts.xlsx ir atvērts Excel programmā vai bloķēts.
    output = available_output_path(source)
    excel = win32.DispatchEx("Excel.Application")
    workbook = None
    saved = False
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(_builtins.str(source), UpdateLinks=0)
        try:
            sheet = workbook.Worksheets("Dati")
        except Exception:
            sheet = workbook.Worksheets(1)

        # Atrodam katra spektra pirmo kolonnu (A, K, U, ...).
        used = sheet.UsedRange
        last_col = used.Column + used.Columns.Count - 1
        starts = []
        for col in range(1, last_col + 1):
            if (_builtins.str(sheet.Cells(1, col).Value).strip().lower() == "data"
                    and is_number(sheet.Cells(FIRST_ROW, col).Value)
                    and is_number(sheet.Cells(FIRST_ROW, col + 1).Value)):
                starts.append(col)
        if not starts:
            raise ValueError("Neatradu spektra datus. Izvēlies sākuma failu 14.10.2026. KD 255.xlsx.")

        last_row = sheet.Cells(sheet.Rows.Count, starts[0]).End(XL_UP).Row
        wave = read_col(sheet, FIRST_ROW, last_row, starts[0])
        (range1_min, range1_max), (range2_min, range2_max) = wavelength_ranges
        range1_rows = [FIRST_ROW+i for i, x in enumerate(wave)
                       if is_number(x) and range1_min <= float(x) <= range1_max]
        range2_rows = [FIRST_ROW+i for i, x in enumerate(wave)
                       if is_number(x) and range2_min <= float(x) <= range2_max]
        if not range1_rows:
            raise ValueError(
                f"Datos nav atrasti viļņa garumi pirmajam intervālam "
                f"{range1_min:g}–{range1_max:g} nm."
            )
        if not range2_rows:
            raise ValueError(
                f"Datos nav atrasti viļņa garumi otrajam intervālam "
                f"{range2_min:g}–{range2_max:g} nm."
            )

        # Laiks sekundēs no pirmā mērījuma.
        time_text = [_builtins.str(sheet.Cells(3, c+4).Text).strip() for c in starts]
        raw_seconds = [to_seconds(t) for t in time_text]
        elapsed, day_add, previous, first = [], 0, None, None
        for i, sec in enumerate(raw_seconds):
            if sec is None:
                elapsed.append(float(i))
                continue
            if previous is not None and sec < previous:
                day_add += 86400
            current = sec + day_add
            if first is None: first = current
            elapsed.append(float(current - first))
            previous = sec

        # Aprēķini atstāti tajā pašā lapā, tukšajās kolonnās aiz datiem.
        # Salīdzinājums ir abu lietotāja izvēlēto intervālu vidējo intensitāšu starpība.
        differences = []
        range1_label = f"Vidējais {range1_min:g}–{range1_max:g} nm"
        range2_label = f"Vidējais {range2_min:g}–{range2_max:g} nm"
        for col in starts:
            y = read_col(sheet, FIRST_ROW, last_row, col+1)
            vals1 = [float(v) for x, v in zip(wave, y)
                     if is_number(x) and is_number(v) and range1_min <= float(x) <= range1_max]
            vals2 = [float(v) for x, v in zip(wave, y)
                     if is_number(x) and is_number(v) and range2_min <= float(x) <= range2_max]
            if not vals1 or not vals2:
                raise ValueError(
                    f"Kādā mērījumā nav derīgu intensitātes vērtību intervāliem "
                    f"{range1_min:g}–{range1_max:g} nm vai {range2_min:g}–{range2_max:g} nm."
                )

            avg1_col, avg2_col, diff_col = col+5, col+6, col+7
            sheet.Cells(15, avg1_col).Value = range1_label
            sheet.Cells(15, avg2_col).Value = range2_label
            sheet.Cells(15, diff_col).Value = "Starpība"

            # Formulas izmanto atlasīto viļņa garumu rindu diapazonus.
            data_col_letter = col_letter(col + 1)
            avg1_addr = f"{data_col_letter}{range1_rows[0]}:{data_col_letter}{range1_rows[-1]}"
            avg2_addr = f"{data_col_letter}{range2_rows[0]}:{data_col_letter}{range2_rows[-1]}"
            sheet.Cells(16, avg1_col).Formula = f"=AVERAGE({avg1_addr})"
            sheet.Cells(16, avg2_col).Formula = f"=AVERAGE({avg2_addr})"
            sheet.Cells(16, diff_col).Formula = f"={cell_ref(16, avg1_col)}-{cell_ref(16, avg2_col)}"
            differences.append(sum(vals1) / len(vals1) - sum(vals2) / len(vals2))

        # Noņem vecos grafikus; visas jaunās diagrammas paliek lapā "Dati".
        for i in range(sheet.ChartObjects().Count, 0, -1):
            sheet.ChartObjects(i).Delete()

        # 1) Visi mērījumu spektri.
        c1 = new_chart(sheet, "C15", "J30", XL_SCATTER,
                       "Visi mērījumu spektri")
        c1.HasLegend = True
        c1.Legend.Position = XL_LEGEND_RIGHT
        for i, col in enumerate(starts):
            s = c1.SeriesCollection().NewSeries()
            # Leģendā laika apzīmējumi ir vienmērīgi: 0s, 20s, 40s, ...
            s.Name = f"{i * 20}s"
            s.XValues = sheet.Range(sheet.Cells(FIRST_ROW, col), sheet.Cells(last_row, col))
            s.Values = sheet.Range(sheet.Cells(FIRST_ROW, col+1), sheet.Cells(last_row, col+1))
            s.MarkerStyle, s.MarkerSize = 8, 3
            try:
                s.MarkerForegroundColor = rgb(COLORS[i % len(COLORS)])
                s.MarkerBackgroundColor = rgb(COLORS[i % len(COLORS)])
            except Exception:
                pass
        format_axes(c1, "Viļņa garums (nm)", "Signāla intensitāte", 150, 1200, 0, 120)

        # 2) Pirmā mērījuma spektrs.
        c2 = new_chart(sheet, "O18", "U29", XL_SCATTER, "Pirmā mērījuma spektrs")
        s = c2.SeriesCollection().NewSeries()
        s.Name = "Signāls"
        s.XValues = sheet.Range(sheet.Cells(FIRST_ROW, starts[0]), sheet.Cells(last_row, starts[0]))
        s.Values = sheet.Range(sheet.Cells(FIRST_ROW, starts[0]+1), sheet.Cells(last_row, starts[0]+1))
        s.MarkerStyle, s.MarkerSize = 8, 3
        format_axes(c2, "Viļņa garums (nm)", "Signāla intensitāte", 150, 1200, 0, 120)

        # 3) Aprēķinātās starpības izmaiņas laikā.
        c3 = new_chart(sheet, "C30", "J45", XL_SCATTER_LINES, "Laikā")
        s = c3.SeriesCollection().NewSeries()
        s.Name = "Starpība"
        s.XValues, s.Values = tuple(elapsed), tuple(differences)
        s.MarkerStyle, s.MarkerSize = 8, 6
        try:
            s.Format.Line.ForeColor.RGB = rgb(COLORS[1])
            s.Format.Line.Weight = 2
            s.MarkerForegroundColor = rgb(COLORS[1])
            s.MarkerBackgroundColor = rgb(COLORS[1])
        except Exception:
            pass
        format_axes(c3, "Laiks no pirmā mērījuma (s)", "Starpība")

        try: excel.CalculateFull()
        except Exception: pass
        sheet.Activate()
        excel.ScreenUpdating = True
        try:
            workbook.SaveAs(Filename=_builtins.str(output), FileFormat=XL_XLSX)
        except Exception as first_save_error:
            # Ja starplaikā cits process izveidoja šo failu, mēģina ar jaunu nosaukumu.
            output = available_output_path(source, numbered=True)
            try:
                workbook.SaveAs(Filename=_builtins.str(output), FileFormat=XL_XLSX)
            except Exception as second_save_error:
                raise RuntimeError(
                    "Excel neizdevās saglabāt apstrādāto failu.\n\n"
                    f"Mēģinātais ceļš: {output}\n\n"
                    "Pārbaudi, vai mape nav tikai lasāma un vai failu atļauts saglabāt. "
                    "Aizver Excel failus ar tādu pašu nosaukumu un mēģini vēlreiz.\n\n"
                    f"Pirmā kļūda: {first_save_error}\n"
                    f"Atkārtotā kļūda: {second_save_error}"
                ) from second_save_error
        saved = True
        excel.Visible = True
        sheet.Activate()
        return output
    finally:
        if not saved:
            try:
                if workbook is not None: workbook.Close(SaveChanges=False)
            except Exception: pass
            try: excel.Quit()
            except Exception: pass


def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Izvēlies sākuma Excel failu",
        filetypes=[("Excel faili", "*.xlsx")])
    if not file_path:
        root.destroy()
        return

    try:
        # Viļņu garumu ievades dialogs ir pirms Excel apstrādes/saglabāšanas.
        wavelength_ranges = ask_wavelength_ranges(root)
        if wavelength_ranges is None:
            root.destroy()
            return

        result = run(file_path, wavelength_ranges)
        messagebox.showinfo(
            "Gatavs",
            "Izveidoti 3 Scatter grafiki lapā Dati.\n\nFails:\n" + _builtins.str(result)
        )
    except Exception as error:
        # Saglabā precīzu traceback arī tad, ja kļūda rodas pašā ievades dialogā.
        details = traceback.format_exc()
        try:
            log_path = Path(file_path).with_name("DWK_kludas_zurnals.txt")
            log_path.write_text(details, encoding="utf-8")
        except Exception:
            pass
        messagebox.showerror(
            "Kļūda",
            "Neizdevās apstrādāt failu vai atvērt ievades logu:\n\n" + _builtins.str(error) +
            "\n\nPilna informācija saglabāta failā DWK_kludas_zurnals.txt.\n\n" + details[-1800:]
        )
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass


if __name__ == "__main__":
    main()
