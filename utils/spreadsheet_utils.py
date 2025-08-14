import os
import subprocess
from tkinter import messagebox

def get_column_from_day(day_of_month):
    """
    Επιστρέφει τη στήλη Excel που αντιστοιχεί σε μια ημέρα του μήνα,
    ξεκινώντας από τη στήλη 'H' για την 1η ημέρα.
    """
    if not (1 <= day_of_month <= 31):
        raise ValueError("Η ημέρα πρέπει να είναι μεταξύ 1 και 31")

    start_index = ord("H") - ord("A")  # 7 (0-based)
    column_index = start_index + (day_of_month - 1)

    def index_to_excel_column(index):
        letters = ""
        while index >= 0:
            letters = chr(index % 26 + 65) + letters
            index = index // 26 - 1
        return letters

    return index_to_excel_column(column_index)

def open_excel(path):
    if os.path.exists(path):
        # Use platform agnostic approach
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            # fallback to start via shell
            subprocess.Popen(["start", "", path], shell=True)
    else:
        messagebox.showwarning("Προσοχή", "Δεν βρέθηκε το αρχείο για άνοιγμα.")