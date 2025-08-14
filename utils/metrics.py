import calendar

def get_metric_rows(ws, base_row):
    """
    Επιστρέφει τις γραμμές για τα 6 metrics με βάση τη σταθερή σειρά τους στο block.
    Δεν ψάχνει τις ετικέτες — τις αντιστοιχεί απευθείας.
    """
    return {
        'ΕΠ.ΩΡΕΣ': base_row,
        'ΝΥΧΤΑ': base_row + 1,
        'ΑΡΓΙΑ': base_row + 2,
        'ΥΠΕΡΕΡΓΑΣΙΑ': base_row + 3,
        'ΥΠΕΡΩΡΙΑ': base_row + 4,
        'ΠΛΗΘΟΣ ΚΥΡΙΑΚΩΝ': base_row + 5
    }

def inspect_sunday_metrics(ws, row_lists, gui=None):
    def generate_day_columns():
        columns = []
        start_index = ord("H") - ord("A")
        for i in range(31):
            index = start_index + i
            col = ""
            while index >= 0:
                col = chr(index % 26 + 65) + col
                index = index // 26 - 1
            columns.append(col)
        return columns

    day_columns = generate_day_columns()

    for row_list in row_lists:
        base_row = row_list[0]
        try:
            metric_rows = get_metric_rows(ws, base_row)
            sunday_row = metric_rows['ΠΛΗΘΟΣ ΚΥΡΙΑΚΩΝ']
            values = []
            start_index = ord("H") - ord("A")
            # iterate once and avoid many attribute lookups
            for offset in range(31):
                col_index = start_index + offset + 1
                val = ws.cell(row=sunday_row, column=col_index).value
                if val not in (None, 0, '', '0'):
                    col_letter = day_columns[offset]
                    values.append(f"{col_letter}: {val}")
            if values:
                message = f"📊 Γραμμή ΠΛΗΘΟΣ ΚΥΡΙΑΚΩΝ ({sunday_row}):\n" + ", ".join(values)
                if gui:
                    gui.show_message(message, level="debug")
        except Exception as e:
            msg = f"⚠️ Σφάλμα για εργαζόμενο στη γραμμή {base_row}: {str(e)}"
            if gui:
                gui.show_message(msg, level="warning")
            else:
                print(msg)

def update_sundays(ws, row_lists, year, month):
    sundays = [day for day in range(1, calendar.monthrange(year, month)[1] + 1)
               if calendar.weekday(year, month, day) == 6]

    start_col_index = ord("H") - ord("A") + 1
    day_to_column = {day: start_col_index + day - 1 for day in sundays}

    for row_list in row_lists:
        if len(row_list) < 6:
            print(f"⚠️ Σφάλμα: row_list δεν έχει 6 γραμμές ➤ {row_list}")
            continue

        base_row = row_list[0]
        try:
            metric_rows = get_metric_rows(ws, base_row)
            sunday_row = metric_rows['ΠΛΗΘΟΣ ΚΥΡΙΑΚΩΝ']

            # vectorized-like test: reuse column indices
            for day, col_index in day_to_column.items():
                worked = any(
                    ws.cell(row=r, column=col_index).value not in (None, 0, 0.0, '', '0')
                    for r in row_list
                )
                ws.cell(row=sunday_row, column=col_index).value = 1 if worked else 0

        except Exception as e:
            print(f"⚠️ Σφάλμα για εργαζόμενο στη γραμμή {base_row}: {str(e)}")