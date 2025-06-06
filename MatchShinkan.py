import tkinter as tk
from tkinter import filedialog, messagebox, IntVar, BooleanVar
import pandas as pd
from collections import defaultdict
import datetime
import os
from openpyxl import Workbook
import win32com.client
from openpyxl.styles import Alignment

#ログ出力
def log(message):
    with open("log.txt", "a", encoding="utf-8") as f:#保存場所
        f.write(f"[{datetime.datetime.now()}] {message}\n")

#希望楽器データ成形
def clean_instrument_field(text):
    if pd.isna(text): return []
    allowed = ['ギター', 'ベース', 'ドラム', 'キーボード', 'その他']#楽器の種類
    parts = [p.strip() for p in str(text).replace('\n', ',').split(',')]
    return [p for p in parts if p in allowed]

#日付➡時間抽出
def extract_availability(df):
    date_columns = [col for col in df.columns[4:] if isinstance(col, str) and '/' in col]
    log(f"[extract_availability] 検出された日付列: {date_columns}")
    availability = []
    for row_idx, (_, row) in enumerate(df.iterrows()):
        slots = []
        for col in date_columns:
            val = row[col]
            if pd.isna(val):
                log(f"[row {row_idx}] {col} は NaN")
                continue
            try:
                if isinstance(val, str):
                    for t in val.split(','):
                        if pd.isna(t):
                            log(f"[row {row_idx}] {col} の中の要素が NaN: {t}")
                            continue
                        t = str(t).strip()
                        if t and t.lower() != 'nan':
                            slots.append((col.strip(), t))
                else:
                    t = str(val).strip()
                    if t and t.lower() != 'nan':
                        slots.append((col.strip(), t))
            except Exception as e:
                log(f"[row {row_idx}] エラー: col={col}, val={val}, type={type(val)}, error={e}")
        log(f"[row {row_idx}] 抽出された時間帯: {slots}")
        availability.append(slots)
    return availability


#個人ごとに希望を纏める関数
def parse_people(df):
    availability = extract_availability(df)

    date_columns = [col for col in df.columns[4:] if isinstance(col, str) and '/' in col]
    log(f"[parse_people] 日付列検出: {date_columns}")
    if not date_columns:
        log("[parse_people] エラー: 日付列が1つも検出されませんでした")
        raise ValueError("日付列が見つかりませんでした")

    instrument_col_index = df.columns.get_loc(date_columns[-1]) + 1
    remarks_col_index = instrument_col_index + 1
    log(f"[parse_people] 楽器列 index: {instrument_col_index}, 備考列 index: {remarks_col_index}")

    people = []
    for i, (_, row) in enumerate(df.iterrows()):
        try:
            name = row[2]
            line = row[3]
            instruments = clean_instrument_field(row[instrument_col_index])
            remarks = row[remarks_col_index] if remarks_col_index < len(row) else ""
            log(f"[row {i}] 名前: {name}, LINE: {line}, 楽器: {instruments}, 備考: {remarks}, availability: {availability[i]}")
            for inst in instruments:
                people.append({
                    "name": name,
                    "line": line,
                    "instrument": inst,
                    "remarks": remarks,
                    "availability": availability[i]
                })
        except Exception as e:
            log(f"[row {i}] parse_people エラー: {e}")
            raise
    return people


# 時間分割処理関数

def split_time_range(time_str, interval_minutes):
    try:
        start_str, end_str = time_str.split('-')
        start = datetime.datetime.strptime(start_str.strip(), '%H:%M')
        end = datetime.datetime.strptime(end_str.strip(), '%H:%M')
        slots = []
        while start + datetime.timedelta(minutes=interval_minutes) <= end:
            slot_end = start + datetime.timedelta(minutes=interval_minutes)
            slots.append(f"{start.strftime('%H:%M')}-{slot_end.strftime('%H:%M')}")
            start = slot_end
        return slots
    except:
        return [time_str]
#講師の時間を分割
def expand_teacher_availability(teachers, interval):
    for t in teachers:
        expanded = []
        for date, time in t["availability"]:
            split_times = split_time_range(time, interval)
            for st in split_times:
                expanded.append((date, st))
        t["availability"] = expanded
    return teachers

#マッチング処理関数
def match(teachers, students, max_per_instrument=1, drum_exclusive=True,
          allow_split=False, split_interval=30, max_pair=2,
          drum_max_per_slot=1, prefer_same_teacher=False, prefer_continuous=False):

    result = defaultdict(list)
    teacher_usage = defaultdict(set)
    student_instr_count = defaultdict(int)
    student_used_slots = defaultdict(set)
    unmatched = defaultdict(list)

    students = sorted(students, key=lambda s: len(s["availability"]))

    if allow_split:
        for t in teachers:
            expanded = []
            for date, time in t["availability"]:
                split_times = split_time_range(time, split_interval)
                for st in split_times:
                    expanded.append((date, st))
            t["availability"] = expanded

    def get_adjacent_slots(slots, target_date, base_time):
        times = sorted(set(t for d, t in slots if d == target_date))
        if base_time not in times:
            return []
        index = times.index(base_time)
        adjacent = []
        if index > 0: adjacent.append(times[index - 1])
        if index < len(times) - 1: adjacent.append(times[index + 1])
        return adjacent

    def assign_slots(target_count):
        for student in students:
            key = (student["name"], student["instrument"])
            if student_instr_count[key] >= target_count:
                continue

            # 割り当て済みの時間帯（あれば）を取得
            used_slots = sorted(list(student_used_slots[student["name"]]))
            available = student["availability"]

            # 時間の近さで並び替える（prefer_continuous が True のときのみ）
            if prefer_continuous and used_slots:
                def time_distance(slot):
                    date, time = slot
                    for u_date, u_time in used_slots:
                        if u_date != date: continue
                        try:
                            t1 = datetime.datetime.strptime(time.split('-')[0], "%H:%M")
                            t2 = datetime.datetime.strptime(u_time.split('-')[0], "%H:%M")
                            return abs((t1 - t2).total_seconds())
                        except:
                            continue
                    return float('inf')  # 日付が違う場合は遠い
                available = sorted(available, key=time_distance)

            matched_this_round = False
            for date, time in available:
                times = split_time_range(time, split_interval) if allow_split else [time]
                for split_time in times:
                    slot_key = (date, split_time)
                    if len(result[slot_key]) >= max_pair:
                        continue
                    for teacher in teachers:
                        if teacher["instrument"] != student["instrument"]: continue
                        if (teacher["name"], slot_key) in teacher_usage: continue
                        if slot_key not in teacher["availability"]: continue

                        instruments_in_slot = {m["teacher"]["instrument"] for m in result[slot_key]}

                        if drum_exclusive:
                            if "ドラム" in instruments_in_slot and teacher["instrument"] != "ドラム":
                                continue
                            if teacher["instrument"] == "ドラム" and any(inst != "ドラム" for inst in instruments_in_slot):
                                continue
                        else:
                            if teacher["instrument"] == "ドラム" and any(inst != "ドラム" for inst in instruments_in_slot):
                                continue
                            if teacher["instrument"] != "ドラム" and "ドラム" in instruments_in_slot:
                                continue

                        # ドラムの最大人数チェック
                        if teacher["instrument"] == "ドラム":
                            drum_count = sum(1 for m in result[slot_key] if m["teacher"]["instrument"] == "ドラム")
                            if drum_count >= drum_max_per_slot:
                                continue
                            #ドラムがなるべく1人になるように割り当てる
                                if len(result[slot_key]) < max_pair:
                                    continue

                        result[slot_key].append({"student": student, "teacher": teacher})
                        teacher_usage[(teacher["name"], slot_key)] = True
                        student_instr_count[key] += 1
                        student_used_slots[student["name"]].add(slot_key)
                        matched_this_round = True
                        break
                    if matched_this_round:
                        break
                if matched_this_round:
                    break

    assign_slots(1)
    if max_per_instrument > 1:
        assign_slots(max_per_instrument)

    for student in students:
        key = (student["name"], student["instrument"])
        if student_instr_count[key] < 1:
            unmatched[student["name"].strip()].append(student)

    for name, entries in unmatched.items():
        for s in entries:
            s["availability"] = [slot for slot in s["availability"] if slot not in student_used_slots[s["name"]]]

    unused_teachers = []
    for t in teachers:
        for d, tslot in t["availability"]:
            slot_key = (d, tslot)
            if (t["name"], slot_key) not in teacher_usage:
                unused_teachers.append({"name": t["name"], "instrument": t["instrument"], "date": d, "time": tslot})

    return result, unmatched, unused_teachers


#エクセル書き込み
def write_excel(result, unmatched, unused_teachers, path, split_mode=1):
    wb = Workbook()
    if split_mode in [1, 3]:
        ws = wb.active
        ws.title = "マッチング結果"
        ws.append(["名前", "LINE", "パート", "日付", "時間", "講師", "備考"])
        for (date, time), matches in sorted(result.items(), key=lambda x: (datetime.datetime.strptime(x[0][0], "%m/%d"), x[0][1])):
            for match in matches:
                s = match["student"]
                t = match["teacher"]
                ws.append([s["name"], s["line"], s["instrument"], date, time, f"{t['name']}({t['instrument'][0]})", s["remarks"]])
        ws.append([])
        ws.append(["-- 講師未割当 --"])
        for name, entries in unmatched.items():
            for s in entries:
                times = [f"{d} {t}" for d, t in s["availability"]]
                ws.append([s["name"], s["line"], s["instrument"], ", ".join(times), "", "", s["remarks"]])
        ws.append([])
        ws.append(["-- 空いている講師一覧 --"])
        for t in unused_teachers:
            ws.append([t["name"], "", t["instrument"], t["date"], t["time"], "", ""])

    if split_mode in [2, 3]:
        dates = sorted(set(d for d, _ in result.keys()))
        for date in dates:
            safe_title = f"{int(date.split('/')[0])}月{int(date.split('/')[1])}日"
            ws = wb.create_sheet(title=safe_title)
            ws.append(["名前", "LINE", "パート", "時間", "講師", "備考"])
            times = sorted(t for d, t in result if d == date)
            for time in times:
                for match in result[(date, time)]:
                    s = match["student"]
                    t = match["teacher"]
                    ws.append([s["name"], s["line"], s["instrument"], time, f"{t['name']}({t['instrument'][0]})", s["remarks"]])
            #割り当てられなかった生徒の処理
            ws.append([])
            ws.append(["-- 未割当生徒 --"])
            for name, entries in unmatched.items():
                for s in entries:
                    avail = [f"{d} {t}" for d, t in s["availability"] if d == date]
                    if avail:
                        ws.append([s["name"], s["line"], s["instrument"], ", ".join(avail), "", s["remarks"]])
            #空いてる時間のある講師
            ws.append([])
            ws.append(["-- 空いている講師一覧 --"])
            for t in unused_teachers:
                if t["date"] == date:
                    ws.append([t["name"], "", t["instrument"], t["time"], "", ""])

    wb.save(path)
    log(f"Excel出力完了: {path}")


#tkGUI処理
class MatchApp:
    def __init__(self, root):
        self.root = root
        self.teacher_file = None
        self.student_file = None
        self.max_teacher_slot = IntVar(value=1)
        self.max_pair = IntVar(value=2)
        self.drum_exclusive = BooleanVar(value=False)
        self.output_excel = BooleanVar(value=True)
        self.output_pdf = BooleanVar(value=True)
        self.max_slots_per_instrument = IntVar(value=1)
        self.enable_split = BooleanVar(value=False)
        self.split_minutes = IntVar(value=30)
        self.output_mode = IntVar(value=1)
        self.prefer_same_teacher = BooleanVar(value=False)
        self.drum_max_per_slot = IntVar(value=1)
        self.prefer_continuous = BooleanVar(value=False)


        root.title("講習マッチング")
        root.geometry("520x600")

        tk.Button(root, text="講師ファイルを選択", command=self.load_teacher).pack(pady=5)
        tk.Button(root, text="生徒ファイルを選択", command=self.load_student).pack(pady=5)

        tk.Label(root, text="① 講師1人につき1コマの生徒人数").pack()
        tk.Spinbox(root, from_=1, to=5, textvariable=self.max_teacher_slot, width=5).pack()

        tk.Label(root, text="② 同時間帯に最大組数").pack()
        tk.Spinbox(root, from_=1, to=5, textvariable=self.max_pair, width=5).pack()

        tk.Checkbutton(root, text="③ ドラムは他楽器と同時不可", variable=self.drum_exclusive).pack()

        tk.Label(root, text="④ ドラムの1コマ最大組数").pack()
        tk.Spinbox(root, from_=1, to=5, textvariable=self.drum_max_per_slot, width=5).pack()

        tk.Label(root, text="⑤ 楽器ごとに最大何枠まで許可").pack()
        tk.Spinbox(root, from_=1, to=10, textvariable=self.max_slots_per_instrument, width=5).pack()

        tk.Checkbutton(root, text="⑥ 割り当て失敗時、時間枠を分割して再試行", variable=self.enable_split).pack()
        tk.Label(root, text="分割単位（分）").pack()
        tk.Radiobutton(root, text="30分ごと", variable=self.split_minutes, value=30).pack()
        tk.Radiobutton(root, text="20分ごと", variable=self.split_minutes, value=20).pack()

        tk.Checkbutton(root, text="⑦ 2枠目もできるだけ同じ講師にする", variable=self.prefer_same_teacher).pack()
        tk.Checkbutton(root, text="⑧ 講習会のコマをできるだけ連続にする", variable=self.prefer_continuous).pack()

        tk.Label(root, text="<出力形式>").pack()
        tk.Radiobutton(root, text="1シートにまとめる", variable=self.output_mode, value=1).pack()
        tk.Radiobutton(root, text="日付ごとに分ける", variable=self.output_mode, value=2).pack()
        tk.Radiobutton(root, text="両方出力", variable=self.output_mode, value=3).pack()

        tk.Checkbutton(root, text="Excelで出力", variable=self.output_excel).pack()

        tk.Button(root, text="マッチング開始", command=self.run, bg="lightgreen").pack(pady=20)

    def load_teacher(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.teacher_file = path
            messagebox.showinfo("読み込み成功", "講師ファイルを読み込みました")
            log(f"講師ファイル読み込み: {path}")

    def load_student(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.student_file = path
            messagebox.showinfo("読み込み成功", "生徒ファイルを読み込みました")
            log(f"生徒ファイル読み込み: {path}")

    def run(self):
        if not self.teacher_file or not self.student_file:
            messagebox.showwarning("エラー", "両ファイルを選択してください")
            return
        try:
            teachers = parse_people(pd.read_excel(self.teacher_file))
            students = parse_people(pd.read_excel(self.student_file))

            result, unmatched, unused_teachers = match(
                teachers,
                students,
                max_per_instrument=self.max_slots_per_instrument.get(),
                drum_exclusive=self.drum_exclusive.get(),
                allow_split=self.enable_split.get(),
                split_interval=self.split_minutes.get(),
                max_pair=self.max_pair.get(),
                prefer_same_teacher=self.prefer_same_teacher.get(),
                prefer_continuous=self.prefer_continuous.get()
            )

            now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

            if self.output_excel.get():
                path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"講習会マッチング表_{now}.xlsx")
                if path:
                    write_excel(result, unmatched, unused_teachers, path, self.output_mode.get())
                    messagebox.showinfo("出力完了", "Excelファイルを出力しました")

        except Exception as e:
            log(f"エラー: {str(e)}")
            messagebox.showerror("エラー", str(e))

#アプリ起動
if __name__ == "__main__":
    root = tk.Tk()
    app = MatchApp(root)
    root.mainloop()
