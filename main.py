import tkinter as tk
from tkinter import ttk, filedialog
from pymodbus.client import ModbusTcpClient
import json

def read_register(ip, port, address):
    client = ModbusTcpClient(ip, port=port)
    client.connect()
    rr = client.read_holding_registers(address)
    client.close()
    if rr and hasattr(rr, "registers") and rr.registers:
        return rr.registers
    else:
        return None

def write_register(ip, port, address, value):
    client = ModbusTcpClient(ip, port=port)
    client.connect()
    wr = client.write_register(address, value)
    client.close()
    if wr and not wr.isError():
        return True
    else:
        return False

def export_json(data, filepath):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def import_json(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise Exception("JSON内容必须为数组")
    for item in data:
        if "addr" not in item or "data" not in item:
            raise Exception("每个对象必须包含 addr 和 data")
    return data

class ModernButton(ttk.Button):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(style="Modern.TButton")

class ModbusApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BMU数据导入导出工具")
        self._fixed_window()
        self._set_style()
        self._init_gui()

    def _fixed_window(self):
        ww, wh = 1280, 750  # 更高，保证全部显示
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - ww) // 2
        y = max((sh - wh) // 3, 20)
        self.geometry(f"{ww}x{wh}+{x}+{y}")
        self.resizable(False, False)

    def _set_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Card.TLabelframe', background="#232C3D", borderwidth=2)
        style.configure('Card.TLabelframe.Label', background="#232C3D", foreground="#7DF9FF", font=("Arial", 14, "bold"))
        # 科技感按钮
        style.configure('Modern.TButton',
            font=("Arial", 13, "bold"),
            background='#00C9A7',
            foreground='#232C3D',
            borderwidth=0,
            padding=8,
            relief="flat"
        )
        style.map('Modern.TButton',
            background=[('active', '#7DF9FF'), ('pressed', '#00A1A7')],
            foreground=[('active', '#232C3D'), ('pressed', '#232C3D')]
        )
        style.configure('TLabel',
            font=("Arial", 12),
            background="#171E29",
            foreground='#7DF9FF')
        style.configure('TEntry',
            font=("Arial", 12),
            fieldbackground="#253347",
            foreground='#FFF')
        style.configure('TProgressbar',
            background='#00C9A7',
            troughcolor='#232C3D',
            bordercolor='#171E29')
        self.option_add("*Font", "Arial 12")
        self.option_add("*Entry.Font", "Arial 12")
        self.configure(bg="#171E29")

    def _init_gui(self):
        main = tk.Frame(self, bg="#171E29")
        main.pack(fill='both', expand=True, padx=24, pady=12)

        card_bg = "#232C3D"
        label_fg = "#7DF9FF"
        entry_bg = "#253347"
        entry_fg = "#FFFFFF"

        # ---- 参数区 ----
        param_frame = ttk.LabelFrame(main, text="Modbus参数设置", style='Card.TLabelframe')
        param_frame.pack(fill='x', pady=12, ipady=4)
        for i in range(6):
            param_frame.grid_columnconfigure(i, weight=1, uniform="col")

        tk.Label(param_frame, text="IP地址:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=0, sticky='e', padx=8, pady=7)
        self.ip_entry = tk.Entry(param_frame, width=17, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.ip_entry.insert(0, "127.0.0.1")
        self.ip_entry.grid(row=0, column=1, sticky='w', padx=8, pady=7)

        tk.Label(param_frame, text="端口:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=2, sticky='e', padx=8, pady=7)
        self.port_entry = tk.Entry(param_frame, width=10, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.port_entry.insert(0, "502")
        self.port_entry.grid(row=0, column=3, sticky='w', padx=8, pady=7)

        tk.Label(param_frame, text="从站地址:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=4, sticky='e', padx=8, pady=7)
        self.unit_entry = tk.Entry(param_frame, width=10, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.unit_entry.insert(0, "1")
        self.unit_entry.grid(row=0, column=5, sticky='w', padx=8, pady=7)

        # ---- 导出区 ----
        export_frame = ttk.LabelFrame(main, text="读取并导出为JSON", style='Card.TLabelframe')
        export_frame.pack(fill='x', pady=12, ipady=4)
        for i in range(6):
            export_frame.grid_columnconfigure(i, weight=1, uniform="col")

        tk.Label(export_frame, text="起始地址:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=0, sticky='e', padx=8, pady=7)
        self.start_entry = tk.Entry(export_frame, width=13, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.start_entry.insert(0, "200")
        self.start_entry.grid(row=0, column=1, sticky='w', padx=8, pady=7)

        tk.Label(export_frame, text="长度:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=2, sticky='e', padx=8, pady=7)
        self.length_entry = tk.Entry(export_frame, width=13, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.length_entry.insert(0, "10")
        self.length_entry.grid(row=0, column=3, sticky='w', padx=8, pady=7)

        self.export_btn = ModernButton(export_frame, text="读取并导出为JSON", width=20, command=self.do_read_and_export)
        self.export_btn.grid(row=0, column=5, padx=10, pady=7)

        # 导出进度条和百分比
        self.export_progress_var = tk.DoubleVar(value=0)
        self.export_progress_bar = ttk.Progressbar(export_frame, variable=self.export_progress_var, maximum=100, mode='determinate')
        self.export_progress_bar.grid(row=1, column=0, columnspan=6, sticky='ew', padx=20, pady=5)
        self.export_progress_label = tk.Label(export_frame, text="进度：0%", bg=card_bg, fg="#00C9A7", font=("Arial", 12))
        self.export_progress_label.grid(row=2, column=0, columnspan=6, sticky='w', padx=20, pady=4)

        self.export_result_label = tk.Label(export_frame, text="", bg=card_bg, font=("Arial", 12), fg=label_fg)
        self.export_result_label.grid(row=3, column=0, columnspan=6, sticky='w', padx=10, pady=6)

        # ---- 导入区 ----
        import_frame = ttk.LabelFrame(main, text="批量导入JSON写入寄存器", style='Card.TLabelframe')
        import_frame.pack(fill='x', pady=12, ipady=4)
        for i in range(6):
            import_frame.grid_columnconfigure(i, weight=1, uniform="col")

        self.json_path_var = tk.StringVar()
        tk.Label(import_frame, text="待写入JSON文件:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=0, sticky='e', padx=8, pady=7)
        self.json_path_entry = tk.Entry(import_frame, textvariable=self.json_path_var, width=32, bg=entry_bg, fg="black", state="readonly", relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.json_path_entry.grid(row=0, column=1, columnspan=3, sticky='we', padx=8, pady=7)
        ModernButton(import_frame, text="选择JSON文件", width=16, command=self._select_json_file).grid(row=0, column=4, padx=10, pady=7)
        ModernButton(import_frame, text="批量写入到寄存器", width=20, command=self.do_write_from_json).grid(row=0, column=5, padx=10, pady=7)

        self.progress_var = tk.DoubleVar(value=0)
        progress_bar = ttk.Progressbar(import_frame, variable=self.progress_var, maximum=100, mode='determinate')
        progress_bar.grid(row=1, column=0, columnspan=6, sticky='ew', padx=20, pady=5)

        self.progress_label = tk.Label(import_frame, text="进度：0%", bg=card_bg, fg="#00C9A7", font=("Arial", 12))
        self.progress_label.grid(row=2, column=0, columnspan=6, sticky='w', padx=20, pady=4)

        self.import_result_label = tk.Label(import_frame, text="", bg=card_bg, font=("Arial", 12), fg=label_fg)
        self.import_result_label.grid(row=3, column=0, columnspan=6, sticky='w', padx=10, pady=6)

        # ---- 单地址写入 ----
        single_frame = ttk.LabelFrame(main, text="单独写入指定寄存器", style='Card.TLabelframe')
        single_frame.pack(fill='x', pady=30, ipady=8)  # 加大pady，保证不遮住
        for i in range(6):
            single_frame.grid_columnconfigure(i, weight=1, uniform="col")

        tk.Label(single_frame, text="地址:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=0, sticky='e', padx=8, pady=7)
        self.addr_entry = tk.Entry(single_frame, width=13, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.addr_entry.insert(0, "200")
        self.addr_entry.grid(row=0, column=1, sticky='w', padx=8, pady=7)

        tk.Label(single_frame, text="值:", bg=card_bg, fg=label_fg, font=("Arial", 12)).grid(row=0, column=2, sticky='e', padx=8, pady=7)
        self.value_entry = tk.Entry(single_frame, width=13, bg=entry_bg, fg=entry_fg, relief="flat", highlightthickness=1, highlightbackground="#00C9A7", justify='center', font=("Arial", 12))
        self.value_entry.insert(0, "1")
        self.value_entry.grid(row=0, column=3, sticky='w', padx=8, pady=7)

        ModernButton(single_frame, text="单独写入该地址", width=20, command=self.do_single_write).grid(row=0, column=5, padx=10, pady=7)
        self.single_result_label = tk.Label(single_frame, text="", bg=card_bg, font=("Arial", 12), fg=label_fg)
        self.single_result_label.grid(row=1, column=0, columnspan=6, sticky='w', padx=10, pady=10)

    def _select_json_file(self):
        filepath = filedialog.askopenfilename(title="选择JSON文件", filetypes=[("JSON files", "*.json")])
        if filepath:
            self.json_path_var.set(filepath)

    def do_read_and_export(self):
        ip = self.ip_entry.get().strip()
        port = int(self.port_entry.get())
        start_addr = int(self.start_entry.get())
        length = int(self.length_entry.get())
        self.export_result_label.config(text="正在读取...", fg="#7DF9FF")
        self.update_idletasks()
        data = []
        for i in range(length):
            d = read_register(ip, port, start_addr + i)
            if d is not None:
                data.append({"addr": start_addr + i, "data": d[0]})
            else:
                data.append({"addr": start_addr + i, "data": None})
            percent = int((i+1)*100/length)
            self.export_progress_var.set(percent)
            self.export_progress_label.config(text=f"进度：{percent}%")
            self.update_idletasks()
        filepath = filedialog.asksaveasfilename(title="保存为JSON文件", defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not filepath:
            self.export_result_label.config(text="取消导出", fg="#999")
            self.export_progress_var.set(0)
            self.export_progress_label.config(text="进度：0%")
            return
        export_json(data, filepath)
        self.export_result_label.config(text=f"导出成功，共{len(data)}条数据", fg="#00ffcc")
        self.export_progress_var.set(0)
        self.export_progress_label.config(text="进度：0%")

    def do_write_from_json(self):
        ip = self.ip_entry.get().strip()
        port = int(self.port_entry.get())
        filepath = self.json_path_var.get()
        if not filepath or not filepath.endswith(".json"):
            self.import_result_label.config(text="请先选择待写入的JSON文件", fg="#ff5e5e")
            self.progress_label.config(text=f"进度：0%")
            return
        try:
            data_list = import_json(filepath)
        except Exception as e:
            self.import_result_label.config(text=f"JSON解析失败: {str(e)}", fg="#ff5e5e")
            self.progress_label.config(text=f"进度：0%")
            return
        total = len(data_list)
        if total == 0:
            self.import_result_label.config(text="JSON无数据", fg="#ff5e5e")
            self.progress_label.config(text=f"进度：0%")
            return
        success = 0
        fail = 0
        self.import_result_label.config(text="正在写入...", fg="#00ffcc")
        self.update_idletasks()
        for idx, item in enumerate(data_list):
            addr = int(item["addr"])
            value = int(item["data"])
            ok = write_register(ip, port, addr, value)
            if ok:
                success += 1
            else:
                fail += 1
            percent = int((idx+1)*100/total)
            self.progress_var.set(percent)
            self.progress_label.config(text=f"进度：{percent}%")
            self.update_idletasks()
        self.import_result_label.config(text=f"写入完成，成功:{success} 失败:{fail}", fg="#00ffcc" if fail==0 else "#ff5e5e")
        self.progress_var.set(0)
        self.progress_label.config(text=f"进度：0%")

    def do_single_write(self):
        ip = self.ip_entry.get().strip()
        port = int(self.port_entry.get())
        addr = int(self.addr_entry.get())
        value = int(self.value_entry.get())
        self.single_result_label.config(text="正在写入...", fg="#00ffcc")
        self.update_idletasks()
        ok = write_register(ip, port, addr, value)
        if ok:
            self.single_result_label.config(text=f"写入成功: 地址{addr}, 值={value}", fg="#00ffcc")
        else:
            self.single_result_label.config(text="写入失败", fg="#ff5e5e")

if __name__ == "__main__":
    app = ModbusApp()
    app.mainloop()