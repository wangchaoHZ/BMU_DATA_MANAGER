import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pymodbus.client import ModbusTcpClient
import json

# ======= 标准底层读写 =======
def read_register(ip, port, address, count):
    client = ModbusTcpClient(ip, port=port)
    client.connect()
    rr = client.read_holding_registers(address=address, count=count)
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

class ModbusApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Modbus数据导入导出工具")
        self.geometry("720x520")
        self.configure(bg="#f4f6fb")
        self._set_style()
        self._init_gui()

    def _set_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Arial', 13), padding=[10, 5], background='#3a7bd5', foreground='#fff', borderwidth=0)
        style.map('TButton', background=[('active', '#5596e6')])
        style.configure('TLabel', font=('Arial', 12), background="#f4f6fb", foreground='#333')
        style.configure('TEntry', font=('Arial', 12))
        self.option_add("*Font", "Arial 12")

    def _init_gui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill='both', expand=True, padx=18, pady=18)

        tab_read = tk.Frame(notebook, bg="#f4f6fb")
        tab_write = tk.Frame(notebook, bg="#f4f6fb")
        notebook.add(tab_read, text="数据导出 (读取并存为JSON)")
        notebook.add(tab_write, text="数据导入 (JSON写入寄存器)")

        self._read_tab(tab_read)
        self._write_tab(tab_write)

    # 数据导出 Tab
    def _read_tab(self, parent):
        frm = tk.Frame(parent, bg="#f4f6fb")
        frm.pack(padx=30, pady=30, fill='both', expand=True)

        tk.Label(frm, text="IP地址:").grid(row=0, column=0, sticky='e')
        ip_entry = tk.Entry(frm, width=15)
        ip_entry.insert(0, "127.0.0.1")
        ip_entry.grid(row=0, column=1, padx=8)

        tk.Label(frm, text="端口:").grid(row=0, column=2, sticky='e')
        port_entry = tk.Entry(frm, width=8)
        port_entry.insert(0, "502")
        port_entry.grid(row=0, column=3, padx=8)

        tk.Label(frm, text="起始地址:").grid(row=1, column=0, sticky='e')
        start_entry = tk.Entry(frm, width=8)
        start_entry.insert(0, "200")
        start_entry.grid(row=1, column=1, padx=8)

        tk.Label(frm, text="长度:").grid(row=1, column=2, sticky='e')
        length_entry = tk.Entry(frm, width=8)
        length_entry.insert(0, "10")
        length_entry.grid(row=1, column=3, padx=8)

        result_label = tk.Label(frm, text="", bg="#f4f6fb", font=('Arial', 12))
        result_label.grid(row=2, column=0, columnspan=4, pady=16)

        def do_read_and_export():
            ip = ip_entry.get().strip()
            port = int(port_entry.get())
            start_addr = int(start_entry.get())
            length = int(length_entry.get())
            data = read_register(ip, port, start_addr, length)
            if data is None:
                result_label.config(text="读取失败或无响应", fg="#e74c3c")
                return
            json_data = [{"addr": start_addr + i, "data": v} for i, v in enumerate(data)]
            filepath = filedialog.asksaveasfilename(title="保存为JSON文件", defaultextension=".json", filetypes=[("JSON files", "*.json")])
            if not filepath:
                return
            export_json(json_data, filepath)
            result_label.config(text=f"导出成功，共{len(json_data)}条数据", fg="#27ae60")

        ttk.Button(frm, text="读取并导出为JSON", command=do_read_and_export).grid(row=3, column=1, columnspan=2, pady=20)

    # 数据导入 Tab
    def _write_tab(self, parent):
        frm = tk.Frame(parent, bg="#f4f6fb")
        frm.pack(padx=30, pady=30, fill='both', expand=True)

        tk.Label(frm, text="IP地址:").grid(row=0, column=0, sticky='e')
        ip_entry = tk.Entry(frm, width=15)
        ip_entry.insert(0, "127.0.0.1")
        ip_entry.grid(row=0, column=1, padx=8)

        tk.Label(frm, text="端口:").grid(row=0, column=2, sticky='e')
        port_entry = tk.Entry(frm, width=8)
        port_entry.insert(0, "502")
        port_entry.grid(row=0, column=3, padx=8)

        self.json_path_var = tk.StringVar()
        tk.Label(frm, text="待写入JSON文件:").grid(row=1, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.json_path_var, width=32, state="readonly").grid(row=1, column=1, columnspan=2, padx=8)

        def select_json_file():
            filepath = filedialog.askopenfilename(title="选择JSON文件", filetypes=[("JSON files", "*.json")])
            if filepath:
                self.json_path_var.set(filepath)

        ttk.Button(frm, text="选择JSON文件", command=select_json_file).grid(row=1, column=3, padx=8)

        result_label = tk.Label(frm, text="", bg="#f4f6fb", font=('Arial', 12))
        result_label.grid(row=2, column=0, columnspan=4, pady=16)

        def do_write_from_json():
            ip = ip_entry.get().strip()
            port = int(port_entry.get())
            filepath = self.json_path_var.get()
            if not filepath or not filepath.endswith(".json"):
                result_label.config(text="请先选择待写入的JSON文件", fg="#e74c3c")
                return
            try:
                data_list = import_json(filepath)
            except Exception as e:
                result_label.config(text=f"JSON解析失败: {str(e)}", fg="#e74c3c")
                return
            success = 0
            fail = 0
            for item in data_list:
                addr = int(item["addr"])
                value = int(item["data"])
                ok = write_register(ip, port, addr, value)
                if ok:
                    success += 1
                else:
                    fail += 1
            result_label.config(text=f"写入完成，成功:{success} 失败:{fail}", fg="#27ae60" if fail==0 else "#e74c3c")

        ttk.Button(frm, text="写入到寄存器", command=do_write_from_json).grid(row=3, column=1, columnspan=2, pady=20)

if __name__ == "__main__":
    app = ModbusApp()
    app.mainloop()