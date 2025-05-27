import tkinter as tk
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
import math
import os
import re

def on_cable_type_change(event):
    selected_type = cable_type_var.get()
    if selected_type == "Single-core":
        placement_label.config(text="Placement Type:")
        placement_label.grid()
        placement_menu.grid()
        trench_combobox['values'] = ['1', '2']
        if trench_combobox.get() not in ['1', '2']:
            trench_combobox.set()  # Varsayılan olarak 1 seç
    else:
        placement_label.grid_remove()
        placement_menu.grid_remove()
        trench_combobox['values'] = [str(i) for i in range(1, 7)]
        if trench_combobox.get() not in trench_combobox['values']:
            trench_combobox.set()

def filter_cables(data):
    # Excel'den verileri oku
    df = pd.read_excel("EE374 Project Cable List.xlsx", header=1)  # <-- BU GEREKLİ

############voltage filtre######################
    voltage_input = data["Rated Voltage Level (V)"]

    if not voltage_input.strip():
        tk.messagebox.showerror("Input Error", "Please enter a Rated Voltage Value (V).")
        return

    try:
        user_voltage = float(voltage_input) / 1000
    except ValueError:
        tk.messagebox.showerror("Input Error", "Rated Voltage must be a number.")
        return None, None


    user_voltage = float(data["Rated Voltage Level (V)"])/1000

    if user_voltage <= 1:
        df = df[df['Voltage level'] == "0.6/1 kV"]
    elif 1 < user_voltage <= 6:
        df = df[df['Voltage level'] == "3.6/6 kV"]
    elif 6 < user_voltage <= 10:
        df = df[df['Voltage level'] == "6/10 kV"]
    elif (10 < user_voltage <= 20) and data["Cable Type"] == "Single-core":
        df = df[df['Voltage level'] == "12/20 kV"]

    elif (10 < user_voltage <= 20) and data["Cable Type"] == "Three-core":
        df = df[df['Voltage level'] == "20.3/35 kV"]
    elif 20 < user_voltage <= 35:
        df = df[df['Voltage level'] == "20.3/35 kV"]
    else:
        result_text.insert(tk.END, "No matching cable voltage level found. Change the voltage to a suitable value.\n")
        return None, None

    #print(user_voltage)
    
    ##################trench corrections###################
    try:
        trench_count = int(data["Parallel Circuits"])
        allowed_trench_factors = {
            1: 1.00,
            2: 0.90,
            3: 0.85,
            4: 0.80,
            5: 0.75,
            6: 0.70
        }
            # Single-core için 1 devre = 3 kablo, 2 devre = 6 kablo
        if data["Cable Type"] == "Single-core":
            trench_count = 3 * int(data["Parallel Circuits"])
        else:
            trench_count = int(data["Parallel Circuits"])
        if trench_count in allowed_trench_factors:
            trench_factor = allowed_trench_factors[trench_count]
        else:
            result_text.insert(tk.END, f"Invalid trench cable count: {trench_count}. Please enter one of: {list(allowed_trench_factors.keys())}\n")
            return None, None


    except (ValueError, TypeError):
        result_text.insert(tk.END, "Invalid input for 'Parallel Circuits'. Please enter a numerical value (e.g., 2)!\n")
        return None, None



    ################temperature corrections#############
    try:
        temp = int(data["Temperature (°C)"])
        allowed_temps = {
            5: 1.15,
            10: 1.10,
            15: 1.05,
            20: 1.00,
            25: 0.95,
            30: 0.90,
            35: 0.85,
            40: 0.80
        }

        if temp in allowed_temps:
            temp_factor = allowed_temps[temp]
        else:
            result_text.insert(tk.END, f"Temperature value {temp}°C is not supported. Please enter one of: {list(allowed_temps.keys())}\n")
            return None, None

        #print(temp_factor)
        #print(trench_factor)
    except (ValueError, TypeError):
        result_text.insert(tk.END, "Invalid temperature. Please enter a numerical value (e.g., 20)!\n")
        return None, None


    df["Corrected Capacity"] = df["Current Capacity (A) at 20°C :."] * temp_factor * trench_factor

    # Gerekli akımı hesapla
    active = float(data["Active Power (kW)"]) if data["Active Power (kW)"] else 0
    reactive = float(data["Reactive Power (kVAR)"]) if data["Reactive Power (kVAR)"] else 0
    apparent = (active**2 + reactive**2)**0.5
    
    try:
        voltage = float(data["Rated Voltage Level (V)"])
    except (ValueError, TypeError):
        result_text.insert(tk.END, "Unvalid voltage. Please, enter a numerical value!\n")
        return None, None


    current = (apparent * 1000) / (voltage * (3)**0.5)
    current = current / temp_factor
    
    
    # suitable cables for the calculated current values
    if data["Cable Type"] == "Three-core":
        allowed_ids = list(range(8, 16)) + list(range(24, 32)) + list(range(39, 46)) + list(range(60, 67))
        current = current / (trench_factor * int(data["Parallel Circuits"]))
        df = df[df.index.isin(allowed_ids)]
        
    else:
        # single-core flat yerleşim seçilmişse sadece 2 devreye kadar
        if data["Placement"] == "Flat" and int(data["Parallel Circuits"]) > 2:
            df = df[0:0]  # Tümü elenir
        allowed_ids = list(range(0, 8)) + list(range(16, 24)) + list(range(32, 39)) + list(range(46, 60))
        current = current / (trench_factor * int(data["Parallel Circuits"]))
        df = df[df.index.isin(allowed_ids)]
    print(current)
    #print(list(range(61,67)))

    ##shows current in the table
    if data["Cable Type"] == "Three-core":
        df = df[df["Current Capacity (A) at 20°C :."] > current]
    else: #if single-core
        if data["Placement"] == "Flat":
            df = df[df["Current Capacity (A) at 20°C ..."] > current]
        else:            
            df = df[df["Current Capacity (A) at 20°C :."] > current]


    
    length = int(data["Cable Length (m)"])
    num_of_circuits = int(data.get("Parallel Circuits", 1))
    active_power = int(data["Active Power (kW)"])
    reactive_power = int(data["Reactive Power (kVAR)"])
    rated_voltage= int(data["Rated Voltage Level (V)"])
    # Excel'den gelen uygun kablolar (ohm/km)
    selected_rows = df.values.tolist()  # df zaten filtrelenmiş uygun kablolar

    # 5. sütunda ohm/km değeri olduğunu varsayalım
    resistances_per_km = [row[5] for row in selected_rows]

    # Direnci hesapla: ohm = (ohm/km) * (length / 1000)
    resistances = [r * (length / 1000) for r in resistances_per_km]

    print("rrr:", resistances)
    ####

    # 1. Kesit alanını Cable Code'dan çıkaran fonksiyon (zaten var)
    def extract_phase_cross_section(s):
        match = re.search(r'[xX](\d+)', str(s))
        if match:
            return int(match.group(1))
        return None

    # 2. Kesit alanı sütununu oluştur (zaten yapmışsın)
    df["Phase Cross Section (mm2)"] = df["Cable code"].apply(extract_phase_cross_section)

    # 3. Kesit alanı listesini al
    phase_cross_sections = df["Phase Cross Section (mm2)"].tolist()

    # 4. Endüktans değerlerini seç (placement ve kablo tipine göre)
    if data["Cable Type"] == "Single-core" and data["Placement"] == "Flat":
        inductances_per_km = [row[6] for row in selected_rows]
    else:
        inductances_per_km = [row[7] for row in selected_rows]
    print("inductances_per_km:", inductances_per_km)

    # 5. Endüktansları uzunluk ve devre sayısına göre hesapla (mm cinsinden uzunluk!)
    inductances = [(l * (length / 1000000)) / num_of_circuits for l in inductances_per_km]
    print("lll:", inductances)

    # 6. Sonuçları yazdır
    for code, ind in zip(df["Cable code"], inductances):
        print(f"{code}: {ind:.6f} H·m/mm²")

    # 1. Kesit alanını Cable Code'dan çıkaran fonksiyon
    def extract_phase_cross_section(s):
        match = re.search(r'[xX](\d+)', str(s))
        if match:
            return int(match.group(1))
        return None
    # 2. Cable Code sütunundan kesit alanı sütununu üret
    df["Phase Cross Section (mm2)"] = df["Cable code"].apply(extract_phase_cross_section)
    # 3. Kesit alanı listesini al
    phase_cross_sections = df["Phase Cross Section (mm2)"].tolist()
    # 4. Uzunluğa göre dirençleri hesapla ve kesit alanına böl
    resistances = [(r * (length / 1000)/num_of_circuits) / cs for r, cs in zip(resistances_per_km, phase_cross_sections)]
    # 5. Sonuçları yazdır
    
    for code, res in zip(df["Cable code"], resistances):
        print(f"{code}: {res:.6f} ohm·m/mm²")


###########


##Voltage-reg
    
    scaled_resistances = [res * active_power/(rated_voltage**2) for res in resistances]
    scaled_inductances = [ind * reactive_power/(rated_voltage**2) for ind in inductances]
    # Sonuçları yazdır
    voltage_reg_terms = [
    res + ind for res, ind in zip(scaled_resistances, scaled_inductances)
]
# ... voltage_reg_terms hesaplandıktan sonra
    df["Voltage Regulation"] = voltage_reg_terms

    # Sonuçları yazdır
    for code, vr in zip(df["Cable code"], voltage_reg_terms):
        print(f"{code}: {vr:.6f} (regulation)")

        
###power-calculations
    active_losses = [res * (current**2) for res in resistances]
    reactive_losses = [reac * (current**2) for reac in inductances]
    for code, p_loss, q_loss in zip(df["Cable code"], active_losses, reactive_losses):
        print(f"{code} - Aktif Kayıp: {p_loss:.6f} W, Reaktif Kayıp: {q_loss:.6f} var")


    # Kablosal bilgiye göre doğru sütun seçilir
    if data["Cable Type"] == "Single-core" and data["Placement"] == "Flat":
        capacity_column = "Current Capacity (A) at 20°C ..."
    elif data["Cable Type"] == "Single-core" and data["Placement"] == "Trefoil":
        capacity_column = "Current Capacity (A) at 20°C :."
    else:
        # Three-core kablolar için ortak sütun varsayalım (gerekirse değiştirilir)
        capacity_column = "Current Capacity (A) at 20°C :."
    df = df[df[capacity_column] >= current]

    df["Voltage Regulation"] = voltage_reg_terms
    df["Active_losses"] = active_losses
    df["Reactive_losses"] = reactive_losses
    
    return df, list(zip(df["Cable code"], voltage_reg_terms)), active_losses, reactive_losses

##################################
cable_df = pd.read_excel("EE374 Project Cable List.xlsx", header=1)

def show_all_cables_initial():
    result_text.delete("1.0", tk.END)
    result_text.insert(tk.END, "All available cables:\n\n")
    if not cable_df.empty:
        result_text.insert(tk.END, cable_df[["Cable ID", "Cable code", "Current Capacity (A) at 20°C ...", "Current Capacity (A) at 20°C :."]
                              ].to_string(index=False, justify="left"))
    else:
        result_text.insert(tk.END, "No cable data found.\n")


def on_submit():
    result_text.delete("1.0", tk.END)  # Önceki çıktıyı temizle
    #result_text.configure(font=("Courier New", 10))

    data = {
        "Load Type": load_type_var.get(),
        "Active Power (kW)": active_power_entry.get(),
        "Reactive Power (kVAR)": reactive_power_entry.get(),
        "Temperature (°C)": temp_combobox.get(),
        "Cable Type": cable_type_var.get(),
        "Placement": placement_var.get() if cable_type_var.get() == "Single-core" else None,
        "Parallel Circuits": trench_combobox.get(),
        "Cable Length (m)": length_entry.get(),
        "Rated Voltage Level (V)": voltage_entry.get(),
    }

    result_df, voltage_reg_terms, Active_losses, Reactive_losses = filter_cables(data)
    
    if result_df is None or result_df.empty:
        result_text.insert(tk.END, "No suitable cable! Check the inputs.\n")
        return

    # Sonuç tablosunu yazdır
    if not result_df.empty:
        result_text.insert(tk.END, result_df[["Cable ID", "Cable code", "Active_losses", "Reactive_losses", "Voltage Regulation"]].to_string(index=False))
        #result_text.insert(tk.END, f"{code}: {vr:.2f}% VR, {ploss:.2f} W aktif kayıp, {qloss:.2f} var reaktif kayıp\n")
       # for code, vr in voltage_reg_terms:
       #     result_text.insert(tk.END, f"{code}: {vr:.6f} (regulation)\n")
    else:
        result_text.insert(tk.END, "Uygun kablo bulunamadı.\n")

# Ana pencere
root = tk.Tk()
root.title("Smart Cable Selection")
root.geometry("1200x900")  # Başlangıç boyutu
#root.configure(bg="#f0f4f8")  # Açık gri-mavi
root.configure(bg="#f7f9fc")  # Açık krem-gri arka plan
# Tema dosyasını yükle
root.tk.call("source", os.path.abspath("forest-dark.tcl"))
ttk.Style().theme_use("forest-dark")  # "forest-light" seçeneği de var# Temayı uygula

# Ana pencereyi dinamik hale getirme
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

# Frame'e stil ekleyelim
style = ttk.Style()
style.configure("TFrame", background="#aecf99") #4e7eed mavi
style.configure("TLabel", background="#aecf99", foreground="#2c3e50", font=("Segoe UI", 10))
style.configure("TButton", background="#3498db", foreground="white", font=("Segoe UI", 10, "bold"))
style.map("TButton",
          background=[('active', '#2980b9')],
          foreground=[('active', 'white')])

# Frame
frame = ttk.Frame(root, padding=15, style="TFrame")
frame.grid(sticky="nsew")
frame.configure(style="TFrame")  # Önce stil belirleyelim



# Frame'i dinamik hale getirme
for i in range(12):  # Frame içindeki satırları yapılandır
    frame.rowconfigure(i, weight=1)
frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)

# Diğer kısımlar, arayüz düzeninize dokunulmadan yerinde bırakıldı
# 1. Yük tipi
load_type_var = tk.StringVar()
ttk.Label(frame, text="Load Type:").grid(row=0, column=0, sticky='w')
ttk.Combobox(frame, textvariable=load_type_var, values=["Industrial", "Residential", "Commercial", "Municipal"], width=15).grid(row=0, column=1)

# 2. Aktif güç
ttk.Label(frame, text="Active Power (kW):").grid(row=1, column=0, sticky='w')
active_power_entry = ttk.Entry(frame)
active_power_entry.grid(row=1, column=1)

# 3. Reaktif güç
ttk.Label(frame, text="Reactive Power (kVAR):").grid(row=2, column=0, sticky='w')
reactive_power_entry = ttk.Entry(frame)
reactive_power_entry.grid(row=2, column=1)

# 4. Ortam sıcaklığı
ttk.Label(frame, text="Environment Temperature (°C):").grid(row=3, column=0, sticky='w')

temperature_options = [str(temp) for temp in range(5, 41, 5)]  # 15, 20, 25, 30, 35
temp_combobox = ttk.Combobox(frame, values=temperature_options, state="readonly")
temp_combobox.grid(row=3, column=1)
temp_combobox.current(0)  # Varsayılan olarak 15 seçili


# 5. Kablo tipi
cable_type_var = tk.StringVar()
cable_type_var.set("")
ttk.Label(frame, text="Cable Type:").grid(row=4, column=0, sticky='w')
cable_type_menu = ttk.Combobox(frame, textvariable=cable_type_var, values=["Single-core", "Three-core"], width=15)
cable_type_menu.grid(row=4, column=1)
cable_type_menu.bind("<<ComboboxSelected>>", on_cable_type_change)

# 6. Placement (yalnızca single-core için)
placement_var = tk.StringVar()
placement_label = ttk.Label(frame, text="")
placement_label.grid(row=5, column=0, padx=5, pady=5, sticky='w')
placement_label.grid_remove()  # Başlangıçta gizle

placement_menu = ttk.Combobox(frame, textvariable=placement_var, values=["Flat", "Trefoil"], width=15)
placement_menu.grid(row=4, column=1, padx=5, pady=5, sticky='w')
placement_menu.grid_remove()  # Başlangıçta gizle


# 7. Trench info
ttk.Label(frame, text="Number of Parallel Circuits:").grid(row=6, column=0, sticky='w')
trench_options = [str(i) for i in range(1, 7)]  # 1'den 6'ya kadar seçenekler
trench_combobox = ttk.Combobox(frame, values=trench_options, state="readonly")
trench_combobox.grid(row=6, column=1)
trench_combobox.current(0)  # Varsayılan olarak 1'i seç

# 8. Kablo uzunluğu
ttk.Label(frame, text="Cable Length (m):").grid(row=7, column=0, sticky='w')
length_entry = ttk.Entry(frame)
length_entry.grid(row=7, column=1)

# 9. Gerilim seviyesi
ttk.Label(frame, text="Rated Voltage Level (V):").grid(row=8, column=0, sticky='w')
voltage_entry = ttk.Entry(frame)
voltage_entry.grid(row=8, column=1)

# Submit butonu
submit_button = ttk.Button(frame, text="Filter", command=on_submit)
submit_button.grid(row=10, column=0, columnspan=2, pady=20)

# Çıktı
result_label = ttk.Label(frame, text="")
result_label.grid(row=11, column=0, columnspan=2)
# GUI aynı kalıyor, sade1ce en alta şu widget ekleniyor:

# Alt frame: sadece çıktı ve scrollbar için
result_frame = ttk.Frame(frame)
result_frame.grid(row=12, column=0, columnspan=2, pady=10, sticky="nsew")

# Scrollbar
scrollbar = ttk.Scrollbar(result_frame, orient="vertical")
scrollbar.pack(side="right", fill="y")

# Text widget
result_text = tk.Text(result_frame, height=25, width=100, bg="#e4eddd", fg="#102aad",
                      font=("Segoe UI", 10), yscrollcommand=scrollbar.set)
result_text.pack(side="left", fill="both", expand=True)

# Scrollbar ile Text'i bağla
scrollbar.config(command=result_text.yview)


show_all_cables_initial()
root.mainloop()
