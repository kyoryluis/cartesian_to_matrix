import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
import pandas as pd
import json
from ttkbootstrap import Style


# Función para convertir grados a radianes
def deg_to_rad(deg):
    return np.deg2rad(deg)


# Función para calcular la matriz de rotación a partir de los ángulos de Euler
def rotation_matrix(roll, pitch, yaw):
    roll = deg_to_rad(roll)
    pitch = deg_to_rad(pitch)
    yaw = deg_to_rad(yaw)

    Rx = np.array([
        [1, 0, 0],
        [0, np.cos(roll), -np.sin(roll)],
        [0, np.sin(roll), np.cos(roll)]
    ])

    Ry = np.array([
        [np.cos(pitch), 0, np.sin(pitch)],
        [0, 1, 0],
        [-np.sin(pitch), 0, np.cos(pitch)]
    ])

    Rz = np.array([
        [np.cos(yaw), -np.sin(yaw), 0],
        [np.sin(yaw), np.cos(yaw), 0],
        [0, 0, 1]
    ])

    R = Rz @ Ry @ Rx
    return R


# Función para calcular la matriz de transformación homogénea
def homogeneous_transformation(x, y, z, roll, pitch, yaw):
    R = rotation_matrix(roll, pitch, yaw)
    T = np.eye(4)
    T[:3, :3] = R
    T[:3, 3] = [x, y, z]
    return T


# Función para actualizar la salida en la GUI
def calcular_matriz():
    try:
        x = float(entry_x.get())
        y = float(entry_y.get())
        z = float(entry_z.get())
        roll = float(entry_roll.get())
        pitch = float(entry_pitch.get())
        yaw = float(entry_yaw.get())

        # Calcular la matriz de transformación
        matriz = homogeneous_transformation(x, y, z, roll, pitch, yaw)

        # Formatear la matriz con 3 decimales
        texto_matriz = "\n".join(["\t".join([f"{val:.3f}" for val in fila]) for fila in matriz])

        # Mostrar la matriz en el label
        output_label.config(text=texto_matriz)

    except ValueError:
        output_label.config(text="Por favor, introduce valores numéricos válidos.")


# Función adicional: Generar matrices de Excel y JSON
def generar_matrices_excel_json(input_excel, output_excel, output_json):
    try:
        df = pd.read_excel(input_excel)
        df.columns = df.columns.str.strip()  # Eliminar espacios en blanco
        matrices = []
        json_data = []

        for index, row in df.iterrows():
            try:
                x = row['X']
                y = row['Y']
                z = row['Z']
                roll = row['Roll']
                pitch = row['Pitch']
                yaw = row['Yaw']
                matriz = homogeneous_transformation(x, y, z, roll, pitch, yaw)
                matriz = np.round(matriz, 3)
                matriz_flat = matriz.flatten()
                matrices.append([row['Numero de posición']] + list(matriz_flat))
                matrix_json = [
                    list(matriz[0]),
                    list(matriz[1]),
                    list(matriz[2]),
                    list(matriz[3])
                ]
                json_data.append({
                    "position": str(row['Numero de posición']).zfill(2),
                    "transform_matrix": matrix_json
                })
            except KeyError as e:
                messagebox.showerror("Error", f"Error: {e}. La columna no existe en la fila {index}.")

        columnas = ['Numero de posición'] + [f'M{i + 1}' for i in range(16)]
        df_matrices = pd.DataFrame(matrices, columns=columnas)
        df_matrices.to_excel(output_excel, index=False)
        with open(output_json, 'w') as json_file:
            json.dump(json_data, json_file, indent=4)
        messagebox.showinfo("Éxito", f"Matrices guardadas en {output_excel} y {output_json}")
    except Exception as e:
        messagebox.showerror("Error", f"Error procesando el archivo: {str(e)}")


# Funciones para seleccionar archivos
def seleccionar_archivo_entrada():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entrada_var.set(archivo)


def seleccionar_archivo_salida_excel():
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        salida_excel_var.set(archivo)


def seleccionar_archivo_salida_json():
    archivo = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Archivos JSON", "*.json")])
    if archivo:
        salida_json_var.set(archivo)


def ejecutar_programa():
    input_excel = entrada_var.get()
    output_excel = salida_excel_var.get()
    output_json = salida_json_var.get()
    if input_excel and output_excel and output_json:
        generar_matrices_excel_json(input_excel, output_excel, output_json)
    else:
        messagebox.showwarning("Advertencia", "Debe seleccionar los archivos de entrada, salida Excel y salida JSON.")


# Crear la ventana principal
root = tk.Tk()
root.title("Matriz de Transformación Homogénea")

# Crear el estilo Material Design
style = ttk.Style()
style.theme_use('clam')
style.configure('TFrame', background='#FAFAFA')
style.configure('TLabel', background='#FAFAFA', foreground='#212121', font=('Arial', 12))
style.configure('TButton', background='#6200EA', foreground='white', font=('Arial', 12, 'bold'), padding=10)
style.map('TButton', background=[('active', '#3700B3')], foreground=[('active', 'white')])
style.configure('TEntry', padding=5, relief='flat', font=('Arial', 12))

# Crear el marco principal
frame = ttk.Frame(root, padding="20")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Etiquetas y campos de entrada
ttk.Label(frame, text="Coordenadas de Traslación").grid(row=0, column=0, columnspan=2, pady=5)
ttk.Label(frame, text="X:").grid(row=1, column=0, sticky=tk.E, pady=5)
entry_x = ttk.Entry(frame)
entry_x.grid(row=1, column=1)

ttk.Label(frame, text="Y:").grid(row=2, column=0, sticky=tk.E, pady=5)
entry_y = ttk.Entry(frame)
entry_y.grid(row=2, column=1)

ttk.Label(frame, text="Z:").grid(row=3, column=0, sticky=tk.E, pady=5)
entry_z = ttk.Entry(frame)
entry_z.grid(row=3, column=1)

ttk.Label(frame, text="Ángulos de Rotación (Grados)").grid(row=4, column=0, columnspan=2, pady=5)
ttk.Label(frame, text="Roll (R):").grid(row=5, column=0, sticky=tk.E, pady=5)
entry_roll = ttk.Entry(frame)
entry_roll.grid(row=5, column=1)

ttk.Label(frame, text="Pitch (P):").grid(row=6, column=0, sticky=tk.E, pady=5)
entry_pitch = ttk.Entry(frame)
entry_pitch.grid(row=6, column=1)

ttk.Label(frame, text="Yaw (W):").grid(row=7, column=0, sticky=tk.E, pady=5)
entry_yaw = ttk.Entry(frame)
entry_yaw.grid(row=7, column=1)

# Botón para calcular la matriz
calculate_button = ttk.Button(frame, text="Calcular Matriz", command=calcular_matriz)
calculate_button.grid(row=8, column=0, columnspan=2, pady=15)

# Etiqueta de salida
output_label = ttk.Label(frame, text="", background="#EEEEEE", relief="solid", padding=10, anchor="center",
                         font=('Arial', 12))
output_label.grid(row=9, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Botón para generar matrices desde Excel
entrada_var = tk.StringVar()
salida_excel_var = tk.StringVar()
salida_json_var = tk.StringVar()

ttk.Button(frame, text="Seleccionar archivo de entrada (Excel)", command=seleccionar_archivo_entrada).grid(row=10,
                                                                                                           column=0,
                                                                                                           pady=5)
ttk.Button(frame, text="Seleccionar archivo de salida (Excel)", command=seleccionar_archivo_salida_excel).grid(row=11,
                                                                                                               column=0,
                                                                                                               pady=5)
ttk.Button(frame, text="Seleccionar archivo de salida (JSON)", command=seleccionar_archivo_salida_json).grid(row=12,
                                                                                                             column=0,
                                                                                                             pady=5)

ttk.Button(frame, text="Crear archivos", command=ejecutar_programa).grid(row=13, column=0, columnspan=2, pady=15)

root.mainloop()
