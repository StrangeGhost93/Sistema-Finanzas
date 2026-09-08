from datetime import datetime
import sqlite3
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ControlFinanzasApp(ctk.CTk):

  def __init__(self):
    super().__init__()
    self.title("Gestor de Finanzas Personales Pro")
    self.geometry("1050x720")

    self.protocol("WM_DELETE_WINDOW", self.on_closing)

    self.init_db()
    self.create_widgets()
    self.actualizar_todo()

  def init_db(self):
    self.conn = sqlite3.connect("finanzas.db")
    self.cursor = self.conn.cursor()
    self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS transacciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                tipo TEXT,
                categoria TEXT,
                monto REAL,
                descripcion TEXT
            )
        """)
    self.conn.commit()

  def create_widgets(self):
    # Panel Izquierdo: Formulario de Registro, Resumen y Exportación
    frame_form = ctk.CTkFrame(self, width=290)
    frame_form.pack(side="left", fill="y", padx=10, pady=10)

    ctk.CTkLabel(
        frame_form, text="Nueva Transacción", font=("Arial", 18, "bold")
    ).pack(pady=10)

    self.tipo_var = ctk.StringVar(value="Gasto")
    ctk.CTkRadioButton(
        frame_form, text="Gasto", variable=self.tipo_var, value="Gasto"
    ).pack(anchor="w", padx=20, pady=4)
    ctk.CTkRadioButton(
        frame_form, text="Ingreso", variable=self.tipo_var, value="Ingreso"
    ).pack(anchor="w", padx=20, pady=4)

    ctk.CTkLabel(
        frame_form, text="Categoría:", font=("Arial", 12, "bold")
    ).pack(anchor="w", padx=20, pady=(10, 2))
    categorias_predefinidas = [
        "Supermercado",
        "Tecnología",
        "Vivienda",
        "Auto",
        "Servicios",
        "Sueldo",
        "Transferencia",
        "Entretenimiento",
        "Otros",
    ]
    self.combo_categoria = ctk.CTkComboBox(
        frame_form, values=categorias_predefinidas
    )
    self.combo_categoria.pack(fill="x", padx=20, pady=5)
    self.combo_categoria.set("Supermercado")

    self.ent_monto = ctk.CTkEntry(frame_form, placeholder_text="Monto ($)")
    self.ent_monto.pack(fill="x", padx=20, pady=10)

    self.ent_desc = ctk.CTkEntry(
        frame_form, placeholder_text="Descripción (Opcional)"
    )
    self.ent_desc.pack(fill="x", padx=20, pady=10)

    btn_guardar = ctk.CTkButton(
        frame_form, text="Registrar Movimiento", command=self.guardar_transaccion
    )
    btn_guardar.pack(fill="x", padx=20, pady=12)

    # Resumen General de Saldos
    ctk.CTkLabel(
        frame_form, text="Resumen del Período", font=("Arial", 14, "bold")
    ).pack(pady=(10, 5))

    self.lbl_ingresos = ctk.CTkLabel(
        frame_form,
        text="Ingresos: $0.00",
        text_color="#2ecc71",
        font=("Arial", 12),
    )
    self.lbl_ingresos.pack(anchor="w", padx=20, pady=2)

    self.lbl_gastos = ctk.CTkLabel(
        frame_form,
        text="Gastos: $0.00",
        text_color="#e74c3c",
        font=("Arial", 12),
    )
    self.lbl_gastos.pack(anchor="w", padx=20, pady=2)

    self.lbl_balance = ctk.CTkLabel(
        frame_form, text="Balance: $0.00", font=("Arial", 14, "bold")
    )
    self.lbl_balance.pack(anchor="w", padx=20, pady=(5, 10))

    # Botones de Acción
    btn_excel = ctk.CTkButton(
        frame_form,
        text="Exportar a Excel (.xlsx)",
        fg_color="#27ae60",
        hover_color="#1e8449",
        command=self.exportar_excel,
    )
    btn_excel.pack(fill="x", padx=20, pady=(10, 5))

    btn_vaciar = ctk.CTkButton(
        frame_form,
        text="Vaciar Base de Datos",
        fg_color="#c0392b",
        hover_color="#962d22",
        command=self.vaciar_base_datos,
    )
    btn_vaciar.pack(fill="x", padx=20, pady=(5, 10))

    # Panel Derecho: Filtros + Pestañas con Gráficos e Historial
    frame_derecho = ctk.CTkFrame(self)
    frame_derecho.pack(side="right", fill="both", expand=True, padx=10, pady=10)

    # Barra Superior de Filtros
    frame_filtros = ctk.CTkFrame(frame_derecho)
    frame_filtros.pack(fill="x", padx=10, pady=(10, 5))

    ctk.CTkLabel(
        frame_filtros, text="Filtrar Período:", font=("Arial", 12, "bold")
    ).pack(side="left", padx=10)

    self.combo_mes = ctk.CTkComboBox(
        frame_filtros,
        width=130,
        values=[
            "Todos",
            "01 - Enero",
            "02 - Febrero",
            "03 - Marzo",
            "04 - Abril",
            "05 - Mayo",
            "06 - Junio",
            "07 - Julio",
            "08 - Agosto",
            "09 - Septiembre",
            "10 - Octubre",
            "11 - Noviembre",
            "12 - Diciembre",
        ],
        command=lambda e: self.actualizar_todo(),
    )
    self.combo_mes.pack(side="left", padx=5, pady=5)
    self.combo_mes.set("Todos")

    anios_disponibles = [
        str(y) for y in range(datetime.now().year - 2, datetime.now().year + 3)
    ]
    self.combo_anio = ctk.CTkComboBox(
        frame_filtros,
        width=90,
        values=["Todos"] + anios_disponibles,
        command=lambda e: self.actualizar_todo(),
    )
    self.combo_anio.pack(side="left", padx=5, pady=5)
    self.combo_anio.set("Todos")

    # Contenedor de Pestañas
    self.tabview = ctk.CTkTabview(frame_derecho)
    self.tabview.pack(fill="both", expand=True, padx=10, pady=5)

    self.tab_graficos = self.tabview.add("Gráficos")
    self.tab_historial = self.tabview.add("Historial de Transacciones")

    self.setup_tab_historial()

  def setup_tab_historial(self):
    frame_acciones = ctk.CTkFrame(self.tab_historial)
    frame_acciones.pack(fill="x", padx=10, pady=5)

    btn_eliminar = ctk.CTkButton(
        frame_acciones,
        text="Eliminar Registro Seleccionado",
        fg_color="#e67e22",
        hover_color="#d35400",
        command=self.eliminar_seleccionado,
    )
    btn_eliminar.pack(side="left", padx=10, pady=5)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "Treeview",
        background="#2b2b2b",
        foreground="white",
        fieldbackground="#2b2b2b",
        rowheight=25,
    )
    style.configure(
        "Treeview.Heading", background="#1f1f1f", foreground="white"
    )
    style.map("Treeview", background=[("selected", "#1f538d")])

    columnas = ("ID", "Fecha", "Tipo", "Categoría", "Monto", "Descripción")
    self.tree = ttk.Treeview(
        self.tab_historial, columns=columnas, show="headings"
    )

    for col in columnas:
      self.tree.heading(col, text=col)
      if col == "ID":
        self.tree.column(col, width=40, anchor="center")
      elif col in ("Tipo", "Monto"):
        self.tree.column(col, width=90, anchor="center")
      elif col == "Fecha":
        self.tree.column(col, width=130, anchor="center")
      else:
        self.tree.column(col, width=150, anchor="w")

    self.tree.pack(fill="both", expand=True, padx=10, pady=10)

  def obtener_df_filtrado(self):
    df = pd.read_sql_query("SELECT * FROM transacciones", self.conn)
    if df.empty:
      return df

    df["fecha_dt"] = pd.to_datetime(df["fecha"])

    anio_sel = self.combo_anio.get()
    mes_sel = self.combo_mes.get()

    if anio_sel != "Todos":
      df = df[df["fecha_dt"].dt.year == int(anio_sel)]

    if mes_sel != "Todos":
      num_mes = int(mes_sel.split(" - ")[0])
      df = df[df["fecha_dt"].dt.month == num_mes]

    return df

  def guardar_transaccion(self):
    tipo = self.tipo_var.get()
    categoria = self.combo_categoria.get().strip()
    monto_str = self.ent_monto.get().strip()
    desc = self.ent_desc.get().strip()

    if not categoria or not monto_str:
      messagebox.showwarning(
          "Campos incompletos",
          "Por favor selecciona una categoría e ingresa el monto.",
      )
      return

    try:
      monto = float(monto_str)
    except ValueError:
      messagebox.showerror(
          "Error de formato", "El monto debe ser un número válido."
      )
      return

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

    self.cursor.execute(
        """
            INSERT INTO transacciones (fecha, tipo, categoria, monto, descripcion)
            VALUES (?, ?, ?, ?, ?)
        """,
        (fecha, tipo, categoria, monto, desc),
    )
    self.conn.commit()

    self.ent_monto.delete(0, "end")
    self.ent_desc.delete(0, "end")

    self.actualizar_todo()

  def eliminar_seleccionado(self):
    selected_item = self.tree.selection()
    if not selected_item:
      messagebox.showinfo(
          "Atención", "Selecciona una fila de la tabla para eliminar."
      )
      return

    val = self.tree.item(selected_item[0], "values")
    trans_id = val[0]

    if messagebox.askyesno(
        "Confirmar", f"¿Deseas eliminar la transacción ID {trans_id}?"
    ):
      self.cursor.execute(
          "DELETE FROM transacciones WHERE id = ?", (trans_id,)
      )
      self.conn.commit()
      self.actualizar_todo()

  def vaciar_base_datos(self):
    if messagebox.askyesno(
        "Advertencia",
        "¿Estás seguro de que deseas eliminar TODOS los registros? Esta"
        " acción no se puede deshacer.",
    ):
      self.cursor.execute("DELETE FROM transacciones")
      self.conn.commit()
      self.actualizar_todo()

  def exportar_excel(self):
    df = self.obtener_df_filtrado()
    if df.empty:
      messagebox.showinfo(
          "Exportación vacía",
          "No hay transacciones registradas para exportar en el período"
          " seleccionado.",
      )
      return

    if "fecha_dt" in df.columns:
      df = df.drop(columns=["fecha_dt"])

    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[
            ("Archivos de Excel", "*.xlsx"),
            ("Todos los archivos", "*.*"),
        ],
        title="Guardar reporte de finanzas",
    )

    if filepath:
      try:
        df.to_excel(filepath, index=False)
        messagebox.showinfo(
            "Éxito", f"Reporte exportado correctamente en:\n{filepath}"
        )
      except Exception as e:
        messagebox.showerror(
            "Error al exportar", f"No se pudo guardar el archivo:\n{e}"
        )

  def actualizar_todo(self):
    df_filtrado = self.obtener_df_filtrado()
    self.actualizar_historial(df_filtrado)
    self.actualizar_resumen_y_graficos(df_filtrado)

  def actualizar_historial(self, df):
    for item in self.tree.get_children():
      self.tree.delete(item)

    if df.empty:
      return

    for _, row in df.sort_values(by="id", ascending=False).iterrows():
      monto_formateado = f"${row['monto']:,.2f}"
      self.tree.insert(
          "",
          "end",
          values=(
              row["id"],
              row["fecha"],
              row["tipo"],
              row["categoria"],
              monto_formateado,
              row["descripcion"],
          ),
      )

  def actualizar_resumen_y_graficos(self, df):
    for widget in self.tab_graficos.winfo_children():
      widget.destroy()

    if df.empty:
      self.lbl_ingresos.configure(text="Ingresos: $0.00")
      self.lbl_gastos.configure(text="Gastos: $0.00")
      self.lbl_balance.configure(text="Balance: $0.00")
      lbl_vacio = ctk.CTkLabel(
          self.tab_graficos,
          text="No hay registros para mostrar gráficos en este período.",
          font=("Arial", 16),
      )
      lbl_vacio.pack(expand=True)
      return

    ingresos = df[df["tipo"] == "Ingreso"]["monto"].sum()
    gastos = df[df["tipo"] == "Gasto"]["monto"].sum()
    balance = ingresos - gastos

    self.lbl_ingresos.configure(text=f"Ingresos: ${ingresos:,.2f}")
    self.lbl_gastos.configure(text=f"Gastos: ${gastos:,.2f}")
    self.lbl_balance.configure(text=f"Balance: ${balance:,.2f}")

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9, 4.5))
    fig.patch.set_facecolor("#2b2b2b")
    ax1.set_facecolor("#2b2b2b")
    ax2.set_facecolor("#2b2b2b")

    # Gráfico 1: Pie Chart de Gastos
    df_gastos = df[df["tipo"] == "Gasto"]
    if not df_gastos.empty:
      gastos_cat = df_gastos.groupby("categoria")["monto"].sum()
      ax1.pie(
          gastos_cat,
          labels=gastos_cat.index,
          autopct="%1.1f%%",
          startangle=140,
          textprops=dict(color="white"),
      )
      ax1.set_title("Distribución de Gastos", color="white")
    else:
      ax1.text(
          0.5,
          0.5,
          "Sin gastos en este período",
          ha="center",
          va="center",
          color="white",
      )
      ax1.axis("off")

    # Gráfico 2: Bar Chart Ingresos vs Gastos
    categorias = ["Ingresos", "Gastos"]
    totales = [ingresos, gastos]
    colores = ["#2ecc71", "#e74c3c"]

    bars = ax2.bar(categorias, totales, color=colores, width=0.4)
    ax2.set_title("Ingresos vs. Gastos", color="white")
    ax2.tick_params(colors="white")
    ax2.spines["bottom"].set_color("white")
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.spines["left"].set_color("white")

    for bar in bars:
      height = bar.get_height()
      ax2.annotate(
          f"${height:,.0f}",
          xy=(bar.get_x() + bar.get_width() / 2, height),
          xytext=(0, 3),
          textcoords="offset points",
          ha="center",
          va="bottom",
          color="white",
      )

    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=self.tab_graficos)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

  def on_closing(self):
    try:
      plt.close("all")
      if hasattr(self, "conn") and self.conn:
        self.conn.close()
    except Exception:
      pass
    finally:
      self.quit()
      self.destroy()
      sys.exit(0)


if __name__ == "__main__":
  app = ControlFinanzasApp()
  app.mainloop()