# -*- coding: utf-8 -*-
"""calculadora.py -- ventana para resolver los ejercicios de Finanzas.

    python calculadora.py      (o doble clic en "Calculadora finanzas.bat")

Cuatro pestañas, una por tipo de ejercicio del curso:

1. Presupuesto operativo (casos I a V del Excel): ventas, producción, costos,
   gastos, precio para una utilidad meta, escenarios y punto de equilibrio.
2. Punto de equilibrio: Qe y $e, escenarios de venta y gráfico.
3. Estado de resultados proyectado a varios años (costeo directo).
4. Análisis vertical y horizontal de un estado de resultados o balance.

Cada pestaña trae los datos de los ejercicios de clase para cargarlos con un
clic, y a la derecha muestra el desarrollo: la fórmula en palabras, la fórmula
con los números y el resultado. Las fórmulas viven en `calculos.py`.

Los números se escriben como en Chile ("1.500.000", "0,45"); también sirve
el punto decimal ("0.45"). Enter o F5 calculan.
"""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import calculos as c
import plantilla_excel
from calculos import fmt

CARPETA = Path(__file__).parent

FUENTE = ("Segoe UI", 10)
MONO = ("Consolas", 10)


class Salida(ttk.Frame):
    """Texto de solo lectura donde se muestra el informe, con botón para guardarlo."""

    def __init__(self, padre: tk.Misc, nombre_archivo: str) -> None:
        super().__init__(padre)
        self.nombre_archivo = nombre_archivo
        barra = ttk.Frame(self)
        barra.pack(fill="x")
        ttk.Label(barra, text="Desarrollo y explicación", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Button(barra, text="Guardar como .txt…", command=self.guardar).pack(side="right")
        ttk.Button(barra, text="Copiar", command=self.copiar).pack(side="right", padx=4)
        self.texto = scrolledtext.ScrolledText(self, font=MONO, wrap="none", padx=8, pady=6,
                                               state="disabled", background="#fcfcfc")
        barra_x = ttk.Scrollbar(self, orient="horizontal", command=self.texto.xview)
        self.texto.configure(xscrollcommand=barra_x.set)
        self.texto.pack(fill="both", expand=True, pady=(4, 0))
        barra_x.pack(fill="x")
        self.texto.tag_configure("titulo", font=("Consolas", 12, "bold"), foreground="#1f3b73")
        self.texto.tag_configure("sub", font=("Consolas", 10, "bold"), foreground="#1f3b73")
        self.texto.tag_configure("cabecera", font=("Consolas", 10, "bold"))
        self.texto.tag_configure("formula", foreground="#444444")
        self.texto.tag_configure("resultado", font=("Consolas", 10, "bold"), foreground="#0b6b2e")
        self.texto.tag_configure("nota", foreground="#8a4b00")
        self.informe: c.Informe | None = None

    def mostrar(self, informe: c.Informe) -> None:
        self.informe = informe
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        for linea, tag in informe.lineas:
            self.texto.insert("end", linea + "\n", tag or ())
        self.texto.configure(state="disabled")
        self.texto.yview_moveto(0)

    def copiar(self) -> None:
        if self.informe:
            self.clipboard_clear()
            self.clipboard_append(self.informe.texto())

    def guardar(self) -> None:
        if not self.informe:
            messagebox.showinfo("Guardar", "Primero presione Calcular.")
            return
        ruta = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=self.nombre_archivo,
                                            filetypes=[("Texto", "*.txt")])
        if ruta:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(self.informe.texto() + "\n")


class Formulario(ttk.Frame):
    """Columna de campos etiquetados. Guarda cada Entry por su clave."""

    def __init__(self, padre: tk.Misc) -> None:
        super().__init__(padre, padding=(0, 0, 10, 0))
        self.campos: dict[str, ttk.Entry] = {}
        self.etiquetas: dict[str, str] = {}
        self.fila = 0

    def seccion(self, texto: str) -> None:
        ttk.Label(self, text=texto, font=("Segoe UI", 10, "bold"), foreground="#1f3b73").grid(
            row=self.fila, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.fila += 1

    def cabecera(self, *textos: str) -> None:
        for i, t in enumerate(textos, 1):
            ttk.Label(self, text=t).grid(row=self.fila, column=i)
        self.fila += 1

    def campo(self, clave: str, etiqueta: str, columnas: int = 1, ancho: int = 11) -> None:
        ttk.Label(self, text=etiqueta).grid(row=self.fila, column=0, sticky="w", pady=0)
        claves = [clave] if columnas == 1 else [f"{clave}{i}" for i in range(columnas)]
        for i, k in enumerate(claves):
            e = ttk.Entry(self, width=ancho, justify="right")
            e.grid(row=self.fila, column=1 + i, padx=2, pady=0,
                   columnspan=3 if columnas == 1 and ancho > 13 else 1, sticky="we")
            self.campos[k] = e
            self.etiquetas[k] = etiqueta + (f" ({c.MESES[i]})" if columnas == 3 else "")
        self.fila += 1

    def poner(self, clave: str, valor) -> None:
        e = self.campos[clave]
        e.delete(0, "end")
        e.insert(0, fmt(valor) if isinstance(valor, (int, float)) else str(valor))

    def texto(self, clave: str) -> str:
        return self.campos[clave].get().strip()

    def numero(self, clave: str, vacio: float | None = None) -> float | None:
        t = self.texto(clave)
        if not t and vacio is not None:
            return vacio
        try:
            return c.num(t)
        except ValueError:
            raise ValueError(f"«{self.etiquetas[clave]}»: '{t}' no es un número válido.")

    def numeros(self, clave: str) -> list[float]:
        try:
            return c.lista(self.texto(clave))
        except ValueError:
            raise ValueError(f"«{self.etiquetas[clave]}»: escriba números separados por ';'.")


def pestaña(notebook: ttk.Notebook, titulo: str) -> tuple[ttk.Frame, ttk.Frame]:
    """Crea una pestaña con panel izquierdo (datos) y derecho (resultado)."""
    marco = ttk.Frame(notebook, padding=8)
    notebook.add(marco, text=titulo)
    # Panel izquierdo con barra de desplazamiento: en pantallas bajas el
    # formulario no cabe entero y el botón Calcular quedaría oculto.
    lienzo = tk.Canvas(marco, highlightthickness=0, borderwidth=0)
    barra = ttk.Scrollbar(marco, orient="vertical", command=lienzo.yview)
    lienzo.configure(yscrollcommand=barra.set)
    izq = ttk.Frame(lienzo, padding=(0, 0, 8, 0))
    lienzo.create_window(0, 0, window=izq, anchor="nw")

    def ajustar(_e=None) -> None:
        lienzo.configure(scrollregion=lienzo.bbox("all"), width=izq.winfo_reqwidth())

    def rueda(e) -> None:
        if izq.winfo_reqheight() > lienzo.winfo_height():
            lienzo.yview_scroll(int(-e.delta / 120), "units")

    izq.bind("<Configure>", ajustar)
    lienzo.bind("<Enter>", lambda _e: lienzo.bind_all("<MouseWheel>", rueda))
    lienzo.bind("<Leave>", lambda _e: lienzo.unbind_all("<MouseWheel>"))
    lienzo.pack(side="left", fill="y")
    barra.pack(side="left", fill="y", padx=(0, 8))
    der = ttk.Frame(marco)
    der.pack(side="left", fill="both", expand=True)
    return izq, der


def selector_casos(padre: tk.Misc, casos, al_cargar) -> None:
    fila = ttk.Frame(padre)
    fila.pack(fill="x", pady=(0, 6))
    ttk.Label(fila, text="Ejercicio de clase:").pack(side="left")
    combo = ttk.Combobox(fila, values=list(casos), state="readonly", width=30)
    combo.pack(side="left", padx=4)
    combo.current(0)
    ttk.Button(fila, text="Cargar datos", command=lambda: al_cargar(combo.get())).pack(side="left")


def botonera(padre: tk.Misc, calcular, limpiar=None) -> None:
    fila = ttk.Frame(padre)
    fila.pack(fill="x", pady=(10, 0))
    ttk.Button(fila, text="Calcular  (Enter)", command=calcular).pack(side="left")
    if limpiar:
        ttk.Button(fila, text="Limpiar", command=limpiar).pack(side="left", padx=6)


# ─────────────────────────── pestaña 1 ───────────────────────────

class PestañaPresupuesto:
    def __init__(self, nb: ttk.Notebook) -> None:
        izq, der = pestaña(nb, " 1. Presupuesto operativo (I–V) ")
        selector_casos(izq, c.CASOS, self.cargar)
        f = self.f = Formulario(izq)
        f.pack(fill="x")
        f.campo("nombre", "Nombre del caso", ancho=28)
        f.campo("unidad", "Se venden (unidad)", ancho=28)
        f.seccion("Ventas y producción")
        f.cabecera(*c.MESES)
        f.campo("ventas", "Ventas proyectadas", columnas=3)
        f.campo("inv_final", "Inventario final deseado", columnas=3)
        f.campo("inv_inicial", "Inventario inicial (enero)")
        f.seccion("Costo por unidad")
        f.campo("mp", "Materia prima / insumos (M.P.)")
        f.campo("mo", "Mano de obra directa (M.O.)")
        f.campo("cif", "Costos indirectos (CIF)")
        f.seccion("Gastos fijos mensuales")
        f.campo("gasto_venta", "Gastos de venta / publicidad")
        f.campo("comisiones", "Comisiones")
        f.campo("sueldo", "Sueldo gerente / administrador")
        f.campo("arriendo", "Arriendo")
        f.campo("servicios", "Servicios (agua, luz, internet)")
        f.campo("depreciacion", "Depreciación")
        f.seccion("Preguntas")
        f.campo("utilidad_objetivo", "Utilidad operativa deseada (trim.)")
        f.campo("baja_precio_pct", "Baja del precio (%)")
        f.campo("alza_mp_pct", "Alza de la materia prima (%)")
        f.campo("nueva_mo", "Nueva máquina: M.O. por unidad")
        f.campo("nueva_depreciacion", "Nueva máquina: depreciación mes")
        botonera(izq, self.calcular)
        self.panel_plantilla(izq)
        self.salida = Salida(der, "presupuesto_operativo.txt")
        self.salida.pack(fill="both", expand=True)
        self.cargar(next(iter(c.CASOS)))

    # ── plantilla Excel ──
    def panel_plantilla(self, padre: tk.Misc) -> None:
        marco = ttk.LabelFrame(padre, text=" Rellenar plantilla Excel ", padding=6)
        marco.pack(fill="x", pady=(12, 0))
        self.plantillas = {p.name: p for p in sorted(CARPETA.glob("*.xlsx")) if not p.name.startswith("~$")}
        ttk.Label(marco, text="Plantilla:").grid(row=0, column=0, sticky="w")
        self.combo_plantilla = ttk.Combobox(marco, values=list(self.plantillas), state="readonly", width=34)
        self.combo_plantilla.grid(row=0, column=1, sticky="we", padx=4)
        self.combo_plantilla.bind("<<ComboboxSelected>>", lambda _e: self.leer_hojas())
        ttk.Button(marco, text="Examinar…", command=self.examinar).grid(row=0, column=2)
        ttk.Label(marco, text="Hoja:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.combo_hoja = ttk.Combobox(marco, state="readonly", width=34)
        self.combo_hoja.grid(row=1, column=1, sticky="we", padx=4, pady=(4, 0))
        ttk.Button(marco, text="Rellenar y guardar…", command=self.rellenar).grid(
            row=2, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(marco, text="Escribe los datos del formulario y deja las fórmulas vivas.\n"
                              "La plantilla original no se modifica: se guarda una copia.",
                  foreground="#777777", justify="left").grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))
        if self.plantillas:
            self.combo_plantilla.current(0)
            self.leer_hojas()

    def examinar(self) -> None:
        ruta = filedialog.askopenfilename(initialdir=CARPETA, filetypes=[("Excel", "*.xlsx *.xlsm")])
        if ruta:
            p = Path(ruta)
            self.plantillas[p.name] = p
            self.combo_plantilla.configure(values=list(self.plantillas))
            self.combo_plantilla.set(p.name)
            self.leer_hojas()

    def leer_hojas(self) -> None:
        ruta = self.plantillas.get(self.combo_plantilla.get())
        try:
            nombres = plantilla_excel.hojas(ruta) if ruta else []
        except Exception as e:  # archivo abierto, dañado, etc.
            messagebox.showerror("Plantilla", f"No se pudo leer la plantilla:\n{e}")
            nombres = []
        self.combo_hoja.configure(values=nombres)
        self.elegir_hoja()

    def elegir_hoja(self) -> None:
        """Propone la hoja que corresponde al caso: 'III. Obra…' → 'Ejercicio III'."""
        nombres = list(self.combo_hoja.cget("values") or [])
        romano = self.f.texto("nombre").split(".")[0].strip()
        sugerida = next((h for h in nombres if h.split()[-1] == romano), None)
        if sugerida:
            self.combo_hoja.set(sugerida)
        elif nombres and self.combo_hoja.get() not in nombres:
            self.combo_hoja.set(nombres[0])

    def rellenar(self) -> None:
        ruta = self.plantillas.get(self.combo_plantilla.get())
        hoja = self.combo_hoja.get()
        if not ruta or not hoja:
            messagebox.showinfo("Plantilla", "Elija una plantilla y una hoja.")
            return
        d = self.datos()
        if d is None:
            return
        nombre = f"{ruta.stem} - {d.nombre or hoja}".replace(".", "").replace("/", "-") + ".xlsx"
        destino = filedialog.asksaveasfilename(initialdir=ruta.parent, initialfile=nombre,
                                               defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not destino:
            return
        if Path(destino).resolve() == ruta.resolve():
            messagebox.showerror("Plantilla", "Guarde con otro nombre: la plantilla original no se sobrescribe.")
            return
        try:
            plantilla_excel.rellenar(ruta, hoja, d, destino)
        except PermissionError:
            messagebox.showerror("Plantilla", "No se pudo guardar: el archivo está abierto en Excel. Ciérrelo e intente otra vez.")
            return
        except Exception as e:
            messagebox.showerror("Plantilla", str(e))
            return
        self.calcular()
        if messagebox.askyesno("Plantilla", f"Excel guardado en:\n{destino}\n\n¿Abrirlo ahora?"):
            os.startfile(destino)

    def cargar(self, nombre: str) -> None:
        d = c.CASOS[nombre]
        for clave, valor in vars(d).items():
            if isinstance(valor, list):
                for i, v in enumerate(valor):
                    self.f.poner(f"{clave}{i}", v)
            else:
                self.f.poner(clave, valor)
        if hasattr(self, "combo_hoja"):
            self.elegir_hoja()
        self.calcular()

    def datos(self) -> c.DatosPresupuesto | None:
        f = self.f
        try:
            return c.DatosPresupuesto(
                nombre=f.texto("nombre"), unidad=f.texto("unidad") or "unidades",
                ventas=[f.numero(f"ventas{i}") for i in range(3)],
                inv_final=[f.numero(f"inv_final{i}", 0) for i in range(3)],
                **{k: f.numero(k, 0) for k in (
                    "inv_inicial", "mp", "mo", "cif", "gasto_venta", "comisiones", "sueldo",
                    "arriendo", "servicios", "depreciacion", "utilidad_objetivo",
                    "baja_precio_pct", "alza_mp_pct", "nueva_mo", "nueva_depreciacion")})
        except ValueError as e:
            messagebox.showerror("Revise los datos", str(e))
            return None

    def calcular(self) -> None:
        d = self.datos()
        if d is None:
            return
        try:
            informe, _ = c.presupuesto_operativo(d)
        except ValueError as e:
            messagebox.showerror("Revise los datos", str(e))
            return
        self.salida.mostrar(informe)


# ─────────────────────────── pestaña 2 ───────────────────────────

class PestañaEquilibrio:
    def __init__(self, nb: ttk.Notebook) -> None:
        izq, der = pestaña(nb, " 2. Punto de equilibrio ")
        selector_casos(izq, c.CASOS_PE, self.cargar)
        f = self.f = Formulario(izq)
        f.pack(fill="x")
        f.seccion("Situación inicial")
        f.campo("cf", "Costos fijos (CF)")
        f.campo("p", "Precio de venta (P)")
        f.campo("cv", "Costo variable unitario (CV)")
        f.seccion("Situación nueva (opcional; vacío = igual)")
        f.campo("cf2", "Nuevos costos fijos")
        f.campo("p2", "Nuevo precio")
        f.campo("cv2", "Nuevo costo variable")
        f.seccion("Escenarios de venta")
        f.campo("vol", "Unidades a evaluar", ancho=28)
        ttk.Label(izq, text="Separe con ';'  (ej.: 300; 500; 1.000)", foreground="#777777").pack(anchor="w")
        botonera(izq, self.calcular, self.limpiar)
        self.grafico = tk.Canvas(der, height=260, background="white", highlightthickness=1,
                                 highlightbackground="#cccccc")
        self.grafico.pack(fill="x", pady=(0, 8))
        self.grafico.bind("<Configure>", lambda _e: self.dibujar())
        self.salida = Salida(der, "punto_equilibrio.txt")
        self.salida.pack(fill="both", expand=True)
        self.datos: dict = {}
        self.cargar(next(iter(c.CASOS_PE)))

    def limpiar(self) -> None:
        for k in self.f.campos:
            self.f.poner(k, "")

    def cargar(self, nombre: str) -> None:
        d = c.CASOS_PE[nombre]
        for k in ("cf", "p", "cv", "cf2", "p2", "cv2"):
            self.f.poner(k, "" if d[k] is None else d[k])
        self.f.poner("vol", "; ".join(fmt(v) for v in d["vol"]))
        self.calcular()

    def calcular(self) -> None:
        f = self.f
        try:
            cf, p, cv = f.numero("cf"), f.numero("p"), f.numero("cv")
            alt = None
            if any(f.texto(k) for k in ("cf2", "p2", "cv2")):
                alt = (f.numero("cf2", cf), f.numero("p2", p), f.numero("cv2", cv))
            vol = f.numeros("vol") if f.texto("vol") else []
            informe, self.datos = c.punto_equilibrio(cf, p, cv, vol, alt)
        except ValueError as e:
            messagebox.showerror("Revise los datos", str(e))
            return
        self.salida.mostrar(informe)
        self.dibujar()

    def dibujar(self) -> None:
        """Ingresos y costos totales versus unidades; el cruce es el equilibrio."""
        g = self.grafico
        g.delete("all")
        escenarios = [(n, d) for n, d in self.datos.items() if d["qe"]]
        if not escenarios:
            g.create_text(10, 10, anchor="nw", text="Sin punto de equilibrio que graficar.", font=FUENTE)
            return
        ancho, alto = max(g.winfo_width(), 300), max(g.winfo_height(), 200)
        m_izq, m_der, m_arr, m_abj = 90, 20, 30, 35
        qmax = 2 * max(d["qe"] for _, d in escenarios)
        ymax = max(max(d["p"] * qmax, d["cf"] + d["cv"] * qmax) for _, d in escenarios) * 1.05

        def x(q): return m_izq + q / qmax * (ancho - m_izq - m_der)
        def y(v): return alto - m_abj - v / ymax * (alto - m_arr - m_abj)

        g.create_line(x(0), y(0), x(qmax), y(0), fill="#888888")
        g.create_line(x(0), y(0), x(0), y(ymax), fill="#888888")
        for i in range(5):
            q, v = qmax * i / 4, ymax * i / 4
            g.create_text(x(q), y(0) + 12, text=fmt(round(q)), font=("Segoe UI", 8), fill="#555555")
            g.create_text(x(0) - 6, y(v), text=fmt(round(v)), anchor="e", font=("Segoe UI", 8), fill="#555555")
        g.create_text(ancho - m_der, alto - 6, text="unidades", anchor="e", font=("Segoe UI", 8), fill="#555555")

        colores = {"Situación inicial": ("#1f5fbf", ()), "Situación nueva": ("#c0392b", (5, 3))}
        leyenda_y = 8
        for nombre, d in escenarios:
            color, guion = colores.get(nombre, ("#333333", ()))
            g.create_line(x(0), y(d["cf"]), x(qmax), y(d["cf"] + d["cv"] * qmax), fill=color, width=2, dash=guion)
            g.create_line(x(0), y(0), x(qmax), y(d["p"] * qmax), fill="#2e8b57", width=2, dash=guion)
            g.create_line(x(0), y(d["cf"]), x(qmax), y(d["cf"]), fill=color, dash=(2, 3))
            px, py = x(d["qe"]), y(d["pe"])
            g.create_oval(px - 4, py - 4, px + 4, py + 4, fill=color, outline="")
            arriba = nombre == "Situación inicial"
            g.create_text(px - 6 if arriba else px + 6, py - 8 if arriba else py + 8,
                          anchor="se" if arriba else "nw", fill=color, font=("Segoe UI", 8, "bold"),
                          text=f"PE {fmt(d['qe'])} u · $ {fmt(d['pe'])}")
            g.create_text(ancho - m_der, leyenda_y, anchor="ne", fill=color, font=("Segoe UI", 8),
                          text=f"{nombre}: costo total (—) y costo fijo (···)")
            leyenda_y += 14
        g.create_text(ancho - m_der, leyenda_y, anchor="ne", fill="#2e8b57", font=("Segoe UI", 8),
                      text="Ingresos por venta (verde)")


# ─────────────────────────── pestaña 3 ───────────────────────────

class PestañaResultados:
    def __init__(self, nb: ttk.Notebook) -> None:
        izq, der = pestaña(nb, " 3. Estado de resultados ")
        selector_casos(izq, c.CASOS_ER, self.cargar)
        f = self.f = Formulario(izq)
        f.pack(fill="x")
        f.campo("anios", "Número de años", ancho=6)
        f.seccion("Ventas (separe los años con ';')")
        f.campo("unidades", "Unidades por año", ancho=28)
        f.campo("precio", "Precio por unidad", ancho=28)
        ttk.Label(f, text="Si pone menos valores que años, se repite el último.", foreground="#777777").grid(
            row=f.fila, column=0, columnspan=4, sticky="w")
        f.fila += 1
        f.seccion("Costos variables por unidad")
        f.campo("mp", "Materia prima")
        f.campo("mo", "Mano de obra")
        f.seccion("Gastos fijos anuales")
        f.campo("arriendo", "Arriendo")
        f.campo("sueldos", "Sueldos")
        f.campo("varios", "Consumos / gastos varios")
        f.seccion("Depreciación")
        f.campo("maquina", "Valor de la maquinaria")
        f.campo("vida", "Vida útil (años)")
        f.campo("residual", "Valor residual")
        f.seccion("No operacionales, deuda e impuesto (por año, con ';')")
        f.campo("no_oper", "Otros no operacionales", ancho=28)
        ttk.Label(f, text="Ganancia en positivo, pérdida en negativo (ej.: -180.000.000; 0)",
                  foreground="#777777").grid(row=f.fila, column=0, columnspan=4, sticky="w")
        f.fila += 1
        f.campo("intereses", "Intereses", ancho=28)
        f.campo("impuesto", "Tasa de impuesto (%)")
        botonera(izq, self.calcular)
        self.salida = Salida(der, "estado_resultados.txt")
        self.salida.pack(fill="both", expand=True)
        self.cargar(next(iter(c.CASOS_ER)))

    def cargar(self, nombre: str) -> None:
        for k, v in c.CASOS_ER[nombre].items():
            self.f.poner(k, v)
        self.calcular()

    def calcular(self) -> None:
        f = self.f
        try:
            informe, _ = c.estado_resultados(
                anios=int(f.numero("anios")), unidades=f.numeros("unidades"), precio=f.numeros("precio"),
                mp=f.numero("mp", 0), mo=f.numero("mo", 0), arriendo=f.numero("arriendo", 0),
                sueldos=f.numero("sueldos", 0), varios=f.numero("varios", 0),
                maquina=f.numero("maquina", 0), vida=f.numero("vida", 0), residual=f.numero("residual", 0),
                no_oper=f.numeros("no_oper") if f.texto("no_oper") else [],
                intereses=f.numeros("intereses") if f.texto("intereses") else [],
                impuesto_pct=f.numero("impuesto", 0))
        except ValueError as e:
            messagebox.showerror("Revise los datos", str(e))
            return
        self.salida.mostrar(informe)


# ─────────────────────────── pestaña 4 ───────────────────────────

class PestañaVH:
    def __init__(self, nb: ttk.Notebook) -> None:
        izq, der = pestaña(nb, " 4. Análisis vertical y horizontal ")
        f = self.f = Formulario(izq)
        f.pack(fill="x")
        f.campo("anio1", "Año 1 (el más antiguo)")
        f.campo("anio2", "Año 2")
        ttk.Label(izq, text="Partidas: una por línea →  Concepto; valor año 1; valor año 2\n"
                            "La primera línea es la base del análisis vertical\n"
                            "(Ventas, o Total activos si es un balance).",
                  foreground="#555555", justify="left").pack(anchor="w", pady=(8, 2))
        self.partidas = tk.Text(izq, width=58, height=20, font=MONO, undo=True)
        self.partidas.pack(fill="both", expand=True)
        fila = ttk.Frame(izq)
        fila.pack(fill="x", pady=(10, 0))
        ttk.Button(fila, text="Calcular  (F5)", command=self.calcular).pack(side="left")
        ttk.Button(fila, text="Cargar ejemplo 2013–2014", command=self.ejemplo).pack(side="left", padx=6)
        self.salida = Salida(der, "analisis_vertical_horizontal.txt")
        self.salida.pack(fill="both", expand=True)
        self.ejemplo()

    def ejemplo(self) -> None:
        self.f.poner("anio1", "2013")
        self.f.poner("anio2", "2014")
        self.partidas.delete("1.0", "end")
        self.partidas.insert("1.0", c.EJEMPLO_VH)
        self.calcular()

    def calcular(self) -> None:
        try:
            filas = c.leer_filas(self.partidas.get("1.0", "end"))
            informe, _ = c.analisis_vh(filas, self.f.texto("anio1") or "Año 1", self.f.texto("anio2") or "Año 2")
        except ValueError as e:
            messagebox.showerror("Revise los datos", str(e))
            return
        self.salida.mostrar(informe)


# ─────────────────────────── ventana ───────────────────────────

def main() -> None:
    raiz = tk.Tk()
    raiz.title("Calculadora de Finanzas · presupuestos, equilibrio y estados de resultado")
    raiz.geometry("1300x800")
    raiz.minsize(1000, 600)
    try:
        raiz.state("zoomed")  # maximizada en Windows
    except tk.TclError:
        pass
    raiz.option_add("*Font", FUENTE)
    nb = ttk.Notebook(raiz)
    nb.pack(fill="both", expand=True, padx=6, pady=6)
    pestañas = [PestañaPresupuesto(nb), PestañaEquilibrio(nb), PestañaResultados(nb), PestañaVH(nb)]

    def calcular_actual(evento=None) -> None:
        # En el cuadro de partidas, Enter es un salto de línea; ahí se calcula con F5.
        if evento is not None and evento.keysym == "Return" and isinstance(evento.widget, tk.Text):
            return
        pestañas[nb.index(nb.select())].calcular()

    raiz.bind("<Return>", calcular_actual)
    raiz.bind("<F5>", calcular_actual)
    raiz.mainloop()


if __name__ == "__main__":
    main()
