# -*- coding: utf-8 -*-
"""calculos.py -- las fórmulas de los ejercicios de Finanzas, sin ventana.

Cada función recibe los datos ya convertidos a números y devuelve un
`Informe`: el resultado paso a paso, con la fórmula en palabras, la misma
fórmula con los números reemplazados y el resultado. La ventana
(`calculadora.py`) solo pide los datos y muestra el informe.

Las fórmulas replican las del Excel "Ejercicio Presupuestos operacionales
I; II; III; IV; V.xlsx" y las de las presentaciones del curso (costeo directo).
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass

MESES = ("Enero", "Febrero", "Marzo")


# ─────────────────────────── números ───────────────────────────

def num(texto: str) -> float:
    """Lee un número escrito como en Chile ("1.500.000", "0,45") o con punto
    decimal ("0.45"). Un punto seguido de grupos de 3 dígitos es de miles."""
    t = str(texto).strip().replace("$", "").replace("%", "").replace(" ", "")
    if not t:
        raise ValueError("vacío")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?[1-9]\d{0,2}(\.\d{3})+", t):
        t = t.replace(".", "")
    return float(t)


def lista(texto: str) -> list[float]:
    """Varios números separados por ';' o espacios: "25.000; 30.000"."""
    return [num(p) for p in re.split(r"[;\s]+", str(texto).strip()) if p]


def fmt(x: float, dec: int = 2) -> str:
    """1234567.8 -> "1.234.567,80"; los enteros van sin decimales."""
    if abs(x - round(x)) < 1e-9:
        dec = 0
    s = f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "0" if s in ("-0", "-0,00") else s


def singular(unidad: str) -> str:
    """"unidades" -> "unidad", "proyectos" -> "proyecto", "m²" -> "m²"."""
    if unidad.endswith("des"):
        return unidad[:-2]
    return unidad[:-1] if unidad.endswith("s") else unidad


def pct(x: float) -> str:
    return fmt(x, 2) + " %"


# ─────────────────────────── informe ───────────────────────────

class Informe:
    """Líneas de texto con una etiqueta de estilo (titulo, sub, formula,
    resultado, nota) que la ventana usa para dar formato."""

    def __init__(self) -> None:
        self.lineas: list[tuple[str, str | None]] = []

    def titulo(self, t: str) -> None:
        self.lineas += [("", None), (t.upper(), "titulo"), ("═" * len(t), "titulo")]

    def sub(self, t: str) -> None:
        self.lineas += [("", None), (t, "sub"), ("─" * len(t), "sub")]

    def p(self, *textos: str, tag: str | None = None) -> None:
        self.lineas += [(t, tag) for t in textos]

    def nota(self, t: str) -> None:
        self.lineas.append(("  → " + t, "nota"))

    def formula(self, nombre: str, palabras: str, numeros: str, resultado: str) -> None:
        sangria = " " * (len(nombre) + 3)
        self.lineas += [
            (f"  {nombre} = {palabras}", "formula"),
            (f"  {sangria}= {numeros}", "formula"),
            (f"  {sangria}= {resultado}", "resultado"),
        ]

    def tabla(self, cabecera: list[str], filas: list[list[str]]) -> None:
        todas = [cabecera] + filas
        anchos = [max(len(f[i]) for f in todas) for i in range(len(cabecera))]

        def linea(f: list[str]) -> str:
            celdas = [f[0].ljust(anchos[0])] + [c.rjust(a) for c, a in zip(f[1:], anchos[1:])]
            return "  " + "   ".join(celdas)

        self.lineas.append((linea(cabecera), "cabecera"))
        self.lineas.append(("  " + "   ".join("-" * a for a in anchos), None))
        self.lineas += [(linea(f), None) for f in filas]

    def extender(self, otro: "Informe") -> None:
        self.lineas += otro.lineas

    def texto(self) -> str:
        return "\n".join(t for t, _ in self.lineas).strip("\n")


# ──────────────────── presupuesto operativo (I–V) ────────────────────

@dataclass
class DatosPresupuesto:
    nombre: str
    unidad: str
    ventas: list[float]          # unidades por mes (3)
    inv_final: list[float]       # inventario final deseado por mes (3)
    inv_inicial: float           # inventario inicial de enero
    mp: float                    # costo unitario de materia prima / insumos
    mo: float                    # costo unitario de mano de obra directa
    cif: float                   # costo indirecto de fabricación unitario
    gasto_venta: float           # mensual
    comisiones: float            # mensual
    sueldo: float                # mensual
    arriendo: float              # mensual
    servicios: float             # mensual
    depreciacion: float          # mensual
    utilidad_objetivo: float     # utilidad operativa trimestral deseada
    baja_precio_pct: float = 10
    alza_mp_pct: float = 10
    nueva_mo: float = 0
    nueva_depreciacion: float = 0


def _caso(nombre, unidad, ventas, costos, venta, admin, utilidad, nueva_mo, nueva_dep):
    return DatosPresupuesto(
        nombre, unidad, list(ventas), [0, 0, 0], 0, *costos,
        venta, 0, *admin, utilidad, 10, 10, nueva_mo, nueva_dep,
    )


# Datos de las presentaciones (Parte II, láminas 19 a 39).
CASOS: dict[str, DatosPresupuesto] = {
    "I. Panadería": _caso(
        "I. Panadería", "unidades", (5000, 4500, 6000), (0.8, 0.6, 0.4),
        500, (1500, 1000, 200, 100), 950, 0.45, 180),
    "II. Proyecto informático": _caso(
        "II. Proyecto informático", "proyectos", (5, 4, 6), (800, 600, 400),
        500, (1500, 1000, 200, 100), 950, 450, 180),
    "III. Obra de construcción": _caso(
        "III. Obra de construcción", "m²", (500, 450, 600), (80000, 60000, 40000),
        500000, (1500000, 1000000, 200000, 100000), 9500000, 45000, 180000),
    "IV. Taller de electromovilidad": _caso(
        "IV. Taller de electromovilidad", "servicios", (50, 45, 60), (80000, 60000, 40000),
        500000, (1500000, 1000000, 200000, 100000), 9500000, 45000, 180000),
    "V. Mantención industrial": _caso(
        "V. Mantención industrial", "servicios", (50, 45, 60), (80000, 60000, 40000),
        500000, (1500000, 1000000, 200000, 100000), 9500000, 45000, 180000),
}


def _estado(inf: Informe, u: str, q: float, precio: float, cvu: float, cf: float) -> float:
    """Estado de resultados de costeo directo. Devuelve el margen operacional."""
    ingresos = q * precio
    costo = q * cvu
    bruto = ingresos - costo
    operacional = bruto - cf
    inf.tabla(
        ["Concepto", u.capitalize(), "Precio / costo", "Total $"],
        [
            ["Ingresos por venta", fmt(q), fmt(precio), fmt(ingresos)],
            ["(−) Costo de venta", fmt(q), fmt(cvu), fmt(-costo)],
            ["(=) Margen bruto", "", "", fmt(bruto)],
            ["(−) Gastos fijos", "", "", fmt(-cf)],
            ["(=) Margen operacional", "", "", fmt(operacional)],
        ],
    )
    return operacional


def _equilibrio(inf: Informe, u: str, cf: float, precio: float, cvu: float, titulo: str):
    """Escribe Qe y $e. Devuelve (Qe, $e) o (None, None) si no existe."""
    inf.p(f"  {titulo}:", tag="cabecera")
    mc = precio - cvu
    if mc <= 0:
        inf.nota(f"No hay punto de equilibrio: el precio ({fmt(precio)}) no cubre el "
                 f"costo variable ({fmt(cvu)}); cada {singular(u)} vendido/a pierde dinero.")
        return None, None
    qe = cf / mc
    pe = cf / (1 - cvu / precio)
    inf.formula("Qe", "Gastos fijos ÷ (Precio − Costo variable unitario)",
                f"{fmt(cf)} ÷ ({fmt(precio)} − {fmt(cvu)}) = {fmt(cf)} ÷ {fmt(mc)}",
                f"{fmt(qe)} {u}  (≈ {fmt(math.ceil(qe - 1e-9))} {u}, redondeando hacia arriba)")
    inf.formula("$e", "Gastos fijos ÷ (1 − Costo variable unitario ÷ Precio)",
                f"{fmt(cf)} ÷ (1 − {fmt(cvu)} ÷ {fmt(precio)}) = {fmt(cf)} ÷ {fmt(1 - cvu / precio, 4)}",
                f"$ {fmt(pe)} en ventas")
    return qe, pe


def presupuesto_operativo(d: DatosPresupuesto) -> tuple[Informe, dict]:
    u = d.unidad or "unidades"
    det = Informe()

    # Paso 1: ventas
    q = sum(d.ventas)
    det.titulo(f"Caso {d.nombre}" if d.nombre else "Presupuesto operativo")
    det.sub("Paso 1 · Presupuesto de ventas (en " + u + ")")
    det.p("  Las ventas proyectadas son el punto de partida: todo lo demás se calcula a partir de ellas.")
    det.tabla(["Concepto", *MESES, "Trimestre"], [[f"Ventas ({u})", *map(fmt, d.ventas), fmt(q)]])
    det.formula("Ventas del trimestre", "Enero + Febrero + Marzo",
                " + ".join(map(fmt, d.ventas)), f"{fmt(q)} {u}")
    det.nota("El precio todavía no se conoce: se calcula en la pregunta 1.")
    if q <= 0:
        raise ValueError("Las ventas del trimestre deben ser mayores que 0.")

    # Paso 2: producción
    inv_ini = [d.inv_inicial, d.inv_final[0], d.inv_final[1]]
    requerido = [v + f for v, f in zip(d.ventas, d.inv_final)]
    produccion = [r - i for r, i in zip(requerido, inv_ini)]
    qp = sum(produccion)
    det.sub("Paso 2 · Presupuesto de producción")
    det.p("  Unidades a producir = Ventas + Inventario final deseado − Inventario inicial")
    det.tabla(["Concepto", *MESES, "Trimestre"], [
        [f"Ventas proyectadas ({u})", *map(fmt, d.ventas), fmt(q)],
        ["(+) Inventario final deseado", *map(fmt, d.inv_final), fmt(d.inv_final[2])],
        ["(=) Total requerido", *map(fmt, requerido), ""],
        ["(−) Inventario inicial", *map(fmt, inv_ini), fmt(d.inv_inicial)],
        ["(=) Unidades a producir", *map(fmt, produccion), fmt(qp)],
    ])
    if not any(d.inv_final) and not d.inv_inicial:
        det.nota("Sin inventario inicial ni final: se produce exactamente lo que se vende.")
    else:
        det.nota("El inventario inicial de cada mes es el inventario final del mes anterior.")

    # Paso 3: costo de producción
    cvu = d.mp + d.mo + d.cif
    costo_mes = [p * cvu for p in produccion]
    det.sub("Paso 3 · Presupuesto de costo de producción")
    det.formula("Costo unitario", "Materia prima + Mano de obra directa + CIF",
                f"{fmt(d.mp)} + {fmt(d.mo)} + {fmt(d.cif)}", fmt(cvu))
    det.tabla(["Concepto", *MESES, "Trimestre"], [
        [f"Unidades a producir", *map(fmt, produccion), fmt(qp)],
        ["× Costo unitario", *[fmt(cvu)] * 3, fmt(cvu)],
        ["(=) Costo de producción $", *map(fmt, costo_mes), fmt(sum(costo_mes))],
    ])
    det.formula("Costo de producción", "Unidades a producir × Costo unitario",
                f"{fmt(qp)} × {fmt(cvu)}", f"$ {fmt(qp * cvu)}")

    # Paso 4: gastos de venta y administración
    items = [("Gastos de venta / marketing", d.gasto_venta), ("Comisiones", d.comisiones),
             ("Sueldo gerente / administrador", d.sueldo), ("Arriendo", d.arriendo),
             ("Servicios (agua, luz, internet)", d.servicios), ("Depreciación", d.depreciacion)]
    gm = sum(v for _, v in items)
    cf = gm * 3
    det.sub("Paso 4 · Presupuesto de gastos de venta y administración")
    det.p("  Son gastos fijos: se pagan igual, se venda mucho o poco.")
    det.tabla(["Concepto", *MESES, "Trimestre"],
              [[n, *[fmt(v)] * 3, fmt(v * 3)] for n, v in items] +
              [["Total gastos fijos", *[fmt(gm)] * 3, fmt(cf)]])
    det.formula("Gastos fijos del trimestre", "Total mensual × 3", f"{fmt(gm)} × 3", f"$ {fmt(cf)}")

    # Pregunta 1: precio para la utilidad deseada
    cv_total = q * cvu
    precio = (d.utilidad_objetivo + cf + cv_total) / q
    det.sub(f"Pregunta 1 · Precio para obtener una utilidad operativa de $ {fmt(d.utilidad_objetivo)}")
    det.p("  Se busca el precio que cubre el costo variable, los gastos fijos y además deja la utilidad deseada:",
          "  Ingresos necesarios = Costo variable total + Gastos fijos + Utilidad deseada")
    det.formula("Ingresos necesarios", "Unidades × Costo unitario + Gastos fijos + Utilidad",
                f"{fmt(q)} × {fmt(cvu)} + {fmt(cf)} + {fmt(d.utilidad_objetivo)}",
                f"$ {fmt(cv_total + cf + d.utilidad_objetivo)}")
    det.formula("Precio", "Ingresos necesarios ÷ Unidades vendidas",
                f"{fmt(cv_total + cf + d.utilidad_objetivo)} ÷ {fmt(q)}",
                f"$ {fmt(precio)} por {singular(u)}")
    det.p("", "  Estado de resultados con ese precio:")
    uo1 = _estado(det, u, q, precio, cvu, cf)
    det.nota(f"Comprobación: el margen operacional da $ {fmt(uo1)}, que es la utilidad pedida.")

    # Pregunta 2: precio a la baja
    precio2 = precio * (1 - d.baja_precio_pct / 100)
    det.sub(f"Pregunta 2 · ¿Qué pasa si el precio baja un {fmt(d.baja_precio_pct)} %?")
    det.formula("Nuevo precio", f"Precio × (1 − {fmt(d.baja_precio_pct)} %)",
                f"{fmt(precio)} × {fmt(1 - d.baja_precio_pct / 100, 4)}", f"$ {fmt(precio2)}")
    uo2 = _estado(det, u, q, precio2, cvu, cf)
    det.nota(f"La utilidad operacional pasa de $ {fmt(uo1)} a $ {fmt(uo2)} "
             f"(cambia $ {fmt(uo2 - uo1)}). Todo lo que baja el ingreso cae directo a la utilidad, "
             f"porque los costos no cambian.")

    # Pregunta 3: alza de materia prima
    mp3 = d.mp * (1 + d.alza_mp_pct / 100)
    cvu3 = mp3 + d.mo + d.cif
    det.sub(f"Pregunta 3 · ¿Qué pasa si la materia prima sube un {fmt(d.alza_mp_pct)} %? (con el precio de la pregunta 1)")
    det.formula("Nueva materia prima", f"Materia prima × (1 + {fmt(d.alza_mp_pct)} %)",
                f"{fmt(d.mp)} × {fmt(1 + d.alza_mp_pct / 100, 4)}", fmt(mp3))
    det.formula("Nuevo costo unitario", "M.P. nueva + M.O. + CIF",
                f"{fmt(mp3)} + {fmt(d.mo)} + {fmt(d.cif)}", fmt(cvu3))
    uo3 = _estado(det, u, q, precio, cvu3, cf)
    det.nota(f"La utilidad operacional pasa de $ {fmt(uo1)} a $ {fmt(uo3)}: el alza de "
             f"{fmt(mp3 - d.mp)} por unidad × {fmt(q)} {u} = $ {fmt((mp3 - d.mp) * q)} menos.")

    # Pregunta 4: nueva máquina
    cvu4 = d.mp + d.nueva_mo + d.cif
    delta_dep = d.nueva_depreciacion - d.depreciacion
    cf4 = cf + delta_dep * 3
    det.sub("Pregunta 4 · Nueva máquina: baja la mano de obra y sube la depreciación (precio de la pregunta 1)")
    det.formula("Nuevo costo unitario", "M.P. + M.O. nueva + CIF",
                f"{fmt(d.mp)} + {fmt(d.nueva_mo)} + {fmt(d.cif)}", fmt(cvu4))
    det.formula("Nuevos gastos fijos", "Gastos fijos + (Depreciación nueva − Depreciación actual) × 3",
                f"{fmt(cf)} + ({fmt(d.nueva_depreciacion)} − {fmt(d.depreciacion)}) × 3", f"$ {fmt(cf4)}")
    uo4 = _estado(det, u, q, precio, cvu4, cf4)
    ahorro = (d.mo - d.nueva_mo) * q
    det.nota(f"La utilidad operacional queda en $ {fmt(uo4)} (antes $ {fmt(uo1)}). "
             f"Se ahorran $ {fmt(ahorro)} en mano de obra y se gastan $ {fmt(delta_dep * 3)} más en depreciación.")

    # Pregunta 5: punto de equilibrio
    det.sub("Pregunta 5 · Punto de equilibrio inicial y final")
    det.p("  Es el nivel de ventas donde la utilidad operacional es 0: los ingresos cubren justo todos los costos.")
    qe1, pe1 = _equilibrio(det, u, cf, precio, cvu, "Inicial (datos originales, precio de la pregunta 1)")
    qe4, pe4 = _equilibrio(det, u, cf4, precio, cvu4, "Final (con la nueva máquina de la pregunta 4)")
    if qe1 is not None and qe4 is not None:
        sentido = "baja" if qe4 < qe1 else "sube"
        det.nota(f"El punto de equilibrio {sentido} de {fmt(qe1)} a {fmt(qe4)} {u} "
                 f"({fmt(qe4 - qe1)}). " + (
                     "La empresa necesita vender menos para no perder: el margen por unidad crece "
                     f"de {fmt(precio - cvu)} a {fmt(precio - cvu4)} y compensa el mayor gasto fijo."
                     if qe4 < qe1 else
                     "La empresa necesita vender más para no perder: el mayor gasto fijo pesa más "
                     "que la mejora del margen por unidad."))

    # Resumen al principio
    res = Informe()
    res.titulo("Resumen")
    res.tabla(["Escenario", "Precio", "Costo unit.", "Gastos fijos", "Utilidad operac."], [
        ["1. Precio para la meta", fmt(precio), fmt(cvu), fmt(cf), fmt(uo1)],
        [f"2. Precio −{fmt(d.baja_precio_pct)} %", fmt(precio2), fmt(cvu), fmt(cf), fmt(uo2)],
        [f"3. M.P. +{fmt(d.alza_mp_pct)} %", fmt(precio), fmt(cvu3), fmt(cf), fmt(uo3)],
        ["4. Nueva máquina", fmt(precio), fmt(cvu4), fmt(cf4), fmt(uo4)],
    ])
    res.p("")
    res.p(f"  Punto de equilibrio inicial: {fmt(qe1) if qe1 else 'no existe'} {u}"
          + (f"  ·  $ {fmt(pe1)}" if pe1 else ""), tag="resultado")
    res.p(f"  Punto de equilibrio final:   {fmt(qe4) if qe4 else 'no existe'} {u}"
          + (f"  ·  $ {fmt(pe4)}" if pe4 else ""), tag="resultado")
    res.extender(det)

    valores = dict(precio=precio, cvu=cvu, cf=cf, uo1=uo1, precio2=precio2, uo2=uo2,
                   cvu3=cvu3, uo3=uo3, cvu4=cvu4, cf4=cf4, uo4=uo4,
                   qe1=qe1, pe1=pe1, qe4=qe4, pe4=pe4)
    return res, valores


# ─────────────────────────── punto de equilibrio ───────────────────────────

# Ejercicios de la Parte I, láminas 60 a 64.
CASOS_PE = {
    "Ej. 1 · Plastic Ltda.": dict(cf=1500000, p=1500, cv=950, cf2=2000000, p2=None, cv2=650, vol=[5000]),
    "Ej. 2 · Chile Posters": dict(cf=2500, p=10, cv=5, cf2=None, p2=None, cv2=None, vol=[300, 500, 1000, 1500, 30000]),
    "Ej. 3 · Importadora International": dict(cf=3500000, p=2000, cv=1850, cf2=5500000, p2=None, cv2=1600, vol=[]),
}


def punto_equilibrio(cf: float, p: float, cv: float, volumenes: list[float],
                     alternativo: tuple[float, float, float] | None = None) -> tuple[Informe, dict]:
    inf = Informe()
    inf.titulo("Punto de equilibrio")
    inf.p("  El punto de equilibrio es la cantidad que hay que vender para que la utilidad operacional sea 0.",
          "  Cada unidad vendida aporta (Precio − Costo variable) para pagar los costos fijos.")
    escenarios = [("Situación inicial", cf, p, cv)]
    if alternativo:
        escenarios.append(("Situación nueva", *alternativo))
    datos = {}
    for nombre, f, pr, c in escenarios:
        inf.sub(f"{nombre}: CF $ {fmt(f)} · P $ {fmt(pr)} · CV $ {fmt(c)}")
        inf.formula("Margen de contribución", "Precio − Costo variable", f"{fmt(pr)} − {fmt(c)}", f"$ {fmt(pr - c)} por unidad")
        qe, pe = _equilibrio(inf, "unidades", f, pr, c, "Punto de equilibrio")
        datos[nombre] = dict(cf=f, p=pr, cv=c, qe=qe, pe=pe)
        vols = list(volumenes) + ([qe] if qe else [])
        if vols:
            inf.p("", "  Estado de resultados para distintos niveles de venta:")
            filas = []
            for v in vols:
                ing, cvt = v * pr, v * c
                uo = ing - cvt - f
                filas.append([fmt(v) + (" (PE)" if qe and v == qe else ""), fmt(ing), fmt(-cvt),
                              fmt(ing - cvt), fmt(-f), fmt(uo), "ganancia" if uo > 1e-6 else "pérdida" if uo < -1e-6 else "equilibrio"])
            inf.tabla(["Unidades", "Ventas", "Costo var.", "Margen contr.", "Costo fijo", "Utilidad", ""], filas)
            inf.nota("Utilidad = Unidades × (Precio − Costo variable) − Costo fijo. Bajo el PE hay pérdida, sobre el PE ganancia.")
    if alternativo:
        a, b = datos["Situación inicial"], datos["Situación nueva"]
        if a["qe"] and b["qe"]:
            inf.sub("Comparación")
            sentido = "sube" if b["qe"] > a["qe"] else "baja"
            inf.p(f"  El punto de equilibrio {sentido} de {fmt(a['qe'])} a {fmt(b['qe'])} unidades "
                  f"({fmt(b['qe'] - a['qe'])}).")
            # Nivel donde ambas situaciones dan la misma utilidad
            dmc = (b["p"] - b["cv"]) - (a["p"] - a["cv"])
            if abs(dmc) > 1e-12:
                qi = (b["cf"] - a["cf"]) / dmc
                if qi > 0:
                    inf.formula("Punto de indiferencia", "(CF nuevo − CF inicial) ÷ (MC nuevo − MC inicial)",
                                f"({fmt(b['cf'])} − {fmt(a['cf'])}) ÷ ({fmt(b['p'] - b['cv'])} − {fmt(a['p'] - a['cv'])})",
                                f"{fmt(qi)} unidades")
                    mejor = "la nueva" if dmc > 0 else "la inicial"
                    inf.nota(f"Sobre {fmt(qi)} unidades conviene {mejor} situación; bajo esa cantidad, la otra.")
    return inf, datos


# ─────────────────────────── estado de resultados anual ───────────────────────────

CASOS_ER = {
    "Ej. 2 años (Parte I, lámina 53)": dict(
        anios=2, unidades="200.000; 250.000", precio="11.000", mp="4.000", mo="3.700",
        arriendo="100.000.000", sueldos="170.000.000", varios="30.000.000",
        maquina="100.000.000", vida="2", residual="0", no_oper="-180.000.000; 0",
        intereses="95.000.000; 47.500.000", impuesto="37"),
    "Ej. 5 años (Parte I, lámina 55)": dict(
        anios=5, unidades="25.000; 30.000", precio="5", mp="1,1", mo="0,9",
        arriendo="3.000", sueldos="5.000", varios="2.000",
        maquina="1.000", vida="4", residual="0", no_oper="0",
        intereses="15.000; 12.775; 10.217; 7.275; 3.891", impuesto="40"),
}


def _completar(valores: list[float], n: int, relleno: float | None) -> list[float]:
    """Alarga la lista a n años: repite el último valor, o rellena con `relleno`."""
    valores = valores[:n]
    extra = (valores[-1] if valores else 0) if relleno is None else relleno
    return valores + [extra] * (n - len(valores))


def estado_resultados(anios: int, unidades: list[float], precio: list[float], mp: float, mo: float,
                      arriendo: float, sueldos: float, varios: float, maquina: float, vida: float,
                      residual: float, no_oper: list[float], intereses: list[float],
                      impuesto_pct: float) -> tuple[Informe, dict]:
    n = int(anios)
    if n < 1:
        raise ValueError("El número de años debe ser al menos 1.")
    q = _completar(unidades, n, None)
    pr = _completar(precio, n, None)
    no = _completar(no_oper, n, 0)
    it = _completar(intereses, n, 0)
    cvu = mp + mo
    gav = arriendo + sueldos + varios
    dep_anual = (maquina - residual) / vida if vida > 0 else 0
    t = impuesto_pct / 100

    inf = Informe()
    inf.titulo(f"Estado de resultados proyectado a {n} año(s) · costeo directo")
    inf.sub("Cómo se calcula cada partida")
    inf.formula("Costo variable unitario", "Materia prima + Mano de obra", f"{fmt(mp)} + {fmt(mo)}", fmt(cvu))
    inf.formula("Gastos de adm. y ventas (GAV)", "Arriendo + Sueldos + Otros gastos",
                f"{fmt(arriendo)} + {fmt(sueldos)} + {fmt(varios)}", f"$ {fmt(gav)} al año")
    inf.formula("Depreciación anual", "(Valor maquinaria − Valor residual) ÷ Vida útil",
                f"({fmt(maquina)} − {fmt(residual)}) ÷ {fmt(vida)}",
                f"$ {fmt(dep_anual)} al año, durante {fmt(vida)} año(s)")
    inf.p("  Ventas = Unidades × Precio · Costo de ventas = Unidades × Costo variable unitario",
          "  Margen operacional = Margen bruto − GAV − Depreciación",
          "  UAII = Margen operacional ± Otros no operacionales · UAI = UAII − Intereses",
          f"  Impuesto = UAI × {fmt(impuesto_pct)} % (si la UAI es negativa no se paga impuesto)", tag="formula")

    col = {k: [] for k in ("ventas", "costo", "bruto", "gav", "dep", "oper", "noop", "uaii", "int", "uai", "imp", "neta")}
    for i in range(n):
        ventas = q[i] * pr[i]
        costo = q[i] * cvu
        dep = dep_anual if i < vida else 0
        oper = ventas - costo - gav - dep
        uaii = oper + no[i]
        uai = uaii - it[i]
        imp = uai * t if uai > 0 else 0
        for k, v in zip(col, (ventas, -costo, ventas - costo, -gav, -dep, oper, no[i], uaii, -it[i], uai, -imp, uai - imp)):
            col[k].append(v)

    nombres = [("Ventas", "ventas"), ("(−) Costo de ventas", "costo"), ("(=) Margen bruto", "bruto"),
               ("(−) Gastos adm. y ventas", "gav"), ("(−) Depreciación", "dep"),
               ("(=) Margen operacional", "oper"), ("(±) Otros no operacionales", "noop"),
               ("(=) Result. antes int. e imptos.", "uaii"), ("(−) Intereses", "int"),
               ("(=) Result. antes impuestos", "uai"), (f"(−) Impuestos {fmt(impuesto_pct)} %", "imp"),
               ("(=) Resultado del ejercicio", "neta")]
    inf.sub("Estado de resultados")
    inf.tabla(["Concepto", *[f"Año {i + 1}" for i in range(n)]],
              [[f"Unidades vendidas", *map(fmt, q)], ["Precio", *map(fmt, pr)]] +
              [[etq, *map(fmt, col[k])] for etq, k in nombres])

    inf.sub("Detalle del año 1")
    inf.formula("Ventas", "Unidades × Precio", f"{fmt(q[0])} × {fmt(pr[0])}", f"$ {fmt(col['ventas'][0])}")
    inf.formula("Costo de ventas", "Unidades × Costo variable unitario", f"{fmt(q[0])} × {fmt(cvu)}", f"$ {fmt(-col['costo'][0])}")
    inf.formula("Margen operacional", "Ventas − Costo de ventas − GAV − Depreciación",
                f"{fmt(col['ventas'][0])} − {fmt(-col['costo'][0])} − {fmt(gav)} − {fmt(-col['dep'][0])}",
                f"$ {fmt(col['oper'][0])}")
    inf.formula("UAI", "Margen operacional ± No operacionales − Intereses",
                f"{fmt(col['oper'][0])} + ({fmt(no[0])}) − {fmt(it[0])}", f"$ {fmt(col['uai'][0])}")
    if col["uai"][0] > 0:
        inf.formula("Impuesto", f"UAI × {fmt(impuesto_pct)} %", f"{fmt(col['uai'][0])} × {fmt(t, 4)}", f"$ {fmt(-col['imp'][0])}")
    else:
        inf.nota("La UAI del año 1 es negativa (pérdida): no se paga impuesto ese año.")
    inf.formula("Resultado del ejercicio", "UAI − Impuesto",
                f"{fmt(col['uai'][0])} − {fmt(-col['imp'][0])}", f"$ {fmt(col['neta'][0])}")
    return inf, col


# ─────────────────────────── análisis vertical y horizontal ───────────────────────────

EJEMPLO_VH = """Ventas; 2.850.000.000; 2.200.000.000
Costo de ventas; -1.780.000.000; -1.540.000.000
Margen comercial; 1.070.000.000; 660.000.000
Gastos de Adm. y Ventas; -558.000.000; -350.000.000
Margen operacional; 512.000.000; 310.000.000
Otros no operacionales; -300.000.000; -180.000.000
Resul. antes ints. e imptos.; 212.000.000; 130.000.000
Intereses; -138.000.000; -95.000.000
Resul. antes imptos.; 74.000.000; 35.000.000
Impuestos; -70.000.000; -25.000.000
Resultado ejercicio; 4.000.000; 10.000.000"""


def leer_filas(texto: str) -> list[tuple[str, float, float]]:
    filas = []
    for n, linea in enumerate(texto.strip().splitlines(), 1):
        if not linea.strip():
            continue
        partes = [p.strip() for p in re.split(r"[;\t]", linea)]
        if len(partes) < 3:
            raise ValueError(f"Línea {n}: escriba 'Concepto; valor año 1; valor año 2'.")
        try:
            filas.append((partes[0], num(partes[1]), num(partes[2])))
        except ValueError:
            raise ValueError(f"Línea {n}: '{linea.strip()}' tiene un número que no se entiende.")
    if not filas:
        raise ValueError("Escriba al menos una partida.")
    return filas


def analisis_vh(filas: list[tuple[str, float, float]], anio1: str, anio2: str) -> tuple[Informe, list]:
    base1, base2 = filas[0][1], filas[0][2]
    if not base1 or not base2:
        raise ValueError("La primera partida (la base, normalmente Ventas o Total activos) no puede ser 0.")
    inf = Informe()
    inf.titulo("Análisis vertical y horizontal")
    inf.p(f"  Vertical: cada partida ÷ la base ({filas[0][0]}) × 100 → cuánto pesa dentro del total de ese año.",
          f"  Horizontal: (valor {anio2} − valor {anio1}) ÷ valor {anio1} × 100 → cuánto creció o cayó.",
          "  (En costos y gastos, que van en negativo, un % horizontal negativo significa que el gasto bajó.)")
    resultado = []
    for nombre, v1, v2 in filas:
        vert1, vert2 = v1 / base1 * 100, v2 / base2 * 100
        hor = (v2 - v1) / v1 * 100 if v1 else None
        resultado.append((nombre, v1, vert1, v2, vert2, v2 - v1, hor))
    inf.sub("Tabla")
    inf.tabla(["Partida", anio1, "% vert.", anio2, "% vert.", "Variación $", "% horiz."],
              [[n, fmt(v1), pct(a), fmt(v2), pct(b), fmt(dv), pct(h) if h is not None else "n/a"]
               for n, v1, a, v2, b, dv, h in resultado])

    n, v1, a, v2, b, dv, h = resultado[1] if len(resultado) > 1 else resultado[0]
    inf.sub("Ejemplo de cálculo")
    inf.formula(f"% vertical {anio1}", f"{n} ÷ {filas[0][0]} × 100", f"{fmt(v1)} ÷ {fmt(base1)} × 100", pct(a))
    if h is not None:
        inf.formula("% horizontal", f"({anio2} − {anio1}) ÷ {anio1} × 100",
                    f"({fmt(v2)} − ({fmt(v1)})) ÷ ({fmt(v1)}) × 100", pct(h))

    inf.sub("Lectura de los resultados")
    for n, v1, a, v2, b, dv, h in resultado:
        if h is None:
            inf.p(f"  • {n}: en {anio1} era 0, no se puede calcular la variación porcentual.")
            continue
        if (v1 < 0) != (v2 < 0) and v2:
            inf.p(f"  • {n} cambió de signo: pasó de {fmt(v1)} a {fmt(v2)}.")
            continue
        mov = "aumentó" if h > 0 else "disminuyó" if h < 0 else "se mantuvo"
        extra = "" if mov == "se mantuvo" else f" un {pct(abs(h))}"
        peso = "" if n == filas[0][0] else f" y pasó de pesar {pct(abs(a))} a {pct(abs(b))} de {filas[0][0].lower()}"
        inf.p(f"  • {n} {mov}{extra}{peso}.")
    return inf, resultado
