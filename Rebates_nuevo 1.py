import pandas as pd
import numpy as np
import re
from datetime import datetime
from datetime import date, timedelta
from decimal import Decimal, getcontext
import tkinter as tk
from tkinter import filedialog, messagebox
getcontext().prec = 28
import os
import sys
from pathlib import Path

# archivos = filedialog.askopenfilenames(title = "Selecciona los archivos")
# print(archivos)

# mes = dict(zip(["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"],
#                ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]))

# mes1 = dict(zip(["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"],
#                 ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]))

# def sel_archivos():
#     archivos = filedialog.askopenfilenames(title = "Selecciona los archivos")
#     logs = []
#     periodo = ["VNA mes actual", "VNA mes pasado", "Remuneracion mes actual", "Aportes mes actual"]
#     actual = []
#     for x in archivos:
#         if re.match(r"vna_(.{3})\d{2}.xlsx", x.split("/")[-1]):
#             res = re.match(r"vna_(.{3})(\d{2}).xlsx", x.split("/")[-1])
#             m = mes.get(res.group(1))
#             a = res.group(2)
#             ma = int(str(a)+ str(m))

#             # messagebox.showinfo("Mes", f"Periodo del proceso: {mes1.get(m)} 20{res.group(2)}")
#             messagebox.showinfo("Mes", f"Periodo del proceso: {ma}")

# sel_archivos()

class Rebates():
    def __init__(self):
        self.vna1 = None
        self.vna2 = None
        self.remu = None
        self.aportes = None
        self.destino = None
        self.tc_manual = None  # TC manual: setear si no hay acceso a la red (ej: Decimal("970"))

        self.homo1, self.homo2 = self.cargar_homo()
    # Funciones relacionadas a la interfaz
    def interfaz(self):
        main = tk.Tk()
        main.geometry("700x500")
        main.title("Rebates")
        main.resizable(False, False)

        self.ar_vna1 = tk.StringVar()
        self.ar_vna2 = tk.StringVar()
        self.ar_remu = tk.StringVar()
        self.ar_aportes = tk.StringVar()
        self.ar_destino = tk.StringVar()

        tk.Label(main, text="Archivo VNA t-1:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=10)
        tk.Label(main, text="Archivo VNA t:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=40)
        tk.Label(main, text="Archivo con remuneraciones:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=70)
        tk.Label(main, text="Archivo con aportes:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=100)
        tk.Label(main, text="Ruta de destino:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=130)
        tk.Label(main, text="* La ruta por defecto es 'Descargas'", font=("Helvetica", 8, "italic"), bg="#ebf5fb").place(x=20*9+10, y=150)

        tk.Entry(main, textvariable=self.ar_vna1, width=60).place(x=20*9+10, y=10)
        tk.Entry(main, textvariable=self.ar_vna2, width=60).place(x=20*9+10, y=40)
        tk.Entry(main, textvariable=self.ar_remu, width=60).place(x=20*9+10, y=70)
        tk.Entry(main, textvariable=self.ar_aportes, width=60).place(x=20*9+10, y=100)
        tk.Entry(main, textvariable=self.ar_destino, width=60).place(x=20*9+10, y=130)

        tk.Button(main, text="...", command=self.sel_vna1).place(x=570, y=10)
        tk.Button(main, text="...", command=self.sel_vna2).place(x=570, y=40)
        tk.Button(main, text="...", command=self.sel_remu).place(x=570, y=70)
        tk.Button(main, text="...", command=self.sel_aportes).place(x=570, y=100)
        tk.Button(main, text="...", command=self.sel_destino).place(x=570, y=130)

        tk.Button(main, text="Generar Archivo con los REBATES", command=self.proceso,
                  bg="green", fg="white", width=40, height=4).place(x=350, y=250, anchor=tk.CENTER)

        main.mainloop()

    def interfaz_aux(self, main):
        self.ar_vna1 = tk.StringVar()
        self.ar_vna2 = tk.StringVar()
        self.ar_remu = tk.StringVar()
        self.ar_aportes = tk.StringVar()
        self.ar_destino = tk.StringVar()

        tk.Label(main, text="Archivo VNA t-1:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=10)
        tk.Label(main, text="Archivo VNA t:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=40)
        tk.Label(main, text="Archivo con remuneraciones:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=70)
        tk.Label(main, text="Archivo con aportes:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=100)
        tk.Label(main, text="Ruta de destino:", font=("bold", 10), width=20, bg="#ebf5fb").place(x=10, y=130)
        tk.Label(main, text="* La ruta por defecto es 'Descargas'", font=("Helvetica", 8, "italic"), bg="#ebf5fb").place(x=20*9+10, y=150)

        tk.Entry(main, textvariable=self.ar_vna1, width=60).place(x=20*9+10, y=10)
        tk.Entry(main, textvariable=self.ar_vna2, width=60).place(x=20*9+10, y=40)
        tk.Entry(main, textvariable=self.ar_remu, width=60).place(x=20*9+10, y=70)
        tk.Entry(main, textvariable=self.ar_aportes, width=60).place(x=20*9+10, y=100)
        tk.Entry(main, textvariable=self.ar_destino, width=60).place(x=20*9+10, y=130)

        tk.Button(main, text="...", command=self.sel_vna1).place(x=570, y=10)
        tk.Button(main, text="...", command=self.sel_vna2).place(x=570, y=40)
        tk.Button(main, text="...", command=self.sel_remu).place(x=570, y=70)
        tk.Button(main, text="...", command=self.sel_aportes).place(x=570, y=100)
        tk.Button(main, text="...", command=self.sel_destino).place(x=570, y=130)

        tk.Button(main, text="Generar Archivo con los REBATES", command=self.proceso,
                  bg="green", fg="white", width=40, height=4).place(x=350, y=250, anchor=tk.CENTER)

    def sel_vna1(self):
        ruta = filedialog.askopenfilename(title= "Seleccione un archivo para VNA 1")
        if ruta:
            self.ar_vna1.set(ruta)
            self.vna1 = ruta

    def sel_vna2(self):
        ruta = filedialog.askopenfilename(title= "Seleccione un archivo para VNA 2")
        if ruta:
            self.ar_vna2.set(ruta)
            self.vna2 = ruta

    def sel_remu(self):
        ruta = filedialog.askopenfilename(title= "Seleccione un archivo de remuneraciones")
        if ruta:
            self.ar_remu.set(ruta)
            self.remu = ruta

    def sel_aportes(self):
        ruta = filedialog.askopenfilename(title= "Seleccione un archivo de aportes")
        if ruta:
            self.ar_aportes.set(ruta)
            self.aportes = ruta

    def sel_destino(self):
        ruta = filedialog.askdirectory(title= "Selecciona una carpeta de destino")
        if ruta:
            self.ar_destino.set(ruta)
            self.destino = ruta

    # Funciones Relacionadas al proceso
    def ffechas(self, fecha):
        x = pd.to_datetime(fecha, format = "%d-%m-%Y")
        base = pd.to_datetime("30-12-1899", format = "%d-%m-%Y")

        return (x - base).days

    def comprobar_ruta(self, vna1 = None, vna2 = None, remu = None, aportes = None):
        try:
            if vna1 is None:
                if self.vna1 is None:
                    raise Exception("Seleccione la ruta de VNA1")
                else:
                    vna1 = self.vna1

            if vna2 is None:
                if self.vna2 is None:
                    raise Exception("Seleccione la ruta de VNA2")
                else:
                    vna2 = self.vna2

            if remu is None:
                if self.remu is None:
                    raise Exception("Seleccione la ruta de remuneraciones")
                else:
                    remu = self.remu

            if aportes is None:
                if self.aportes is None:
                    raise Exception("Seleccione una ruta de aportes")

        except Exception as e:
            messagebox.showerror("Error", f">Ha ocurrido el siguiente error<\n{e}")

    def cargar_homo(self):
        ar = "Z:/TRADICIONALES/REPORTES FFMM-AFI/HOMOLOGACION/TH_REBATES.xlsx"
        try:
            homo1 = pd.read_excel(ar, sheet_name="TH1")
            homo2 = pd.read_excel(ar, sheet_name="TH2")
            return homo1, homo2
        except:
            pass

        # Fallback local: usa TIPO 1.xlsx si no hay conexion a red
        script_dir = os.path.dirname(os.path.abspath(globals().get("__file__", sys.argv[0])))
        ar_local = os.path.join(script_dir, "TIPO 1.xlsx")
        print(f"Buscando archivo en: {ar_local}")
        try:
            tipo = pd.read_excel(ar_local, dtype=str)
            for c in tipo.columns:
                tipo[c] = tipo[c].str.strip()

            homo1 = tipo[["Fondo", "Fondo-Serie", "AFECTO-EXENTO"]].copy()
            homo1.columns = ["completo", "Fondo-Serie", "IVA"]
            homo1 = homo1.dropna(subset=["Fondo-Serie"]).reset_index(drop=True)

            homo2 = tipo[["Cod Realais", "Moneda", "Nombre", "Run"]].copy()
            homo2.columns = ["Codigo", "Moneda", "Nombre", "RUN"]
            homo2 = homo2.drop_duplicates("Codigo").dropna(subset=["Codigo"]).reset_index(drop=True)

            print("AVISO: usando homologacion local (TIPO 1.xlsx). Conectese a la VPN para usar TH_REBATES.xlsx")
            return homo1, homo2
        except:
            raise Exception("No se encontro TH_REBATES.xlsx en red ni TIPO 1.xlsx local")

    def cargar_tc(self, mes:int = None):
        if mes is None:
            raise Exception("Ingrese el numero de mes")

        if not os.path.exists("//TRADICIONALES/REPORTES FFMM-AFI/VALIDACION DIARIA DE CUOTA"):
            print("AVISO: ruta de TC no accesible por UNC, intentando por letra de unidad...")

        tc = None
        for x in ["z:", "y:", "x:", "w:"]:
            try:
                tc = pd.read_excel(f"{x}" + r"\TRADICIONALES\REPORTES FFMM-AFI\VALIDACION DIARIA DE CUOTA\Tipo de cambio.xlsx")
                break
            except:
                continue

        # fallback: buscar localmente junto al script
        if tc is None:
            local_tc = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tipo de cambio.xlsx")
            if os.path.exists(local_tc):
                print(f"AVISO: usando TC local ({local_tc})")
                tc = pd.read_excel(local_tc)

        if tc is None:
            if self.tc_manual is not None:
                print(f"AVISO: usando TC manual = {self.tc_manual}")
                return Decimal(str(self.tc_manual))
            raise Exception("No se pudo cargar el archivo de Tipo de Cambio. Verifique conexion a la red o setee self.tc_manual")

        tc["Fecha"] = [pd.to_datetime(x, dayfirst=True) for x in tc["Fecha"]]
        tc = tc.dropna()
        tc["Mes"] = [x.month for x in tc["Fecha"]]
        tc = tc.loc[tc["Mes"] == mes,:].reset_index(drop = True)

        return Decimal(str(tc.iloc[-1,1]))

    def cargar_vna(self, vna1 = None, vna2 = None):
        # comprobamos la existencia de los archivos
        if vna1 is None:
            if self.vna1 is None:
                raise Exception("Seleccione la ruta de VNA1")
            else:
                vna1 = self.vna1

        if vna2 is None:
            if self.vna2 is None:
                raise Exception("Seleccione la ruta de VNA1")
            else:
                vna2 = self.vna2

        # cargamos los datos
        t1 = pd.read_excel(vna1, dtype = str)
        t2 = pd.read_excel(vna2, dtype = str)
        vna = pd.concat([t1,t2])
        vna = vna.reset_index(drop = True)

        # eliminamos los espacios de mas
        for c in vna.columns.tolist():
            vna[c] = vna[c].str.strip()

        # juntamos ambos archivos y le sumamos un dia a la fecha
        for x in vna["FechaProceso"].unique().tolist():
            vna.loc[vna["FechaProceso"] == x, "FechaProceso"] = (pd.to_datetime(x) + timedelta(days=1)).strftime("%Y-%m-%d")

        # Convertir a datetime para operaciones de fecha
        vna["FechaProceso"] = pd.to_datetime(vna["FechaProceso"])

        # determinamos y filtramos por el mes del medio
        vna = vna.sort_values("FechaProceso", ascending=True)
        vna.insert(0, "Mes", 0)
        vna["Mes"] = vna["FechaProceso"].dt.month
        mes = vna["FechaProceso"].dt.month.unique()[1]
        vna = vna.loc[vna["Mes"] == mes,].reset_index(drop=True)

        # Volver a string para Excel
        vna["FechaProceso"] = vna["FechaProceso"].dt.strftime("%d-%m-%Y")

        vna["a_Rut"] = ""
        vna["a_DV"] = ""
        for i, x in enumerate(vna["RutParticipe"]):
            y = x.split("-")
            vna.loc[i,"a_Rut"] = y[0]
            vna.loc[i, "a_DV"] = y[1]

        # aplicamos formato decimal a la columan de SaldoCuotas
        vna["SaldoCuotas"] = vna["SaldoCuotas"].apply(Decimal)

        vna["a_vna"] = [str(self.ffechas(a)) + str(b)+ str(c) for a, b, c in zip(vna["FechaProceso"],vna["CodigoSerie"], vna["a_Rut"])]

        # reseteamos index
        vna = vna.reset_index(drop = True)

        return vna

    def cargar_remu(self, remu: str = None):
        if remu is None:
            if self.remu is None:
                raise Exception("Seleccione la ruta de VNA1")
            else:
                remu = self.remu

        # importamos
        remu = pd.read_excel(remu, dtype = str)
        for c in remu.columns:
            remu[c] = remu[c].str.strip()

        # realizamos cruces de informacion
        remu["a_FS"] = ""

        # Lookup case-insensitive: evita fallas por diferencias de mayusculas/minusculas
        # entre el archivo de remuneraciones y la tabla de homologacion (ej: "Serie" vs "SERIE")
        homo1_lower = {k.lower(): v for k, v in zip(self.homo1["completo"], self.homo1["Fondo-Serie"]) if pd.notna(k)}

        for x in remu["Fondo"].unique().tolist():
            try:
                remu.loc[remu["Fondo"] == x, "a_FS"] = homo1_lower[x.lower()]
            except:
                print(x)

        for x in remu["Fecha"].unique().tolist():
            remu.loc[remu["Fecha"] == x, "Fecha"] = pd.to_datetime(x).strftime("%d-%m-%Y")

        remu["a_remu"] = [x + str(self.ffechas(y)) for x, y in zip(remu["a_FS"], remu["Fecha"])]
        remu["Patrimonio_Afecto"] = remu["Patrimonio_Afecto"].apply(Decimal)
        remu["Remuneracion"] = remu["Remuneracion"].apply(Decimal)

        remu["a_Fondo"] = [f"F{x[:2]}" if x else "" for x in remu["a_FS"]]

        for x in remu["a_Fondo"].unique().tolist():
            if not x:
                continue
            res = self.homo2.loc[self.homo2["Codigo"] == x, "Moneda"]
            if len(res) > 0:
                remu.loc[remu["a_Fondo"] == x, "Moneda"] = res.iloc[0]

        remu["a_Fecha"] = [pd.to_datetime(x, dayfirst = True) for x in remu["Fecha"]]
        remu["a_Mes"] = [x.month for x in remu["a_Fecha"]]
        mes = max(remu["a_Mes"])

        # tipo de cambio: solo se carga si hay fondos USD con remuneracion
        remu = remu.reset_index(drop = True)
        remu_usd_total = remu.loc[remu["Moneda"] == "USD", "Remuneracion"].apply(float).sum()
        hay_usd = remu_usd_total != 0
        u_tc = self.cargar_tc(mes) if hay_usd else Decimal("1")

        for i, x in enumerate(remu["Moneda"]):
            if x == "USD":
                remu.loc[i, "Remuneracion"] = remu.loc[i, "Remuneracion"] * u_tc
                remu.loc[i, "Patrimonio_Afecto"] = remu.loc[i, "Patrimonio_Afecto"] * u_tc
            else:
                continue

        return remu

    def carga_aportes(self, aportes: str = None):
        if aportes is None:
            if self.aportes is None:
                raise Exception("Seleccione la ruta de VNA1")
            else:
                aportes = self.aportes

        aportes = pd.read_excel(aportes, dtype = str)
        for c in aportes.columns:
            aportes[c] = aportes[c].str.strip()

        aportes = aportes.dropna(subset="Rut_Participe").reset_index(drop = True)
        aportes["Cuotas"] = aportes["Cuotas"].apply(Decimal)

        # aportes.columns.tolist()
        f1 = (aportes["Origen_Mov"] == "INV") & (aportes["Fondo_Madre"].str.contains("F01|F50|F42|F61|F45|F65"))
        f2 = (aportes["Origen_Mov"] == "TRF") & (aportes["Tipo_Mov"] == "I")

        aportes = aportes.loc[f1|f2,].reset_index(drop = True)

        # aportes["Fondo"] = aportes["Fondo"].apply(lambda x: x.strip())

        for x in aportes["Fecha"].unique().tolist():
            aportes.loc[aportes["Fecha"] == x, "Fecha"] = pd.to_datetime(x).strftime("%d-%m-%Y")

        aportes["a_cruce"] = [str(int(a))+str(b).strip()+str(int(c))+str(self.ffechas(d)) for a, b, c, d in zip(aportes["Rut_Participe"], aportes["Fondo"], aportes["Cuenta"], aportes["Fecha"])]

        return aportes

    def proceso(self, vna1=None, vna2=None, remu=None, aportes=None):
        self.comprobar_ruta(vna1, vna2, remu, aportes)

        d_vna     = self.cargar_vna(vna1, vna2)
        d_remu    = self.cargar_remu(remu)
        d_aportes = self.carga_aportes(aportes)

        # Convertir a float para operaciones vectorizadas
        d_vna["SaldoCuotas"]        = d_vna["SaldoCuotas"].apply(float)
        d_remu["Remuneracion"]      = d_remu["Remuneracion"].apply(float)
        d_remu["Patrimonio_Afecto"] = d_remu["Patrimonio_Afecto"].apply(float)
        d_aportes["Cuotas"]         = d_aportes["Cuotas"].apply(float)

        # --- Tabla base desde VNA ---
        run_map = dict(zip(self.homo2["Codigo"], self.homo2["RUN"]))

        entre = pd.DataFrame({
            "Fecha":             d_vna["FechaProceso"].values,
            "Cod.Realais":       ["F" + x.split("-")[0] for x in d_vna["CodigoSerie"].str.strip()],
            "RUN":               "",
            "Nombre Fondo":      d_vna["NombreFondoMadre"].values,
            "MonedaFondo":       d_vna["MonedaFondo"].values,
            "Serie":             d_vna["CodigoSerie"].str.strip().values,
            "Rut Aportante":     d_vna["a_Rut"].values,
            "DV":                d_vna["a_DV"].values,
            "Nombre Aportante":  d_vna["NombreParticipe"].values,
            "Cuenta":            d_vna["Cuenta"].values,
            "SaldoCuotasInicial": d_vna["SaldoCuotas"].values,
            "AportesTipo1":      0.0,
        })
        entre["RUN"] = entre["Cod.Realais"].map(run_map)

        # Clave de cruce con aportes
        entre["a_cruce"] = (
            entre["Rut Aportante"].astype(str).apply(lambda x: str(int(x)))
            + entre["Serie"].str.strip()
            + entre["Cuenta"].astype(str)
            + entre["Fecha"].apply(lambda x: str(self.ffechas(x)))
        )

        # --- Match aportes (vectorizado) ---
        aportes_sum = d_aportes.groupby("a_cruce")["Cuotas"].sum()
        entre["AportesTipo1"] = entre["a_cruce"].map(aportes_sum).fillna(0.0)

        # --- Filas de aportes sin match en VNA ---
        f_aportes = d_aportes[~d_aportes["a_cruce"].isin(set(entre["a_cruce"]))].copy()
        if len(f_aportes) > 0:
            fg = f_aportes.groupby("a_cruce").agg(
                Fecha            =("Fecha",             "first"),
                Nombre_Fondo     =("Nombre_Fondo_Madre","first"),
                MonedaFondo      =("Moneda",            "first"),
                Serie            =("Fondo",             "first"),
                Rut_Aportante    =("Rut_Participe",     "first"),
                DV               =("Dv",                "first"),
                Nombre_Aportante =("Nombre_Partici",    "first"),
                Cuenta           =("Cuenta",            "first"),
                Cuotas           =("Cuotas",            "sum"),
                Cod_Realais      =("Fondo_Madre",       "first"),
            ).reset_index()
            extra = pd.DataFrame({
                "Fecha":             fg["Fecha"].values,
                "Cod.Realais":       fg["Cod_Realais"].values,
                "RUN":               fg["Cod_Realais"].map(run_map).values,
                "Nombre Fondo":      fg["Nombre_Fondo"].values,
                "MonedaFondo":       fg["MonedaFondo"].values,
                "Serie":             fg["Serie"].str.strip().values,
                "Rut Aportante":     fg["Rut_Aportante"].apply(lambda x: str(int(float(x)))).values,
                "DV":                fg["DV"].values,
                "Nombre Aportante":  fg["Nombre_Aportante"].values,
                "Cuenta":            fg["Cuenta"].values,
                "SaldoCuotasInicial": 0.0,
                "AportesTipo1":      fg["Cuotas"].values,
                "a_cruce":           fg["a_cruce"].values,
            })
            entre = pd.concat([entre, extra], ignore_index=True)

        dif_a = entre["AportesTipo1"].sum() - d_aportes["Cuotas"].sum()
        print("Monto de aportes coincide" if abs(dif_a) < 0.01
              else f"Monto de aportes no coincide por {dif_a:.4f}")

        entre = entre.sort_values(["Fecha", "Serie"]).reset_index(drop=True)
        entre["SaldoCuotas"] = entre["SaldoCuotasInicial"] + entre["AportesTipo1"]

        # --- Distribucion de remuneracion y patrimonio (vectorizado) ---
        entre["a_remu"] = (
            entre["Serie"].str.strip()
            + entre["Fecha"].apply(lambda x: str(self.ffechas(x)))
        )

        remu_tot  = d_remu.groupby("a_remu")[["Remuneracion","Patrimonio_Afecto"]].sum().rename(
            columns={"Remuneracion":"t_remu","Patrimonio_Afecto":"t_patri"}).reset_index()
        cuota_tot = entre.groupby("a_remu")["SaldoCuotas"].sum().rename("t_cuota").reset_index()

        entre = entre.merge(remu_tot,  on="a_remu", how="left")
        entre = entre.merge(cuota_tot, on="a_remu", how="left")
        entre[["t_remu","t_patri","t_cuota"]] = entre[["t_remu","t_patri","t_cuota"]].fillna(0.0)

        mask_cuota = entre["t_cuota"] != 0.0
        entre["Remuneración"]     = 0.0
        entre["Patrimonio Afecto"] = 0.0
        entre.loc[mask_cuota, "Remuneración"]      = (
            entre.loc[mask_cuota, "SaldoCuotas"] / entre.loc[mask_cuota, "t_cuota"]
            * entre.loc[mask_cuota, "t_remu"]
        )
        entre.loc[mask_cuota, "Patrimonio Afecto"] = (
            entre.loc[mask_cuota, "SaldoCuotas"] / entre.loc[mask_cuota, "t_cuota"]
            * entre.loc[mask_cuota, "t_patri"]
        )
        entre = entre.drop(columns=["t_remu","t_patri","t_cuota"])

        dif_r = entre["Remuneración"].sum() - d_remu["Remuneracion"].sum()
        dif_p = entre["Patrimonio Afecto"].sum() - d_remu["Patrimonio_Afecto"].sum()
        print("Monto de remuneraciones coincide" if abs(dif_r) < 0.01
              else f"Monto de remuneraciones NO coincide por {dif_r:.4f}")
        print("Monto de Patrimonio coincide" if abs(dif_p) < 0.01
              else f"Monto de Patrimonio NO coincide por {dif_p:.4f}")

        # --- IVA / Neto / Exento (vectorizado) ---
        iva_map = dict(zip(self.homo1["Fondo-Serie"], self.homo1["IVA"]))
        entre["Aplica IVA"] = entre["Serie"].map(iva_map)

        mask_s = entre["Aplica IVA"] == "S"
        mask_n = entre["Aplica IVA"] == "N"

        entre["Neto"]   = 0.0
        entre["Exento"] = 0.0
        entre.loc[mask_s, "Neto"]   = entre.loc[mask_s, "Remuneración"] / 1.19
        entre.loc[mask_n, "Exento"] = entre.loc[mask_n, "Remuneración"]
        entre["TOTAL"] = entre["Remuneración"]
        entre["IVA"]   = 0.0
        entre.loc[mask_s, "IVA"] = entre.loc[mask_s, "TOTAL"] - entre.loc[mask_s, "Neto"]

        t_sum = entre["Neto"].sum() + entre["Exento"].sum() + entre["IVA"].sum()
        dif_iva = entre["Remuneración"].sum() - t_sum
        print("Neto+Exento+IVA coincide" if abs(dif_iva) < 0.01
              else f"Neto+Exento+IVA NO coincide por {dif_iva:.4f}")

        entre["Tipo Series"] = np.where(mask_s, "No APV", "APV")

        # --- Consolidar duplicados (entre_v2) ---
        entre["id"] = (
            entre["Fecha"].apply(lambda x: str(self.ffechas(x)))
            + entre["Serie"]
            + entre["Rut Aportante"].astype(str)
        )

        num_cols = ["SaldoCuotasInicial","AportesTipo1","SaldoCuotas",
                    "Patrimonio Afecto","Remuneración","Neto","Exento","IVA","TOTAL"]
        str_cols = ["Fecha","Cod.Realais","RUN","Nombre Fondo","MonedaFondo",
                    "Serie","Rut Aportante","DV","Nombre Aportante","Tipo Series"]

        entre_v2 = entre.groupby("id", sort=False).agg(
            **{c: (c, "sum")   for c in num_cols},
            **{c: (c, "first") for c in str_cols},
        ).reset_index(drop=True)

        entre_v2 = entre_v2[str_cols + num_cols]
        entre_v2["Remuneración Neta"] = entre_v2["Neto"] + entre_v2["Exento"]
        entre_v2 = entre_v2.rename(columns={"Neto": "Afecto"})

        # --- Columna 01-si-Cp (cartera propia) ---
        entre_v2["01-si-Cp"] = entre_v2["Serie"].str.upper().str.contains("CP", na=False)

        # --- Orden Interno ---
        serie_u  = entre_v2["Serie"].str.upper().str.strip()
        cod_u    = entre_v2["Cod.Realais"].str.upper().str.strip()
        rut_s    = entre_v2["Rut Aportante"].astype(str).str.strip()
        nom_u    = entre_v2["Nombre Fondo"].str.upper().fillna("")
        tipo     = entre_v2["Tipo Series"]

        cond_404 = (
            entre_v2["01-si-Cp"]                                                          # cartera propia (CP en serie)
            | (serie_u.str.startswith("01-SI") & ~serie_u.str.contains("AFP", na=False)) # seguros de vida 01-si
            | (rut_s == "87908100")                                                       # RUT 87908100
            | (cod_u.isin(["F42","F45","F61","F65"]) & (tipo == "No APV"))               # fondos 42/45/61/65 seguros de vida
            | nom_u.str.contains("SURA ASSET", na=False)                                 # Sura Asset completo
            | serie_u.str.startswith("01-AFP")                                            # 01-AFP
            | (cod_u == "F61")                                                            # F61 completo
        )
        cond_420 = ~cond_404 & (rut_s != "76011193")
        cond_406 = ~cond_404 & (rut_s == "76011193") & (tipo == "APV")
        cond_407 = ~cond_404 & (rut_s == "76011193") & (tipo == "No APV")

        entre_v2["Orden Interno"] = np.select(
            [cond_404, cond_420, cond_406, cond_407],
            [404,       420,      406,      407],
            default=0
        )

        # --- Diferencias (vectorizado) ---
        vali = d_remu.groupby(["a_remu","Fecha","a_FS"])[["Remuneracion","Patrimonio_Afecto"]].sum().reset_index()
        vali.columns = ["a_remu","fecha","serie","remu_remu","patri_remu"]
        remu_entre  = entre.groupby("a_remu")["Remuneración"].sum()
        patri_entre = entre.groupby("a_remu")["Patrimonio Afecto"].sum()
        vali["remu_entre"]  = vali["a_remu"].map(remu_entre).fillna(0.0)
        vali["patri_entre"] = vali["a_remu"].map(patri_entre).fillna(0.0)
        vali["dif_remu"]    = (vali["remu_remu"]  - vali["remu_entre"]).astype(float)
        vali["dif_patri"]   = (vali["patri_remu"] - vali["patri_entre"]).astype(float)
        for c in ["remu_remu","patri_remu","remu_entre","patri_entre"]:
            vali[c] = vali[c].astype(float)

        # --- Resumen por Orden Interno ---
        resumen = (
            entre_v2
            .groupby("Orden Interno")[["Afecto","Exento","IVA","TOTAL","Remuneración Neta"]]
            .sum()
            .reset_index()
        )

        ### EXPORTAMOS ###
        ruta = str(Path.home() / "Downloads") + "/Rebates1.xlsx"
        with pd.ExcelWriter(ruta) as w:
            entre_v2.to_excel(w, sheet_name="Entregable",  index=False)
            vali.to_excel(    w, sheet_name="Diferencias", index=False)
            resumen.to_excel( w, sheet_name="Resumen",     index=False)

        return entre_v2


CARPETA = r"C:\Users\Administrador\Desktop\Rebates"

ar1 = CARPETA + r"\12-2025_VNA 1.xlsx"       # VNA mes anterior
ar2 = CARPETA + r"\01-2026_VNA 1.xlsx"       # VNA mes actual
ar3 = CARPETA + r"\01-2026_Original 1.xlsx"  # Remuneraciones
ar4 = CARPETA + r"\01-2026_I.xlsx"           # Aportes

# --- Para correr el proceso directamente sin interfaz ---
# r = Rebates()
# r.tc_manual = 997  # TC CLP/USD de enero 2026 (solo si no hay VPN)
# t = r.proceso(vna1=ar1, vna2=ar2, remu=ar3, aportes=ar4)
# print(t)

# --- Para abrir la interfaz grafica ---
Rebates().interfaz()
