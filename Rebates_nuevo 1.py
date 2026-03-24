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

    def proceso(self, vna1 = None, vna2 = None, remu = None, aportes = None):
        # comprobamos los archivos
        self.comprobar_ruta(vna1, vna2, remu, aportes)

        # cargamos vna
        d_vna = self.cargar_vna(vna1, vna2)

        # cargamos remuneraciones
        d_remu = self.cargar_remu(remu)

        # cargamos aportes
        d_aportes = self.carga_aportes(aportes)

        # Generamos el entregable
        entre = pd.DataFrame(columns = ['Fecha', 'Cod.Realais', 'RUN', 'Nombre Fondo', 'MonedaFondo', 'Serie', 'Rut Aportante', 'DV', 'Nombre Aportante', 'Cuenta', 'SaldoCuotasInicial', 'AportesTipo1', 'SaldoCuotas', 'Remuneración', 'Patrimonio Afecto', 'Aplica IVA', 'Neto', 'Exento', 'IVA', 'TOTAL', 'Tipo Series'])

        entre["SaldoCuotasInicial"] = d_vna["SaldoCuotas"].apply(Decimal)
        entre["Fecha"] = d_vna["FechaProceso"]
        entre["Nombre Fondo"] = d_vna["NombreFondoMadre"]
        entre["MonedaFondo"] = d_vna["MonedaFondo"]
        entre["Serie"] = [x.strip() for x in d_vna["CodigoSerie"]]
        entre["Rut Aportante"] = d_vna["a_Rut"]
        entre["DV"] = d_vna["a_DV"]
        entre["Nombre Aportante"] = d_vna["NombreParticipe"]
        entre["Cuenta"] = d_vna["Cuenta"]

        # agregamos Cod.Realais y RUN, con las tablas de homologacion
        for x in entre["Serie"].unique().tolist():
            x1 = f"F{x.split('-')[0]}"
            try:
                entre.loc[entre["Serie"] == x, "Cod.Realais"] = x1
            except:
                print(x)

        for x in entre["Cod.Realais"].unique().tolist():
            try:
                entre.loc[entre["Cod.Realais"] == x, "RUN"] = self.homo2.loc[self.homo2["Codigo"] == x, "RUN"].iloc[0]
            except:
                print(x)

        # agregamos el ID para cruzar los aportes
        entre["a_cruce"] = [str(int(a)) + str(b).strip() + str(c) + str(self.ffechas(d)) for a, b, c, d in zip(entre["Rut Aportante"], entre["Serie"], entre["Cuenta"], entre["Fecha"])]
        entre = entre.reset_index(drop = True)

        for i, x in enumerate(entre["a_cruce"]):
            try:
                entre.loc[i, "AportesTipo1"] = d_aportes.loc[d_aportes["a_cruce"] == x, "Cuotas"].apply(Decimal).sum()
            except:
                entre.loc[i, "AportesTipo1"] = "0"

        f_aportes = d_aportes.loc[[i for i, x in enumerate(d_aportes["a_cruce"]) if not x in entre["a_cruce"].tolist()],].reset_index(drop = True)

        # agregamos los aportes faltantes al entregable
        for i in range(0, len(f_aportes)):
            j = len(entre)
            entre.loc[j, "Fecha"] = f_aportes.loc[i, "Fecha"] #
            entre.loc[j, "Nombre Fondo"] = f_aportes.loc[i, "Nombre_Fondo_Madre"]#
            entre.loc[j, "MonedaFondo"] = f_aportes.loc[i, "Moneda"] #
            entre.loc[j, "Serie"] = f_aportes.loc[i, "Fondo"].strip() #
            entre.loc[j, "Rut Aportante"] = str(int(f_aportes.loc[i, "Rut_Participe"])) #
            entre.loc[j, "DV"] = f_aportes.loc[i, "Dv"] #
            entre.loc[j, "Nombre Aportante"] = f_aportes.loc[i, "Nombre_Partici"]
            entre.loc[j, "Cuenta"] = f_aportes.loc[i, "Cuenta"] #
            entre.loc[j, "AportesTipo1"] = f_aportes.loc[i, "Cuotas"]
            entre.loc[j, "SaldoCuotasInicial"] = Decimal("0") #

            # agregamos Cod.Realais y RUN, con las tablas de homologacion
            try:
                # Cod.Realais
                entre.loc[j, "Cod.Realais"] = f_aportes.loc[i, "Fondo_Madre"]
                # RUN
                entre.loc[j, "RUN"] = self.homo2.loc[self.homo2["Codigo"] == entre.loc[j, "Cod.Realais"], "RUN"].iloc[0]
            except:
                print(f"No se encontro {entre.loc[j, 'Nombre Fondo']}")

        # cambiamos de formato a las columnas de cuotas y aportes
        entre["SaldoCuotasInicial"] = entre["SaldoCuotasInicial"].apply(Decimal)
        entre["AportesTipo1"] = entre["AportesTipo1"].apply(Decimal)

        # ordenamos las data por fecha y serie
        entre = entre.sort_values(["Fecha", "Serie"], ascending = [True, True]).reset_index(drop = True)

        # sumamos los cuotas con los aportes
        entre["SaldoCuotas"] = entre["SaldoCuotasInicial"] + entre["AportesTipo1"]

        # creamos ID para distribuir la remuneracion y el patrimonio
        entre["a_remu"] = [str(a).strip()+str(self.ffechas(b)) for a, b in zip(entre["Serie"], entre["Fecha"])]

        for i, x in enumerate(entre["a_remu"]):
            t_remu = d_remu.loc[d_remu["a_remu"] == x, "Remuneracion"].sum()
            t_patri = d_remu.loc[d_remu["a_remu"] == x, "Patrimonio_Afecto"].sum()
            t_cuota = entre.loc[entre["a_remu"] == x, "SaldoCuotas"].sum()

            cuotas = entre.loc[i, "SaldoCuotas"]

            if t_cuota != 0:
                entre.loc[i, "Remuneración"] = (cuotas/t_cuota)*t_remu
                entre.loc[i, "Patrimonio Afecto"] = (cuotas/t_cuota)*t_patri
            else:
                entre.loc[i, "Remuneración"] = Decimal("0")
                entre.loc[i, "Patrimonio Afecto"] = Decimal("0")

        ### COMPROBAMOS LAS DISTRIBUCIONES ###
        if entre["AportesTipo1"].sum() == d_aportes["Cuotas"].sum():
            print("Monto de aportes coincide")
        else:
            dif = entre["AportesTipo1"].sum() - d_aportes["Cuotas"].sum()
            print(f"Monto de aportes no coincide por {dif}")

        if entre["Remuneración"].sum() == d_remu["Remuneracion"].sum():
            print("Monto de remuneraciones coincide")
        else:
            dif = entre["Remuneración"].sum() - d_remu["Remuneracion"].sum()
            print(f"Monto de remuneraciones NO coincide por {dif}")


        if entre["Patrimonio Afecto"].sum() == d_remu["Patrimonio_Afecto"].sum():
            print("Monto de Patrimonio coincide")
        else:
            dif = entre["Patrimonio Afecto"].sum() - d_remu["Patrimonio_Afecto"].sum()
            print(f"Monto de Patrimonio NO coincide por {dif}")

        # Agregamos si aplica IVA
        for x in entre["Serie"].unique().tolist():
            try:
                entre.loc[entre["Serie"] == x, "Aplica IVA"] = self.homo1.loc[self.homo1["Fondo-Serie"] == x, "IVA"].iloc[0]
            except:
                print(x)

        # asignamos a las columnas Neto Exento IVA y Total el formato Decimal
        entre["Neto"] = Decimal("0")
        entre["Exento"] = Decimal("0")
        entre["IVA"] = Decimal("0")
        entre["TOTAL"] = Decimal("0")

        # Calculamos el Neto
        for i, x in enumerate(entre["Aplica IVA"]):
            if x == "S":
                entre.loc[i, "Neto"] = Decimal(str(entre.loc[i, "Remuneración"])) / Decimal("1.19")
            else:
                entre.loc[i, "Neto"] = Decimal("0")

        # Calculamos el exento
        for i, x in enumerate(entre["Aplica IVA"]):
            if x == "S":
                entre.loc[i, "Exento"] = Decimal("0")
            elif x == "N":
                entre.loc[i, "Exento"] = entre.loc[i, "Remuneración"]

        # agregamos el total
        entre["TOTAL"] = entre["Remuneración"]

        # Agregamos el monto del IVA
        for i in range(0, len(entre)):
            if entre.loc[i, "Neto"] != Decimal("0"):
                entre.loc[i, "IVA"] = entre.loc[i, "TOTAL"] - entre.loc[i, "Neto"]
            else:
                entre.loc[i, "IVA"] = Decimal("0")

        ### COMPROBAMOS QUE EL IVA SE HAYA DISTRIBUIDO CORRECTAMENTE ###
        t_sum = entre["Neto"].sum() + entre["Exento"].sum() + entre["IVA"].sum()

        if entre["Remuneración"].sum() == t_sum:
            print("Monto de remuneracion coincide con neto, exento e IVA")
        else:
            dif = entre["Remuneración"].sum() - t_sum
            print(f"Monto de remuneracion coincide con neto, exento e IVA NO coincide por {dif}")

        # Identificamos si son series APV o NO-APV
        entre["Tipo Series"] = ["No APV" if x == "S" else "APV" for x in entre["Aplica IVA"]]

        ### ENTREGABLE V2 ###
        entre["id"] = [str(self.ffechas(a)) + str(b) + str(c) for a, b, c in zip(entre["Fecha"], entre["Serie"], entre["Rut Aportante"])]

        # creamos un nuevo dataframe para dejar solo una observacion diaria por cliente
        entre_v2 = pd.DataFrame(columns = ['Fecha', 'Cod.Realais', 'RUN', 'Nombre Fondo', 'MonedaFondo', 'Serie', 'Rut Aportante', 'DV', 'Nombre Aportante', 'Tipo Series', 'SaldoCuotasInicial', 'AportesTipo1', 'SaldoCuotas', 'Patrimonio Afecto', 'Remuneración', 'Neto', 'Exento', 'IVA', 'TOTAL'])


        for x in entre["id"].unique().tolist():
            j = len(entre_v2)
            for c in entre_v2.columns:
                if c in ['SaldoCuotasInicial', 'AportesTipo1', 'SaldoCuotas', 'Patrimonio Afecto', 'Remuneración', 'Neto', 'Exento', 'IVA', 'TOTAL']:
                    entre_v2.loc[j, c] = entre.loc[entre["id"] == x, c].sum()
                else:
                    entre_v2.loc[j, c] = entre.loc[entre["id"] == x, c].iloc[0]


        # agregamos columnas de formato
        entre_v2["Remuneración Neta"] = entre_v2["Neto"] + entre_v2["Exento"]
        entre_v2 = entre_v2.rename(columns = {"Neto":"Afecto"})

        entre_v2.insert(0, "LLAVE", "")
        entre_v2.insert(1, "RUNFONDO", "")
        entre_v2.insert(2, "SERIEFONDO","")


        ### diferencias ###
        vali = pd.DataFrame(d_remu.groupby(["a_remu", "Fecha", "a_FS"])[["Remuneracion", "Patrimonio_Afecto"]].sum()).reset_index()
        vali.columns = ["a_remu", "fecha", "serie", "remu_remu", "patri_remu"]
        for i, x in enumerate(vali["a_remu"]):
            vali.loc[i, "remu_entre"] = entre.loc[entre["a_remu"] == x, "Remuneración"].sum()
            vali.loc[i, "patri_entre"] = entre.loc[entre["a_remu"] == x, "Patrimonio Afecto"].sum()
            vali.loc[i, "dif_remu"] = vali.loc[i, "remu_remu"] - vali.loc[i, "remu_entre"]
            vali.loc[i, "dif_patri"] = vali.loc[i, "patri_remu"] - vali.loc[i, "patri_entre"]

        for x in vali.iloc[:,3:].columns.tolist():
            vali.loc[:,x] = vali.loc[:,x].astype(float)

        ### tabla con informacion relevante ###
        # tabla1 = remu.groupby(["a_Fondo", "a_FS"])["Remune"]

        # ENTREGRA_FINAL Pasamos de formato Decimal a float
        entre_final = entre_v2.copy()
        for c in entre_final.columns:
            if c in ["SaldoCuotasInicial", "AportesTipo1", "SaldoCuotas", "Remuneración", "Patrimonio Afecto", "Afecto", "Exento", "IVA", "TOTAL", "Remuneración Neta"]:
                entre_final[c] = entre_final[c].astype(float)

        # ENTREGA_BRUTA
        for c in entre.columns:
            if c in ["SaldoCuotasInicial", "AportesTipo1", "SaldoCuotas", "Remuneración", "Patrimonio Afecto", "Neto", "Exento", "IVA", "TOTAL"]:
                entre[c] = entre[c].astype(float)

        ### EXPORTAMOS ###
        ruta = str(Path.home() / "Downloads") + "/Rebates1.xlsx"
        with pd.ExcelWriter(ruta) as w:
            entre.iloc[:,0:21].to_excel(w, sheet_name="Plantilla", index = False)
            entre_final.to_excel(w, sheet_name="Entregable", index = False)
            vali.to_excel(w, sheet_name="Diferencias", index = False)

        return entre_final


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
