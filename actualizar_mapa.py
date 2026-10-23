from pathlib import Path
import zipfile

carpeta = Path(__file__).resolve().parent
excel = carpeta / "MapaIndustrialUTNFRD.xlsx"

if excel.exists():
    with zipfile.ZipFile(excel) as archivo:
        print("¡Python pudo abrir tu Excel!")

    print("Archivo encontrado:", excel.name)
else:
    print("No encontré el Excel.")
    print("Debe estar en esta carpeta:", carpeta)
    print("Y llamarse: MapaIndustrialUTNFRD.xlsx")
import xml.etree.ElementTree as ET

if excel.exists():
    with zipfile.ZipFile(excel) as archivo:
        ns = {
            "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        }

        textos = []

        if "xl/sharedStrings.xml" in archivo.namelist():
            contenido = archivo.read("xl/sharedStrings.xml")
            raiz = ET.fromstring(contenido)

            for elemento in raiz.findall("m:si", ns):
                texto = "".join(
                    parte.text or ""
                    for parte in elemento.findall(".//m:t", ns)
                )
                textos.append(texto)

        contenido = archivo.read("xl/worksheets/sheet1.xml")
        hoja = ET.fromstring(contenido)

        nombres = []

        for fila in hoja.findall("m:sheetData/m:row", ns):
            if fila.get("r") == "1":
                continue

            for celda in fila.findall("m:c", ns):
                referencia = celda.get("r", "")
                columna = "".join(
                    letra for letra in referencia if letra.isalpha()
                )

                if columna != "A":
                    continue

                valor = celda.find("m:v", ns)
                nombre = ""

                if celda.get("t") == "s" and valor is not None:
                    nombre = textos[int(valor.text)]
                elif celda.get("t") == "inlineStr":
                    nombre = "".join(
                        parte.text or ""
                        for parte in celda.findall(".//m:t", ns)
                    )
                elif valor is not None:
                    nombre = valor.text or ""

                if nombre.strip():
                    nombres.append(nombre.strip())

        print("\nCantidad de fichas:", len(nombres))
        print("\nPrimeras cinco:")

        for nombre in nombres[:5]:
            print("-", nombre)
import json

if excel.exists():
    def leer_celda(celda):
        tipo = celda.get("t")
        valor = celda.find("m:v", ns)

        if tipo == "s" and valor is not None:
            return textos[int(valor.text)]

        if tipo == "inlineStr":
            return "".join(
                parte.text or ""
                for parte in celda.findall(".//m:t", ns)
            )

        return valor.text if valor is not None else ""


    columnas = {}
    empresas = []

    for fila in hoja.findall("m:sheetData/m:row", ns):
        datos = {}

        for celda in fila.findall("m:c", ns):
            referencia = celda.get("r", "")
            columna = "".join(
                letra for letra in referencia if letra.isalpha()
            )

            if columna not in [
                "A", "B", "C", "D", "E", "F",
                "G", "H", "I", "J", "K", "L"
            ]:
                continue

            valor = leer_celda(celda)

            if fila.get("r") == "1":
                columnas[columna] = str(valor).strip()
            elif columna in columnas:
                datos[columnas[columna]] = valor

        if fila.get("r") == "1":
            continue

        if not datos.get("Nombre", "").strip():
            continue
        print("Datos leídos:", datos)
        for campo in ["Latitud", "Longitud"]:
            coordenada = datos.get(campo, "")
            datos[campo] = float(str(coordenada).replace(",", "."))

        empresas.append(datos)

    salida = carpeta / "empresas.json"

    salida.write_text(
        json.dumps(empresas, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    print("\nSe creó empresas.json con", len(empresas), "fichas.")