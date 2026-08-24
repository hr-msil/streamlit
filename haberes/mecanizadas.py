from pypdf import PdfReader
import re
import pandas as pd
import unicodedata


PATRON = re.compile(r'\b\d{3}\.\d\b')

PATRON_NOMINAL = re.compile(
    r'^\s*(F?\d+/\d+)\s+([A-ZÁÉÍÓÚÑ ]+?)\s+(?:T P|T E|S P|T J|S J|P P|P J|T D|S D|T M|P M|S M)\s+(\d+\.\d+)\s+(.+?)\s+(\d+\.\d+)\s+NETO'
)

PATRON_DATO = re.compile(
    r'(?P<codigo>\d+\.\d+)\s+'
    r'(?P<descripcion>.*?)\s+'
    r'(?P<importe>-?\s*\d+\.\d+)\s*$'
)

PATRON_ORG = re.compile(
    r'TIPO ORG\s*:\s*(.*?)\s{3,}(.*)$'
)

PATRON_FECHA_AFECTADA = re.compile(
    r'AFEC:\s*(\d{6})'
)

PATRON_CATEGORIA = re.compile(
    r'(CATEG\.\s*[A-Z0-9.]+(?:\s+[A-Z0-9.]+)*?(?:\s+\d+\.\d{2})?)(?=\s{2,}|$)'
)

PATRON_ANTIGUEDAD = re.compile(
    r'ANT\.\s*([0-9]{1,2}/[0-9]{2})'
)

PATRON_SUPLENCIA = re.compile(
    r'SUPLE\s+A:\s*([0-9]+/\d+)'
)

PATRON_FECHA_SUPLENCIA  = re.compile(
    r'(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})'
)

def normalize_filename(text: str, separator: str = "_") -> str:
    # 1. Decompose unicode characters (e.g., 'é' becomes 'e' + accent)
    text = unicodedata.normalize('NFKD', text)
    # 2. Encode to ASCII bytes ignoring non-ASCII characters, then decode back to string
    text = text.encode('ascii', 'ignore').decode('ascii')
    # 3. Lowercase the text
    text = text.lower()
    # 4. Replace spaces and existing hyphens/underscores with the chosen separator
    text = re.sub(r'[\s\-_]+', separator, text)
    # 5. Remove any character that isn't alphanumeric or the separator
    text = re.sub(r'[^a-z0-9' + re.escape(separator) + r']', '', text)
    # 6. Trim leading/trailing separators
    return text.strip(separator)


def leer_mecanica(nombre_archivo: str = r"C:\Users\mmaurer\Desktop\EMAP - JUNIO.pdf") -> tuple[pd.DataFrame, pd.DataFrame, str]:
    '''
    Función que recibe un archivo PDF e itera linea por linea para extraer todos los datos necesarios a partir de las regex especificadas.
    Devuelve un DataFrame con los agentes y sus respectivos importes y descripciones correspondientes. Además del nombre que va a tener el archivo resultante, correspondiente
    al nombre de la institución correspondiente.
    '''
    # Abrir el archivo PDF en modo de lectura binaria
    reader = PdfReader(nombre_archivo)

    terminar = False

    fila_nombre = None
    contador_global = 0

    filas = []
    fila_pendiente = None
    filas_datos_persona = []

    nombre = None
    identificador = None
    fecha_afectada = None
    categoria = None
    antiguedad = None
    suplencia = None

    # Extraer el texto página por página
    for i, page in enumerate(reader.pages):
        texto = page.extract_text()
        pagina = i + 1

        #Itero linea por linea
        for j,linea in enumerate(texto.splitlines()):
            contador_global += 1
                    
            if "TOTAL DEL DISTRITO" in linea:
                terminar = True
                break

            if pagina == 1:
                match_org = PATRON_ORG.search(linea)
                if match_org:
                    tipo_org = match_org.group(1).strip()
                    institucion = match_org.group(2).strip()

            if fila_nombre is not None:
                offset = contador_global - fila_nombre

                if offset ==  1:
                    match_fecha_afectada = PATRON_FECHA_AFECTADA.search(linea)
                    match_categoria = PATRON_CATEGORIA.search(linea)
                    fecha_afectada = match_fecha_afectada.group(1) if match_fecha_afectada else None
                    categoria = re.sub(r'\s+', ' ', match_categoria.group(1)).strip() if match_categoria else None
                
                elif offset == 2:
                    match_antiguedad = PATRON_ANTIGUEDAD.search(linea)
                    antiguedad = match_antiguedad.group(1) if match_antiguedad else None
                    if "SIN HABERES" in linea:
                        continue
                    if "NO SUBVENCIONADO" in linea:
                        continue
                    
                elif offset == 3:
                    match_suplencia = PATRON_SUPLENCIA.search(linea)
                    suplencia = match_suplencia.group(1) if match_suplencia else None
        
                elif offset == 4:
                    match_fecha_afectada_sup = PATRON_FECHA_SUPLENCIA.search(linea)
                    fecha_desde = match_fecha_afectada_sup.group(1) if match_fecha_afectada_sup else None
                    fecha_hasta = match_fecha_afectada_sup.group(2) if match_fecha_afectada_sup else None
                    fila_datos_persona = {
                                            "DNI": identificador,
                                            "NOMBRE": nombre,
                                            "FECHA": fecha_afectada,
                                            "CATEGORIA": categoria,
                                            "ANTIGUEDAD": antiguedad,
                                            "SUPLENCIA": suplencia,
                                            "FECHA DESDE": fecha_desde,
                                            "FECHA HASTA": fecha_hasta
                    
                                        }
                    filas_datos_persona.append(fila_datos_persona)
            
            match = PATRON.search(linea)
            
            if match:
                if match.group() == "011.0":
                    match_2 = PATRON_NOMINAL.search(linea)   
                    if match_2:

                        identificador = match_2.group(1)
                        nombre = match_2.group(2).strip()
                        codigo = match_2.group(3)
                        descripcion = match_2.group(4).strip()
                        importe = match_2.group(5)
                        fila_nombre = contador_global

    

                else:

                    match_dato = PATRON_DATO.search(linea[match.start():])

                    if match_dato:
                        codigo = match_dato.group("codigo")
                        descripcion = match_dato.group("descripcion").strip()
                        importe = match_dato.group("importe").replace(" ", "")

                fila = {
                    "DNI": identificador,
                    "NOMBRE": nombre,
                    "CODIGO": codigo,
                    "DESCIPCION": descripcion,
                    "IMPORTE": importe,

                }

                filas.append(fila)

            if terminar: break


    df = pd.DataFrame(filas)
    df_datos = pd.DataFrame(filas_datos_persona)
    df["IMPORTE"] = df["IMPORTE"].astype('float')
    nombre_archivo_res = normalize_filename(text = f"{tipo_org}_{institucion}", separator = "_") 

   
    #with pd.ExcelWriter(f"{tipo_org}_{institucion}.xlsx") as writer:
     #   df.to_excel(writer, sheet_name="IMPORTES", index=False)
      #  df_datos.to_excel(writer, sheet_name="DATOS", index=False)
       # print(f'Archivo generado con nombre: {tipo_org}_{institucion}.xlsx')

    return df, df_datos, nombre_archivo_res


