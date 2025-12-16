from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import pandas as pd
import tempfile
from io import StringIO # <-- Importación necesaria para manejar el contenido en memoria

app = FastAPI()

# Función que procesa los archivos CSV (No necesita cambio)
def procesar_archivos(valores_buffer, personas_buffer): # <-- Ahora espera buffers (StringIO)
    # Pandas puede leer directamente desde StringIO
    valores = pd.read_csv(valores_buffer)
    personas = pd.read_csv(personas_buffer)

    # ... (El resto de tu lógica de pandas permanece igual) ...
    # No se muestra el código completo de Pandas aquí por brevedad, pero
    # asume que el resto de 'procesar_archivos' sigue intacto.

    personas['CELULAR_1'] = personas['CELULAR_1'].astype(str).str.replace('.0', '', regex=False)

    agrupado = personas.groupby('numero_poliza').agg({
        'clave_agente': 'first',
        'codigo_producto': 'first',
        'fecha_emision': 'first',
        'nombre_producto': 'first',
        'nombre_opcion_poliza': 'first',
        'tipo_documento': 'first',
        'NUMERO_DOCUMENTO': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
        'NOMBRE': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
        'CORREO_1': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
        'CELULAR_1': lambda x: next(iter(sorted(set(x.dropna().astype(str)))), ''),
        'FECHA_PROCESO': 'first'
    }).reset_index()

    resultado = pd.merge(
        agrupado,
        valores[['numero_poliza', 'PRIMA']],
        on='numero_poliza',
        how='left'
    ).rename(columns={
        'PRIMA': 'Prima_totalizada',
        'FECHA_PROCESO': 'FECG'
    })

    def expandir_columna(df, columna_base, prefijo):
        df[columna_base] = df[columna_base].fillna('')
        listas = df[columna_base].apply(lambda x: [i.strip() for i in x.split(',') if i.strip()])
        max_items = listas.apply(len).max()
        nuevas = pd.DataFrame(listas.tolist(), columns=[f"{prefijo}_{i+1}" for i in range(max_items)])
        return pd.concat([df.drop(columns=[columna_base]), nuevas], axis=1)

    resultado = expandir_columna(resultado, 'NUMERO_DOCUMENTO', 'Numero_documento')
    resultado['telefono'] = resultado['CELULAR_1']

    columnas_finales = [
        'emisiones', 'revision agente', 'codigo_producto', 'clave_agente', 'numero_poliza',
        'fecha_emision', 'nombre_producto', 'nombre_opcion_poliza', 'Prima_totalizada',
        'CORREO_1', 'Numero_documento_1', 'tipo_documento', 'Numero_documento_2',
        'FECG', 'telefono'
    ]

    for c in columnas_finales:
        if c not in resultado.columns:
            resultado[c] = ''

    return resultado[columnas_finales]


# Endpoint para procesar archivos y devolver CSV (CORREGIDO)
@app.post("/procesar")
async def procesar(
    valores: UploadFile = File(...),
    personas: UploadFile = File(...)
):
    # 1. Leer el contenido del archivo de forma asíncrona
    # Nota: .read() lee todo el contenido en memoria (cuidado con archivos gigantes)
    valores_content = await valores.read()
    personas_content = await personas.read()

    # 2. Convertir el contenido binario a un buffer de texto (StringIO) para Pandas
    # Esto es crucial para que pd.read_csv funcione correctamente con el contenido en memoria
    valores_buffer = StringIO(valores_content.decode('utf-8'))
    personas_buffer = StringIO(personas_content.decode('utf-8'))

    # 3. Llamar a la función de procesamiento con los buffers
    df = procesar_archivos(valores_buffer, personas_buffer)

    # 4. Guardar en archivo temporal
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
    # Nota: Se usa 'utf-8-sig' para asegurar que Excel lea correctamente los acentos
    df.to_csv(tmp.name, index=False, encoding="utf-8-sig")

    # 5. Devolver archivo CSV real
    return FileResponse(tmp.name, filename="resultado.csv", media_type="text/csv")