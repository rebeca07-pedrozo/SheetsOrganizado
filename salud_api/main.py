from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import pandas as pd
import tempfile
from io import StringIO 

app = FastAPI()

def procesar_archivos(valores_buffer, personas_buffer): 
    valores = pd.read_csv(valores_buffer)
    personas = pd.read_csv(personas_buffer)


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


@app.post("/procesar")
async def procesar(
    valores: UploadFile = File(...),
    personas: UploadFile = File(...)
):

    valores_content = await valores.read()
    personas_content = await personas.read()


    valores_buffer = StringIO(valores_content.decode('utf-8'))
    personas_buffer = StringIO(personas_content.decode('utf-8'))

    df = procesar_archivos(valores_buffer, personas_buffer)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
    df.to_csv(tmp.name, index=False, encoding="utf-8-sig")

    return FileResponse(tmp.name, filename="resultado.csv", media_type="text/csv")