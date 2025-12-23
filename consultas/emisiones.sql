WITH leads_ordenados AS (
  SELECT 
    KEY_ID,
    TIPO_DOCUMENTO,
    FECHA_LEAD,
    FUENTE_LEAD,
    MEDIO_LEAD,
    CAMPANA_LEAD,
    PRODUCTO_SELECCIONADO,
    NOMBRE_FORMULARIO,
    ROW_NUMBER() OVER(
      PARTITION BY KEY_ID 
      ORDER BY FECHA_LEAD DESC
    ) as ranking_reciente,
    COUNT(*) OVER(PARTITION BY KEY_ID) as total_contactos
  FROM sb-ecosistemaanalitico-lago.contact_center.t_seguimiento_leads_mercadeo
  WHERE PRODUCTO_SELECCIONADO = 'PORTAFOLIO TRANQUILIDAD EN VIDA'
    AND KEY_ID IS NOT NULL
)

SELECT 
  KEY_ID AS identificador,
  TIPO_DOCUMENTO,
  FECHA_LEAD AS fecha_contacto_mas_reciente,
  FUENTE_LEAD,
  MEDIO_LEAD,
  CAMPANA_LEAD,
  PRODUCTO_SELECCIONADO,
  NOMBRE_FORMULARIO,
  total_contactos
FROM leads_ordenados
WHERE ranking_reciente = 1 
ORDER BY FECHA_LEAD DESC