# Privacidad y preparación de datos

Antes de adjuntar un documento, comparte solo los campos necesarios para resolver la tarea. Para probar una aplicación o reproducir un error, utiliza datos ficticios que conserven la estructura del caso.

## Qué revisar antes de compartir un archivo

| Material | Datos que conviene retirar o sustituir |
| --- | --- |
| Facturas y extractos | DNI, NIF, IBAN, direcciones, contactos y referencias internas que no sean necesarias. |
| Hojas de cálculo | Hojas ocultas, comentarios, columnas auxiliares y datos de terceros ajenos a la prueba. |
| Código y registros | Contraseñas, claves de API, tokens, cadenas de conexión y rutas con datos personales. |
| Capturas y PDF | Firmas, códigos QR, identificadores visibles y texto seleccionable bajo zonas cubiertas. |

Crea una copia independiente y abre el archivo resultante antes de compartirlo. Comprueba que los datos retirados tampoco aparecen al buscar texto, copiar contenido o mostrar hojas y columnas ocultas.

## Cómo preparar datos ficticios para una conciliación

1. Conserva los nombres de las columnas y el tipo de cada dato.
2. Sustituye terceros, referencias, fechas e importes por valores inventados.
3. Usa el mismo tercero ficticio en ambos archivos cuando quieras representar una coincidencia.
4. Reproduce el patrón que necesitas probar. Por ejemplo, crea dos registros contables con el mismo importe y fecha para comprobar cómo se presenta una ambigüedad.
5. Recalcula los totales y anota qué cruces esperas encontrar.

| Campo | Valor de prueba |
| --- | --- |
| Tercero | Proveedor Ejemplo A |
| Factura | F-TEST-001 |
| Fecha | 04/11/2026 |
| Importe | 125,00 € |
| Concepto | Material de oficina |

Si la tarea requiere usar información real, revisa las reglas de tu organización y las opciones de datos del servicio antes de subir el archivo. [Controles de datos de ChatGPT](https://help.openai.com/en/articles/7730893-data-controls-in-chatgpt)
