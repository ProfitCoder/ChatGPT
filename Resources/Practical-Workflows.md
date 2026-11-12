# Flujos de trabajo prácticos

Estos ejemplos muestran cómo preparar una entrada, utilizar una herramienta y verificar el resultado. Todos los datos de ejemplo son ficticios.

## 1. Diagnosticar un error en un formulario

**Situación.** Un formulario rechaza un nombre vacío, pero la vista no muestra el mensaje de validación.

1. Reproduce el fallo enviando el formulario con el campo vacío.
2. Reúne el fragmento de validación, la llamada que prepara los datos para la vista y el código que muestra el mensaje.
3. Pide que se siga el recorrido del error desde su creación hasta la vista y se localice el primer punto donde deja de estar disponible.
4. Aplica el cambio mínimo propuesto y repite la prueba con un nombre vacío, uno formado solo por espacios y uno válido.

**Resultado esperado.** Las dos entradas vacías generan un mensaje en el campo correspondiente. La entrada válida sigue el flujo habitual del formulario.

## 2. Preparar un repaso a partir de apuntes

**Situación.** Quieres practicar consultas SQL con `JOIN`, `GROUP BY` y `COUNT`.

1. Facilita tus apuntes y un esquema reducido con dos tablas relacionadas.
2. Pide tres ejercicios progresivos y solicita que las soluciones aparezcan al final.
3. Resuelve las consultas y ejecútalas con datos que incluyan un registro sin coincidencias y otro con varias.
4. Contrasta el número de filas y los recuentos. Si obtienes un resultado distinto, pide que se explique la diferencia fila por fila.

**Resultado esperado.** Puedes justificar qué registros conserva cada unión y por qué cambia el recuento al agrupar.

## 3. Revisar una conciliación bancaria

**Situación.** Quieres comprobar cómo revisar un cruce claro y otro con más de un candidato. Los importes de ambas tablas se muestran como salidas positivas. Al preparar los Excel de prueba, utiliza la convención de signos y las columnas de tus archivos de entrada habituales.

### Extracto bancario de ejemplo

| ID | Fecha | Importe | Tercero | Concepto |
| --- | --- | ---: | --- | --- |
| B-001 | 04/11/2026 | 125,00 € | Papelería Ejemplo | Factura F-101 |
| B-002 | 05/11/2026 | 80,00 € | Servicios Ejemplo | Transferencia |

### Registro contable de ejemplo

| ID | Fecha | Importe | Tercero | Concepto |
| --- | --- | ---: | --- | --- |
| C-001 | 04/11/2026 | 125,00 € | Papelería Ejemplo | Factura F-101 |
| C-002 | 05/11/2026 | 80,00 € | Servicios Ejemplo | Factura F-102 |
| C-003 | 05/11/2026 | 80,00 € | Servicios Ejemplo | Factura F-103 |

### Prueba y revisión

1. Prepara dos archivos de prueba a partir de las tablas y sigue las instrucciones de inicio del [conciliador bancario](../Apps/App%20Sobre%20conciliacion%20Bancaria%20%28via%20Python%29/README_ULTRA10.txt).
2. Revisa el informe final. `B-001` y `C-001` forman un cruce justificable por fecha, importe, tercero y referencia.
3. Para `B-002`, revisa `C-002` y `C-003`. Ambos tienen la misma fecha, importe y tercero. El concepto bancario no identifica la factura, así que hace falta un justificante para resolver el cruce.
4. Comprueba los importes por separado. El cruce justificable suma **125,00 €**. Quedan **80,00 €** por identificar en el banco y dos candidatos contables de **80,00 €** cada uno.

**Criterio de aceptación.** Cada movimiento marcado como conciliado debe tener una correspondencia justificable. Mantén pendiente el pago ambiguo hasta disponer de una referencia, un justificante o un detalle de la remesa que lo aclare.
