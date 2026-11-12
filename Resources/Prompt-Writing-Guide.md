# Guía para escribir prompts

Un prompt bien planteado permite que el asistente entienda qué debe hacer y que tú puedas comprobar si lo ha hecho. Incluye los detalles que cambian la respuesta esperada.

## Plantilla reutilizable

```text
Objetivo: [qué necesito conseguir]
Contexto: [para quién, para qué y en qué entorno]
Material: [texto, código, datos o imagen de referencia]
Restricciones: [qué debe respetarse y qué puede modificarse]
Entrega: [formato, idioma y nivel de detalle]
Comprobación: [cómo sabré que el resultado cumple el objetivo]
```

## Ejemplo de texto

**Petición inicial**

> Arréglame este formulario.

**Petición preparada**

```text
Revisa el procesamiento de un formulario en PHP 8. El formulario debe
mostrar un error junto al campo nombre cuando se envía vacío.

Te adjunto UsuarioRequest.php, el controlador y la vista. Conserva la
estructura actual y propón el cambio mínimo necesario.

Indica en qué archivo aplicar cada cambio y muestra el fragmento que
debo sustituir. Explica cómo comprobar el resultado con un nombre vacío
y con un nombre válido.
```

La segunda petición describe el fallo observable, delimita los archivos y define dos pruebas para comprobar la corrección.

## Ejemplo de imagen

**Petición inicial**

> Hazme una imagen para la portada de una web.

**Petición preparada**

```text
Genera una imagen horizontal 16:9 para la cabecera de una web de
asesoría contable. Muestra un escritorio luminoso con una libreta y un
portátil cerrado. Coloca los objetos en el tercio derecho y deja la
mitad izquierda despejada, con un fondo azul oscuro uniforme, para
superponer el título en blanco.

Usa iluminación natural y un estilo fotográfico sobrio. Entrega la
imagen sin texto, logotipos ni datos visibles en los documentos.
```

## Cómo pedir una revisión

Cuando una respuesta se acerque al resultado que quieres, identifica qué parte te sirve y qué ajuste necesitas. Por ejemplo, «Conserva el asunto y el primer párrafo. Acorta el cierre a una frase y deja clara la fecha límite del viernes».

Consulta el [catálogo de prompts](../Prompts/) para adaptar ejemplos a otras tareas.
