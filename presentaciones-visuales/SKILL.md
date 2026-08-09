---
name: presentaciones-visuales
description: >
  Crea presentaciones visuales en formato HTML, modernas, claras y listas para usar en navegador,
  a partir de una idea, esquema, documento, transcripción, PowerPoint o cualquier contenido que
  el usuario quiera convertir en diapositivas. Úsala siempre que el usuario quiera crear una
  presentación, slides, deck o diapositivas — ya sea desde cero o convirtiendo contenido
  existente. También actívala si el usuario dice "hazlo visual", "crea un deck", "convierte esto
  en slides", "prepara una presentación para explicar esto", "mejora esta presentación" o
  cualquier variante. Si hay un contenido que transformar en presentación visual, usa esta skill
  aunque el usuario no lo llame explícitamente "presentación". Cubre casos para reuniones,
  clases, charlas, pitch, propuestas comerciales, formaciones y vídeos.
---

# Presentaciones Visuales

Crea presentaciones HTML autocontenidas, con diseño profesional, narrativa clara y calidad
suficiente para usar en reuniones, clases, vídeos o propuestas comerciales.

---

## Proceso paso a paso

### 1. Analiza el input del usuario

Lee el contenido o idea proporcionada. Extrae:
- **Tema y objetivo** de la presentación
- **Público al que va dirigida** (si se puede inferir)
- **Número aproximado de slides** necesarias
- **Tono**: formal, divulgativo, comercial, educativo, técnico...
- **Uso previsto**: reunión, vídeo, formación, pitch, clase, propuesta

Si falta información crítica (público, objetivo, duración), pregunta antes de generar.
Si no es crítica, asume lo razonable y continúa — no bloquees el proceso.

---

### 2. Define la estructura narrativa

Toda presentación debe tener un arco claro:

| Sección | Propósito |
|---|---|
| Apertura / portada | Captar atención, enmarcar el tema |
| Contexto / problema | Por qué importa esto |
| Desarrollo | El contenido principal, ideas clave |
| Ejemplos o datos | Prueba, evidencia, caso real |
| Cierre / conclusión | Resumen o llamada a la acción |

Adapta este esquema al tipo de presentación. Un pitch no es lo mismo que una formación.

---

### 3. Elige un estilo visual

Escoge el estilo más adecuado al tema y al público. Si el usuario especifica marca, colores o
estilo propios, prioriza siempre su guía visual.

| Estilo | Cuándo usarlo |
|---|---|
| Profesional minimalista | Propuestas, reuniones corporativas |
| Tecnológico / oscuro | IA, producto digital, startups |
| Educativo / claro | Formaciones, clases, explicaciones |
| Premium / editorial | Marca personal, lujo, consultoría |
| Creativo | Agencias, diseño, contenido |
| Corporativo | Empresas grandes, informes internos |

---

### 4. Reglas de diseño que debes aplicar siempre

**Estructura visual:**
- Una sola idea principal por slide
- Jerarquía clara: título grande → subtítulo → contenido → nota opcional
- Márgenes amplios, espacio en blanco generoso
- Tamaño de letra legible (mínimo 16px para cuerpo, 28-48px para títulos)

**Variedad de layouts:**
No repitas el mismo layout en todas las slides. Usa una mezcla de:
- Slide centrada con título grande + frase
- Dos columnas (concepto + explicación)
- Tarjetas horizontales o verticales
- Lista con iconos
- Slide de dato grande / número clave
- Timeline o proceso en pasos
- Comparativa dos opciones
- Slide de cierre con CTA o frase impactante

**Elementos visuales permitidos:**
- Iconos SVG inline o emojis minimalistas como apoyo visual
- Líneas separadoras, fondos de color en secciones
- Números grandes para estadísticas o pasos
- Bloques de color para destacar conceptos clave
- Progress bar o numeración de slides

**Prohibido:**
- Párrafos largos en slides
- Más de 5-6 puntos en una lista
- Colores aleatorios sin coherencia
- Fondos con imágenes externas (no se garantiza que carguen)
- Estilo infantil salvo que se pida expresamente

---

### 5. Genera el HTML

El output debe ser un **único archivo HTML autocontenido**:
- CSS incluido dentro del mismo archivo (en `<style>`)
- Sin dependencias externas de imágenes o fuentes (usar Google Fonts solo si es seguro)
- Navegación por slides con teclado (← →) o botones visibles
- Diseño responsive si es posible
- Transiciones suaves y profesionales entre slides (opcional pero recomendable)

**Estructura técnica mínima:**
```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Nombre de la presentación</title>
  <style>
    /* Variables CSS para el tema de color */
    :root {
      --primary: #...;
      --accent: #...;
      --bg: #...;
      --text: #...;
    }
    /* Reset y base */
    /* Estilos de slides */
    /* Navegación */
    /* Animaciones suaves */
  </style>
</head>
<body>
  <!-- slides como secciones o divs -->
  <!-- Navegación con flechas o dots -->
  <script>
    // Lógica de navegación entre slides
  </script>
</body>
</html>
```

---

### 6. Presentaciones para vídeo

Si el usuario indica que la presentación es para vídeo o pantalla:
- Usa frases muy cortas, una por slide
- Elementos grandes, tipografía en tamaño máximo
- Fondo de color sólido o gradiente
- Sin listas largas — máximo 3 puntos por slide
- Animaciones de entrada para cada elemento si es posible

---

### 7. Presentaciones para reunión o formación

Si el uso es reunión, clase o formación:
- Prioriza claridad y orden sobre espectacularidad visual
- Incluye más contenido por slide que en formato vídeo
- Numeración de slides visible
- Espacio para notas del presentador si el usuario lo pide

---

## Formato de salida

Cuando generes una presentación, entrega siempre este formato:

**Resumen de enfoque:**
> [Objetivo de la presentación, público asumido, estilo visual elegido y por qué]

**Estructura de slides:**
> [Lista numerada con los títulos de las diapositivas]

**Presentación HTML:**
> [Código HTML completo, listo para abrir en navegador]

**Recomendaciones opcionales:**
> [Ajustes de marca, colores, imágenes, logos o mejoras que el usuario podría aplicar]

---

## Casos especiales

**Si el usuario aporta un documento o transcripción larga:**
Extrae las ideas principales, no intentes meter todo el texto en las slides. Resume, prioriza, jerarquiza.

**Si el usuario aporta un PowerPoint o esquema:**
Respeta la estructura y los mensajes clave, pero mejora el diseño y la jerarquía visual.

**Si el usuario no da estilo:**
Elige el estilo más adecuado al tema, explícalo brevemente en el resumen de enfoque y úsalo de forma coherente.

**Si el usuario da una guía de marca:**
Respeta colores, fuentes y tono por encima de cualquier otra preferencia de diseño.

**Si el contenido necesita datos o fuentes que no tienes:**
Indica claramente qué datos faltan. No inventes cifras.
