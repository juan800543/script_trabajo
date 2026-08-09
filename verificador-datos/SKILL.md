---
name: verificador-datos
description: >
  Revisa textos, posts, guiones, artículos, informes o respuestas de IA para comprobar si las afirmaciones son correctas, necesitan matiz o pueden ser falsas, exageradas o no verificables. Úsala cuando el usuario pida verificar, comprobar, fact-checkear, revisar datos, detectar errores o validar información. Actívala con frases como: "verifica este texto", "comprueba si esto es verdad", "revisa si hay errores", "pásalo por el verificador", "haz fact check", "comprueba las afirmaciones", "dime si esto está exagerado", "revisa este guion antes de publicarlo", "verifica este post antes de subirlo". Compatible con: posts de LinkedIn, guiones de YouTube, artículos, emails, informes, transcripciones, noticias, resúmenes de IA y documentos técnicos. No esperes palabras exactas — si hay un texto con afirmaciones que verificar antes de publicar o usar, activa esta skill.
---

# Verificador de Datos

Skill para detectar errores, exageraciones, afirmaciones no verificables y datos incorrectos en cualquier texto antes de publicarlo, enviarlo o usarlo.

---

## Proceso obligatorio

Sigue estos pasos en orden:

**1. Lee el texto completo** antes de hacer ninguna valoración.

**2. Extrae todas las afirmaciones factuales verificables**: datos numéricos, fechas, nombres, estadísticas, funciones de herramientas, precios, disponibilidades, rankings, leyes, claims de salud o finanzas, afirmaciones técnicas, comparaciones.

**3. Separa opinión de hecho**: distingue afirmaciones comprobables de valoraciones subjetivas. Las opiniones no se verifican, pero sí se señala si se presentan como hechos.

**4. Verifica cada afirmación usando**:
- El documento o material fuente si el usuario lo ha aportado
- La transcripción u original si está disponible en el contexto
- Búsqueda web si la herramienta está disponible (úsala para afirmaciones cambiantes: precios, fechas, rankings, disponibilidad de herramientas)
- Razonamiento cuidadoso basado en conocimiento disponible si no hay acceso web

**5. Clasifica cada afirmación** según las categorías definidas abajo.

**6. Señala errores, exageraciones y frases que necesitan matiz.**

**7. Propón correcciones concretas** para cada afirmación problemática.

**8. Entrega el informe completo** en el formato especificado.

---

## Clasificación de afirmaciones

| Categoría | Definición |
|-----------|-----------|
| ✅ **Correcta** | Respaldada claramente por las fuentes disponibles |
| 🟡 **Mayormente correcta** | Cierta en general, pero necesita algún matiz |
| ⚠️ **Dudosa / No verificable** | Sin evidencia suficiente o sin fuente clara |
| 🔶 **Exagerada** | Tiene base real, pero formulada de forma demasiado rotunda |
| ❌ **Incorrecta** | Contradice la información disponible o parece falsa |
| 💬 **Opinión / No factual** | Valoración subjetiva, no requiere verificación |

---

## Formato de respuesta

Devuelve siempre este informe completo:

---

### 1. Resumen general
Evaluación breve del texto: limpio, necesita matices menores, tiene errores importantes, no publicar sin revisar, etc.

### 2. Tabla de verificación

| Afirmación | Clasificación | Evidencia o motivo | Corrección sugerida |
|-----------|--------------|-------------------|-------------------|
| [cita textual o paráfrasis de la afirmación] | [emoji + categoría] | [por qué se clasifica así] | [texto corregido o "—" si no aplica] |

### 3. Errores o riesgos principales
Lista breve de los puntos más importantes a corregir, ordenados por impacto en la credibilidad.

### 4. Versión corregida
Reescribe únicamente las frases problemáticas en su contexto. Si el usuario pide el texto completo corregido, proporciónalo.

### 5. Recomendación final
Una de estas cuatro opciones con explicación breve:
- ✅ **Publicar tal cual**
- 🟡 **Publicar con cambios menores**
- 🔶 **Revisar antes de publicar**
- ❌ **No publicar sin verificar**

---

## Reglas de comportamiento

- **No marques como falso algo solo porque no puedas comprobarlo.** Si no hay fuente, clasifícalo como no verificable.
- **No cambies opiniones subjetivas** salvo que se presenten como hechos objetivos.
- **No seas alarmista.** El tono debe ser útil y constructivo, no intimidatorio.
- **No alargues el informe** más de lo necesario. Prioriza claridad sobre exhaustividad.
- **Prioriza los errores que afectan a la credibilidad** del usuario o de su marca.
- **Sé especialmente cuidadoso** con: herramientas de IA (funciones, precios, disponibilidad), estadísticas, fechas, rankings, leyes, afirmaciones de salud o finanzas, y claims técnicos.
- **Si hay acceso web**, úsalo para verificar afirmaciones cambiantes. No dependas solo del conocimiento de entrenamiento para precios, versiones de software o datos recientes.
- **Si no hay acceso web**, indica al inicio del informe: *"Verificación basada en el material disponible y conocimiento de entrenamiento. Afirmaciones cambiantes (precios, versiones, disponibilidad) deben comprobarse manualmente."*
- **Mantén el idioma original del texto revisado.** Si el texto está en inglés, el informe va en inglés. Si está en español, en español.
- **Si el texto es muy largo** (más de 800 palabras), prioriza las afirmaciones con mayor riesgo y avisa al usuario de que el análisis se ha centrado en los puntos críticos.

---

## Áreas de especial atención

Cuando el texto contenga alguno de estos elementos, aplica criterio más estricto:

- **Herramientas de IA**: funcionalidades que pueden haber cambiado, modelos que pueden haber sido actualizados, precios y planes de suscripción, integraciones disponibles
- **Estadísticas y datos numéricos**: porcentajes de mejora, cifras de usuarios, comparativas de rendimiento
- **Fechas y cronologías**: lanzamientos, actualizaciones, eventos
- **Claims de productividad o eficiencia**: "X veces más rápido", "ahorra el 80% del tiempo"
- **Afirmaciones sobre competidores**: comparativas que pueden ser inexactas o sesgadas
- **Contenido de salud, finanzas o legal**: especialmente sensible por el impacto potencial

---

## Ejemplos de uso

El usuario puede pegar directamente el texto a verificar, o puede decir:

- "Antes de subir este guion, pásalo por el verificador"
- "¿Hay algo incorrecto en este post?"
- "Verifica las estadísticas de este artículo"
- "Comprueba si lo que dice aquí sobre ChatGPT es verdad"
- "Haz fact check de este email antes de enviarlo"

En todos estos casos, sigue el proceso completo y devuelve el informe en el formato definido.
