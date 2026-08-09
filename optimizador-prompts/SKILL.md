---
name: optimizador-prompts
description: Transforma ideas desordenadas, prompts mal escritos, notas rápidas, dictados por voz o instrucciones incompletas en prompts claros, estructurados y listos para usar en herramientas de IA. Úsala siempre que el usuario quiera mejorar, ordenar, optimizar, reescribir o estructurar un prompt o una idea para convertirla en instrucción para una IA. Activa esta skill cuando el usuario diga cosas como "mejora este prompt", "conviértelo en un prompt", "ordena esta idea", "hazme un prompt para…", "optimiza esto para Claude / ChatGPT / Gemini / Midjourney / Sora / Claude Code / n8n / Make", "te voy a dictar una idea desordenada", o cualquier variante. No esperes que el usuario use la palabra exacta "prompt" — si hay una idea o instrucción que necesita ser estructurada para una IA, usa esta skill.
---

# Optimizador de Prompts

Convierte ideas caóticas, notas rápidas o instrucciones incompletas en prompts profesionales, precisos y listos para usar.

---

## Proceso

### 1. Detecta la herramienta objetivo

Antes de generar el prompt, identifica **para qué herramienta o modelo** se va a usar:
- Claude, ChatGPT, Gemini (conversacionales/texto)
- Claude Code u otras herramientas de programación
- Midjourney, Flux, Stable Diffusion, Firefly (imagen)
- Sora, Kling, Runway (vídeo)
- n8n, Make, Zapier (automatizaciones)

Si no queda claro, pregunta antes de continuar.

### 2. Extrae los componentes del prompt

Del input del usuario, identifica y extrae:
1. **Objetivo real**: ¿qué quiere conseguir?
2. **Contexto relevante**: rol, situación, datos de partida
3. **Tarea concreta**: acción específica que debe ejecutar la IA
4. **Especificaciones**: detalles técnicos, restricciones, tono, estilo
5. **Formato de salida**: cómo debe presentarse el resultado
6. **Criterios de calidad**: qué hace que el resultado sea bueno
7. **Cosas a evitar**: errores frecuentes, restricciones, exclusiones
8. **Verificación final**: si aplica, instrucción de autocomprobación

Si falta información **imprescindible**, pregunta. Si no es crítica, asume lo razonable y continúa.

### 3. Construye el prompt final

Estructura el prompt con estas secciones (adapta según la herramienta y complejidad):

```
[CONTEXTO Y ROL]
Quién es la IA, desde qué perspectiva actúa y en qué situación.

[TAREA CONCRETA]
Qué debe hacer exactamente.

[ESPECIFICACIONES]
Detalles importantes: tono, estilo, longitud, idioma, público objetivo, restricciones.

[CRITERIOS DE CALIDAD]
Cómo saber si el resultado es bueno.

[FORMATO DE RESPUESTA]
Cómo debe estructurarse la salida: listas, párrafos, tablas, JSON, etc.

[VERIFICACIÓN FINAL] (opcional)
Instrucción de autocomprobación antes de responder.
```

---

## Adaptación por herramienta

### Claude / ChatGPT / Gemini (texto y conversacional)
- Prioriza contexto claro, pasos definidos y formato de salida explícito.
- Añade instrucción de rol si el caso lo requiere.
- Si es un sistema de prompts, separa el system prompt del user message.

### Claude Code / herramientas de programación
- Incluye: objetivo del código, lenguaje/framework, estructura del proyecto si se conoce.
- Añade restricciones técnicas, comportamiento esperado, archivos afectados.
- Define criterios de validación y casos límite.
- Especifica si debe explicar el código o solo entregarlo.

### Midjourney / Flux / Stable Diffusion / herramientas de imagen
- Estructura: sujeto principal → composición → estilo visual → iluminación → encuadre → relación de aspecto → mood.
- Añade elementos obligatorios y lista de elementos a evitar (negative prompts si aplica).
- Adapta el formato a la herramienta (Midjourney usa `--ar`, `--style`; SD usa `negative prompt:`).

### Sora / Kling / Runway / herramientas de vídeo
- Incluye: escena de apertura, movimiento de cámara, acción principal, progresión visual.
- Añade estilo, duración aproximada, ambiente sonoro si aplica.
- Especifica continuidad visual si es parte de una secuencia.

### n8n / Make / Zapier / automatizaciones
- Incluye: trigger/desencadenante, inputs, pasos del flujo en orden, herramientas conectadas, output esperado.
- Añade casos límite, manejo de errores y qué debe pasar si falla un paso.

---

## Formato de respuesta

Devuelve **siempre** en este orden:

---

**Prompt optimizado:**

[Prompt final limpio, listo para copiar y pegar]

---

**Cambios principales realizados:**

[Lista breve de 3 a 5 mejoras aplicadas: qué faltaba, qué se ordenó, qué se añadió]

---

**Dudas opcionales:** *(solo si hay información que mejoraría mucho el resultado)*

[Preguntas concretas, máximo 2-3]

---

## Reglas

- No inventes detalles críticos que el usuario no ha dado.
- No cambies el objetivo original del usuario.
- No hagas el prompt más largo de lo necesario.
- No uses lenguaje corporativo, genérico o artificioso si el usuario busca naturalidad.
- No metas disclaimers ni advertencias innecesarias dentro del prompt.
- No des varias versiones salvo que el usuario lo pida explícitamente.
- No expliques metodologías ni teorías sobre prompting si el usuario no lo ha pedido.
- El resultado debe ser inmediatamente usable.
