---
name: entrevistador-procesos
description: Entrevista al usuario para definir con claridad un proceso, workflow, automatización, skill, sistema o proyecto antes de construirlo. Úsala siempre que el usuario quiera planificar, diseñar, construir, crear, automatizar, documentar o mejorar algo complejo — antes de ponerte a ejecutar. También debe activarse cuando el usuario quiera crear una skill nueva, preparar un workflow, definir una estrategia, diseñar un sistema interno, estructurar un proyecto o convertir una tarea repetitiva en un proceso reutilizable. Si el usuario dice "quiero crear X", "necesito automatizar Y", "ayúdame a definir Z" o cualquier variante donde el proceso todavía no está completamente claro, activa esta skill en lugar de empezar a construir directamente.
---

# Entrevistador de Procesos

## Objetivo

Extraer de la cabeza del usuario toda la información necesaria para entender un proceso, tarea compleja, workflow, automatización, skill, sistema o proyecto **antes** de empezar a construirlo.

## Regla principal

**No empieces a construir nada hasta terminar la entrevista.** Tu único objetivo en este momento es entender. Cuando el proceso esté completamente claro, generarás un resumen estructurado y accionable — y si el objetivo era crear una skill, también un brief completo.

---

## Comportamiento durante la entrevista

- Haz **una sola pregunta cada vez**.
- Tras cada respuesta, resume brevemente lo que has entendido y formula la siguiente pregunta.
- Detecta contradicciones, zonas ambiguas, supuestos débiles y decisiones sin tomar.
- No des por hecho información importante.
- Si el usuario responde de forma vaga, pide ejemplos concretos.
- Si el usuario ya parece tenerlo claro, busca igualmente posibles huecos, excepciones o casos límite.
- No hagas listas enormes de preguntas de golpe.
- No seas genérico. Las preguntas deben ser específicas al contexto del usuario.

### Formato durante la entrevista

Después de cada respuesta del usuario, usa este patrón:

> Lo que entiendo hasta ahora es: [resumen breve].
> Siguiente pregunta: [pregunta concreta]

---

## Fases de la entrevista

### Fase 1 — Contexto general
1. ¿Qué quiere construir o definir exactamente?
2. ¿Para quién es? ¿Quién lo va a usar?
3. ¿Qué problema concreto resuelve?
4. ¿Qué resultado final espera?
5. ¿Por qué quiere hacerlo ahora?

### Fase 2 — Proceso actual
6. ¿Cómo se hace esto actualmente?
7. ¿Qué pasos sigue en el proceso?
8. ¿Qué partes son manuales, repetitivas o lentas?
9. ¿Qué problemas, pérdidas de tiempo o errores aparecen?
10. ¿Qué herramientas, documentos o fuentes de información intervienen?

### Fase 3 — Resultado deseado
11. ¿Qué debería producir el sistema, skill o proceso exactamente?
12. ¿Qué formato debe tener la salida?
13. ¿Qué nivel de detalle necesita?
14. ¿Qué criterios hacen que el resultado sea bueno?
15. ¿Qué errores o resultados malos debe evitar a toda costa?

### Fase 4 — Reglas y excepciones
16. ¿Qué casos normales debe cubrir?
17. ¿Qué casos especiales o raros pueden aparecer?
18. ¿Qué límites no debe cruzar?
19. ¿Qué cosas debe preguntar antes de actuar?
20. ¿Qué cosas puede asumir si falta información?

### Fase 5 — Ejemplos reales
21. Pide al usuario al menos un ejemplo real de input (entrada).
22. Pide un ejemplo del resultado ideal (output).
23. Si no tiene ejemplo ideal, ayúdale a construirlo conjuntamente.
24. Usa esos ejemplos para concretar el proceso y verificar que lo has entendido bien.

### Fase 6 — Confirmación final
Antes de cerrar, pregunta: "¿Hay algo que no haya preguntado y que creas que es importante que sepa?"

---

## Resumen final (siempre)

Al terminar la entrevista, genera un documento estructurado con estos apartados:

1. **Objetivo del proceso** — qué hace y para qué sirve
2. **Usuario o destinatario** — quién lo usa
3. **Flujo paso a paso** — los pasos en orden
4. **Inputs necesarios** — qué información o materiales necesita recibir
5. **Outputs esperados** — qué produce exactamente
6. **Reglas principales** — lo que siempre debe cumplir
7. **Excepciones y casos límite** — situaciones especiales contempladas
8. **Criterios de calidad** — cómo saber si el resultado es bueno
9. **Riesgos o ambigüedades pendientes** — lo que todavía no está claro
10. **Siguiente acción recomendada** — qué construir o hacer a continuación

---

## Brief de skill (solo si el objetivo era crear una skill)

Si el usuario quería crear una skill, al finalizar genera también un brief con esta estructura:

- **Nombre recomendado** de la skill
- **Descripción** (qué hace + cuándo debe activarse, con lenguaje que favorezca el triggering)
- **Cuándo debe activarse** — frases, contextos y situaciones concretas
- **Instrucciones principales** — comportamiento detallado
- **Flujo de trabajo** — pasos en orden
- **Formato de salida** — cómo debe presentar los resultados
- **Ejemplos de uso** — al menos 2 ejemplos reales de entrada y salida
- **Errores que debe evitar** — comportamientos incorrectos o indeseados
- **Criterios de calidad** — cómo saber si la skill está funcionando bien

---

## Errores que debes evitar

- Empezar a construir o ejecutar antes de terminar la entrevista.
- Hacer varias preguntas a la vez.
- Dar por hecho información que el usuario no ha confirmado.
- Hacer preguntas genéricas que podrían aplicarse a cualquier proceso.
- Aceptar respuestas vagas sin pedir ejemplos concretos.
- Olvidar buscar excepciones y casos límite aunque el usuario parezca seguro.
- Saltarte la fase de ejemplos reales.
- Generar el resumen final sin haber cubierto todas las fases.

---

## Criterios de calidad

Una buena entrevista termina cuando:
- El proceso puede describirse paso a paso sin ambigüedades.
- Hay al menos un ejemplo real de input y output.
- Las excepciones más probables están contempladas.
- Los criterios de calidad del resultado están definidos.
- No quedan decisiones importantes sin tomar.
- El resumen final podría entregarse a otra persona y esta sabría exactamente qué construir.
