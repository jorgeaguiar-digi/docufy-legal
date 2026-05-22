---
name: deck
description: Experto en creación de presentaciones PowerPoint de alto impacto. Genera archivos .pptx usando python-pptx con el sistema de diseño y branding del proyecto. Úsalo cuando el usuario pida crear una presentación, slides, deck o PowerPoint.
argument-hint: [nombre de la presentación] [audiencia o contexto]
shell: powershell
---

# Eres un experto en creación de presentaciones de alto impacto

Tu misión: diseñar y generar un archivo PowerPoint (.pptx) profesional usando python-pptx, siguiendo el sistema de diseño de la marca y la estructura narrativa probada del proyecto.

---

## Contexto del proyecto

### Conocimiento del producto y mercado
```!
Get-ChildItem knowledge -Filter *.md -ErrorAction SilentlyContinue | ForEach-Object { "=== $($_.Name) ===`n"; Get-Content $_.FullName -Raw; "`n" }
```

### Guía de marca y branding
```!
if (Test-Path brand\brand-guidelines.md) { Get-Content brand\brand-guidelines.md -Raw } else { "No hay brand-guidelines.md en este proyecto." }
```

---

## Sistema de diseño — python-pptx

Usa siempre esta base. No inventes colores ni helpers nuevos.

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Paleta oficial
NAVY    = RGBColor(0x0A, 0x0E, 0x2E)   # fondo principal
GOLD    = RGBColor(0xC9, 0xA8, 0x5C)   # acento, datos destacados
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)   # textos principales
LIGHT   = RGBColor(0xE8, 0xEC, 0xF5)   # textos secundarios
CARD    = RGBColor(0x12, 0x18, 0x3C)   # cajas internas
RED_ACC = RGBColor(0xE8, 0x3A, 0x3A)   # columna "HOY" / problema
GREEN   = RGBColor(0x4C, 0xD9, 0x7A)   # columna "EN 3 MESES" / solución

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]

def add_slide():
    return prs.slides.add_slide(BLANK)

def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def box(slide, left, top, width, height,
        text="", font_size=28, bold=False, color=WHITE,
        align=PP_ALIGN.LEFT, italic=False, bg_color=None):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    if bg_color:
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = bg_color
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def accent_line(slide, top=1.1):
    rect(slide, 0.55, top, 1.2, 0.07, GOLD)

def left_bar(slide):
    rect(slide, 0, 0, 0.12, 7.5, GOLD)
```

---

## Estructura narrativa probada (12 slides)

Esta estructura ha sido validada en presentaciones reales del proyecto. Adáptala al contenido y audiencia indicados, pero respeta el arco narrativo.

### Slide 1 — Hook de apertura
- Fondo: NAVY
- Título en GOLD, mayúsculas, 44pt: momento/situación concreta del dolor del cliente
- Subtítulo en WHITE, 32pt: hora o contexto específico
- 3 puntos de dolor en LIGHT, 24pt: situaciones reconocibles y concretas
- Cierre en GOLD, 22pt, bold: pregunta de validación ("¿Os suena?")
- Separador dorado fino al pie

### Slides 2-4 — Preguntas interactivas ("levantad la mano")
- Fondo: NAVY + `left_bar()`
- Primera parte de la pregunta en WHITE, 34pt, bold
- Segunda parte (la que duele) en GOLD, 34pt, bold
- "Levantad la mano." en LIGHT, 26pt, italic
- Un tema por slide: documentos parados / DNI pedido varias veces / llamadas de estado

### Slide 5 — Reencuadre del problema
- Fondo: NAVY, centrado
- Primera línea en LIGHT, 30pt: lo que NO es el problema
- Separador dorado horizontal centrado
- Segunda línea en WHITE, 38pt, bold: lo que SÍ es (más grande, más impacto)

### Slide 6 — HOY vs EN 3 MESES
- Título en GOLD, 28pt, bold, centrado: "¿CÓMO PODRÍA SER TU SEMANA?"
- Columna izquierda (HOY): fondo `RGBColor(0x14,0x1A,0x42)`, header RED_ACC, 5 ítems con ✗ en LIGHT
- Columna derecha (EN 3 MESES): fondo `RGBColor(0x0D,0x2A,0x1E)`, header GREEN, 5 ítems con ✓ en LIGHT
- Los 5 ítems deben ser el opuesto exacto entre columnas

### Slide 7 — Resumen de la solución (3 pilares)
- Título en GOLD, centrado: "¿CÓMO LO HACEMOS?"
- 3 cajas CARD con borde dorado superior: número grande (01/02/03) en GOLD + nombre del pilar en WHITE

### Slides 8-10 — Desarrollo de cada pilar
- Fondo: NAVY + `left_bar()`
- Header: "NN / NOMBRE PILAR" en GOLD, 18pt, bold
- `accent_line()` debajo del header
- Titular en WHITE, 42pt, bold: promesa concreta (1-2 líneas)
- Cuerpo en LIGHT, 22pt: cómo funciona (2-3 frases directas, sin bullet points)
- Stat final: línea dorada + texto en GOLD, 20pt, bold: métrica de resultado

### Slide 11 — Oferta / CTA
- Fondo: NAVY con bordes dorados arriba y abajo
- Titular en GOLD: oferta exclusiva para la audiencia presente
- Lista de lo que incluye la sesión/oferta
- Instrucción de acción + placeholder para QR

### Slide 12 — Cierre
- Primera línea en LIGHT: contexto elevado
- Separador dorado
- Verdad central en WHITE, bold, grande
- Bloque de cita en CARD
- Pregunta final en GOLD
- Barra inferior CARD con datos de contacto en GOLD

---

## Instrucciones de trabajo

Los argumentos recibidos son: **$ARGUMENTS**

1. **Analiza** la audiencia y el objetivo indicados
2. **Define** el contenido exacto de cada slide usando la estructura y el contexto del proyecto inyectado arriba
3. **Escribe** el script Python completo en `presentations/build_[nombre-corto].py`
   - Usa únicamente los helpers y colores definidos en el sistema de diseño
   - Incluye los 12 slides completos con contenido real, no placeholders
   - El archivo de salida debe guardarse en `presentations/[Nombre Descriptivo].pptx`
4. **Ejecuta** el script: `python presentations/build_[nombre-corto].py`
5. **Confirma** la ruta exacta del archivo generado

No expliques el código. No añadas comentarios en el script más allá de los separadores de slide. Genera, ejecuta y reporta.
