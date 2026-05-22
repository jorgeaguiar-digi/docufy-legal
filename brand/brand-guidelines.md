# Docufy Legal — Brand Guidelines

## Paleta de colores

| Nombre | Hex | Uso |
|---|---|---|
| Navy (primario) | `#0A0E2E` | Fondo principal de todas las piezas |
| Gold (acento) | `#C9A85C` | Nombre de marca, datos destacados, separadores, números clave |
| White | `#FFFFFF` | Textos principales sobre fondo navy |
| Light (off-white) | `#E8ECF5` | Textos secundarios / body sobre fondo navy |
| Light Gray | `#EBEBEB` | Fondo neutro claro (reverso tarjeta, variante clara) |
| Red Accent | `#E83A3A` | Uso puntual para señalar el problema / estado negativo |
| Green Accent | `#4CD97A` | Uso puntual para señalar la solución / estado positivo |
| Dark Card | `#12183C` | Cajas interiores sobre fondo navy (contraste sutil) |

### En python-pptx
```python
NAVY    = RGBColor(0x0A, 0x0E, 0x2E)
GOLD    = RGBColor(0xC9, 0xA8, 0x5C)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT   = RGBColor(0xE8, 0xEC, 0xF5)
LGRAY   = RGBColor(0xEB, 0xEB, 0xEB)
RED_ACC = RGBColor(0xE8, 0x3A, 0x3A)
GREEN   = RGBColor(0x4C, 0xD9, 0x7A)
CARD    = RGBColor(0x12, 0x18, 0x3C)
```

---

## Logotipo

- **Ícono:** columna clásica (referencia jurídica) con una curva de carretera atravesándola (referencia al tráfico). Comunica la intersección de derecho + accidentes de tráfico.
- **Tipografía del logo:** "DOCUFY" en mayúsculas, sans-serif geométrica bold (similar a Montserrat ExtraBold). "LEGAL" en mayúsculas, tracking amplio, peso regular.
- **Versiones:**
  - Sobre fondo navy: ícono gold + "DOCUFY" white + "LEGAL" white
  - Sobre fondo claro: ícono gold + "DOCUFY" navy + "LEGAL" navy
- **Posición en slides:** esquina superior derecha, o centrado en slides de apertura/cierre.
- **Archivos:** `brand/Docufy Legal - logo (3).png`

---

## Tipografía

| Uso | Estilo | Características |
|---|---|---|
| Títulos principales | Sans-serif bold, mayúsculas | Tamaño grande, tracking normal |
| Subtítulos / claims | Sans-serif bold o regular | Mezcla mayúsculas y minúsculas |
| "LEGAL" del logo | Mayúsculas, tracking amplio (+200) | Peso ligero o regular |
| Body / descripción | Sans-serif regular | Legible, tamaño 18-22pt en slides |
| Datos/stats | Sans-serif extrabold | Muy grande (48-72pt), color gold |

**Familias recomendadas para slides:** Montserrat (títulos) + Open Sans o Lato (body).

---

## Estilo visual

- **Mood:** profesional, premium, confianza. No tecnológico frío — cercano y directo.
- **Contraste alto:** fondo navy muy oscuro con texto blanco y acentos dorados.
- **Jerarquía clara:** un mensaje dominante por slide, apoyado por un dato o subtexto.
- **Fotografía:** profesionales del sector legal (mujer/hombre en entorno de despacho), estilo editorial, tono neutro-cálido. Usada como elemento de soporte, no protagonista.
- **Sin decoración excesiva:** no hay gradientes, sombras complejas ni patrones. El diseño es limpio y estructurado.
- **Separadores dorados:** líneas horizontales finas en gold (`rect` de 0.07" de alto) para dividir secciones dentro de un slide.
- **Barra lateral gold:** rectángulo vertical fino (0.12" de ancho) en el borde izquierdo de slides de desarrollo/contenido.

---

## Elementos de copy recurrentes

### Claims y mensajes de marca
- "Automatiza tu despacho de accidentes de tráfico"
- "El 80% no es Derecho. Es perseguir al cliente."
- "De caos a pro en 3 pasos."
- "Docufy Legal lo automatiza. Hasta 80% menos tiempo administrativo."

### Métricas clave (usar en slides de resultado)
- **−80%** tiempo por expediente
- **15-25h** al mes recuperadas
- **30 días** para implementarlo
- **80%** de clientes entregan documentación completa en menos de 48h
- **De 30 min** por escrito **a menos de 10 segundos**
- **−90%** errores humanos en documentación enviada a aseguradoras

### Estructura narrativa del flyer (probada)
1. EL PROBLEMA — headline impactante
2. LA SOLUCIÓN — 3 pasos numerados
3. EL RESULTADO — 3 métricas en columnas

---

## Datos de contacto (usar en slide de cierre)
- **Web:** docufylegal.com
- **Email:** info@docufylegal.com
- **Teléfono:** +34 681 801 356
- **Persona:** Jorge Aguiar

---

## Archivos de referencia en esta carpeta

| Archivo | Descripción |
|---|---|
| `Docufy Legal - logo (3).png` | Logo sobre fondo navy (versión principal) |
| `Tarjeta Visita.png` | Tarjeta anverso — logo sobre fondo claro |
| `Tarjeta Visita (1).png` | Tarjeta reverso — datos de contacto sobre navy |
| `flyer docufy.png` | Flyer anverso — hook + CTA con fotografía |
| `flyer docufy (1).png` | Flyer reverso — problema / solución / resultado |
