# powerpoint-templates
Presentaciones PowerPoint editables y personalizadas

## Business Review

[`Business_Review_Presentation.pptx`](Business_Review_Presentation.pptx) contiene
las ocho diapositivas en el orden solicitado, sin numeración visible. Los textos,
tarjetas, métricas y tablas son objetos nativos editables en PowerPoint, no imágenes.

### Generar y actualizar cada mes

Requiere Python 3.9 o posterior:

```bash
cd /home/runner/work/powerpoint-templates/powerpoint-templates
python -m pip install -r requirements.txt
python generate_presentation.py
```

En `generate_presentation.py`, modifica el bloque inicial **Monthly content**:

- `PERIOD`: periodo del informe.
- `SLIDES`: valores, porcentajes, fechas, textos, estados y tablas de cada slide.
- `EVENTS`, `CONSULTANTS`, `COUNTRIES` y `SECTORS`: listas compartidas.
- `COLORS`: paleta de la presentación.

Los valores se guardan como texto para conservar formatos como `€22k+`,
`~€100k` y `35 / 80`. Actualiza también los porcentajes, totales y las menciones
repetidas en las tarjetas cuando cambien: no se recalculan automáticamente.
Se conservan los datos proporcionados, incluido `5 Strategic Sectors` frente
a seis sectores descritos y los pipelines de `€93k` y `€92k` en sus respectivas slides.

Ejecuta nuevamente el script para sobrescribir el `.pptx` junto al generador,
independientemente del directorio desde el que lo ejecutes. Guarda antes cualquier
edición manual de PowerPoint que quieras conservar: la regeneración no la importa.
Para incorporar otra slide, añade una entrada a `SLIDES`; su posición en la lista
determina el orden. Puedes modificar las tarjetas y su número de columnas sin
cambiar el renderizador. Si añades más texto, revisa que quepa en la diapositiva.

El diseño utiliza las descripciones suministradas y la paleta indicada; no se
disponía de las imágenes originales para reproducirlas exactamente. Los tres
highlights de pipeline son textos descriptivos basados en los stages proporcionados.
