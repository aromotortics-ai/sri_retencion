# Tablero de conciliación en vivo y reorganización del proyecto

Fecha: 2026-10-27

## Intención y por qué

Contabilidad concilia a mano comprobantes del SRI contra Odoo. La app web ya
cruza los datos, pero el flujo actual es "subir todo → esperar → descargar un
Excel → abrirlo para ver qué pasó". Contabilidad quiere **ver en la misma
pantalla** qué va conciliando mientras sube los TXT, y **descargar un Excel
de observaciones** que se reparte a los departamentos para que corrijan en
Odoo. Todo lo revisa contabilidad; el Excel es solo el medio para comunicar
las correcciones.

Además, el código está desordenado (unos 25 `.py` sueltos en la raíz, Excels y
TXT de prueba mezclados, un flujo antiguo que ya nadie necesita). Se aprovecha
para ordenarlo.

## Objetivos

1. Dos pantallas, **Compras** y **Ventas**, con el mismo tablero en vivo:
   - Zona para soltar/seleccionar TXT, con la lista de archivos cargados y el
     estado de cada uno (en cola, procesando, listo, con error).
   - Tarjetas por tipo de comprobante (Facturas, Notas de crédito, Notas de
     débito, Liquidaciones —solo Ventas—, Retenciones) con "N de M coinciden"
     y barra de avance; se actualizan con cada TXT que entra.
   - Tabla de excepciones (todo lo que no es "Coincide") con filtros por
     estado y por tipo, y búsqueda por número de documento.
   - Botón **Descargar observaciones (Excel)**.
2. Reorganizar el repositorio en carpetas coherentes (ver "Estructura").
3. Retirar la pestaña antigua "Retenciones" y los scripts que solo ella usa.

## No objetivos

- No se guarda historial entre sesiones ni se mantiene estado por usuario.
- No hay botón "Volver a conciliar": tras corregir en Odoo, el usuario vuelve
  a subir los TXT (decisión del usuario).
- No hay marcado manual de casos ("revisado", notas por fila): la cola de
  revisión (opción B del boceto) se descartó.
- No se cambia la lógica de cruce, los estados ni las reglas de elegibilidad
  descritas en `CLAUDE.md`. Este trabajo cambia presentación, flujo de carga y
  organización.
- Sin cambios en Odoo (la lectura sigue siendo XML-RPC solo lectura).

## Anti-objetivos (sería un fracaso aunque "funcione")

- Que el resultado en pantalla y el Excel descargado muestren cifras
  distintas.
- Que se pierda algún estado, hoja o regla existente de Compras/Ventas al
  reorganizar.
- Que el despliegue en Easypanel (Docker) deje de arrancar por rutas movidas.

## Decisiones del usuario

- **Pantalla = fotografía del momento.** Se suben los TXT, se concilia contra
  Odoo tal como está ahora y se muestra el resultado. Cuando el departamento
  corrige en Odoo, se **vuelven a subir los TXT** y lo corregido aparece como
  "Coincide". Odoo se lee de nuevo cada vez.
- **Disposición "A – Tablero en vivo"** (un solo panel: carga arriba, tarjetas
  por tipo, tabla de excepciones y botón de descarga abajo).
- Solo dos pantallas: Compras y Ventas. La pestaña "Retenciones" antigua (TXT
  + `reten_odoo.xlsx`) y sus scripts se retiran. Las retenciones recibidas
  siguen dentro de Compras y las emitidas dentro de Ventas.
- Reorganizar tanto las carpetas como las pantallas.

## Comportamiento acordado

- Cada TXT que se agrega **recalcula** el resultado con todo lo cargado hasta
  ese momento (el conjunto completo, no solo el nuevo). Mientras recalcula, la
  pantalla conserva el resultado anterior y marca que está actualizando.
- Quitar un TXT de la lista también recalcula.
- Los límites actuales de cantidad de TXT por envío se mantienen (Compras 5,
  Ventas 400).
- Las advertencias que hoy van en el Excel (falta un tipo, período largo,
  XML no disponible, retenciones sin clave) se muestran también en pantalla,
  arriba del tablero.
- Estados y colores: los mismos de `CLAUDE.md`.

## Excel de observaciones (propuesta mía, el usuario no la contradijo)

Contiene **solo las excepciones**, no los "Coincide". Una hoja por tipo de
problema (Solo en SRI, Solo en Odoo, Diferencias, Factura incorrecta, Anulados,
XML no disponible…), cada fila con documento, fecha, proveedor/cliente,
valores SRI y Odoo, y una columna **"Qué corregir en Odoo"** con texto claro
(por ejemplo "Registrar la retención en Odoo", "Corregir el número de factura
sustento"). El Excel completo actual (con fórmulas y todas las hojas) se
conserva como segunda descarga.

## Estructura de carpetas (propuesta mía; el detalle lo decide quien construya)

```
app/
  web/         FastAPI, autenticación, plantillas, estáticos
  sri/         sri_session, descarga de XML, lectura de TXT, caché
  odoo/        odoo_client
  compras/     cruce, Excel y pipeline de Compras (incluye retenciones recibidas)
  ventas/      cruce, Excel y pipeline de Ventas (incluye retenciones emitidas)
  retenciones/ lógica compartida de retenciones (cruce y Excel por documento)
scripts/       CLI que se conservan: conciliar_emitidas.py, mover_emitidos.py,
               setup_usuario.py
tests/         misma estructura que app/
docs/
```

Excels y TXT de prueba (`test_*.xlsx`, `verify_output.xlsx`, `claves_*.txt`,
etc.) salen de la raíz o se eliminan si son basura; `cache/`, `jobs/` y `data/`
siguen siendo carpetas de datos montadas como volumen. Hay que actualizar
`Dockerfile`, `CLAUDE.md`, `README` si existe, e imports de los tests.

## Retiro del flujo antiguo

Se retira la pestaña "Retenciones" y la ruta `POST /upload` junto con lo que
solo ella usa (`pipeline.py`, `build_workbook_final.py`, `conciliar_odoo.py` y,
tras comprobarlo por imports, `parse_recibidos_txt.py`,
`descargar_comprobantes_sri.py`, `parse_comprobante_xml.py`,
`conciliar_recibidas.py` si nadie más los importa). Se **conserva**
`conciliar_emitidas.py` (CLI independiente) y `mover_emitidos.py`. Cualquier
módulo que Compras/Ventas importe se queda. Se actualiza `CLAUDE.md`.

## Restricciones

- Python 3.12, FastAPI, despliegue Docker en Easypanel; variables `ODOO_*` en
  el entorno. Lectura de Odoo solo lectura, sin credenciales en errores.
- Convenciones de Excel de `CLAUDE.md` (fórmulas, Arial, recálculo con
  LibreOffice) aplican al Excel completo.
- El repositorio tiene cambios sin confirmar en `pipeline_compras.py`,
  `pipeline_ventas.py`, `xml_cache.py` y sus tests; deben confirmarse o
  guardarse antes de mover archivos.
- Las pruebas existentes (`python -m pytest tests -q`) deben seguir pasando
  tras la reorganización.

## Decisiones dejadas a quien construya

- Mecanismo técnico de actualización en vivo (sondeo de `/status`, SSE, etc.)
  y de recálculo al agregar/quitar TXT (p. ej. un job nuevo con el conjunto
  completo de archivos del usuario en la sesión).
- Nombres finales de paquetes y archivos, y cómo repartir el código común.
- Diseño visual fino dentro de la disposición A; si usar o no un framework JS
  (hoy es HTML + JS simple).
- Redacción exacta de la columna "Qué corregir en Odoo".

## Decisiones tomadas por mí tras la revisión del constructor (marcadas para que las revises)

- **Dónde vive el conjunto de TXT:** en el servidor, en un job por sesión de trabajo del usuario. El navegador sube cada TXT una sola vez y luego agrega/quita por id; no se re-suben los 400 archivos en cada cambio. El job sigue limpiándose a las 2 h. No se exige aislamiento entre usuarios más allá de la autenticación actual (cookie de sesión); varios contadores a la vez usan jobs distintos.
- **Costo del recálculo:** se agrupan (debounce) las altas rápidas en un solo recálculo, se cancela el recálculo superado por uno más nuevo, y Odoo se vuelve a leer en cada recálculo para reflejar correcciones recientes. Los XML de retenciones siguen usando el caché compartido. Sin meta de latencia fija; la pantalla muestra "actualizando" y conserva el resultado anterior.
- **Estado por archivo:** indica solo validez de lectura (TXT ilegible o de tipo desconocido = error) más cuántas filas aportó. Un archivo con error se excluye con una advertencia y el resto se concilia igual.
- **Una sola fuente de cifras:** pantalla, Excel de observaciones y Excel completo salen de la misma estructura de resultado en Python; la pantalla nunca lee el Excel recalculado. Las tarjetas cuentan documentos, igual que el Resumen actual.
- **Qué entra en el Excel de observaciones:** todo lo accionable, incluido "XML no disponible" (volver a intentar luego). **No** entran "Coincide" ni "Informativo" ni "Duplicado anulado" (no requieren corrección); sí se ven en la tabla de pantalla bajo su filtro. "Duplicado en Odoo" sí entra (hay que anular uno).
- **Columna "Qué corregir en Odoo":** texto fijo por estado, más el detalle de diferencia solo donde el cruce ya lo expone (sin tocar la lógica de cruce). Un solo archivo para todos los departamentos, con la columna de estado para filtrar.
- **Reorganización:** `cache/`, `jobs/` y `data/` se quedan en la raíz (volúmenes de Easypanel); el nombre `app/` es sugerencia, no obligación.
