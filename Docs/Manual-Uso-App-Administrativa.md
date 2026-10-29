# Manual de uso — App Remuneración Quincenal UAESP

> Guía práctica para el área administrativa.
> Versión de la aplicación: 1.0.

---

## 1. Qué es la App 

1. La App arma el archivo de la **remuneración quincenal** de los 5 operadores del servicio de aseo de Bogotá.
2. **Le entra** la carpeta de fuentes del período y la **plantilla oficial**.
3. **Le sale** un archivo Excel nuevo llamado `Remuneración AAAAMM-# Total.xlsx`.
4. La App **copia la plantilla y pega solo valores**. Nunca borra ni cambia las fórmulas. La **base del período** (hojas, código de período y la columna de la quincena) **se prepara automáticamente** sobre la copia; usted no prepara nada a mano.
5. Usted **abre el archivo final en Excel** para que recalcule, lo revisa y lo guarda.

---

## 2. Qué necesita antes de empezar

Tenga a mano estas tres cosas:

| Qué | Ejemplo |
|---|---|
| Carpeta de fuentes del período | `REMUNERACION 2026072` |
| Plantilla OFICIAL (el molde, con las fórmulas) | `Formato Remuneracion 2026xx-x Total.xlsx` (en `Docs/Plantilla estandarizada/`) |
| Carpeta de salida | Una carpeta **vacía**, donde quedará el resultado |

> **Ojo:** la plantilla es el **molde en blanco** (solo estructura y fórmulas). Use siempre la **plantilla oficial**; la base del período (hojas, código de período y columna de la quincena) la prepara la App automáticamente. No use como plantilla un archivo ya completado.

La carpeta de fuentes debe tener esta estructura (organización estándar):

```
REMUNERACION 2026072/
├── 1-Promoambiental/
├── 2-Lime/
├── 3-Ciudad Limpia/
├── 4-Bogotá Limpia/
├── 5-Área Limpia/
├── Conciliaciones/          ← 5 archivos de conciliación por empresa de facturación
│   ├── Conjunta ENEL-072026.xlsx
│   ├── Conjunta ENERBIT-072026.xlsx
│   ├── Conjunta Otros -072026.xlsx
│   ├── Conjunta Recip -072026.xlsx
│   └── Directa-072026.xlsx
└── R10_Remuneracion_2026072.xlsx   ← El R10 del período, al nivel de las carpetas
                                       de los operadores
```

Cada operador tiene su carpeta:

- `1-Promoambiental`
- `2-Lime`
- `3-Ciudad Limpia`
- `4-Bogotá Limpia`
- `5-Área Limpia`

La carpeta `Conciliaciones/` debe traer los 5 archivos de conciliación, uno por empresa de
facturación: **ENEL, ENERBIT, Otros (EAAB + Ciudad Limpia), Recip (EAAB Reciprocidad) y
Directa (Banco de Occidente)**. Sus nombres empiezan con `Conjunta ENEL`, `Conjunta ENERBIT`,
`Conjunta Otros`, `Conjunta Recip` o `Directa`.

Y a nivel de la carpeta del período, junto a las carpetas de los operadores, va el archivo
**R10** del período: `R10_Remuneracion_AAAAMMQ.xlsx` (ejemplo: `R10_Remuneracion_2026072.xlsx`).

Para qué usa la App esos dos insumos, en sencillo:

- **Conciliaciones:** de cada archivo abre su pestaña `RESUMEN MES`, toma la columna de la quincena
  que está procesando (1.ª o 2.ª) y la copia tal cual en las primeras 5 pestañas del archivo final
  (`Recaudo EAAB Reciprocidad, ENEL, ENERBIT, EAAB + Ciud Limp, Directa Occidente`). No suma ni resta.
- **R10:** no se pega en ningún lado. La App calcula el total por operador con sus propias cuentas
  y al final lo **compara con el R10**: si cuadra (diferencia de centavos), prosigue; si no, avisa
  en qué operador falló.

Recuerde el período:

- **1.ª Quincena:** días 1 al 15 del mes. La carpeta termina en `1` (ejemplo: `REMUNERACION 2026071`).
- **2.ª Quincena:** días 16 al 31 del mes. La carpeta termina en `2` (ejemplo: `REMUNERACION 2026072`).

---

## 3. Ejemplo completo, paso a paso

Ejemplo real: período **202607**, **2.ª quincena**, en modo **Procesar todos**.
Al final, el archivo se llama `Remuneración 202607-2 Total.xlsx`.

### Paso 0 — Abra la App

Abra la App. Verá tres tarjetas: **Período**, **Rutas** y **Ejecución**.

### Paso 1 — Elija el período

En la tarjeta **Período**:

- **Año:** elija `2026`.
- **Mes:** elija `Julio`.
- **Quincena:** elija `2.ª Quincena`.

Qué debe ver: bajo el título de la ventana aparece `Período 2026072 · 2.ª Quincena`.

### Paso 2 — Carpeta de fuentes

En la tarjeta **Rutas**:

- Clic en el botón **`...`** que está junto a "Carpeta fuentes:".
- Seleccione la carpeta `REMUNERACION 2026072` y confirme.

Qué debe ver: la caja "Carpeta fuentes:" muestra la ruta completa.

### Paso 3 — Plantilla

En la tarjeta **Rutas**:

- Clic en el botón **`...`** junto a "Plantilla:".
- Seleccione la **plantilla oficial** — `Formato Remuneracion 2026xx-x Total.xlsx` (en `Docs/Plantilla estandarizada/`) — y confirme. La App prepara la base del período sola: no necesita una plantilla por período.

Qué debe ver: la caja "Plantilla:" muestra la ruta del archivo.

### Paso 4 — Carpeta de salida

En la tarjeta **Rutas**:

- Clic en el botón **`...`** junto a "Carpeta salida:".
- Seleccione la carpeta de salida y confirme.

Qué debe ver: la caja "Carpeta salida:" muestra la ruta de la carpeta.

### Paso 5 — Elija el modo

En la tarjeta **Ejecución**:

- Marque la casilla **Procesar todos**.

Qué debe ver: el combo **ASE:** se apaga (queda deshabilitado). Es normal: "Procesar todos" usa los 5 operadores.

### Paso 6 — Ejecute

- Clic en el botón azul **▶ Ejecutar**.

Qué debe ver: la barra de progreso empieza a llenarse. Debajo de la barra, el texto dice `Iniciando…` y luego va nombrando el operador que está procesando. A la derecha de la barra sube el porcentaje: `0 %`, `25 %`, `50 %` y así hasta `100 %`.

### Paso 7 — Confirme que terminó

Qué debe ver: la barra llega a `100 %`. Abajo, en la barra de estado, aparece `Completado — archivo listo en <su carpeta de salida>`.

También: el botón **Ver logs** se activa (antes estaba apagado).

### Paso 8 — Abra el resultado en Excel

- Clic en **Abrir salida** para ir a la carpeta.
- Abra `Remuneración 202607-2 Total.xlsx` en Excel.
- Deje que Excel recalcule (verá que las fórmulas se completan), revise y guarde.

---

## 4. Los 4 botones

| Botón | Para qué sirve | Cuándo usarlo |
|---|---|---|
| **▶ Ejecutar** | Procesa el período y genera el archivo de salida. | Cuando ya eligió el período y las 3 rutas. |
| **Limpiar** | Deja todo en cero: vacía las 3 rutas, vuelve el período a hoy y el ASE al primero. | Cuando terminó un caso y quiere empezar otro desde cero. No borra archivos. |
| **Ver logs** | Abre la ventana **Detalle técnico — logs** con el detalle paso a paso. Dentro tiene **Copiar** y **Cerrar**. | Solo se activa al terminar. Úselo para enviar el detalle a soporte si algo salió mal. |
| **Abrir salida** | Abre la carpeta de salida en el Explorador. | Para buscar el archivo que se generó. |

---

## 5. Si algo sale mal

| Qué ve usted | Qué pasó | Qué hace usted |
|---|---|---|
| Mensaje "Debe seleccionar la carpeta de fuentes, la plantilla y la carpeta de salida antes de ejecutar." | Le falta una de las 3 rutas. | Confirme el mensaje y complete las tres rutas. |
| Aviso de que la salida no puede ser la plantilla | La ruta de salida coincide con la plantilla. | Elija otra carpeta de salida, distinta a la plantilla. |
| Pregunta "El archivo '...' ya existe. ¿Desea sobrescribirlo?" | Ya hay un archivo con ese nombre en la carpeta de salida. | **Sí** para reemplazarlo, **No** para cancelar y revisar antes. |
| Mensaje "Seleccione primero una carpeta de salida válida." (al usar Abrir salida) | No hay carpeta de salida o la ruta ya no existe. | Elija de nuevo la "Carpeta salida:" con el botón `...`. |
| Mensaje "Todavía no hay contenido para copiar." (dentro de Ver logs) | Abrió Ver logs sin haber ejecutado nada. | Cierre la ventana, ejecute un período y vuelva a intentar. |
| Mensaje "No se encontró el archivo de conciliación de …" (nombra la empresa) | Falta uno de los 5 archivos dentro de `Conciliaciones/`. | Verifique que la carpeta del período tenga la carpeta `Conciliaciones/` con los 5 archivos y vuelva a ejecutar. |
| Mensaje "No se encontró R10_Remuneracion_…" (nombra el período y la carpeta) | Falta el archivo R10 en la carpeta del período. | Verifique que el R10 del período esté junto a las carpetas de los operadores y vuelva a ejecutar. |
| Mensaje de diferencias **por operador** entre el DetRetri calculado y el R10 (nombra el período, el archivo y los dos valores) | El total por operador que calculó la App no cuadra con el R10 del período. | Revise la novedad en **Ver logs**. Si la corrida quedó marcada `(sin validar)`, consulte con soporte **antes de entregar** el archivo. |
| Mensaje de integridad del pegado del R10 (nombra una **hoja** y una **celda** con el valor esperado y el escrito) | La App copia las hojas `DetRetri`/`DetValiRetri` del R10 del período y, al releerlas, una celda no quedó como el R10. | No entregue el archivo: consulte con soporte indicando la hoja y la celda del mensaje. **Puede volver a ejecutar** sobre una copia limpia de la plantilla oficial. |
| Mensaje de validación que **nombra una hoja y una celda** (ejemplo: `ASE 5 · VALIDACION_TOTAL · C15`) | La App revisa que las validaciones del archivo cierren con los valores que ella misma escribió, y encontró una que **no cierra**. | **No es un defecto de la corrida: la App falla a propósito** para no entregar un archivo con validaciones en falso. Anote la hoja y la celda que nombra el mensaje y avise a soporte con el detalle de **Ver logs**. |
| Ventana **Divergencia declarada — reconocimiento para esta corrida** | La App encontró una **diferencia conocida** entre el reporte oficial y el manual del administrativo, cuya decisión no está en manos de la App. No es un error del cálculo: es una deuda de otro. | Lea el monto, de quién es (responsable) y la evidencia. Si acepta cerrar la corrida con esa diferencia, marque la casilla **"Acepto cerrar la corrida con la divergencia…"** y pulse **Aceptar y cerrar**. Si no la acepta, pulse **Volver sin cerrar** (la corrida no se cierra). |
| Tras **Volver sin cerrar**, la barra de estado dice `No completado — pendiente de reconocimiento explícito` | No aceptó la diferencia, así que la corrida no cerró y no se entregó el archivo. | Vuelva a ejecutar cuando quiera revisarla: la App **vuelve a mostrar la misma diferencia** (nunca la acepta sola ni la recuerda de una corrida a otra). |
| Aparece un archivo **`Remuneración AAAAMM-# Total - Divergencias declaradas.txt`** junto al Excel de salida | Es el **reporte de divergencias declaradas**: el registro que la App deja cuando usted acepta cerrar con una diferencia declarada. Dice qué se aceptó, el monto, de quién es y por qué, y desde cuándo deja de aplicar. | Consérvelo **junto al archivo de salida** y entréguelo con él. **Aceptar no es "arreglar"**: la diferencia sigue viva hasta que el responsable (UAESP) la resuelva. No la dé por corregida. |

---

## 6. Lista de chequeo final

- [ ] El período es el correcto (Año, Mes y Quincena).
- [ ] Las 3 rutas están seleccionadas (fuentes, plantilla oficial y salida).
- [ ] La carpeta del período tiene: las 5 carpetas de operadores + la carpeta `Conciliaciones/` con sus 5 archivos + el archivo R10 del período.
- [ ] El modo es el correcto (**Procesar todos** o un solo ASE).
- [ ] La barra llegó a `100 %` y abajo dice `Completado`.
- [ ] El archivo tiene el nombre oficial `Remuneración AAAAMM-# Total.xlsx`.
- [ ] Abrí el archivo en Excel, recalculó y lo guardé.

---

## 7. Cierre: qué entregar y a quién

- **Entregue** el archivo de salida `Remuneración AAAAMM-# Total.xlsx` a la persona responsable del área que solicitó la liquidación.
- **Si todo salió bien:** avise que el archivo está listo e indique el período.
- **Si hubo un error:** avise a soporte con el detalle. Para eso: clic en **Ver logs**, luego en **Copiar**, y pegue ese texto en el correo o mensaje.
- **Si tiene dudas del resultado:** revise el archivo en Excel antes de entregarlo.
