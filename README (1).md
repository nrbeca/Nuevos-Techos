# Distribución del Nuevo Techo AC01 por UR

App en Streamlit para distribuir el Nuevo Techo de cada Unidad Responsable (UR)
entre sus partidas, con base en el peso (%) de cada partida dentro del
Monto Anual de esa UR.

## Qué hace

1. Subes el Excel base (mismo formato que `Base_para_calculo_AC01_Central_y_OREF.xlsx`),
   con las hojas:
   - **Base Estimación**: detalle de partidas por UR (col. A `UR` ... col. O `Monto Anual`).
   - **Nuevos Techos**: techo nuevo por UR (col. A `UR`, col. B `Nuevo Techo`).
2. Seleccionas la UR en el menú lateral.
3. La app calcula:
   - Total del Monto Anual de la UR.
   - Porcentaje de cada partida dentro del total.
   - Nuevo Techo por partida = Nuevo Techo de la UR × Porcentaje.
   - Nuevo techo redondeado = `ROUND(Nuevo Techo, 0)`.
4. Descargas un Excel con **fórmulas vivas** (no valores fijos) para poder
   verificar el cálculo directamente en Excel, igual que la hoja "Ejemplo"
   original.

## Archivos

- `app.py` — código de la app.
- `requirements.txt` — dependencias.

## Cómo desplegar en Streamlit Community Cloud

1. Crea un repositorio en GitHub (o usa uno existente, p. ej. `nrbeca/ac01`)
   y sube `app.py` y `requirements.txt` a la raíz del repo.
2. Entra a https://share.streamlit.io y da clic en **New app**.
3. Selecciona el repositorio, la rama (`main`) y el archivo principal `app.py`.
4. Deploy. Streamlit instalará automáticamente lo que está en `requirements.txt`.

No requiere secretos ni tokens: todo el procesamiento ocurre en memoria a
partir del archivo que subas en cada sesión.

## Uso local (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```
