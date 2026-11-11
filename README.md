# Contratos del Ayuntamiento de Pinto

Descarga diaria y automática de los contratos del Ayuntamiento de Pinto (licitaciones y contratos menores) a partir de los **datos abiertos de la Plataforma de Contratación del Sector Público** (Ministerio de Hacienda).

- **Excel actualizado:** pestaña *Releases* → `datos` → `contratos_pinto.xlsx`
  Enlace fijo: `https://github.com/<USUARIO>/<REPOSITORIO>/releases/download/datos/contratos_pinto.xlsx`
- **Datos en bruto:** `datos_contratos_pinto.json` (última versión de cada expediente).

## Cómo funciona

El flujo `.github/workflows/actualizar.yml` se ejecuta cada día:

1. Lee los feeds ATOM de la Plataforma (licitaciones y menores).
2. Se queda con los órganos del Ayuntamiento de Pinto.
3. Actualiza `datos_contratos_pinto.json`.
4. Regenera el Excel y lo publica en la release `datos`.

## Carga inicial del histórico

*Actions* → *Actualizar contratos de Pinto* → *Run workflow* → escribir el año de inicio (p. ej. `2018`) → *Run workflow*.

## Ajustar el filtro de órganos

Al principio de `pinto_contratos.py`: `FILTRO_NOMBRE`, `IDS_ORGANOS` y `EXCLUIR_NOMBRES`.
