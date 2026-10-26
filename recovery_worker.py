"""Worker aislado para Sincronización de Recuperación.

Se ejecuta como subproceso desde sync_manager para aislar completamente COM/Excel.
Recibe un JSON por stdin con {db_path, tmp_excel, sheet, table} y hace:

  1. Leer de SQLite las filas modificadas hoy (timestamp_modificacion).
  2. Abrir el Excel temporal con live hooking (xw.Book) en una ventana VISIBLE.
  3. UPDATE de filas existentes e INSERT de filas nuevas al final.
  4. Guardar y cerrar limpiamente (sin instancias invisibles).

Imprime un JSON por stdout: {ok, insertadas, actualizadas, error}.
"""

import json
import sys

import database_manager
import etl_processor
import fpa_validations

_CLAVE_DB = ("cedis", "sf", "planta_obra", "articulo", "modalidad")
_DB_TO_EXCEL_POS = {v: k for k, v in fpa_validations.EXCEL_TO_DB_COLUMNS.items()}


def _norm_clave(fila):
    return tuple(
        fpa_validations._norm_modalidad_db(fila.get(c)) if c == "modalidad"
        else fpa_validations._norm_id(fila.get(c))
        for c in _CLAVE_DB
    )


def main():
    resultado = {"ok": False, "insertadas": 0, "actualizadas": 0, "error": ""}
    try:
        payload = json.loads(sys.stdin.read())
        db_path = payload["db_path"]
        tmp_excel = payload["tmp_excel"]
        sheet_name = payload.get("sheet", etl_processor.EXCEL_SHEET_NAME)
        table_name = payload.get("table", etl_processor.EXCEL_TABLE_NAME)
    except Exception as e:
        resultado["error"] = f"Payload inválido: {e}"
        print(json.dumps(resultado))
        return

    import xlwings as xw

    conn = database_manager.create_connection(db_path)
    filas_hoy = database_manager.get_rows_modificadas_hoy(conn) if conn else []
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass

    wb = None
    try:
        # Live hooking: usa el libro ya abierto (si existe) o ábrelo VISIBLE.
        # Evita instancias invisibles (xw.App(visible=False)), que son las que
        # quedan colgadas con procesos EXCEL.EXE zombi.
        wb = xw.Book(tmp_excel)
        app = wb.app

        # Suspender refresco/alertas mientras escribimos; se restaura en finally.
        try:
            app.display_alerts = False
            app.screen_updating = False
        except Exception:
            pass

        sht = wb.sheets[sheet_name]

        lo = None
        for candidate in sht.api.ListObjects:
            if candidate.Name == table_name:
                lo = candidate
                break
        if lo is None:
            raise RuntimeError(f"Tabla '{table_name}' no encontrada en el Excel descargado.")

        first_row = lo.Range.Row
        last_row = lo.Range.Row + lo.Range.Rows.Count - 1
        last_col = lo.Range.Column + lo.Range.Columns.Count - 1

        data = sht.range((first_row + 1, 1), (last_row, last_col)).value
        if data is None:
            data = []
        elif last_row == first_row + 1:
            data = [data]

        excel_index = {}
        for offset, vals in enumerate(data):
            if vals is None:
                continue
            clave = tuple(
                fpa_validations._norm_id(vals[fpa_validations.EXCEL_ID_POS[c] - 1])
                if c != "modalidad"
                else fpa_validations._norm_modalidad_db(vals[fpa_validations.EXCEL_ID_POS[c] - 1])
                for c in _CLAVE_DB
            )
            excel_index[clave] = first_row + 1 + offset

        filas_nuevas = []
        for fila in filas_hoy:
            clave = _norm_clave(fila)
            num_fila = excel_index.get(clave)
            if num_fila is not None:
                for db_col, pos in _DB_TO_EXCEL_POS.items():
                    sht.range((num_fila, pos)).value = fila.get(db_col, "")
                resultado["actualizadas"] += 1
            else:
                filas_nuevas.append(fila)

        if filas_nuevas:
            fila_destino = last_row + 1
            matriz = []
            for fila in filas_nuevas:
                fila_excel = [None] * last_col
                for db_col, pos in _DB_TO_EXCEL_POS.items():
                    fila_excel[pos - 1] = fila.get(db_col, "")
                matriz.append(fila_excel)
            sht.range((fila_destino, 1)).value = matriz
            resultado["insertadas"] = len(filas_nuevas)

        wb.save()
        resultado["ok"] = True
    except Exception as e:
        resultado["error"] = str(e)
    finally:
        if wb is not None:
            try:
                app = wb.app
                app.display_alerts = True
                app.screen_updating = True
            except Exception:
                pass
        # Cerramos SOLO el libro temporal (no la app de Excel) para liberar
        # el archivo y que sync_manager pueda reemplazar el local. El Excel
        # visible del usuario (si es distinto) permanece intacto.
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

    print(json.dumps(resultado))


if __name__ == "__main__":
    main()
