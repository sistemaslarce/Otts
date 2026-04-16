from __future__ import annotations

import sys
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import verificar_horimetros as vh


class ClasificaTests(unittest.TestCase):
    def test_detecta_horas_disminuidas(self) -> None:
        estado, dias, diff, max_h = vh.clasifica(
            100,
            110,
            datetime(2026, 4, 16),
            datetime(2026, 4, 15),
        )
        self.assertEqual(estado, "Horas disminuidas")
        self.assertEqual(dias, 1)
        self.assertEqual(diff, -10)
        self.assertEqual(max_h, 24)

    def test_detecta_exceso_en_horimetro(self) -> None:
        exceso = next(estado for estado in vh.COL if estado.startswith("Exceso"))
        estado, dias, diff, max_h = vh.clasifica(
            200,
            100,
            datetime(2026, 4, 16),
            datetime(2026, 4, 14),
        )
        self.assertEqual(estado, exceso)
        self.assertEqual(dias, 2)
        self.assertEqual(diff, 100)
        self.assertEqual(max_h, 48)


class LimpiarResolutionTests(unittest.TestCase):
    def test_elimina_lineas_no_utiles(self) -> None:
        texto = "\n".join(
            [
                "WORK PERFORMED: Cambio de filtro",
                "all ok",
                "Carlos Perez (2026-04-16)",
                "2026-04-16",
                "Cambio de filtro",
            ]
        )

        limpio = vh.limpiar_resolution(
            texto,
            tecnico="Carlos Perez",
            fecha=datetime(2026, 4, 16),
        )

        self.assertEqual(limpio, "Cambio de filtro")

    def test_conserva_texto_util(self) -> None:
        limpio = vh.limpiar_resolution("  Ajuste general  ")
        self.assertEqual(limpio, "Ajuste general")


class PrepararDfQcTests(unittest.TestCase):
    def test_limpia_columnas_presentacionales(self) -> None:
        tecnico_col = vh.QC_COLS[2]
        serie_col = vh.QC_COLS[3]
        horimetro_col = vh.QC_COLS[5]
        faltas_col = vh.QC_COLS[6]
        rows = [
            {
                "ERROR": "Correcto",
                "Call ID": 1,
                tecnico_col: "tecnico",
                serie_col: "SN-1",
                "Fecha de Cierre": datetime(2026, 4, 16).date(),
                horimetro_col: 123,
                faltas_col: "",
                "Estado OT": "Único",
            }
        ]

        df = vh.preparar_df_qc(rows)

        self.assertEqual(df.loc[0, "ERROR"], "")
        self.assertEqual(df.columns.tolist(), vh.QC_COLS)

    def test_es_error_relevante_para_sin_horimetro_reciente(self) -> None:
        faltas_col = vh.QC_COLS[6]
        sin_horimetro = next(estado for estado in vh.COL if estado.startswith("Sin "))
        fila = pd.Series({"ERROR": sin_horimetro, faltas_col: "1"})
        self.assertTrue(vh.es_error_relevante(fila))


if __name__ == "__main__":
    unittest.main()
