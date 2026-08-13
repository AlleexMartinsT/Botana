import unittest

from config import PLANILHAS
from sheet_registry import DEFAULT_SHEET_IDS, sheet_ids_from_environment
from sheets_writer import atualizarPlanilha


class _WorksheetThatLosesAppendedRows:
    def get_values(self):
        return self.get_all_values()

    def get_all_values(self):
        return [
            [
                "Vencimento",
                "Descrição",
                "NF",
                "Valor Total",
                "Qtd Parcelas",
                "Parcela",
                "Valor Parcela",
            ]
        ]

    def update(self, values, range_name, value_input_option):
        return {"updates": {"updatedRows": 1}}


class _WorksheetThatPersistsRows:
    def __init__(self):
        self.rows = [
            [
                "Vencimento",
                "Descrição",
                "NF",
                "Valor Total",
                "Qtd Parcelas",
                "Parcela",
                "Valor Parcela",
            ]
        ]
        self.updated_ranges = []

    def get_values(self):
        return self.rows

    def update(self, values, range_name, value_input_option):
        self.updated_ranges.append(range_name)
        row_number = int(range_name.split(":", 1)[0][1:])
        while len(self.rows) < row_number:
            self.rows.append([])
        self.rows[row_number - 1] = values[0]
        return {"updates": {"updatedRows": 1}}


class _HorizonteSpreadsheet:
    title = "Contas a Receber Horizonte 2026"
    url = "https://docs.google.com/spreadsheets/d/eh-2026"

    def __init__(self, worksheet):
        self._worksheet = worksheet

    def worksheet(self, title):
        self.last_requested_worksheet = title
        return self._worksheet


class SheetRegistryTests(unittest.TestCase):
    def test_provides_2027_receivable_sheets_when_environment_is_not_updated(self):
        sheet_ids = sheet_ids_from_environment({})

        self.assertEqual(
            sheet_ids["MVA"]["2027"],
            "1PyKble2HQUEA3EeL4RDIKdRu_9NkrEwDUh-SDGC70-Y",
        )
        self.assertEqual(
            sheet_ids["EH"]["2027"],
            "12Fb8oVxTI12tigbl56IV5snoGB8SnFBqlMThNs_vrCo",
        )

    def test_allows_environment_to_override_a_2027_sheet(self):
        sheet_ids = sheet_ids_from_environment({"SHEET_MVA_2027": "configured-sheet"})

        self.assertEqual(sheet_ids["MVA"]["2027"], "configured-sheet")
        self.assertEqual(
            sheet_ids["EH"]["2027"],
            DEFAULT_SHEET_IDS["EH"]["2027"],
        )

    def test_does_not_report_inserted_when_the_appended_row_cannot_be_read_back(self):
        planilha = _HorizonteSpreadsheet(_WorksheetThatLosesAppendedRows())
        planilha.url = f"https://docs.google.com/spreadsheets/d/{PLANILHAS['EH']['2026']}"
        dados = {
            "vencimento": "11/09/2026",
            "descricao": "FACOM BLT 7505-9",
            "nf": "22618",
            "qtdParcelas": 1,
            "numParcela": "1ª Parcela",
            "valorTotal": 206.0,
            "valorParcela": 206.0,
        }

        resultado = atualizarPlanilha(planilha, dados, gc=None)

        self.assertFalse(resultado["ok"])
        self.assertFalse(resultado["inserted"])
        self.assertEqual(resultado["reason"], "append_unverified")

    def test_writes_into_the_first_empty_row_of_the_main_financial_table(self):
        aba = _WorksheetThatPersistsRows()
        planilha = _HorizonteSpreadsheet(aba)
        planilha.url = f"https://docs.google.com/spreadsheets/d/{PLANILHAS['EH']['2026']}"
        dados = {
            "vencimento": "11/09/2026",
            "descricao": "FACOM BLT 7505-9",
            "nf": "22618",
            "qtdParcelas": 1,
            "numParcela": "1ª Parcela",
            "valorTotal": 206.0,
            "valorParcela": 206.0,
        }

        resultado = atualizarPlanilha(planilha, dados, gc=None)

        self.assertTrue(resultado["ok"])
        self.assertTrue(resultado["inserted"])
        self.assertEqual(aba.updated_ranges, ["A2:I2"])
