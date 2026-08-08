import unittest

from sheet_registry import DEFAULT_SHEET_IDS, sheet_ids_from_environment


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
