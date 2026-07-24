from pathlib import Path
import unittest

from word_table_parser import RawTable, combine_summary_tables, extract_word_tables, parse_summary_table


class EgfrDocumentTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        path = Path(__file__).with_name("Table_14.310.8_eGFR_Change_from_Baseline.docx")
        cls.tables = extract_word_tables(path.read_bytes(), path.name)
        cls.data = combine_summary_tables(cls.tables)

    def test_all_tables_are_combined(self) -> None:
        self.assertEqual(len(self.tables), 3)
        self.assertEqual(len(self.data), 90)
        self.assertEqual(self.data["visit"].nunique(), 10)

    def test_parameters_are_discovered_from_section_rows(self) -> None:
        self.assertListEqual(
            self.data["parameter"].drop_duplicates().tolist(),
            ["Total", "eGFR >= 60 ml/min/1.73m2", "eGFR < 60 ml/min/1.73m2"],
        )

    def test_total_group_is_retained(self) -> None:
        self.assertEqual(
            set(self.data["group"]),
            {"Surgical", "Non-surgical", "Total / TransCon PTH"},
        )


class RowStatisticDocumentTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        path = Path(__file__).with_name("Table_14_2_2_1_SkybriGHt.docx")
        tables = extract_word_tables(path.read_bytes(), path.name)
        cls.data = combine_summary_tables(tables)

    def test_row_oriented_statistics_are_parsed(self) -> None:
        self.assertEqual(len(self.data), 9)
        self.assertEqual(set(self.data["metric"]), {"Mean", "SE", "Median"})
        self.assertEqual(set(self.data["visit"]), {"Enrollment"})

    def test_group_headers_and_counts_are_preserved(self) -> None:
        self.assertEqual(
            set(self.data["group"]),
            {"Took ADHD Stimulant", "Did not Take ADHDStimulant", "Total"},
        )
        mean_rows = self.data[self.data["metric"] == "Mean"]
        self.assertEqual(set(mean_rows["n"].astype(int)), {44, 189, 233})


class ClinicalListingTableTest(unittest.TestCase):
    def test_n_percent_listing_is_tidy_and_hierarchical(self) -> None:
        table = RawTable(
            7,
            [
                ["System Organ Class / Preferred Term", "Surgical (N=121)", "Non-surgical (N=31)", "Total (N=152)"],
                ["Infections and infestations", "85 (70.2)", "22 (71.0)", "107 (70.4)"],
                ["COVID-19", "53 (43.8)", "12 (38.7)", "65 (42.8)"],
                ["Nasopharyngitis", "13 (10.7)", "6 (19.4)", "19 (12.5)"],
            ],
        )
        data = parse_summary_table(table)
        self.assertEqual(len(data), 6)
        self.assertEqual(set(data["metric"]), {"Percent"})
        self.assertEqual(set(data["parameter"]), {"Infections and infestations"})
        total = data[(data["visit"] == "COVID-19") & (data["group"] == "Total")].iloc[0]
        self.assertEqual(total["n"], 65)
        self.assertEqual(total["value"], 42.8)
        self.assertEqual(total["layout"], "categorical")


if __name__ == "__main__":
    unittest.main()
