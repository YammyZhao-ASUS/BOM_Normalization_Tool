import unittest

from capacitor_normalizer import CapacitorNormalizer
from normalizer import ResistorNormalizer
from value_extractor import ValueExtractor


class ResistorNormalizerTests(unittest.TestCase):
    def test_normalizes_common_resistor_notations(self):
        parser = ResistorNormalizer()
        test_cases = {
            "4K7": 4700,
            "4.7K": 4700,
            "4700": 4700,
            "2R2": 2.2,
            "R22": 0.22,
            "0R": 0,
            "10M": 10000000,
            "1m OHM": 0.001,
        }

        for input_value, expected in test_cases.items():
            with self.subTest(input_value=input_value):
                self.assertEqual(parser.normalize(input_value), expected)

    def test_normalizes_extended_capacitor_notations(self):
        parser = CapacitorNormalizer()
        test_cases = {
            ".1uF": "100nF",
            "100N": "100nF",
            "104": "100nF",
            "1mF": "1mF",
            "1F": "1F",
        }

        for input_value, expected in test_cases.items():
            with self.subTest(input_value=input_value):
                self.assertEqual(parser.normalize(input_value), expected)

    def test_extracts_structured_engineering_attributes(self):
        extractor = ValueExtractor()
        description = "RES 10K OHM 1/16W (0402) 1% THICK FILM"

        self.assertEqual(extractor.extract_resistor_value(description), "10K OHM")
        self.assertEqual(extractor.extract_power_rating(description), "0.0625W")
        self.assertEqual(extractor.extract_tolerance(description), "1%")
        self.assertEqual(extractor.extract_material(description, "R"), "THICK FILM")
        self.assertEqual(extractor.extract_voltage("MLCC 10UF 6V3 X5R"), "6.3V")


if __name__ == "__main__":
    unittest.main()
