from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestValidations:
    """Tests for data validation feature (v0.10.0)"""

    def test_list_validation(self):
        """Dropdown list validation"""
        df = pd.DataFrame({"Status": ["Open", "Closed"], "Value": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                validations={"Status": {"type": "list", "values": ["Open", "Closed", "Pending"]}},
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert len(ws.data_validations.dataValidation) > 0
                dv = ws.data_validations.dataValidation[0]
                # Assert the type and that the dropdown values actually landed.
                assert dv.type == "list"
                assert "Open" in dv.formula1 and "Pending" in dv.formula1
                wb.close()
        finally:
            os.unlink(path)

    def test_number_validation(self):
        """Whole number range validation"""
        df = pd.DataFrame({"Score": [85, 90]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert len(ws.data_validations.dataValidation) > 0
                dv = ws.data_validations.dataValidation[0]
                # Assert the type and the min/max bounds, not just presence.
                assert dv.type == "whole"
                assert dv.formula1 == "0"
                assert dv.formula2 == "100"
                wb.close()
        finally:
            os.unlink(path)

    def test_validation_with_messages(self):
        """Validation with input and error messages"""
        df = pd.DataFrame({"Value": [50]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                validations={
                    "Value": {
                        "type": "decimal",
                        "min": 0,
                        "max": 100,
                        "input_title": "Enter Value",
                        "input_message": "Must be between 0 and 100",
                        "error_title": "Invalid",
                        "error_message": "Value out of range",
                    }
                },
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert len(ws.data_validations.dataValidation) > 0
                dv = ws.data_validations.dataValidation[0]
                assert dv.promptTitle == "Enter Value"
                assert dv.errorTitle == "Invalid"
                wb.close()
        finally:
            os.unlink(path)

    def test_validation_pattern_matching(self):
        """Validation with column pattern"""
        df = pd.DataFrame({"score_a": [80], "score_b": [90], "name": ["Test"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, validations={"score_*": {"type": "whole_number", "min": 0, "max": 100}}
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # Should have validations on the score columns
                assert len(ws.data_validations.dataValidation) > 0
                wb.close()
        finally:
            os.unlink(path)
