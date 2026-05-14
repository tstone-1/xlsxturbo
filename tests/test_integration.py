from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestAllFeaturesCombined:
    """Test using all new features together"""

    def test_all_features_df_to_xlsx(self):
        """All features work together in df_to_xlsx"""
        df = pd.DataFrame(
            {"Name": ["Alice", "Bob"], "Score": [95, 87], "Grade": ["A", "B"]}
        )
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                autofit=True,
                table_style="Medium2",
                table_name="StudentScores",
                column_widths={"_all": 30, 0: 20},
                header_format={"bold": True, "bg_color": "#4F81BD", "font_color": "white"},
                freeze_panes=True,
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert "StudentScores" in ws.tables
                # Note: table_style overrides header_format styling
                # This is expected Excel behavior - tables have their own header styles
                wb.close()
        finally:
            os.unlink(path)

    def test_all_features_dfs_to_xlsx(self):
        """All features work together in dfs_to_xlsx"""
        df1 = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        df2 = pd.DataFrame({"X": ["a", "b"], "Y": ["c", "d"]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (
                        df1,
                        "Numbers",
                        {
                            "table_style": "Medium2",
                            "table_name": "NumbersTable",
                            "header_format": {"bold": True},
                            "column_widths": {"_all": 15},
                        },
                    ),
                    (
                        df2,
                        "Letters",
                        {
                            "table_style": "Medium9",
                            "table_name": "LettersTable",
                            "header_format": {"bg_color": "#FF0000"},
                        },
                    ),
                ],
                path,
                autofit=True,
                freeze_panes=True,
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert "NumbersTable" in wb["Numbers"].tables
                assert "LettersTable" in wb["Letters"].tables
                wb.close()
        finally:
            os.unlink(path)

class TestV10AllFeatures:
    """Tests combining v0.10.0 features"""

    def test_all_new_features_together(self):
        """All v0.10.0 features work together"""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [85, 92]})
        path = get_temp_path()
        import base64

        png_data = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
        img_path = get_temp_path().replace(".xlsx", ".png")
        try:
            with open(img_path, "wb") as f:
                f.write(png_data)

            xlsxturbo.df_to_xlsx(
                df,
                path,
                comments={"A1": "Names column", "B1": {"text": "Scores", "author": "Test"}},
                validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
                rich_text={"D1": [("Legend:", {"bold": True}), " student scores"]},
                images={"E1": img_path},
            )
            assert os.path.exists(path)
        finally:
            if os.path.exists(path):
                os.unlink(path)
            if os.path.exists(img_path):
                os.unlink(img_path)

    def test_new_features_with_dfs_to_xlsx(self):
        """New features work with dfs_to_xlsx"""
        df1 = pd.DataFrame({"A": [1, 2]})
        df2 = pd.DataFrame({"B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (df1, "Sheet1", {"comments": {"A1": "First sheet header"}}),
                    (df2, "Sheet2", {"validations": {"B": {"type": "whole_number", "min": 0, "max": 10}}}),
                ],
                path,
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)
