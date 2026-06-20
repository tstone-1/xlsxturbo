import zipfile

from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestImages:
    """Tests for images feature (v0.10.0)"""

    def test_image_simple_path(self):
        """Image with simple path"""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        # Create a small test image (1x1 white pixel PNG)
        import base64

        # Smallest valid PNG (1x1 white pixel)
        png_data = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
        img_path = get_temp_path().replace(".xlsx", ".png")
        try:
            with open(img_path, "wb") as f:
                f.write(png_data)

            xlsxturbo.df_to_xlsx(df, path, images={"D1": img_path})
            # The image must actually land in the package, not just produce a file.
            with zipfile.ZipFile(path) as zf:
                media = [n for n in zf.namelist() if n.startswith("xl/media/")]
                assert media, "no embedded image found in xl/media/"
                assert any(n.endswith(".png") for n in media)
        finally:
            if os.path.exists(path):
                os.unlink(path)
            if os.path.exists(img_path):
                os.unlink(img_path)

    def test_image_with_options(self):
        """Image with scaling options"""
        df = pd.DataFrame({"A": [1]})
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
                images={"B5": {"path": img_path, "scale_width": 2.0, "scale_height": 2.0}},
            )
            with zipfile.ZipFile(path) as zf:
                media = [n for n in zf.namelist() if n.startswith("xl/media/")]
                assert media, "no embedded image found in xl/media/"
                # A drawing relationship must anchor the image to the sheet.
                assert any(n.startswith("xl/drawings/") for n in zf.namelist())
        finally:
            if os.path.exists(path):
                os.unlink(path)
            if os.path.exists(img_path):
                os.unlink(img_path)

class TestCheckboxes:
    """Tests for checkboxes feature (v0.13.0)"""

    def test_checkbox_simple_bool(self):
        """Checkboxes with plain bool values"""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, checkboxes={"D1": True, "D2": False})
            assert os.path.exists(path)
            wb = load_workbook(path)
            ws = wb.active
            # Checkboxes render as boolean TRUE/FALSE in cells
            assert ws["D1"].value is True
            assert ws["D2"].value is False
        finally:
            os.unlink(path)

    def test_checkbox_dict_form(self):
        """Checkbox with dict specifying checked state"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, checkboxes={"B2": {"checked": True}})
            assert os.path.exists(path)
            wb = load_workbook(path)
            ws = wb.active
            assert ws["B2"].value is True
        finally:
            os.unlink(path)

    def test_checkbox_with_format(self):
        """Checkbox with optional cell format"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                checkboxes={"C3": {"checked": True, "format": {"bg_color": "#C6EFCE", "bold": True}}},
            )
            assert os.path.exists(path)
            wb = load_workbook(path)
            ws = wb.active
            assert ws["C3"].value is True
        finally:
            os.unlink(path)

    def test_checkbox_missing_checked_key_raises(self):
        """Dict form without 'checked' raises a clear error"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, checkboxes={"B2": {"format": {"bold": True}}})
            assert False, "Should have raised"
        except ValueError as e:
            assert "checked" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_checkbox_format_not_dict_raises(self):
        """'format' field present but not a dict raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, checkboxes={"B2": {"checked": True, "format": "not a dict"}}
            )
            assert False, "Should have raised"
        except TypeError as e:
            assert "format" in str(e) and "dict" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_checkbox_invalid_cell_ref_raises(self):
        """Invalid cell reference raises"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, checkboxes={"not_a_ref": True})
            assert False, "Should have raised"
        except ValueError:
            pass
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_checkbox_wrong_value_type_raises(self):
        """Non-bool, non-dict value raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, checkboxes={"B2": "not_a_bool"})
            assert False, "Should have raised"
        except TypeError as e:
            assert "checkboxes" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_checkbox_with_dfs_to_xlsx_per_sheet(self):
        """Checkboxes work via per-sheet options in dfs_to_xlsx"""
        df1 = pd.DataFrame({"A": [1, 2]})
        df2 = pd.DataFrame({"B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (df1, "S1", {"checkboxes": {"D1": True}}),
                    (df2, "S2", {"checkboxes": {"D1": False}}),
                ],
                path,
            )
            assert os.path.exists(path)
            wb = load_workbook(path)
            assert wb["S1"]["D1"].value is True
            assert wb["S2"]["D1"].value is False
        finally:
            os.unlink(path)

    def test_checkbox_combined_with_other_features(self):
        """Checkboxes coexist with other features on the same sheet"""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Active": [True, False]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                checkboxes={"D1": True, "D2": False},
                comments={"A1": "Names"},
                validations={"Name": {"type": "text_length", "min": 1, "max": 100}},
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_checkbox_constant_memory_warns(self):
        """constant_memory=True with checkboxes emits RuntimeWarning"""
        import warnings

        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                xlsxturbo.df_to_xlsx(
                    df, path, constant_memory=True, checkboxes={"B2": True}
                )
                assert len(w) == 1
                assert issubclass(w[0].category, RuntimeWarning)
                assert "checkboxes" in str(w[0].message)
        finally:
            os.unlink(path)

class TestTextboxes:
    """Tests for textboxes feature (v0.14.0)"""

    def test_textbox_simple_string(self):
        """Textbox with bare string value writes file successfully"""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, textboxes={"C2": "Simple note"})
            with zipfile.ZipFile(path) as zf:
                assert "xl/drawings/drawing1.xml" in zf.namelist()
                drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
                assert "Simple note" in drawing
        finally:
            os.unlink(path)


class TestCharts:
    """Tests for native Excel charts"""

    def test_single_series_chart(self):
        """Charts create native chart XML with title and data table"""
        df = pd.DataFrame({"Month": ["Jan", "Feb", "Mar"], "Sales": [120, 145, 160]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                charts={
                    "D2": {
                        "type": "bar",
                        "data_range": "Sheet1!$B$2:$B$4",
                        "categories_range": "Sheet1!$A$2:$A$4",
                        "title": "Monthly Sales",
                        "width": 720,
                        "height": 480,
                        "show_data_table": True,
                        "legend_position": "bottom",
                    }
                },
            )
            with zipfile.ZipFile(path) as zf:
                names = zf.namelist()
                assert "xl/charts/chart1.xml" in names
                chart_xml = zf.read("xl/charts/chart1.xml").decode("utf-8")
                assert "Monthly Sales" in chart_xml
                assert "<c:dTable>" in chart_xml
                assert "<c:barChart>" in chart_xml
                # The data_range and categories_range must actually land in the
                # chart XML, not just produce a well-formed but empty chart.
                assert "Sheet1!$B$2:$B$4" in chart_xml
                assert "Sheet1!$A$2:$A$4" in chart_xml
        finally:
            os.unlink(path)

    def test_multi_series_chart(self):
        """Charts support explicit multiple series"""
        df = pd.DataFrame(
            {
                "Month": ["Jan", "Feb", "Mar"],
                "Sales": [120, 145, 160],
                "Margin": [32, 41, 48],
            }
        )
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                charts={
                    "E2": {
                        "type": "column",
                        "series": [
                            {"name": "Sales", "values_range": "Sheet1!$B$2:$B$4"},
                            {"name": "Margin", "values_range": "Sheet1!$C$2:$C$4"},
                        ],
                        "categories_range": "Sheet1!$A$2:$A$4",
                        "title": "Quarter Results",
                        "show_legend": False,
                    }
                },
            )
            with zipfile.ZipFile(path) as zf:
                chart_xml = zf.read("xl/charts/chart1.xml").decode("utf-8")
                assert chart_xml.count("<c:ser>") == 2
                assert "Quarter Results" in chart_xml
                assert "<c:legend>" not in chart_xml
                # Each series' values range and the shared categories range
                # must reach the XML.
                assert "Sheet1!$B$2:$B$4" in chart_xml
                assert "Sheet1!$C$2:$C$4" in chart_xml
                assert "Sheet1!$A$2:$A$4" in chart_xml
        finally:
            os.unlink(path)

    def test_charts_with_dfs_to_xlsx_per_sheet(self):
        """Charts work via per-sheet options in dfs_to_xlsx"""
        df1 = pd.DataFrame({"Month": ["Jan", "Feb"], "Sales": [10, 20]})
        df2 = pd.DataFrame({"Month": ["Jan", "Feb"], "Sales": [30, 40]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (
                        df1,
                        "North",
                        {
                            "charts": {
                                "D2": {
                                    "type": "line",
                                    "data_range": "North!$B$2:$B$3",
                                    "categories_range": "North!$A$2:$A$3",
                                }
                            }
                        },
                    ),
                    (
                        df2,
                        "South",
                        {
                            "charts": {
                                "D2": {
                                    "type": "line",
                                    "data_range": "South!$B$2:$B$3",
                                    "categories_range": "South!$A$2:$A$3",
                                }
                            }
                        },
                    ),
                ],
                path,
            )
            with zipfile.ZipFile(path) as zf:
                assert "xl/charts/chart1.xml" in zf.namelist()
                assert "xl/charts/chart2.xml" in zf.namelist()
        finally:
            os.unlink(path)

    def test_chart_invalid_type_raises_error(self):
        """Invalid chart types raise a clear error"""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown chart type"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    charts={"D2": {"type": "not_a_chart", "data_range": "Sheet1!$A$2:$A$3"}},
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_chart_series_unknown_key_raises(self):
        """A typo in a series-item key is rejected, not silently dropped"""
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="unknown option"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    charts={
                        "D2": {
                            "type": "column",
                            "series": [
                                # 'categorie_range' is a typo for 'categories_range'
                                {
                                    "values_range": "Sheet1!$B$2:$B$4",
                                    "categorie_range": "Sheet1!$A$2:$A$4",
                                }
                            ],
                        }
                    },
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_chart_constant_memory_warns(self):
        """constant_memory=True with charts emits RuntimeWarning"""
        import warnings

        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    constant_memory=True,
                    charts={"D2": {"type": "bar", "data_range": "Sheet1!$A$2:$A$3"}},
                )
                assert len(w) == 1
                assert issubclass(w[0].category, RuntimeWarning)
                assert "charts" in str(w[0].message)
        finally:
            os.unlink(path)

    def test_textbox_dict_form(self):
        """Textbox dict form with required 'text' key"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, textboxes={"B2": {"text": "Hello"}})
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_textbox_with_all_options(self):
        """Textbox with every supported option"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                textboxes={
                    "B2": {
                        "text": "Annotated",
                        "width": 200,
                        "height": 100,
                        "x_offset": 10,
                        "y_offset": 5,
                        "font": {
                            "name": "Arial",
                            "size": 14,
                            "bold": True,
                            "italic": True,
                            "underline": True,
                            "color": "#FF0000",
                        },
                        "fill_color": "#F0F0F0",
                        "line_color": "#000000",
                        "alt_text": "A textbox annotation",
                    }
                },
            )
            with zipfile.ZipFile(path) as zf:
                drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
                assert "Annotated" in drawing
                # alt_text, fill color, and font color must reach the XML.
                assert "A textbox annotation" in drawing
                assert "F0F0F0" in drawing.upper()
                assert "FF0000" in drawing.upper()
        finally:
            os.unlink(path)

    def test_textbox_font_partial(self):
        """Partial font options work (only some keys set)"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, textboxes={"B2": {"text": "T", "font": {"bold": True}}}
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_textbox_missing_text_raises(self):
        """Dict form without 'text' raises ValueError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, textboxes={"B2": {"width": 100}})
            assert False, "Should have raised"
        except ValueError as e:
            assert "text" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_unknown_option_raises(self):
        """Unknown top-level option raises with the list of valid keys"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, textboxes={"B2": {"text": "T", "bogus": 1}}
            )
            assert False, "Should have raised"
        except ValueError as e:
            assert "bogus" in str(e) and "Valid" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_unknown_font_option_raises(self):
        """Unknown font sub-option raises"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                textboxes={"B2": {"text": "T", "font": {"weight": "heavy"}}},
            )
            assert False, "Should have raised"
        except ValueError as e:
            assert "weight" in str(e) and "font" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_font_not_dict_raises(self):
        """'font' present but not a dict raises ValueError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, textboxes={"B2": {"text": "T", "font": "bold"}}
            )
            assert False, "Should have raised"
        except ValueError as e:
            assert "font" in str(e) and "dict" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_offsets_only(self):
        """Offsets without size/font still writes file (exercises insert_shape_with_offset path)"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                textboxes={"B2": {"text": "T", "x_offset": 15, "y_offset": 25}},
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_textbox_wrong_value_type_raises(self):
        """Non-string, non-dict value raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, textboxes={"B2": 123})
            assert False, "Should have raised"
        except TypeError as e:
            assert "textboxes" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_invalid_color_raises(self):
        """Invalid color string surfaces the parse error"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                textboxes={"B2": {"text": "T", "fill_color": "not_a_color"}},
            )
            assert False, "Should have raised"
        except ValueError as e:
            assert "color" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_invalid_cell_ref_raises(self):
        """Invalid cell reference raises"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, textboxes={"not_a_ref": "T"})
            assert False, "Should have raised"
        except ValueError:
            pass
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_textbox_with_dfs_to_xlsx_per_sheet(self):
        """Textboxes work via per-sheet options in dfs_to_xlsx"""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (df1, "S1", {"textboxes": {"C2": "Sheet 1 note"}}),
                    (df2, "S2", {"textboxes": {"C2": {"text": "Sheet 2", "font": {"bold": True}}}}),
                ],
                path,
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_textbox_combined_with_other_features(self):
        """Textboxes coexist with images, checkboxes, comments on the same sheet"""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [85, 92]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                textboxes={"D2": {"text": "Notes", "width": 150, "height": 60}},
                checkboxes={"E1": True},
                comments={"A1": "Names column"},
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_textbox_constant_memory_warns(self):
        """constant_memory=True with textboxes emits RuntimeWarning"""
        import warnings

        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                xlsxturbo.df_to_xlsx(
                    df, path, constant_memory=True, textboxes={"B2": "note"}
                )
                assert len(w) == 1
                assert issubclass(w[0].category, RuntimeWarning)
                assert "textboxes" in str(w[0].message)
        finally:
            os.unlink(path)
