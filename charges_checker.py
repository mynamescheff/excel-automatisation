def check_file_conditions(excel_file_name, cell_b20_value, cell_c20_value):
    is_condition_1 = cell_b20_value == 18 and cell_c20_value == "GBP"
    is_condition_2 = "AQA" in excel_file_name.upper() and cell_b20_value == 15

    if is_condition_1 or is_condition_2:
        return True, None
    else:
        mismatched_values = {
            "cell_b20_value": cell_b20_value,
            "cell_c20_value": cell_c20_value,
        }
        return False, mismatched_values