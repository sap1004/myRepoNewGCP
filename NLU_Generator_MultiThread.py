import pandas as pd
import configparser
from concurrent.futures import ThreadPoolExecutor, as_completed

from google.cloud.dialogflowcx_v3beta1.services.sessions import SessionsClient
from google.cloud.dialogflowcx_v3beta1.types import session as cx_session
from google.api_core.client_options import ClientOptions

from dfcx_scrapi.core.conversation import DialogflowConversation
from dfcx_scrapi.core.sessions import Sessions


# ─── Helpers ────────────────────────────────────────────────────────────────
def get_config_string(config, section, option, default_value):
    try:
        return config.get(section, option)
    except (configparser.NoSectionError, configparser.NoOptionError):
        return default_value


def clean_value(value):
    if value is None:
        return ""
    s = str(value).strip()
    if "E+" in s.upper():
        try:
            return str(int(float(s)))
        except Exception:
            pass
    if s.endswith(".0"):
        s = s[:-2]
    return s


def extract_tin_extracted(params: dict) -> str:
    print(params)
    if not isinstance(params, dict):
        return ""

    for key in ("tin_extracted", "Tin_Extracted", "TIN_EXTRACTED"):
        if key in params:
            val = params[key]
            raw = (
                str(val.get("original") or val.get("resolved"))
                if isinstance(val, dict)
                else str(val)
            )
            return clean_value(raw.split(",")[0].strip())

    for key, val in params.items():
        if key.lower() == "tin_extracted":
            raw = (
                str(val.get("original") or val.get("resolved"))
                if isinstance(val, dict)
                else str(val)
            )
            return clean_value(raw.split(",")[0].strip())

    return ""


def extract_params_from_map(parameters) -> dict:
    params_dict = {}
    for key, value in parameters.items():
        if hasattr(value, "string_value"):
            params_dict[key] = value.string_value
        elif hasattr(value, "number_value"):
            params_dict[key] = value.number_value
        elif hasattr(value, "bool_value"):
            params_dict[key] = value.bool_value
        elif hasattr(value, "struct_value"):
            params_dict[key] = dict(value.struct_value)
        elif hasattr(value, "list_value"):
            params_dict[key] = list(value.list_value)
        else:
            params_dict[key] = str(value)
    return params_dict


def save_to_excel(results: pd.DataFrame, result_file: str):
    writer = pd.ExcelWriter(result_file, engine="openpyxl")
    results.to_excel(writer, index=False, sheet_name="Results")

    worksheet = writer.sheets["Results"]

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            if cell.value is not None:
                cell.number_format = "@"
                cell.value = str(cell.value)

    for col in worksheet.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        worksheet.column_dimensions[col[0].column_letter].width = max_len + 4

    writer.close()
    print(f"\n✅ Saved: {result_file} | Rows: {len(results)}")


# ─── Row Processing ─────────────────────────────────────────────────────────
def process_row(
    i, row, total,
    agent_id, cx_client, sessions,
    utterances_col, intent_column,
    test_case_column, expected_output_col,
    intent_results
):
    utterance = clean_value(row[utterances_col])
    test_case = clean_value(row[test_case_column])
    expected_output = clean_value(row[expected_output_col])

    params_dict = {}

    try:
        session_id = sessions.build_session_id(agent_id=agent_id)

        cx_client.detect_intent(
            request=cx_session.DetectIntentRequest(
                session=session_id,
                query_input=cx_session.QueryInput(
                    text=cx_session.TextInput(text="Hi"),
                    language_code="en",
                ),
            )
        )

        response = cx_client.detect_intent(
            request=cx_session.DetectIntentRequest(
                session=session_id,
                query_input=cx_session.QueryInput(
                    text=cx_session.TextInput(text=utterance),
                    language_code="en",
                ),
            )
        )

        if response.query_result.parameters:
            params_dict = extract_params_from_map(
                response.query_result.parameters
            )

    except Exception as e:
        print(f"[ERROR] Row {i+1}: {e}")

    tin_extracted = extract_tin_extracted(params_dict)

    return i, {
        test_case_column: test_case,
        "utterance": utterance,
        "tin_extracted": tin_extracted,
        "Expected Normalized Output": expected_output,
        "Matching": tin_extracted == expected_output,
    }


# ─── Main ───────────────────────────────────────────────────────────────────
def main():
    config = configparser.ConfigParser()
    config.read("Manisha_config1.properties")

    agent_id = config.get("dialogflow", "agent_id")
    excel_file = config.get("input", "excel_file")
    sheet_name = config.get("input", "sheet_name")

    flow_display_name = config.get("input", "flow_display_name")
    page_display_name = config.get("input", "page_display_name")

    utterances_col = config.get("input", "Utterances")
    intent_column = config.get("input", "intent_column")
    test_case_column = get_config_string(
        config, "input", "test_case_column", "Test Case name"
    )

    expected_output_col = "Expected Normalized Output"
    result_file = config.get("output", "result_file")

    start_row = int(get_config_string(config, "input", "start_row", "0") or 0)
    end_row_raw = get_config_string(config, "input", "end_row", "").strip()
    end_row = int(end_row_raw) if end_row_raw else None

    sheet_name = int(sheet_name) if sheet_name.isdigit() else sheet_name

    df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
    df = df.reset_index(drop=True)

    total_rows = len(df)

    if end_row is not None:
        df = df.iloc[start_row:end_row + 1]
    else:
        df = df.iloc[start_row:]

    df = df.reset_index(drop=True)

    if df.empty:
        print("[WARN] No rows to process after applying start/end row.")
        return

    total = len(df)

    print(f"[DEBUG] Row range: {start_row} → {end_row or total_rows - 1}")
    print(f"[DEBUG] Rows to run: {total}")

    location = agent_id.split("/locations/")[1].split("/")[0]

    cx_client = SessionsClient(
        client_options=ClientOptions(
            api_endpoint=f"{location}-dialogflow.googleapis.com"
        )
    )

    conversation = DialogflowConversation(agent_id=agent_id)
    sessions = Sessions(agent_id=agent_id)

    test_set = pd.DataFrame({
        "flow_display_name": [flow_display_name] * total,
        "page_display_name": [page_display_name] * total,
        "utterance": df[utterances_col],
        "inject_parameters": "",
        "end_user_metadata": "",
    })

    intent_results = conversation.run_intent_detection(
        test_set, 10, 100
    ).reset_index(drop=True)

    rows = [None] * total

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(
                process_row,
                i, row, total,
                agent_id, cx_client, sessions,
                utterances_col, intent_column,
                test_case_column, expected_output_col,
                intent_results
            ): i
            for i, row in df.iterrows()
        }

        for future in as_completed(futures):
            i, row_data = future.result()
            rows[i] = row_data

    results = pd.DataFrame(rows)
    save_to_excel(results, result_file)

    print("\n✅ Script completed successfully")


if __name__ == "__main__":
    main()