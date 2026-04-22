import pandas as pd
import configparser
import json
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
    """
    Converts any value to clean string.
    Handles scientific notation like 3.86E+11 → 316104332181
    Removes trailing .0 from floats.
    """
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


def extract_member_id(params: dict) -> str:
    """Extracts MemberID from parameters dict and strips ,VALID or any suffix."""
    if not isinstance(params, dict):
        return ""

    for key in ("MemberID", "memberId", "memberid", "member_id", "MEMBERID"):
        if key in params:
            val = params[key]
            if isinstance(val, dict):
                raw = str(val.get("original") or val.get("resolved") or "")
            else:
                raw = str(val)
            return clean_value(raw.split(",")[0].strip())

    # Case-insensitive fallback
    for key, val in params.items():
        if key.lower().replace("_", "") == "memberid":
            if isinstance(val, dict):
                raw = str(val.get("original") or val.get("resolved") or "")
            else:
                raw = str(val)
            return clean_value(raw.split(",")[0].strip())

    return ""


def extract_params_from_map(parameters) -> dict:
    """Safely converts MapComposite parameters to plain Python dict."""
    params_dict = {}
    for key, value in parameters.items():
        if hasattr(value, 'string_value'):
            params_dict[key] = value.string_value
        elif hasattr(value, 'number_value'):
            params_dict[key] = value.number_value
        elif hasattr(value, 'bool_value'):
            params_dict[key] = value.bool_value
        elif hasattr(value, 'struct_value'):
            params_dict[key] = dict(value.struct_value)
        elif hasattr(value, 'list_value'):
            params_dict[key] = list(value.list_value)
        else:
            params_dict[key] = str(value)
    return params_dict


def save_to_excel(results: pd.DataFrame, result_file: str):
    """
    Saves results to Excel with all columns forced as TEXT format
    so long numbers like 316104332181 are never converted to 3.86E+11.
    """
    writer = pd.ExcelWriter(result_file, engine='openpyxl')
    results.to_excel(writer, index=False, sheet_name='Results')

    workbook  = writer.book
    worksheet = writer.sheets['Results']

    # ✅ Force ALL cells to TEXT format
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            if cell.value is not None:
                cell.number_format = '@'
                cell.value = str(cell.value)

    # Auto-fit column widths
    for col in worksheet.columns:
        max_length = max(
            len(str(cell.value)) if cell.value else 0
            for cell in col
        )
        worksheet.column_dimensions[col[0].column_letter].width = max_length + 4

    writer.close()
    print(f"\n✅ Saved: {result_file} | Rows: {len(results)}")


def process_row(
    i, row, total,
    agent_id, cx_client, sessions,
    utterances_col, intent_column,
    test_case_column, expected_output_col,
    intent_results
):
    """Processes a single row — called in parallel via ThreadPoolExecutor."""

    utterance       = clean_value(row[utterances_col])
    expected_intent = clean_value(row[intent_column])
    test_case_name  = clean_value(row[test_case_column])
    expected_output = clean_value(row[expected_output_col])
    detected_intent = str(intent_results["detected_intent"].iloc[i]) \
        if "detected_intent" in intent_results.columns else ""

    print(f"[DEBUG] ── Row {i+1}/{total} | Utterance: {utterance}")

    params_dict = {}
    try:
        # ✅ Fresh session per row
        session_id = sessions.build_session_id(agent_id=agent_id)
        print(f"[DEBUG] Row {i+1} | Session ID: {session_id}")

        # ── Turn 1: Send "Hi" ─────────────────────────────────────────────
        turn1_request = cx_session.DetectIntentRequest(
            session=session_id,
            query_input=cx_session.QueryInput(
                text=cx_session.TextInput(text="Hi"),
                language_code="en"
            )
        )
        turn1_response = cx_client.detect_intent(request=turn1_request)
        print(f"[DEBUG] Row {i+1} | Turn 1 page: {turn1_response.query_result.current_page.display_name}")

        # ── Turn 2: Send actual utterance in SAME session ─────────────────
        turn2_request = cx_session.DetectIntentRequest(
            session=session_id,
            query_input=cx_session.QueryInput(
                text=cx_session.TextInput(text=utterance),
                language_code="en"
            )
        )
        turn2_response = cx_client.detect_intent(request=turn2_request)
        query_result   = turn2_response.query_result

        print(f"[DEBUG] Row {i+1} | Turn 2 page: {query_result.current_page.display_name}")

        if query_result.parameters:
            params_dict = extract_params_from_map(query_result.parameters)
            #print(f"[DEBUG] Row {i+1} | ✅ params_dict: {params_dict}")
        else:
            print(f"[DEBUG] Row {i+1} | ⚠️  parameters empty after Turn 2")

    except Exception as e:
        print(f"[ERROR] Row {i+1} failed: {e}")

    # ── Extract MemberID and compare ──────────────────────────────────────
    member_id    = extract_member_id(params_dict)
    intent_match = expected_intent == detected_intent
    matching     = member_id == expected_output

    print(f"[DEBUG] Row {i+1} | extracted_member_id       : '{member_id}'")
    print(f"[DEBUG] Row {i+1} | Expected Normalized Output : '{expected_output}'")
    print(f"[DEBUG] Row {i+1} | Matching                   : {matching}")
    print(f"[DEBUG] Row {i+1} | intent_match               : {intent_match}")

    return i, {
        test_case_column:             test_case_name,
        "utterance":                  utterance,
        #"expected_intent":            expected_intent,
        #"detected_intent":            detected_intent,
        #"intent_match":               intent_match,
        #"all_parameters":             json.dumps(params_dict),
        "extracted_member_id":        member_id,
        "Expected Normalized Output": expected_output,
        "Matching":                   matching,
    }


# ─── Main ───────────────────────────────────────────────────────────────────
def main():
    print("\n[DEBUG] ── Starting Script ──────────────────────────────────────")

    config = configparser.ConfigParser()
    config.read("Manisha_config1.properties")

    agent_id          = config.get("dialogflow", "agent_id")
    excel_file        = config.get("input", "excel_file")
    sheet_name_raw    = config.get("input", "sheet_name")

    try:
        sheet_name = int(sheet_name_raw)
    except ValueError:
        sheet_name = sheet_name_raw

    flow_display_name   = config.get("input", "flow_display_name")
    page_display_name   = config.get("input", "page_display_name")
    utterances_col      = config.get("input", "Utterances")
    intent_column       = config.get("input", "intent_column")
    test_case_column    = get_config_string(config, "input", "test_case_column", "Test Case name")
    result_file         = config.get("output", "result_file")
    expected_output_col = "Expected Normalized Output"

    # ── Make sure result_file ends with .xlsx ──────────────────────────────
    if not result_file.endswith(".xlsx"):
        result_file = result_file.replace(".csv", ".xlsx")
        if not result_file.endswith(".xlsx"):
            result_file = result_file + ".xlsx"

    # ── Extract location from agent_id ─────────────────────────────────────
    location = agent_id.split("/locations/")[1].split("/")[0]
    print(f"[DEBUG] agent_id    : {agent_id}")
    print(f"[DEBUG] location    : {location}")
    print(f"[DEBUG] flow        : {flow_display_name}")
    print(f"[DEBUG] page        : {page_display_name}")
    print(f"[DEBUG] result_file : {result_file}")

    # ── Initialize clients ─────────────────────────────────────────────────
    print("\n[DEBUG] Initializing DialogflowConversation...")
    conversation = DialogflowConversation(agent_id=agent_id)

    print("[DEBUG] Initializing Sessions...")
    sessions = Sessions(agent_id=agent_id)

    print("[DEBUG] Initializing CX SessionsClient...")
    client_options = ClientOptions(
        api_endpoint=f"{location}-dialogflow.googleapis.com"
    )
    cx_client = SessionsClient(client_options=client_options)
    print("[DEBUG] All clients initialized ✅")

    # ── Read Excel ─────────────────────────────────────────────────────────
    print(f"\n[DEBUG] Reading Excel: {excel_file}, sheet: {sheet_name}")
    df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
    print(f"[DEBUG] Excel loaded — shape: {df.shape}")
    print(f"[DEBUG] Columns: {df.columns.tolist()}")

    for col in [utterances_col, intent_column, test_case_column, expected_output_col]:
        if col not in df.columns:
            raise ValueError(f"[ERROR] Missing column '{col}' in Excel")
        print(f"[DEBUG] ✅ Column found: '{col}'")

    df = df.reset_index(drop=True)
    total = len(df)

    # ── Run intent detection via dfcx_scrapi ───────────────────────────────
    test_set = pd.DataFrame({
        "flow_display_name": [flow_display_name] * total,
        "page_display_name": [page_display_name] * total,
        "utterance":         df[utterances_col],
        "inject_parameters": "",
        "end_user_metadata": ""
    })

    print(f"\n[DEBUG] Running run_intent_detection for {total} utterances...")
    intent_results = conversation.run_intent_detection(
        test_set, 10, 100
    ).reset_index(drop=True)
    print(f"[DEBUG] Intent results columns: {intent_results.columns.tolist()}")

    # ── Parallel processing with ThreadPoolExecutor ────────────────────────
    print(f"\n[DEBUG] Starting parallel processing with 10 workers...")
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

        completed = 0
        for future in as_completed(futures):
            i, result_row = future.result()
            rows[i] = result_row
            completed += 1
            print(f"[DEBUG] ✅ Completed {completed}/{total} rows")

    # ── Build Results ──────────────────────────────────────────────────────
    results = pd.DataFrame(rows)
    print(f"\n[DEBUG] Results shape: {results.shape}")

    found = results[results["extracted_member_id"] != ""]
    print(f"\n🎯 MemberID extracted in {len(found)}/{len(results)} rows")
    print(results[[
        "utterance",
        #"detected_intent",
        #"intent_match",
        "extracted_member_id",
        "Expected Normalized Output",
        "Matching"
    ]].to_string())

    # ── Save to Excel with TEXT format ─────────────────────────────────────
    save_to_excel(results, result_file)
    print("[DEBUG] ── Script Complete ───────────────────────────────────────\n")


if __name__ == "__main__":
    main()